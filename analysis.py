# -*- coding: utf-8 -*-
import re
import json
import base64
import threading
from datetime import datetime, timezone, timedelta
from typing import List, Dict, Any, Tuple, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed

import openpyxl
from openai import OpenAI
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow

# =========================
# 사용자 설정
# =========================
PERIOD = "after:2025/01/01 before:2025/12/31"  # Gmail period 필터에 의존
LABEL_NAME = "student"
PROF_EMAIL = "jwpark12@sungshin.ac.kr"

TEMPLATE_XLSX = "counsel_excelupload.xlsx"
OUTPUT_XLSX = "output.xlsx"

# ✅ GPT-5 계열 (권한 있어야 함)
PRIMARY_MODEL = "gpt-5-nano"
FALLBACK_MODEL = "gpt-5-mini"  # 없으면 None으로

# 병렬 처리 개수 (레이트리밋/안정성 고려해 4 추천)
MAX_WORKERS = 4

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

VALID_CODES = {
    "CF01", "CF02", "CF03", "CF04", "CF05", "CF06", "CF07", "CF08", "CF09",
    "CF10", "CF11", "CF12", "CF13", "CF14", "CF15", "CF16", "CF17", "CF18",
    "CF19", "CF22", "CF23", "CF24", "CF25",
}

# =========================
# Gmail 인증/조회
# =========================
def authenticate_gmail():
    flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
    creds = flow.run_local_server(port=0)
    return build("gmail", "v1", credentials=creds)

def get_threads(service, label_name: str, period: str) -> List[Dict[str, Any]]:
    query = f"label:{label_name} {period}"
    threads: List[Dict[str, Any]] = []

    resp = service.users().threads().list(userId="me", q=query).execute()
    threads.extend(resp.get("threads", []))
    while "nextPageToken" in resp:
        resp = service.users().threads().list(
            userId="me", q=query, pageToken=resp["nextPageToken"]
        ).execute()
        threads.extend(resp.get("threads", []))
    return threads

# =========================
# 본문 추출
# =========================
def _extract_plain_text_from_payload(payload: Dict[str, Any]) -> str:
    def decode(data: str) -> str:
        try:
            return base64.urlsafe_b64decode(data).decode("utf-8", errors="ignore")
        except Exception:
            return ""

    if "data" in payload.get("body", {}):
        return decode(payload["body"]["data"])

    parts = payload.get("parts", [])
    for part in parts:
        if part.get("mimeType") == "text/plain" and "data" in part.get("body", {}):
            return decode(part["body"]["data"])

    for part in parts:
        if part.get("parts"):
            nested = _extract_plain_text_from_payload(part)
            if nested.strip():
                return nested

    for part in parts:
        if part.get("mimeType") == "text/html" and "data" in part.get("body", {}):
            return decode(part["body"]["data"])

    return ""

def get_thread_messages(service, thread_id: str) -> Tuple[str, List[Dict[str, Any]]]:
    thread = service.users().threads().get(
        userId="me", id=thread_id, format="full"
    ).execute()
    messages = thread.get("messages", [])

    subject = ""
    if messages:
        for h in messages[0].get("payload", {}).get("headers", []):
            if h.get("name") == "Subject":
                subject = h.get("value", "")
                break

    out: List[Dict[str, Any]] = []
    for m in messages:
        headers = {
            h["name"].lower(): h.get("value", "")
            for h in m.get("payload", {}).get("headers", [])
            if "name" in h
        }
        body = _extract_plain_text_from_payload(m.get("payload", {})).strip()

        out.append({
            "internal_ms": int(m.get("internalDate", "0")),
            "from": headers.get("from", ""),
            "to": headers.get("to", ""),
            "cc": headers.get("cc", ""),
            "body": body,
        })

    # 시간순 정렬
    out.sort(key=lambda x: x["internal_ms"])
    for i, msg in enumerate(out):
        msg["idx"] = i

    return subject, out

def is_prof_message(from_header: str) -> bool:
    return PROF_EMAIL.lower() in (from_header or "").lower()

def ms_to_kst(ms: int) -> datetime:
    kst = timezone(timedelta(hours=9))
    return datetime.fromtimestamp(ms / 1000, tz=timezone.utc).astimezone(kst)

# =========================
# OpenAI 호출/파싱 (nano -> mini fallback)
# =========================
def _clean_json_block(text: str) -> str:
    t = text.strip()
    if t.startswith("```json"):
        t = t[7:].strip()
    if t.startswith("```"):
        t = t[3:].strip()
    if t.endswith("```"):
        t = t[:-3].strip()
    return t

def call_llm_json(client: OpenAI, prompt: str) -> Dict[str, Any]:
    last_err: Optional[Exception] = None
    for model in [PRIMARY_MODEL, FALLBACK_MODEL]:
        if not model:
            continue
        try:
            resp = client.chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": prompt}],
                response_format={"type": "json_object"},
            )
            content = (resp.choices[0].message.content or "").strip()
            if not content:
                raise ValueError("빈 응답")
            return json.loads(_clean_json_block(content))
        except Exception as e:
            last_err = e
            continue
    raise RuntimeError(f"LLM 호출 실패: {last_err}")

# =========================
# 요약 검증(후처리 호출 제거: 위반이면 스킵)
# =========================
BAD_PATTERNS = [
    r"합니다", r"드립니다", r"입니다", r"하세요", r"했습니다", r"하셨다", r"부탁", r"감사",
    r"감사합니다", r"부탁드립니다",
]

def is_bad_summary(text: str) -> bool:
    t = (text or "").strip()
    if not t:
        return True
    if any(re.search(p, t) for p in BAD_PATTERNS):
        return True
    if len(t) > 220:
        return True
    return False

# =========================
# 프롬프트 (1스레드 = 1건)
# =========================
def build_one_record_prompt(subject: str, msgs: List[Dict[str, Any]]) -> str:
    parts: List[str] = []
    for m in msgs:
        role = "PROF" if is_prof_message(m["from"]) else "OTHER"
        dt = ms_to_kst(m["internal_ms"]).strftime("%Y-%m-%d %H:%M:%S")
        body = (m.get("body") or "").strip()
        if not body:
            continue
        parts.append(
            f"[{m['idx']}] ({dt}) {role}\nFROM: {m['from']}\nBODY:\n{body}\n"
        )
    joined = "\n---\n".join(parts)

    return f"""
너는 '학생 이메일 상담 기록'을 엑셀로 정리하는 도우미다.
이 스레드는 "상담 1건"으로만 처리한다(주제 여러 개여도 한 건으로 묶음).

[필수 조건]
- 교수({PROF_EMAIL})가 보낸 답장 메시지가 반드시 있어야 함. 없으면 item 생성 금지.

[요약 규칙]
- student_request_summary / prof_reply_summary는 각각 2문장 이내로 매우 짧게 작성.
- 반드시 음슴체만 사용(~함/~됨/~요청함/~문의함/~안내함/~원함/~필요함/~받고싶음).
- 아래 표현이 하나라도 나오면 실패: 합니다, 드립니다, 입니다, 하세요, 했습니다, 하셨다, 부탁, 감사
- 인사/감사/서명/링크/원문 복붙 금지. 핵심만 재서술.

[상담유형]
- 아래 CF코드 중 하나로 category_code를 선택. 모르면 CF08.
{", ".join(sorted(list(VALID_CODES)))}

반환 JSON(반드시 정확히):
{{
  "item": {{
    "student_name": "홍길동",        // 성명 없으면 null
    "student_id": "20231234",       // 8자리 없으면 ""
    "category_code": "CF10",
    "student_request_summary": "....",
    "prof_reply_summary": "...."
  }}
}}

스레드 제목: {subject}

메시지들:
{joined}
""".strip()

# =========================
# 시간 반올림(30분)
# =========================
def round_to_nearest_30_minutes(dt: datetime) -> datetime:
    total = dt.hour * 60 + dt.minute
    rounded = 30 * round(total / 30)
    hour = (rounded // 60) % 24
    minute2 = rounded % 60
    return dt.replace(hour=hour, minute=minute2, second=0, microsecond=0)

# =========================
# 엑셀 저장
# =========================
def save_to_excel(records: List[Dict[str, Any]], template_xlsx: str, output_xlsx: str):
    wb = openpyxl.load_workbook(template_xlsx)
    ws = wb["Sheet1"]

    for r in records:
        ws.append([
            r.get("학번", ""),              # 학번(빈칸 허용)
            r.get("성명", ""),              # 성명(필수)
            "이메일",                       # 상담형태(고정)
            r.get("상담일", ""),            # YYYY-MM-DD
            r.get("상담시작시간", ""),      # HH:MM
            r.get("상담종료시간", ""),      # HH:MM
            r.get("상담유형", "CF08"),      # CF코드
            "",                             # 장소(비필수)
            r.get("학생상담신청내용", ""),   # 요약
            r.get("교수답변내용", ""),       # 요약(필수)
            "Y",                            # 공개여부(고정)
        ])

    wb.save(output_xlsx)
    print(f"[OK] Saved: {output_xlsx}")

# =========================
# 스레드 1개 처리 (병렬 작업 단위)
# =========================
def process_one_thread(task: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """
    task keys: i, total, subject, msgs, short_subject
    return: record dict or None
    """
    subject = task["subject"]
    msgs = task["msgs"]

    # 교수 메시지(본문 있는 것) 없으면 제외(LLM 호출도 안 함)
    prof_msgs = [
        m for m in msgs
        if is_prof_message(m.get("from", "")) and (m.get("body") or "").strip()
    ]
    if not prof_msgs:
        return None

    # 상담 시간 기준: "가장 이른 교수 답장"을 대표로 사용
    rep_msg = min(prof_msgs, key=lambda x: x["internal_ms"])
    rep_dt = ms_to_kst(rep_msg["internal_ms"])
    start_dt = round_to_nearest_30_minutes(rep_dt)
    end_dt = start_dt + timedelta(minutes=30)

    # 스레드별로 클라이언트 생성(스레드 세이프 이슈 회피)
    client = OpenAI()

    prompt = build_one_record_prompt(subject, msgs)
    result = call_llm_json(client, prompt)

    item = result.get("item")
    if not isinstance(item, dict):
        return None

    name = item.get("student_name")
    if not name or not str(name).strip():
        return None  # 성명 없으면 제외

    student_sum = str(item.get("student_request_summary", "")).strip()
    prof_sum = str(item.get("prof_reply_summary", "")).strip()

    # 후처리 재작성 없음: 규칙 위반이면 그냥 스킵
    if is_bad_summary(student_sum) or is_bad_summary(prof_sum):
        return None
    if not prof_sum:
        return None  # 답변 없으면 상담 아님

    code = str(item.get("category_code", "CF08")).strip()
    if code not in VALID_CODES:
        code = "CF08"

    sid = item.get("student_id", "")
    sid = "" if sid is None else str(sid).strip()

    return {
        "학번": sid,
        "성명": str(name).strip(),
        "상담일": start_dt.strftime("%Y-%m-%d"),
        "상담시작시간": start_dt.strftime("%H:%M"),
        "상담종료시간": end_dt.strftime("%H:%M"),
        "상담유형": code,
        "학생상담신청내용": student_sum,
        "교수답변내용": prof_sum,
    }

# =========================
# 메인
# =========================
def main():
    service = authenticate_gmail()
    threads = get_threads(service, LABEL_NAME, PERIOD)
    total = len(threads)
    print(f"[INFO] threads matched query: {total}")

    # 1) Gmail에서 스레드 메시지를 먼저 모음
    prepared: List[Dict[str, Any]] = []
    for i, t in enumerate(threads, start=1):
        thread_id = t["id"]
        subject, msgs = get_thread_messages(service, thread_id)

        short_subject = (subject or "").strip()
        if len(short_subject) > 60:
            short_subject = short_subject[:60] + "..."

        if i == 1 or i % 10 == 0 or i == total:
            print(f"[FETCH] {i}/{total} thread='{short_subject}' msgs={len(msgs)}")

        prepared.append({
            "i": i,
            "total": total,
            "subject": subject or "",
            "msgs": msgs,
            "short_subject": short_subject,
        })

    print(f"[INFO] fetched threads: {len(prepared)}. start LLM with {MAX_WORKERS} workers")

    # 2) LLM 병렬 처리
    records: List[Dict[str, Any]] = []
    lock = threading.Lock()
    done_count = 0

    def log(msg: str):
        with lock:
            print(msg)

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        future_map = {ex.submit(process_one_thread, task): task for task in prepared}

        for fut in as_completed(future_map):
            task = future_map[fut]
            i = task["i"]
            short_subject = task["short_subject"]

            try:
                rec = fut.result()
            except Exception as e:
                log(f"[WARN] {i}/{total} '{short_subject}' failed: {e}")
                rec = None

            if rec:
                records.append(rec)

            done_count += 1
            if done_count % 10 == 0 or done_count == total:
                log(f"[PROGRESS] done {done_count}/{total}, rows={len(records)}")

    # 3) 엑셀 저장
    save_to_excel(records, TEMPLATE_XLSX, OUTPUT_XLSX)
    print(f"[DONE] rows appended: {len(records)}")

if __name__ == "__main__":
    main()
