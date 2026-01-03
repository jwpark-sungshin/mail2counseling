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


# ======================================================
# 설정
# ======================================================
PERIOD = "after:2025/01/01 before:2025/12/31"
LABEL_NAME = "student"
PROF_EMAIL = "jwpark12@sungshin.ac.kr"

TEMPLATE_XLSX = "counsel_excelupload.xlsx"
OUTPUT_XLSX = "output.xlsx"

# ✅ mini 기본
PRIMARY_MODEL = "gpt-5-mini"
FALLBACK_MODEL = "gpt-5-nano"

MAX_WORKERS = 4
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

COUNSEL_TYPE_DEFAULT = "3"  # 상담형태 기본값(텍스트): 이메일 = "3"
PUBLIC_YN_DEFAULT = "Y"

# 명칭-코드 매핑(사용자 제공)
CF_CATEGORIES = {
    "학업": "CF01",
    "전공": "CF02",
    "장학금": "CF03",
    "진로": "CF04",
    "생활": "CF05",
    "휴학": "CF06",
    "자퇴": "CF07",
    "기타": "CF08",
    "멘토링 장학금": "CF09",
    "수강": "CF10",
    "성적": "CF11",
    "입학": "CF12",
    "진학": "CF13",
    "창업": "CF14",
    "사회봉사": "CF15",
    "건강": "CF16",
    "학술활동": "CF17",
    "논문": "CF18",
    "입학사정관": "CF19",
    "취업": "CF22",
    "대외활동": "CF23",
    "교환학생": "CF24",
    "현장실습": "CF25",
}
VALID_CODES = set(CF_CATEGORIES.values())


# ======================================================
# Gmail
# ======================================================
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


# ======================================================
# 메시지 파싱
# ======================================================
def _extract_plain_text_from_payload(payload: Dict[str, Any]) -> str:
    def decode(data: str) -> str:
        try:
            return base64.urlsafe_b64decode(data).decode("utf-8", errors="ignore")
        except Exception:
            return ""

    if "data" in payload.get("body", {}):
        return decode(payload["body"]["data"])

    for part in payload.get("parts", []):
        if part.get("mimeType") == "text/plain" and "data" in part.get("body", {}):
            return decode(part["body"]["data"])
        if part.get("parts"):
            nested = _extract_plain_text_from_payload(part)
            if nested.strip():
                return nested

    for part in payload.get("parts", []):
        if part.get("mimeType") == "text/html" and "data" in part.get("body", {}):
            return decode(part["body"]["data"])

    return ""


def get_thread_messages(service, thread_id: str) -> Tuple[str, List[Dict[str, Any]]]:
    thread = service.users().threads().get(userId="me", id=thread_id, format="full").execute()
    messages = thread.get("messages", [])

    subject = ""
    if messages:
        for h in messages[0]["payload"]["headers"]:
            if h["name"] == "Subject":
                subject = h["value"]
                break

    out: List[Dict[str, Any]] = []
    for m in messages:
        headers = {h["name"].lower(): h["value"] for h in m["payload"]["headers"]}
        body = _extract_plain_text_from_payload(m["payload"]).strip()

        out.append({
            "internal_ms": int(m["internalDate"]),
            "from": headers.get("from", ""),
            "body": body,
        })

    out.sort(key=lambda x: x["internal_ms"])
    for i, msg in enumerate(out):
        msg["idx"] = i

    return subject, out


def is_prof_message(from_header: str) -> bool:
    return PROF_EMAIL.lower() in (from_header or "").lower()


def ms_to_kst(ms: int) -> datetime:
    return datetime.fromtimestamp(ms / 1000, tz=timezone.utc).astimezone(
        timezone(timedelta(hours=9))
    )


# ======================================================
# 학번 추출: 이메일 우선, LLM 후순위 (둘 중 하나만 성공해도 사용)
# ======================================================
EMAIL_RE = re.compile(r"([A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})", re.I)


def _extract_email_addr(from_header: str) -> str:
    if not from_header:
        return ""
    m = EMAIL_RE.search(from_header)
    return (m.group(1) if m else from_header).strip().lower()


def extract_student_id_from_email(from_header: str) -> str:
    """
    @sungshin.ac.kr 메일의 로컬파트가 8자리 숫자면 학번으로 간주
    From 예: "홍길동 <20201234@sungshin.ac.kr>" / "20201234@sungshin.ac.kr"
    """
    addr = _extract_email_addr(from_header)
    if not addr.endswith("@sungshin.ac.kr"):
        return ""
    local = addr.split("@", 1)[0].strip()
    return local if re.fullmatch(r"\d{8}", local) else ""


def normalize_student_id_from_llm(student_id: Any) -> str:
    sid = "" if student_id is None else str(student_id).strip()
    return sid if re.fullmatch(r"\d{8}", sid) else ""


# ======================================================
# OpenAI
# ======================================================
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
                raise ValueError("empty response")
            return json.loads(content)
        except Exception as e:
            last_err = e
            continue
    raise RuntimeError(f"LLM call failed: {last_err}")


# ======================================================
# 요약 검증 (후처리 LLM 재작성 없음: 위반이면 스킵)
# ======================================================
BAD_PATTERNS = [
    r"합니다", r"드립니다", r"입니다", r"하세요", r"했습니다", r"하셨다",
    r"부탁", r"감사", r"감사합니다", r"부탁드립니다",
    r"하고자", r"하였다", r"했다", r"하라",
]

def is_bad_summary(text: str) -> bool:
    t = (text or "").strip()
    if not t:
        return True
    if len(t) > 240:
        return True
    return any(re.search(p, t) for p in BAD_PATTERNS)


# ======================================================
# 프롬프트 (학생-교수 상담 검증 + 1스레드=1건)
# ======================================================
def build_prompt(subject: str, msgs: List[Dict[str, Any]]) -> str:
    blocks = []
    for m in msgs:
        role = "PROF" if is_prof_message(m["from"]) else "OTHER"
        dt = ms_to_kst(m["internal_ms"]).strftime("%Y-%m-%d %H:%M")
        body = (m.get("body") or "").strip()
        if not body:
            continue
        blocks.append(f"[{m['idx']}] {dt} {role}\nFROM: {m['from']}\n{body}")

    joined = "\n---\n".join(blocks)
    cf_lines = "\n".join([f"- {name}: {code}" for name, code in CF_CATEGORIES.items()])

    return f"""
너는 '학생 이메일 상담 기록'을 엑셀로 정리하는 도우미다.

[1] 먼저 이 스레드가 "학생 1명 ↔ 교수({PROF_EMAIL})" 상담 대화인지 판정하라.
- 학생은 보통 상담/질문/요청을 하고, 교수는 답변/안내를 하는 형태임
- 단순 공지, 행정담당/조교/시스템 알림, 외부 업체 메일이면 학생 상담 아님
- 확신이 없으면 학생 상담 아님(false)

[2] 학생 상담이 맞다면 이 스레드는 "상담 1건"으로만 요약하라(주제 여러 개여도 1건).

[필수 조건]
- 교수 메시지(답변)와 학생 메시지(질문/요청) 둘 다 존재해야 함
- 교수 답변이 없거나 학생 메시지가 없으면 is_student_thread=false

[요약 규칙]
- 반드시 음슴체만 사용
- student_request_summary / prof_reply_summary 각각 2문장 이내
- 아래 표현 포함 금지: 합니다, 드립니다, 입니다, 하세요, 했습니다, 하셨다, 부탁, 감사, 했다, 하였다, 하고자, 하라
- 인사/감사/서명/링크/원문 복붙 금지, 핵심만 재서술

[상담유형 분류(명칭-코드)]
{cf_lines}

category_code 규칙:
- 위 목록의 코드 중 하나만 선택해서 반환
- 애매하면 기타(CF08)

반환 JSON(반드시 정확히):
{{
  "is_student_thread": true/false,
  "item": {{
    "student_name": "홍길동",                // 없으면 ""
    "student_id": "20231234",               // 8자리 없으면 ""
    "category_code": "CF10",
    "student_request_summary": "...",
    "prof_reply_summary": "..."
  }}
}}

제목: {subject}

메시지들:
{joined}
""".strip()


# ======================================================
# 시간 계산
# ======================================================
def round_30(dt: datetime) -> datetime:
    total = dt.hour * 60 + dt.minute
    rounded = 30 * round(total / 30)
    return dt.replace(
        hour=(rounded // 60) % 24,
        minute=rounded % 60,
        second=0,
        microsecond=0,
    )


# ======================================================
# 엑셀 저장
# ======================================================
def save_to_excel(records: List[Dict[str, Any]]):
    wb = openpyxl.load_workbook(TEMPLATE_XLSX)
    ws = wb.active

    for r in records:
        ws.append([
            r.get("학번", ""),                 # 학번(빈칸 허용)
            r.get("성명", ""),                 # 성명(필수)
            COUNSEL_TYPE_DEFAULT,              # 상담형태 기본값 "3"(텍스트)
            r.get("상담일", ""),
            r.get("상담시작시간", ""),
            r.get("상담종료시간", ""),
            r.get("상담유형", "CF08"),
            "",                                 # 장소(비필수)
            r.get("학생상담신청내용", ""),
            r.get("교수답변내용", ""),
            PUBLIC_YN_DEFAULT,                  # "Y"
        ])

    wb.save(OUTPUT_XLSX)
    print(f"[OK] Saved: {OUTPUT_XLSX}")


# ======================================================
# 스레드 처리 (병렬 작업 단위)
# ======================================================
def process_thread(task: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    subject = task["subject"]
    msgs = task["msgs"]

    # (A) 교수 메시지 존재 확인 (답변 기준 시간 산정용)
    prof_msgs = [m for m in msgs if is_prof_message(m["from"]) and (m.get("body") or "").strip()]
    if not prof_msgs:
        return None

    # (B) 학생 메시지 존재 확인: 교수 아닌 발신자의 본문 메시지 최소 1개
    student_msgs = [m for m in msgs if (not is_prof_message(m["from"])) and (m.get("body") or "").strip()]
    if not student_msgs:
        return None  # 학생 미참여(교수만 보냄) 제외

    # 상담시간: "가장 이른 교수 답장" 기준
    rep = min(prof_msgs, key=lambda x: x["internal_ms"])
    start = round_30(ms_to_kst(rep["internal_ms"]))
    end = start + timedelta(minutes=30)

    # 학번: 이메일 우선(학생 메시지 발신자에서)
    sid_email = ""
    for m in student_msgs:
        sid_email = extract_student_id_from_email(m["from"])
        if sid_email:
            break

    client = OpenAI()
    result = call_llm_json(client, build_prompt(subject, msgs))

    # (C) 라벨 오지정 검증: 학생 상담 아니면 제외
    if not bool(result.get("is_student_thread")):
        return None

    item = result.get("item")
    if not isinstance(item, dict):
        return None

    name = str(item.get("student_name", "")).strip()
    if not name:
        return None  # 이름 없으면 제외(요구사항)

    # 학번: 이메일이 최우선, 없으면 LLM값 사용(둘 중 하나만 성공해도 활용)
    sid_llm = normalize_student_id_from_llm(item.get("student_id", ""))
    sid = sid_email if sid_email else sid_llm

    student_sum = str(item.get("student_request_summary", "")).strip()
    prof_sum = str(item.get("prof_reply_summary", "")).strip()

    if is_bad_summary(student_sum) or is_bad_summary(prof_sum):
        return None
    if not prof_sum:
        return None  # 답변 없으면 상담 아님(요구사항)

    code = str(item.get("category_code", "CF08")).strip()
    if code not in VALID_CODES:
        code = "CF08"

    return {
        "학번": sid,
        "성명": name,
        "상담일": start.strftime("%Y-%m-%d"),
        "상담시작시간": start.strftime("%H:%M"),
        "상담종료시간": end.strftime("%H:%M"),
        "상담유형": code,
        "학생상담신청내용": student_sum,
        "교수답변내용": prof_sum,
    }


# ======================================================
# 메인 + 진행 로그
# ======================================================
def main():
    service = authenticate_gmail()
    threads = get_threads(service, LABEL_NAME, PERIOD)
    total = len(threads)
    print(f"[INFO] threads matched query: {total}")

    # 1) Gmail에서 스레드 메시지 수집 (진행 로그)
    prepared: List[Dict[str, Any]] = []
    for i, t in enumerate(threads, 1):
        subject, msgs = get_thread_messages(service, t["id"])
        short_subject = (subject or "").strip()
        if len(short_subject) > 60:
            short_subject = short_subject[:60] + "..."

        if i == 1 or i % 10 == 0 or i == total:
            print(f"[FETCH] {i}/{total} thread='{short_subject}' msgs={len(msgs)}")

        prepared.append({
            "i": i,
            "total": total,
            "subject": subject or "",
            "short_subject": short_subject,
            "msgs": msgs,
        })

    print(f"[INFO] fetched threads: {len(prepared)}. start LLM with {MAX_WORKERS} workers")

    # 2) 병렬 처리 + 진행 로그
    records: List[Dict[str, Any]] = []
    lock = threading.Lock()
    done_count = 0

    def log(msg: str):
        with lock:
            print(msg)

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        futures = {ex.submit(process_thread, task): task for task in prepared}

        for fut in as_completed(futures):
            task = futures[fut]
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
    save_to_excel(records)
    print(f"[DONE] rows appended: {len(records)}")


if __name__ == "__main__":
    main()

