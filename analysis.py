# -*- coding: utf-8 -*-
import re
import json
import base64
import threading
from datetime import datetime, timezone, timedelta
from typing import List, Dict, Any, Tuple, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
import os
import argparse
import sys

import openpyxl
from openai import OpenAI
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request


# ======================================================
# 설정
# ======================================================
PERIOD = ""
LABEL_NAME = "student"
PROF_EMAIL = ""

TEMPLATE_XLSX = "counsel_excelupload.xlsx"
OUTPUT_XLSX = "output.xlsx"

PRIMARY_MODEL = "gpt-5-mini"  # config/args로 덮어씀

MAX_WORKERS = 4
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

COUNSEL_TYPE_DEFAULT = "3"
PUBLIC_YN_DEFAULT = "Y"

CF_CATEGORIES = {
    "학업": "CF01", "전공": "CF02", "장학금": "CF03", "진로": "CF04",
    "생활": "CF05", "휴학": "CF06", "자퇴": "CF07", "기타": "CF08",
    "멘토링 장학금": "CF09", "수강": "CF10", "성적": "CF11",
    "입학": "CF12", "진학": "CF13", "창업": "CF14", "사회봉사": "CF15",
    "건강": "CF16", "학술활동": "CF17", "논문": "CF18", "입학사정관": "CF19",
    "취업": "CF22", "대외활동": "CF23", "교환학생": "CF24", "현장실습": "CF25",
}
VALID_CODES = set(CF_CATEGORIES.values())


# ======================================================
# config.json
# ======================================================
BASE_DIR = Path(__file__).resolve().parent
def _config_path() -> Path:
    return BASE_DIR / "config.json"

def load_config() -> Dict[str, Any]:
    default_cfg = {
        "label_name": "student",
        "prof_email": "",
        "openai_api_key": "",
        "primary_model": "gpt-5-mini",
    }

    p = _config_path()
    if not p.exists():
        p.write_text(json.dumps(default_cfg, ensure_ascii=False, indent=2), encoding="utf-8")
        return default_cfg.copy()

    try:
        cfg = json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        p.write_text(json.dumps(default_cfg, ensure_ascii=False, indent=2), encoding="utf-8")
        return default_cfg.copy()

    changed = False
    for k, v in default_cfg.items():
        if k not in cfg:
            cfg[k] = v
            changed = True

    if changed:
        p.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")

    return cfg

def save_config(cfg: Dict[str, Any]) -> None:
    _config_path().write_text(
        json.dumps(cfg, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )


# ======================================================
# OpenAI 모델 유틸
# ======================================================
def get_available_models(api_key: str) -> List[str]:
    client = OpenAI(api_key=api_key)
    models = client.models.list()
    return sorted(m.id for m in models.data)

def print_available_models(api_key: str) -> None:
    models = get_available_models(api_key)
    print("[AVAILABLE MODELS]")
    for m in models:
        print(m)


# ======================================================
# Gmail
# ======================================================
def authenticate_gmail():
    creds_path = BASE_DIR / "credentials.json"
    token_path = BASE_DIR / "token.json"

    if not creds_path.exists():
        raise SystemExit(f"[ERROR] credentials.json 파일이 없습니다: {creds_path}")

    creds: Optional[Credentials] = None
    if token_path.exists():
        try:
            creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)
        except Exception:
            creds = None

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(creds_path), SCOPES)
            creds = flow.run_local_server(port=0)
        token_path.write_text(creds.to_json(), encoding="utf-8")

    return build("gmail", "v1", credentials=creds, cache_discovery=False)


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
# 학번 추출
# ======================================================
EMAIL_RE = re.compile(r"([A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})", re.I)

def _extract_email_addr(from_header: str) -> str:
    if not from_header:
        return ""
    m = EMAIL_RE.search(from_header)
    return (m.group(1) if m else from_header).strip().lower()

def extract_student_id_from_email(from_header: str) -> str:
    addr = _extract_email_addr(from_header)
    if not addr.endswith("@sungshin.ac.kr"):
        return ""
    local = addr.split("@", 1)[0].strip()
    return local if re.fullmatch(r"\d{8}", local) else ""

def normalize_student_id_from_llm(student_id: Any) -> str:
    sid = "" if student_id is None else str(student_id).strip()
    return sid if re.fullmatch(r"\d{8}", sid) else ""


# ======================================================
# OpenAI 호출
# ======================================================
def call_llm_json(client: OpenAI, prompt: str) -> Dict[str, Any]:
    resp = client.chat.completions.create(
        model=PRIMARY_MODEL,
        messages=[{"role": "user", "content": prompt}],
        response_format={"type": "json_object"},
    )
    content = (resp.choices[0].message.content or "").strip()
    if not content:
        raise RuntimeError("LLM empty response")
    return json.loads(content)


# ======================================================
# 요약 검증
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
# 프롬프트
# ======================================================
def build_prompt(subject: str, msgs: List[Dict[str, Any]]) -> str:
    blocks = []
    for m in msgs:
        role = "PROF" if is_prof_message(m["from"]) else "OTHER"
        dt_ = ms_to_kst(m["internal_ms"]).strftime("%Y-%m-%d %H:%M")
        body = (m.get("body") or "").strip()
        if not body:
            continue
        blocks.append(f"[{m['idx']}] {dt_} {role}\nFROM: {m['from']}\n{body}")

    joined = "\n---\n".join(blocks)
    cf_lines = "\n".join([f"- {name}: {code}" for name, code in CF_CATEGORIES.items()])

    return f"""
너는 '학생 이메일 상담 기록'을 엑셀로 정리하는 도우미다.

[1] 먼저 이 스레드가 "학생 1명 ↔ 교수({PROF_EMAIL})" 상담 대화인지 판정하라.
- 학생은 보통 상담/질문/요청을 하고, 교수는 답변/안내를 하는 형태임
- 단순 공지, 행정담당/조교/시스템 알림, 외부 업체 메일이면 학생 상담 아님

[2] 학생 상담이 맞다면 이 스레드는 "상담 1건"으로만 요약하라(주제 여러 개여도 1건).

[필수 조건]
- 교수 메시지(답변)와 학생 메시지(질문/요청) 둘 다 존재해야 함
- 교수 답변이 없거나 학생 메시지가 없으면 is_student_thread=false

[요약 규칙]
- 반드시 음슴체만 사용
- student_request_summary / prof_reply_summary 각각 2문장 이내
- 아래 표현 포함 금지: 합니다, 드립니다, 입니다, 하세요, 했습니다, 하셨다, 부탁, 감사, 했다, 하였다, 하고자, 하라
- 인사/감사/서명/링크/원문 복붙 금지, 핵심만 재서술

[교수답변 요약 스타일 규칙]
- prof_reply_summary는 "교수 본인이 작성한 말"처럼 작성해야 함(1인칭/발화자 관점)
- '교수는', '{PROF_EMAIL} 교수는', 'XXX 교수는' 같은 3인칭 표현 금지
- '안내함/말함/설명함' 같은 서술자 표현도 금지
- 예시(좋음): "수강변경은 수강정정기간에 처리 가능하다고 안내함. 필요 서류는 메일로 보내달라고 요청함"
- 예시(나쁨): "교수는 수강정정기간에 처리 가능하다고 안내함"

[상담유형 분류(명칭-코드)]
{cf_lines}

category_code 규칙:
- 위 목록의 코드 중 하나만 선택해서 반환
- 애매하면 기타(CF08)

반환 JSON(반드시 정확히):
{{
  "is_student_thread": true/false,
  "item": {{
    "student_name": "홍길동",
    "student_id": "20231234",
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
def round_30(dt_: datetime) -> datetime:
    total = dt_.hour * 60 + dt_.minute
    rounded = 30 * round(total / 30)
    return dt_.replace(
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
            r.get("학번", ""),
            r.get("성명", ""),
            COUNSEL_TYPE_DEFAULT,
            r.get("상담일", ""),
            r.get("상담시작시간", ""),
            r.get("상담종료시간", ""),
            r.get("상담유형", "CF08"),
            "",
            r.get("학생상담신청내용", ""),
            r.get("교수답변내용", ""),
            PUBLIC_YN_DEFAULT,
        ])

    wb.save(OUTPUT_XLSX)
    print(f"[OK] Saved: {OUTPUT_XLSX}")


# ======================================================
# 스레드 처리 (병렬 작업 단위)
# ======================================================
OPENAI_API_KEY = ""  # main에서 결정해 주입

def process_thread(task: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    subject = task["subject"]
    msgs = task["msgs"]

    prof_msgs = [m for m in msgs if is_prof_message(m["from"]) and (m.get("body") or "").strip()]
    if not prof_msgs:
        return None

    student_msgs = [m for m in msgs if (not is_prof_message(m["from"])) and (m.get("body") or "").strip()]
    if not student_msgs:
        return None

    rep = min(prof_msgs, key=lambda x: x["internal_ms"])
    start = round_30(ms_to_kst(rep["internal_ms"]))
    end = start + timedelta(minutes=30)

    sid_email = ""
    for m in student_msgs:
        sid_email = extract_student_id_from_email(m["from"])
        if sid_email:
            break

    client = OpenAI(api_key=OPENAI_API_KEY)
    result = call_llm_json(client, build_prompt(subject, msgs))

    if not bool(result.get("is_student_thread")):
        return None

    item = result.get("item")
    if not isinstance(item, dict):
        return None

    name = str(item.get("student_name", "")).strip()
    if not name:
        return None

    sid_llm = normalize_student_id_from_llm(item.get("student_id", ""))
    sid = sid_email if sid_email else sid_llm

    student_sum = str(item.get("student_request_summary", "")).strip()
    prof_sum = str(item.get("prof_reply_summary", "")).strip()

    if is_bad_summary(student_sum) or is_bad_summary(prof_sum):
        return None
    if not prof_sum:
        return None

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
# 메인 + 진행 로그 (현재 로직 유지 + 기능 복구)
# ======================================================
def main():
    global PERIOD, PROF_EMAIL, LABEL_NAME, OPENAI_API_KEY, PRIMARY_MODEL

    class MyHelpFormatter(argparse.HelpFormatter):
        def __init__(self, prog):
            super().__init__(prog, max_help_position=40, width=100)

    ap = argparse.ArgumentParser(formatter_class=MyHelpFormatter)
    ap.add_argument("--period", help="대상 기간. (e.g., --period \"after:2025/01/01 before:2025/12/31\")")
    ap.add_argument("--prof-email", help="교수자 이메일. (e.g., --prof-email \"jwpark12@sungshin.ac.kr\")", default=None)
    ap.add_argument("--openai-api-key", help="OpenAI API Key. (e.g., --openai-api-key \"sk-proj-XXXXXXXXXXXXXX\")", default=None)
    ap.add_argument("--label-name", help="대상 이메일 라벨. (e.g., --label-name \"student\")", default=None)
    ap.add_argument("--model", help="사용할 OpenAI 모델. (e.g., --model \"gpt-5-mini\")", default=None)
    ap.add_argument("--list-models", action="store_true", help="사용 가능한 OpenAI 모델 목록 출력")
    args = ap.parse_args()

    cfg = load_config()

    # 1) OpenAI API Key 결정 (period 없어도 설정 가능)
    OPENAI_API_KEY = (
        (args.openai_api_key or "").strip()
        or (os.environ.get("OPENAI_API_KEY", "") or "").strip()
        or (cfg.get("openai_api_key", "") or "").strip()
    )

    # 모델 목록 출력
    if args.list_models:
        if not OPENAI_API_KEY:
            ap.error("--list-models에는 OpenAI API Key가 필요합니다. (--openai-api-key 또는 OPENAI_API_KEY)")
        print_available_models(OPENAI_API_KEY)
        return

    # 2) config 갱신(Period 없이도 가능)
    LABEL_NAME = (args.label_name.strip() if args.label_name is not None else cfg.get("label_name", "student")).strip()
    if not LABEL_NAME:
        LABEL_NAME = "student"

    PROF_EMAIL = (args.prof_email.strip() if args.prof_email is not None else cfg.get("prof_email", "")).strip()

    requested_model = (args.model.strip() if args.model is not None else cfg.get("primary_model", "gpt-5-mini")).strip()
    if not requested_model:
        requested_model = "gpt-5-mini"

    # 3) 모델 검증 (키 있을 때만)
    if OPENAI_API_KEY:
        available_models = get_available_models(OPENAI_API_KEY)
        if requested_model not in available_models:
            print("[ERROR] 요청한 모델을 사용할 수 없습니다:", requested_model)
            print("\n사용 가능한 모델:")
            for m in available_models:
                print(" ", m)
            sys.exit(1)

    PRIMARY_MODEL = requested_model

    # 4) config 저장 (키/이메일/라벨/모델)
    if OPENAI_API_KEY:
        cfg["openai_api_key"] = OPENAI_API_KEY
    cfg["label_name"] = LABEL_NAME
    cfg["prof_email"] = PROF_EMAIL
    cfg["primary_model"] = PRIMARY_MODEL
    save_config(cfg)

    print("[CONFIG] saved:")
    print(f"  label_name     = {LABEL_NAME}")
    print(f"  prof_email     = {PROF_EMAIL or '(empty)'}")
    print(f"  primary_model  = {PRIMARY_MODEL}")
    print(f"  openai_api_key = {'set' if OPENAI_API_KEY else 'not set'}")

    # 5) config 완성 여부
    config_complete = bool(cfg.get("openai_api_key") and cfg.get("prof_email") and cfg.get("primary_model"))
    if not config_complete:
        print("[INFO] 설정이 아직 완성되지 않았습니다. (period 없이 종료)")
        print("       필요한 항목: openai_api_key, prof_email, primary_model")
        return

    # 6) 실행 모드: period 필수
    PERIOD = (args.period or "").strip()
    if not PERIOD:
        ap.error("설정이 완료되었으므로 실제 실행에는 --period가 필요합니다.")

    print(f"[CONFIG] period='{PERIOD}' (not saved)")

    # 7) Gmail 처리 + 기존 로직 복구
    service = authenticate_gmail()
    threads = get_threads(service, LABEL_NAME, PERIOD)
    total = len(threads)
    print(f"[INFO] threads matched query: {total}")

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

    save_to_excel(records)
    print(f"[DONE] rows appended: {len(records)}")


if __name__ == "__main__":
    main()
