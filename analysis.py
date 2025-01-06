# -*- coding: utf-8 -*-
from openai import OpenAI
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import base64
import email
import openpyxl
import math
import pandas as pd
from datetime import datetime
import json

# Gmail API SCOPES
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

# 1. Gmail API 인증
def authenticate_gmail():
    creds = None
    flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
    creds = flow.run_local_server(port=0)
    return build('gmail', 'v1', credentials=creds)

# 2. Gmail 대화(Conversation) 가져오기
def get_threads(service, query):
    query += ' label:inbox'  # inbox 레이블 필터 추가
    threads = []
    response = service.users().threads().list(userId='me', q=query).execute()
    if 'threads' in response:
        threads.extend(response['threads'])

    while 'nextPageToken' in response:
        response = service.users().threads().list(userId='me', q=query, pageToken=response['nextPageToken']).execute()
        threads.extend(response['threads'])

    return threads

# 3. 특정 thread의 모든 메일 가져오기
def get_thread_messages(service, thread_id):
    thread = service.users().threads().get(userId='me', id=thread_id).execute()
    messages = thread.get('messages', [])
    if len(messages) <= 1:  # 메시지가 하나만 있는 경우 처리하지 않음
        return "", ""
    thread_body = ""
    thread_subject = ""

    if messages:
        for header in messages[0]['payload']['headers']:
            if header['name'] == 'Subject':
                thread_subject = header['value']
                break

    for message in messages:
        payload = message['payload']
        body = ''
        if 'data' in payload['body']:
            body = payload['body']['data']
        else:
            parts = payload.get('parts', [])
            for part in parts:
                if part['mimeType'] == 'text/plain' and 'data' in part['body']:
                    body = part['body']['data']
                    break

        thread_body += base64.urlsafe_b64decode(body).decode('utf-8', errors='ignore') + "\n"

    return thread_body, thread_subject

# 4. ChatGPT를 사용하여 대화 단위 분석
def analyze_thread_with_chatgpt(thread_body):
    client = OpenAI()
    prompt = f"""나는 컴퓨터공학과 박지웅 교수야. 아래는 이메일 대화 내용이야. 이 대화가 학생과의 대화인지 판단해서 예 또는 아니오로 대답해줘. 내가 학생이라는 호칭을 사용해야 학생이야. XXX 학생이라는 내용이 있을 경우에만 학생과의 대화로 인정해줘. 학생이 맞다면 상담 요청 내용 또는 질문과 그에 대한 나의 답변을 요약해줘. 요약은 음슴체로 작성해줘. 각각의 요약은 400자 이내로 작성해줘. 학과 정보나, 학번 정보가 없다면 빈칸으로 남겨줘. 학번은 보통 8자리 숫자야. 상담유형은 학업, 전공, 장학금, 진로, 생활, 멘토링 장학금, 수강, 성적, 입학, 진학, 창업, 사회봉사, 건강, 학술활동, 논문, 입학사정관, 취업, 대외활동, 교환학생, 현장실습 중 가장 적절한 것을 골라줘. 아래 내용을 json 형태로 반환해줘.
- 학생 여부: 예/아니오
- 학과
- 학번
- 이름
- 상담유형
- 상담요청 내용 또는 질문
- 답변

이메일 대화:
{thread_body}
"""
    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            #model="gpt-4o",
            messages=[
                {"role": "user", "content": prompt}
            ]
        )
        content = response.choices[0].message.content.strip()

        # 응답이 비어 있는 경우 처리
        if not content:
            raise ValueError("ChatGPT 응답이 비어 있습니다.")

        return content
    except Exception as e:
        print(f"ChatGPT API 호출 실패: {e}")
        print(f"Prompt 내용: {prompt[:500]}...")  # 프롬프트 일부 출력
        return None

# ChatGPT 응답을 정리하고 JSON으로 파싱
def parse_json_response(response_content):
    try:
        # 응답 내용에서 ```json과 ``` 제거
        response_content = response_content.strip()
        if response_content.startswith("```json"):
            response_content = response_content[7:]  # ```json 제거
        if response_content.endswith("```"):
            response_content = response_content[:-3]  # ``` 제거

        # JSON 형식으로 변환
        return json.loads(response_content)
    except json.JSONDecodeError as e:
        print(f"JSON 디코딩 실패: {e}")
        print(f"정리된 응답 내용: {response_content}")
        return None

# 5. 30분 단위로 시간 반올림
def round_to_nearest_30_minutes(hour, minute):
    """시간과 분을 30분 단위로 반올림"""
    total_minutes = hour * 60 + minute
    rounded_minutes = 30 * round(total_minutes / 30)  # 30분 단위로 반올림
    return divmod(rounded_minutes, 60)  # (hour, minute)

# 6. 데이터 저장
def save_to_custom_excel(data, input_file, output_file):
    # 기존 input.xlsx 파일 로드
    wb = openpyxl.load_workbook(input_file)
    sheet = wb.active

    # 데이터 추가
    for record in data:
        # 답장 날짜와 시간 정보 계산
        if record.get('메일 답장의 년') and record.get('메일 답장의 월') and record.get('메일 답장의 일'):
            reply_year = record['메일 답장의 년']
            reply_month = f"{record['메일 답장의 월']}월"  # 월 추가
            reply_date = record['메일 답장의 일']
            start_hour, start_min = round_to_nearest_30_minutes(
                int(record.get('메일 답장의 시간', 0)),
                int(record.get('메일 답장의 분', 0))
            )
            end_hour, end_min = round_to_nearest_30_minutes(start_hour, start_min + 30)
        else:
            reply_year = reply_month = reply_date = start_hour = start_min = end_hour = end_min = ""

        # 데이터 입력
        sheet.append([
            reply_year, reply_month, reply_date,  # year, month, date
            start_hour, start_min,               # startHour, startMin
            end_hour, end_min,                   # endHour, endMin
            record.get('학과(전공)', ''),         # Major
            record.get('학번', ''),              # ID
            record.get('이름', ''),              # name
            record.get('상담요청 내용 또는 질문', ''), # request
            record.get('상담유형', ''),           # purpose
            "이메일",                             # type
            record.get('답변', '')                # opinion
        ])

    # 새로운 파일로 저장
    wb.save(output_file)
    print(f"데이터가 '{output_file}' 파일에 저장되었습니다.")

# 실행 코드 수정
if __name__ == "__main__":
    # Gmail API 인증 및 서비스 초기화
    service = authenticate_gmail()

    # 2024년 동안 주고받은 대화 가져오기
    query = 'after:2024/12/30 before:2025/01/01'
    threads = get_threads(service, query)

    results = []

    for thread in threads:
        thread_id = thread['id']
        thread_body, thread_subject = get_thread_messages(service, thread_id)
        # `thread_body`가 비어 있는 경우 건너뜀
        if not thread_body.strip():
            continue

        print(f"Processing thread: {thread_subject}")  # 현재 처리 중인 대화 제목 출력

        # ChatGPT를 사용하여 대화 단위 분석
        analysis_result = analyze_thread_with_chatgpt(thread_body)
        if not analysis_result:
            print(f"분석 실패 for thread: {thread_subject if thread_subject else 'No Subject'} (응답 없음)")
            continue

        # JSON 형식으로 변환
        parsed_result = parse_json_response(analysis_result)
        if not parsed_result:
            continue  # JSON 파싱 실패 시 건너뛰기

        # 학생 여부 확인
        if parsed_result.get('학생 여부') == "예":

            # 예시 답장 날짜 설정 (임의로 첫 메시지 날짜 사용)
            try:
                thread_detail = service.users().threads().get(userId='me', id=thread_id).execute()
                messages = thread_detail.get('messages', [])
                if not messages:
                    raise ValueError("No messages found in thread.")

                headers = messages[0]['payload']['headers']
                date_header = next((header['value'] for header in headers if header['name'] == 'Date'), None)
                if not date_header:
                    raise ValueError("Date header not found.")
                date_obj = email.utils.parsedate_to_datetime(date_header)
                
                print(parsed_result)

                # 정리된 결과 저장
                results.append({
                    '메일 답장의 년': date_obj.year,
                    '메일 답장의 월': date_obj.month,
                    '메일 답장의 일': date_obj.day,
                    '메일 답장의 시간': date_obj.hour,
                    '메일 답장의 분': date_obj.minute,
                    '학과(전공)': parsed_result.get('학과', ''),
                    '학번': parsed_result.get('학번', ''),
                    '이름': parsed_result.get('이름', ''),
                    '상담요청 내용 또는 질문': parsed_result.get('상담요청 내용 또는 질문', ''),
                    '상담유형': parsed_result.get('상담유형', ''),
                    '답변': parsed_result.get('답변', '')  # 상담 내용 저장
                })
            except Exception as e:
                print(f"Error extracting date for thread: {thread}")
                continue
        else:
            print(f"학생 여부 == 아니오")

    # 기존 input.xlsx 파일을 수정하여 저장
    save_to_custom_excel(results, 'input.xlsx', 'output.xlsx')
