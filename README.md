# mail2counseling

**mail2counseling**은 gmail에 있는 학생–교수 상담 메일 중 **교수가 라벨로 지정한 메일만**을 대상으로,
OpenAI API(유료)를 이용해 상담 내용을 요약하고, 학교 상담 시스템(SunShine)에 업로드 가능한 엑셀 파일로 자동 변환해주는 도구입니다.

> 교수 개인 PC에서 실행하는 **로컬 스크립트**이며,  
> 각 교수가 GitHub에서 clone 해서 **자기 계정으로 독립적으로 사용**하는 것을 전제로 설계되었습니다.

---

## 먼저 알아야 할 것 (필수)

### 1️⃣ OpenAI 계정 + 크레딧 필요

이 도구는 상담 메일을 요약하기 위해 **OpenAI API**를 사용합니다.

- OpenAI 계정과 API Key가 필요합니다.
- API 사용을 위해 **크레딧 충전**이 필요합니다. (최소 \$5부터 충전 가능)

**실제 사용량 기준 예시:**

- 상담 메일 스레드 약 **200건 (1년치)**
- 모델: **gpt-5-mini**
- 예상 비용: **약 \$1 이하**

👉 즉, **\$5를 한 번 충전하면 여러 해 동안 사용 가능**한 수준입니다.

> 메일 길이와 모델에 따라 실제 비용은 달라질 수 있으나,  
> 일반적인 교수–학생 상담 메일 기준으로는 비용 부담이 매우 낮습니다.

---

### 2️⃣ 개인정보 처리

- Gmail 메일 수집, 분류, 엑셀 저장은 **교수 개인 PC**(**로컬 환경**)에서 수행됩니다.
- 메일 본문 내용은 **요약 및 분류를 위해 OpenAI API로 전송**됩니다.
- OpenAI API는 요청 처리를 위해 텍스트를 일시적으로 처리하며,
  본 프로그램은 별도의 외부 서버나 데이터베이스에 메일을 저장하지 않습니다.

#### ⚠️ 주의 사항
- 메일 내용이 외부 API(OpenAI)로 전송되는 것이 부담되는 경우,
  본 도구 사용을 권장하지 않습니다.

---

## 주요 기능

- 지정 기간에 대한 Gmail 라벨 기반 상담 메일 자동 수집
- 학생–교수 상담 여부 자동 판별
- 학생 요청 / 교수 답변 요약
- 상담유형 자동 분류 (학교 코드 체계 대응)
- 엑셀 업로드 양식 자동 생성
- 설정/인증 정보 **프로젝트 디렉터리 내 저장** (self-contained)

---

## 디렉터리 구조

```text
mail2counseling/
├─ m2c.py                   # 메인 스크립트
├─ credentials.json         # Google OAuth (Desktop app)
├─ config.json              # 자동 생성됨 (설정 저장)
├─ token.json               # 자동 생성됨 (OAuth 토큰)
├─ counsel_excelupload.xlsx # 엑셀 템플릿
├─ output.xlsx              # 결과 파일
└─ .gitignore
```

## 사전 준비

### 1️⃣ Gmail 라벨링 (필수)

이 도구는 **Gmail 라벨이 붙은 메일만 처리**합니다.  
따라서 **학생과의 상담 메일에 라벨을 먼저 붙이는 작업이 필요**합니다.

자동 라벨링을 사용할 필요는 없습니다.  
**Gmail에서 제공하는 검색 + 필터 기능을 활용해 수동으로 라벨링**하면 충분합니다.


#### 권장 방법 (간단)

1. Gmail 검색창에서 필터링 규칙을 활용하여 상담 메일 후보를 검색  
   예: 교수님 @sungshin.ac.kr -공지 -안내 -설문 -noreply


2. 검색 결과를 보면서  
- 실제 상담 메일만 선택
- 라벨(`student` 등) 수동 적용

3. 결과를 확인하며 포함/제외 단어를 조금씩 추가·조정


#### 참고 사항 (중요)

- 라벨링이 다소 부정확해도 문제없음
- **LLM이 실제로 학생–교수 상담 메일인지 한 번 더 판별**
- 상담이 아니라고 판단되면 자동으로 제외됨

즉, 라벨은 *대략적인 후보 추리기* 용도로만 사용됩니다.
   

### 2️⃣ Python

Python 3.10 이상 권장

    python --version


### 3️⃣ 필수 패키지 설치

가상환경 사용을 권장합니다.

    python -m venv .venv
    source .venv/bin/activate
    pip install openai openpyxl google-api-python-client google-auth google-auth-oauthlib


## Google OAuth 설정 (credentials.json 만들기)

⚠️ 중요  
OAuth Client는 반드시 Desktop app 타입이어야 합니다.  

### A. GCP 프로젝트 생성 / 선택

- Google Cloud Console 접속 (https://console.cloud.google.com/)
- 상단 프로젝트 선택 → 새 프로젝트 생성 또는 기존 프로젝트 선택


### B. Gmail API 활성화

- APIs & Services → Library
- Gmail API 검색 → Enable


### C. OAuth 동의 화면 설정

- APIs & Services → OAuth consent screen
- User Type 선택
  - Google Workspace 계정이면 보통 Internal
  - 불가하면 External
- App name 등 최소 항목 입력 후 저장


### D. OAuth Client ID 생성

- APIs & Services → Credentials
- Create Credentials → OAuth client ID
- Application type: Desktop app
- Name 입력 (예: mail2counseling)
- Create


### E. credentials.json 다운로드

- 생성된 OAuth Client의 Download JSON 클릭
- 파일명을 credentials.json으로 변경
- 프로젝트 루트 디렉터리에 위치

      mail2counseling/credentials.json
      mail2counseling/m2c.py


## 기본 사용 흐름

### 1️⃣ 최초 설정 (실행 없이 설정만)

OpenAI API Key 저장

    python m2c.py --openai-api-key sk-xxxxxxxx

교수 이메일 저장

    python m2c.py --prof-email abc123@sungshin.ac.kr

Gmail 라벨 설정 (기본값: student)

    python m2c.py --label-name student

OpenAI 모델 설정

    python m2c.py --model gpt-5-mini


### 2️⃣ 사용 가능한 OpenAI 모델 확인

    python m2c.py --list-models

OpenAI API Key가 먼저 설정되어 있어야 합니다.


### 3️⃣ 실제 실행 (period 필수)

설정이 모두 완료되면, 실행 시에 --period가 필요합니다.

    python m2c.py --period "after:2025/01/01 before:2025/12/31"

- 최초 1회 실행 시 Google OAuth 인증 필요
- 이후에는 token.json을 재사용하여 자동 로그인
- 결과 파일: output.xlsx


## Gmail 검색 기간 (--period) 예시

Gmail search query 형식을 그대로 사용합니다.

    after:2025/01/01 before:2025/12/31
    after:2025/03/01
    before:2025/06/30


## config.json 설명

자동 생성되며 프로젝트 디렉터리에 저장됩니다.

    {
      "label_name": "student",
      "prof_email": "abc123@sungshin.ac.kr",
      "openai_api_key": "sk-...",
      "primary_model": "gpt-5-mini"
    }

- period는 저장하지 않으며 항상 인자로 입력


## 문제 해결

### 브라우저가 자동으로 열리지 않을 때

- 터미널에 출력되는 OAuth URL을 복사하여 브라우저에 직접 붙여넣기

### redirect_uri_mismatch 오류

- OAuth Client 타입이 Desktop app인지 확인
- Web application이면 Desktop app으로 새로 생성 후 credentials.json 교체


### 설정 초기화

    rm config.json token.json
