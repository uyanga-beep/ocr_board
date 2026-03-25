# Role

당신은 건설 현장의 업무 자동화를 돕는 시니어 Python 개발자이자 AI 엔지니어입니다. 주어진 요구사항을 바탕으로 완성도 높은 웹 애플리케이션 코드를 작성해 주세요.



# 프로젝트 개요

- **프로젝트명**: 건설 현장 동산보드판 OCR 자동화 및 데이터 통합 관리 시스템

- **목적**: 현장에서 촬영한 보드판 사진을 업로드하면, AI가 텍스트를 추출/교정하고 엑셀 파일로 변환합니다. 변환된 엑셀 행(Row)에는 원본 사진의 썸네일이 자동으로 삽입되어야 합니다.



# 기술 스택

- **Frontend/UI**: Streamlit (빠른 웹 프로토타이핑 및 사용성 확보)

- **Backend**: Python

- **OCR API**: Upstage Document Parse API (비정형/손글씨 텍스트 추출용)

- **LLM / 구조화**: Google Gemini API (`gemini-2.5-flash`) — 추출된 텍스트를 JSON 형태로 정제 및 건설 용어 오타 교정

- **Excel 처리**: `openpyxl`, `Pillow` (데이터 작성 및 이미지 삽입, 셀 크기 조절)

- **이미지 전처리**: OpenCV (`opencv-python-headless`) — 보드판 흰색 박스 영역 감지·크롭



# 주요 기능 및 구현 로직



## 0. 이미지 전처리 — 흰 박스 크롭

- 원본 사진에서 OpenCV를 사용해 **흰색 직사각형 영역(보드판 표)을 감지**하고, 해당 영역만 크롭한다.
- 배경의 자재 라벨, 철근, 바닥 글자 등 **노이즈를 제거**하여 OCR 정확도를 높인다.
- 흰 박스를 찾지 못하면 원본 이미지를 그대로 사용한다.



## 1. 이미지 업로드 및 UI (Streamlit)

- 사용자가 여러 장의 사진(보드판 이미지)을 한 번에 업로드할 수 있는 File Uploader 구현.
- 업로드된 이미지를 화면에 갤러리 형태로 미리보기 제공.
- "변환 시작" 버튼을 누르면 전체 처리 파이프라인 가동.



## 2. Vision OCR 처리 (Upstage Document Parse API)

- 크롭된 이미지를 Upstage API (`https://api.upstage.ai/v1/document-digitization`)로 전송하여 텍스트 데이터 추출.
- `model=document-parse`, `ocr=force` 옵션 사용.
- API Key는 `.env`로 관리 (`UPSTAGE_API_KEY`).
- 회사망 SSL 오류 대응: `verify=False` 적용.



## 3. 데이터 구조화 및 현장 용어 교정 (Gemini LLM)

- OCR로 추출된 Raw Text를 **Google Gemini API** (`gemini-2.5-flash`)에 프롬프트와 함께 전달하여 정형화된 JSON으로 변환.

- **추출 항목 (4개 컬럼만 인식)**:
  - `project_name`: 공사명
  - `category`: 공종
  - `location`: 위치
  - `details`: 내용

- **Gemini 호출 방식**: REST API 직접 호출 (`requests`, `verify=False`)
  - 엔드포인트: `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent`
  - `responseMimeType: "application/json"` 으로 JSON 응답 강제.
  - API Key는 `.env`로 관리 (`GEMINI_API_KEY`).
  - 429 (요청 제한) 시 자동 재시도 (최대 5회, 점진적 대기).

- **프롬프트 요구사항**: "너는 건설 현장 데이터 교정 AI야. OCR로 추출된 텍스트 중 인식 오류(예: 숫자 오기입, 필체에 따른 글자 깨짐)를 문맥에 맞게 수정하고, 특히 건설 전문 용어 사전을 기반으로 오타를 자동으로 교정해. 결과는 반드시 JSON 포맷으로 출력해."



## 4. 엑셀 변환 및 썸네일 이미지 매칭 (핵심 기능)

- 구조화된 JSON 데이터를 바탕으로 엑셀(.xlsx) 파일 생성.

- **Excel 구조 (6열)**:
  - A열: 사진 (썸네일)
  - B열: 파일명
  - C열: 공사명
  - D열: 공종
  - E열: 위치
  - F열: 내용

- **이미지 삽입 로직 (`openpyxl` 사용)**:
  - 원본 이미지를 `Pillow`를 사용하여 썸네일 크기(120×120 픽셀)로 리사이징.
  - 리사이징된 이미지를 A열의 각 행(Row)에 삽입.
  - 이미지가 셀 안에 맞게 들어가도록 행 높이(Row height)와 A열 너비(Column width)를 동적으로 조절.



## 5. 작업 로그 기록

- 매번 실행할 때마다 `work_log.txt`에 아래 정보를 **누적 기록**:
  - 실행 일시
  - 처리한 사진 수
  - OCR 사용 토큰/페이지 수
  - LLM(Gemini) 사용 토큰 수 (입력 + 출력)
  - 예상 소모 비용 (Upstage OCR + Gemini 각각)



# 단계별 개발 요청 사항

아래 순서대로 코드를 작성해 주세요. 한 번에 모든 코드를 주지 말고, 단계별로 확인하며 진행합시다.

1. **Step 1**: 프로젝트 폴더 구조를 제안하고, 필요한 라이브러리(`requirements.txt`)를 작성해 주세요.

2. **Step 2**: Streamlit을 이용한 기본 UI(파일 업로드 및 미리보기) 코드를 작성해 주세요 (`app.py`).

3. **Step 3**: Upstage API를 연동하여 이미지를 보내고 텍스트를 받아오는 모듈을 만들어 주세요 (`ocr_utils.py`). API Key는 `.env` 처리. **이미지 전처리(흰 박스 크롭)**를 OCR 전에 적용.

4. **Step 4**: 받아온 텍스트를 **Gemini API**를 통해 JSON으로 구조화하는 함수를 만들어 주세요. (건설 용어 교정 프롬프트 포함, 컬럼은 **공사명·공종·위치·내용** 4개만)

5. **Step 5**: 추출된 데이터를 바탕으로 `openpyxl`을 사용해 이미지가 포함된 엑셀 파일을 생성하고 다운로드 버튼을 제공하는 로직을 추가해 주세요. **매 실행마다 `work_log.txt`에 사진 수·토큰 수·비용을 기록.**
