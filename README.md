# 엑셀 정렬 & A4 인쇄용 이미지 생성 웹앱 (Streamlit)

엑셀을 업로드해서 두 가지 방식으로 분류(정렬)하고, 정렬된 엑셀(서식·색상·필터 그대로 유지)과
A4 인쇄용 이미지를 생성하는 웹앱입니다.

## 기능

- **대량 분류**: A열 → I열 → J열 → K열 순서로 순차 정렬
- **소량 분류**: F열 → G열 → A열 → K열 순서로 순차 정렬 후, H열 값이 같은 행이 기준 개수(50/40/30개
  중 선택, 기본 50) 이상이면 그 그룹만 별도 시트로 자동 분리
- **정렬된 엑셀 다운로드**: 원본 파일의 행(XML row) 자체를 재배치하는 방식이라 셀 서식·배경색·필터·
  열 너비 등 원본 정보가 전혀 손실되지 않습니다.
- **A4 인쇄용 이미지 생성**: 시트(그룹)마다 15개씩 좌우 2단으로 배치, 번호는 시트마다 1번부터 다시
  시작. A열 배경색을 그대로 재현하고, F·G열 글자가 길면 자동으로 폰트 크기를 줄여 겹침/잘림을 방지합니다.

## 로컬에서 실행하기

```bash
pip install -r requirements.txt
streamlit run app.py
```

브라우저가 자동으로 열리지 않으면 터미널에 표시된 `Local URL`(보통 http://localhost:8501)로 접속하세요.

## GitHub에 올리고 Streamlit Community Cloud에 배포하기

Claude(저)는 GitHub·Streamlit 계정 로그인이 필요한 작업을 대신 수행할 수 없어서(계정 인증이 필요한
영역이라 이 세션에서는 접근할 수 없습니다), 아래 순서대로 직접 진행해주시면 됩니다. 전부 무료로 가능하고
5~10분 정도 걸립니다.

### 1단계. GitHub 저장소 만들기

1. https://github.com 에 로그인 (계정이 없으면 무료 가입)
2. 오른쪽 위 `+` → `New repository` 클릭
3. Repository name 입력 (예: `excel-sort-print`), Public 선택 → `Create repository`
4. 저장소가 만들어지면 나오는 화면에서 **"uploading an existing file"** 링크 클릭
   (또는 저장소 페이지의 `Add file` → `Upload files`)
5. 이 폴더 안의 모든 파일(`app.py`, `xlsx_utils.py`, `image_utils.py`, `requirements.txt`,
   `packages.txt`, `.gitignore`, `README.md`, `test_core.py`)을 통째로 끌어다 놓기
   (또는 `Choose your files`로 선택)
6. 아래 `Commit changes` 클릭

> Git 명령어에 익숙하다면 다음과 같이 올려도 됩니다.
> ```bash
> git init
> git add .
> git commit -m "엑셀 정렬 & 인쇄 이미지 생성 웹앱"
> git branch -M main
> git remote add origin https://github.com/<내계정>/<저장소이름>.git
> git push -u origin main
> ```

### 2단계. Streamlit Community Cloud에 배포하기

1. https://share.streamlit.io 접속 → GitHub 계정으로 로그인
2. `Create app` (또는 `New app`) 클릭
3. `Repository`에서 방금 만든 저장소 선택, `Branch`는 `main`, `Main file path`는 `app.py` 입력
4. `Deploy` 클릭 → 1~2분 기다리면 배포 완료

배포가 끝나면 `https://<앱이름>-<임의문자>.streamlit.app` 형태의 링크가 생기고, 이 **링크를 공유하면
누구나 접속해서 함께 사용**할 수 있습니다. (Streamlit Cloud는 방문자마다 세션이 분리되어 동시에 여러
명이 각자 다른 파일을 올려도 서로 섞이지 않습니다.)

### 한글 폰트 관련 (중요)

A4 이미지 안의 한글이 깨지지 않으려면 서버에 한글 폰트가 설치되어 있어야 합니다. 이 저장소에는
`packages.txt` 파일에 `fonts-noto-cjk`가 이미 적혀 있어서, Streamlit Cloud가 배포할 때 자동으로
한글 폰트를 설치합니다. **`packages.txt` 파일이 저장소 최상위(= `app.py`와 같은 위치)에 반드시
포함되어 있어야 합니다.** (파일 업로드 시 실수로 빠뜨리지 않도록 확인해주세요.)

배포 후 앱 화면 위쪽에 "한글 폰트를 찾지 못했습니다" 경고가 뜨면 `packages.txt`가 누락되었거나
아직 적용되지 않은 것이니, Streamlit Cloud 앱 관리 화면에서 `Reboot app`을 눌러보세요.

### 코드나 UI를 수정하고 싶을 때

GitHub 저장소에서 파일을 수정(웹에서 연필 아이콘 클릭 후 편집 → Commit)하면 Streamlit Cloud가
자동으로 변경 사항을 감지해서 몇 초~몇십 초 안에 앱을 다시 배포합니다. 별도 작업이 필요 없습니다.

## 파일 구성

| 파일 | 설명 |
|---|---|
| `app.py` | Streamlit UI (업로드, 분류 버튼, 다운로드, 이미지 갤러리) |
| `xlsx_utils.py` | 엑셀 읽기, 정렬 로직, 서식 보존 다중 시트 저장 로직 |
| `image_utils.py` | A4(300dpi) 인쇄용 이미지 렌더링 로직 |
| `requirements.txt` | 파이썬 패키지 목록 |
| `packages.txt` | Streamlit Cloud 배포 시 설치할 시스템 패키지(한글 폰트) |
| `test_core.py` | 핵심 로직(정렬/서식보존/이미지생성)을 실제 엑셀로 검증하는 독립 테스트 스크립트 |

## 정렬 방식에 대한 참고

"A열 정렬 → I열 정렬 → J열 정렬 → K열 정렬" 은 Excel에서 각 열을 **차례대로 독립적인 정렬**로
적용한 것과 동일하게 동작합니다(안정 정렬). 그 결과 마지막에 적용한 열(대량 분류는 K열, 소량 분류는
K열)이 최종적으로 가장 우선순위가 높은 정렬 기준이 됩니다. 이 방식은 실제 사용자가 로컬 Excel에서
같은 순서로 정렬한 결과와 전체 행이 정확히 일치하는 것을 확인해 검증했습니다.
