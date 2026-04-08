# NT 서비스센터 — 리드타임 대시보드

엑셀 파일을 업로드하면 자동으로 대시보드가 생성되어 웹에 공개됩니다.

## 최초 설정 (1회만)

### 1단계: GitHub에 새 저장소 만들기

1. https://github.com/new 접속
2. Repository name: `nt-dashboard` 입력
3. **Private** 선택 (사내 전용일 경우) 또는 Public
4. **Create repository** 클릭

### 2단계: 이 폴더를 GitHub에 업로드

1. 생성된 저장소 페이지에서 **"uploading an existing file"** 클릭
2. 이 폴더의 **모든 파일과 폴더**를 드래그 앤 드롭
   - `.github/` 폴더 (workflows 포함)
   - `data/` 폴더
   - `build.py`
   - `template.html`
   - `.gitignore`
   - `README.md`
3. **Commit changes** 클릭

> 주의: `.github` 폴더는 숨김 폴더입니다.
> 탐색기에서 보이지 않으면 상단 메뉴 > 보기 > 숨김 항목 체크

### 3단계: GitHub Pages 활성화

1. 저장소 페이지 > **Settings** 탭
2. 왼쪽 메뉴에서 **Pages** 클릭
3. **Source**를 **GitHub Actions**로 선택
4. 저장 (Save)

### 4단계: 첫 빌드 실행

1. 저장소 페이지 > **Actions** 탭
2. **Build Dashboard & Deploy** 워크플로우 클릭
3. 오른쪽 **Run workflow** > **Run workflow** 클릭
4. 약 1~2분 후 초록색 체크가 나오면 완료

### 5단계: 웹페이지 확인

빌드 완료 후 아래 주소로 접속:
```
https://[GitHub아이디].github.io/nt-dashboard/
```

## 데이터 업데이트 방법

### 매월 데이터 갱신

1. GitHub 저장소 페이지 접속
2. `data/전월/` 폴더 클릭
3. 기존 파일 삭제 후 전월 엑셀 2개 업로드
4. `data/현월/` 폴더 클릭
5. 기존 파일 삭제 후 현월 엑셀 2개 업로드
6. **Commit changes** 클릭
7. 자동으로 빌드 → 1~2분 후 웹페이지 갱신

### 파일 업로드 방법 (상세)

1. 저장소에서 `data/현월/` 폴더로 이동
2. 기존 파일 클릭 > 오른쪽 **...** > **Delete file** > Commit
3. **Add file** > **Upload files** 클릭
4. 새 엑셀 파일 드래그 앤 드롭
5. **Commit changes** 클릭

### 전년 트렌드 데이터

`data/전년/01/` ~ `data/전년/12/` 폴더에 월별 `서비스_리드타임_현황.xlsx`를 업로드하면 심화분석 탭의 월별 트렌드가 자동 업데이트됩니다.

## 폴더 구조

```
nt-dashboard/
├── .github/workflows/deploy.yml  ← 자동 빌드 설정
├── build.py                       ← 빌드 스크립트
├── template.html                  ← 대시보드 템플릿
├── data/
│   ├── 현월/                      ← 현재 월 데이터
│   │   ├── 서비스_리드타임_현황.xlsx
│   │   └── RoReport.xlsx
│   ├── 전월/                      ← 전월 데이터
│   │   ├── 서비스_리드타임_현황.xlsx
│   │   └── RoReport.xlsx
│   └── 전년/                      ← 월별 트렌드
│       ├── 01/ ~ 12/
└── output/                        ← 자동 생성 (건드리지 않음)
```

## 로컬에서 실행

```bash
pip install pandas openpyxl
python build.py
```
생성된 `output/서비스_리드타임_대시보드.html`을 브라우저에서 열기
