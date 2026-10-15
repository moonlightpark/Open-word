# Oword — DOCX 문서 편집기

Microsoft Word 문서(.docx)를 열고, 편집하고, 저장하는 오픈 소스 편집기입니다.
[Opptx](https://github.com/moonlightpark/Open-Power-Point-Editor)(PPTX 편집기)와 같은 구조로 만들었습니다. 이 구조는 [HOP](https://github.com/golbin/hop)(HWP 편집기)을 참고한 것입니다.
**문서 엔진, 렌더러, 명령 기반 UI, 데스크톱 셸**의 네 층으로 되어 있습니다. 설계는 [docs/architecture.md](docs/architecture.md)에 있습니다.

## 설치

[Releases](https://github.com/moonlightpark/Open-word/releases)에서 운영체제에 맞는 파일을 받습니다.

| 운영체제 | 파일 | 참고 |
|---|---|---|
| Windows 10/11 (x64) | `Oword_x.y.z_x64-setup.exe` 또는 `.msi` | 서명하지 않은 빌드입니다. SmartScreen 경고가 뜨면 **추가 정보 → 실행**을 누릅니다. WebView2가 없으면 설치 중에 자동으로 받습니다. |
| macOS (Apple Silicon) | `Oword_x.y.z_aarch64.dmg` | 임시 서명(ad-hoc)이라 처음 실행할 때 앱을 **우클릭 → 열기**로 실행합니다. |
| macOS (Intel) | `Oword_x.y.z_x64.dmg` | 위와 같습니다. |
| Linux (x64) | `.AppImage`, `.deb`, `.rpm` | WebKitGTK 4.1이 필요합니다(Ubuntu 22.04 이상, Fedora 36 이상 등). AppImage는 `chmod +x` 후 실행합니다. |

릴리즈 빌드는 `v*` 태그를 푸시하면 GitHub Actions([`.github/workflows/release.yml`](.github/workflows/release.yml))가 자동으로 만듭니다.

## 실행

```bash
npm install
npm run dev          # http://localhost:5175 (샘플: /?file=./sample.docx)
npm test             # 엔진 테스트 (vitest + jsdom)
npm run build        # dist/ 웹 빌드

npx tauri dev        # 데스크톱 앱(Tauri 2) 개발 실행
npx tauri build      # 설치 파일 만들기
```

샘플 문서와 빈 문서는 `python3 scripts/make-fixtures.py`(python-docx 필요)로 다시 만들 수 있습니다.

## 기능

- **열기/저장**: .docx 열기(대화상자, 끌어다 놓기), 저장, 다른 이름으로 저장, 인쇄/PDF
  - 손대지 않은 문단, 표, 그림, 필드, 책갈피, 각주, 머리글/바닥글, 그 밖의 XML 파트는 그대로 보존됩니다.
- **렌더링**
  - 스타일 상속(문서 기본값 → 표 스타일 → 문단 스타일 → 글자 스타일 → 직접 서식), 테마 글꼴과 색
  - 글자: 글꼴, 크기, 진하게, 기울임, 밑줄(종류별), 취소선, 위/아래 첨자, 색, 형광펜, 음영, 대문자, 숨은 글자
  - 문단: 정렬, 들여쓰기/내어쓰기, 문단 위·아래 간격, 줄 간격(배수/최소/고정), 테두리, 음영, 탭
  - 목록: 글머리표와 번호(1. 가. ① i. A. 등 여러 형식, 9수준)
  - 표: 셀 병합(가로/세로), 테두리(표 스타일 포함), 셀 음영, 셀 여백, 세로 정렬
  - 그림(인라인/떠 있는 그림), 글상자, 쪽 번호 필드, 머리글/바닥글(첫 쪽·짝수 쪽 구분)
  - **쪽 나누기 보기**: 용지 크기와 여백대로 쪽을 나누어 보여 주고, 쪽마다 머리글과 바닥글을 그립니다. 제목은 다음 문단과 같은 쪽에 둡니다.
- **편집**
  - 글자 입력(한글 IME 포함), Enter로 문단 나누기, Shift+Enter 줄 바꿈, Tab(탭 · 목록 수준 · 다음 셀)
  - 글자 모양: 서식 바(스타일, 글꼴, 크기, 가/가/가/가, 첨자, 글자 색, 형광펜), 커서만 있을 때는 Word처럼 단어 전체 또는 다음 입력에 적용
  - 문단 모양 대화상자, 정렬, 줄 간격, 글머리표/문단 번호, 들여쓰기/내어쓰기, 스타일(표준, 제목 1–3 등)
  - 삽입: 표, 그림(파일, 붙여넣기, 끌어다 놓기), 쪽 나누기
  - 그림 크기 조절(모서리 끌기), 표 줄/칸 넣기·지우기
  - 복사/붙여넣기(문서 안에서는 서식 유지), 찾기/바꾸기, 되돌리기/다시 실행(300단계)
  - 쪽 설정(용지 · 방향 · 여백)
- **보기**: 확대/축소(Ctrl+휠), 쪽 너비/쪽 전체 맞춤, 눈금자(여백과 들여쓰기 표시), 탐색 창(제목 목록), 문단 부호, 글자 수

## 구조

```
src/engine/          문서 엔진 (UI 의존성 없음)
  package.ts         OPC 컨테이너: zip, 관계, 콘텐츠 형식, undo용 스냅샷
  document.ts        DocxDocument: 문단/표/그림/목록/쪽 설정 명령 + 트랜잭션 undo
  model.ts           문단 내용 모델(항목과 길이), 런 나누기/합치기, 속성 쓰기(스키마 순서 유지)
  styles.ts          스타일·테마 상속 계산
  numbering.ts       목록 정의와 번호 매기기
  render.ts          문서 → 편집 가능한 HTML 렌더러 (1pt = 20 twip = 12700 EMU)
  edit.ts            DOM 위치 ↔ 모델 위치, 편집한 DOM → w:p 반영
src/ui/              편집기 UI (HOP · Opptx와 같은 메뉴/툴바/서식 바/눈금자/상태바)
  editor.ts          contenteditable 편집기, 쪽 나누기 배치, 선택 영역 복원
src-tauri/           데스크톱 셸 (네이티브 파일 대화상자, 임시 파일 후 rename 방식 저장)
tests/               엔진 테스트 + 샘플 문서
```

## 제약

- 쪽 나누기는 문단 단위로 합니다. 한 쪽보다 긴 문단이나 표는 쪽 경계를 넘어 그려질 수 있습니다.
- 떠 있는 그림과 글상자는 문단 안의 인라인 개체로 보입니다(저장할 때 위치 정보는 보존됨).
- 차트, SmartArt, EMF/WMF 그림은 자리 표시 상자로 보입니다(보존됨).
- 각주/미주 본문, 메모, 변경 내용 추적은 보여 주지 않습니다(보존됨). 머리글/바닥글은 읽기 전용입니다.
- 줄바꿈은 브라우저 레이아웃을 쓰기 때문에 Word와 약간 다를 수 있습니다.

## 참고한 프로젝트

- **[Opptx](https://github.com/moonlightpark/Open-Power-Point-Editor)** (MIT License): 같은 저장소 구조, OPC 패키지 처리, 명령 체계, 대화상자, 데스크톱 셸, 릴리즈 워크플로를 그대로 이어받았습니다.
- **[HOP](https://github.com/golbin/hop)** (golbin/hop, MIT License): Opptx가 참고한 구조(문서 엔진 · 렌더러 · 명령 기반 UI · Tauri 셸)와 화면 배치, `category:action` 명령 체계, 임시 파일에 쓴 뒤 바꿔치는 저장 방식. HOP과 rhwp의 소스 코드는 포함하지 않았습니다.

## 라이선스

[MIT License](LICENSE) © 2026 moonlightpark

배포 앱에 포함된 서드파티 소프트웨어의 라이선스는 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)에 있습니다. 의존성이 바뀌면 `npm run notices`로 다시 만듭니다.

## 상표

Oword는 Microsoft Corporation과 관계가 없고, Microsoft가 보증하지 않습니다.
Microsoft와 Word는 Microsoft Corporation의 상표입니다.
DOCX 형식은 공개 표준인 ECMA-376 / ISO/IEC 29500(Office Open XML)에 따라 구현했습니다.
