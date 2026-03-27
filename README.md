# VibrantSheets

VibrantSheets는 브라우저에서 동작하는 스프레드시트 앱입니다. 빠른 편집 성능과 직관적인 UI, 엑셀 호환성을 목표로 합니다.

## 최신 진행 현황
최신 상태는 아래 문서를 기준으로 합니다.
- documents/PROCESS.md
- documents/STATUS.md

## 주요 기능
- 그리드/편집: Ready/Edit/Enter 모드, 범위 선택, Fill Handle, 복사/붙여넣기
- 스타일: 굵게/기울임/밑줄/취소선, 글자색/배경색, 정렬, 폰트/크기
- 테두리: 9종 테두리, 내부선, 병합 경계 우선순위 처리
- 데이터 포맷: General/Number/Text/Currency(KRW/USD)/Percentage/Date
- 다중 시트: 추가/삭제/이름 변경/이동
- 찾기/바꾸기: 옵션(대소문자 구분/정확히 일치) 및 순차 이동
- 파일 입출력: VSHT/XLSX/CSV 저장·불러오기
- 인쇄: 인쇄 모달/프리뷰, 인쇄영역 기반 출력
- 이미지: 삽입/이동/리사이즈/삭제, 저장/불러오기

## 엑셀 호환성
- 기본 셀 크기 기준 동기화 (64 x 22 px)
- XLSX 라운드트립: 병합/서식/통화 포맷 보존
- 이미지: 엑셀 저장/불러오기 크기 보존(EXT 기반)

## 실행 방법
- `index.html`을 최신 브라우저에서 열면 실행됩니다.
- 파일 열기/저장 기능은 Chromium 계열(Chrome/Edge)에서 원활합니다.

## 문서
- 진행 프로세스: documents/PROCESS.md
- 프로젝트 구조: documents/PROJECT_STRUCTURE.md
- 체크리스트: documents/REGRESSION_CHECKLIST.md
- 트러블슈팅: documents/TROUBLESHOOTING.md
- 계획: documents/plan.md

## 기술 스택
- Language: Vanilla JavaScript (ES6+)
- Styling: Vanilla CSS
- Library: ExcelJS (XLSX 입출력)
- API: File System Access API
