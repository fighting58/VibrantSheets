# VibrantSheets

VibrantSheets는 Vanilla JavaScript/CSS로 만든 웹 기반 스프레드시트 앱입니다.  
빠른 그리드 상호작용, 리본 UI, 파일 입출력, 데이터 포맷팅을 중심으로 Excel 스타일 경험을 제공합니다.

## 현재 구현된 기능
- Grid & Interaction: Ready/Edit/Enter 모드, 무한 스크롤 행 추가, 범위 선택, Fill Handle, 복사/붙여넣기
- 스타일링: 굵게/기울임/밑줄/취소선, 텍스트/배경 색, 정렬, 폰트/크기
- 행/열 작업: 삽입/삭제, 데이터/스타일/포맷 재계산
- 선택 기능: 헤더 선택/전체 선택/연속 범위 확장
- 데이터 포맷팅: General/Currency(KRW)/Percentage/Date
- 다중 시트: 추가/삭제/이름 변경/이동 및 시트별 데이터 유지
- 찾기/바꾸기: 검색/치환/하이라이트/순차 이동
- 함수 엔진(간단): SUM/CONCAT/LEFT/RIGHT/MID, 참조/순환 참조 감지
- 파일 입출력: VSHT/XLSX/CSV 저장 및 불러오기

## 최근 업데이트 (완료)
- **셀 테두리 시스템 (Phase 11):**
  - 9종 아이콘 기반 직관적 배치 제어 (상/하/좌/우/외곽/전체/내부가로/내부세로/내부사각)
  - SVG 아이콘 기반 커스텀 선 스타일 선택기 (점선, 점선-굵게, 이중선 등 지원)
  - 미러링 렌더링(Mirror Rendering)을 적용하여 그리드 라인에 의한 테두리 잘림 현상 완벽 해결
- **IME 입력 최적화:**
  - 포커스 시 '전체 선택(Select All)' 전략을 도입하여 한글 첫 글자 분리 현상(rㅏ) 제거 및 매끄러운 입력 보장
- **UI 리밸런싱:**
  - 상단 리본 메뉴 및 아이콘 크기를 컴팩트하게 재조정하여 작업 공간 확대

## 예정 기능 (Upcoming)
- Phase 12: 자동 저장 및 세션 복구 (localStorage)
- Phase 13: 인쇄 및 PDF 내보내기 설정
- Phase 14: 이미지 삽입 및 자유 드로잉 레이어

## 문서
- 진행 프로세스: `documents/PROCESS.md`
- 구조 설명: `documents/PROJECT_STRUCTURE.md`
- 기존 계획: `documents/plan.md`

## 기술 스택
- Language: Vanilla JavaScript (ES6+)
- Styling: Vanilla CSS
- Library: ExcelJS (XLSX 입출력)
- API: File System Access API

## 실행 방법
`index.html`을 최신 브라우저에서 열면 실행됩니다.  
파일 저장/불러오기 경험은 Chromium 계열(Chrome, Edge)에서 가장 안정적입니다.
