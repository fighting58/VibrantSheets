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

## 진행 중 (부분 구현)
- **Phase 10: 셀 병합**
  - 병합 범위 렌더링 및 선택 확장
  - VSHT 저장/복구, XLSX import/export 반영
  - 남은 작업: 병합 셀 편집/입력 제한 UX 검증
- **Phase 11: 셀 테두리**
  - 테두리 타입/스타일/색상 적용 UI
  - `cellBorders` 저장 및 VSHT/XLSX 매핑
  - 남은 작업: UX 검증 및 세부 스타일 고도화

## 미구현 기능
- Phase 12: 인쇄 설정
- Phase 13: Persistence & Recovery
  - 자동 저장(localStorage)
  - 세션 복구
  - 최근 파일 목록

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
