# Excel-like Application (Project "VibrantSheets") Planning

A modern, high-performance, and visually stunning web-based spreadsheet application.

## 1. 개요 (Overview)
본 프로젝트는 웹 브라우저에서 동작하는 엑셀과 유사한 고성능 스프레드시트 애플리케이션을 구축하는 것을 목표로 합니다. 사용자에게 부드러운 가독성과 직관적인 인터페이스를 제공하며, 프리미엄급 디자인 요소를 적용합니다.

## 2. 기술 스택 (Tech Stack)
- **Core**: HTML5, Vanilla JavaScript (ES6+)
- **Styling**: Vanilla CSS3 (Custom Properties, Flexbox, Grid)
- **Rendering Engine**: HTML5 Canvas (그리드 및 셀 렌더링 성능 최적화를 위함)
- **Design System**: Glassmorphism, Dark Mode Support, Modern Typography (Inter/Roboto)

## 3. 핵심 기능 (Core Features)

### 3.1 그리드 시스템 (Grid System) ✅
- **무한 스크롤(Lazy Loading)**: 스크롤에 따라 행을 동적으로 추가하여 DOM 부하 최소화 및 성능 최적화
- 행(Row) 및 열(Column) 헤더 표시 (A, B, C... / 1, 2, 3...)
- 셀 선택, 다중 선택 지원 (Range Selection)

### 3.2 셀 편집 및 인터랙션 (Cell Editing & Interaction) ✅
- **Excel 스타일 상태 모델**: Ready (선택), Edit (더블 클릭/F2), Enter (즉시 타이핑/덮어쓰기) 상태 구현
- **인라인 편집**: `contenteditable`을 이용한 실시간 편집 및 방향키/엔터/탭 네비게이션
- **시각적 선택 피드백**: 전용 오버레이 레이어를 통한 3px 두께의 선명한 셀 선택 테두리 구현
- **스타일링 시스템**: 툴바를 통한 서식(굵게, 기울임, 밑줄, 취소선), 색상(글자, 배경), 정렬(L, C, R) 지원 ✅
- **IME 최적화**: '상시 편집(Always-Editable)' 모델을 통해 한글 입력 시 자음/모음 분리 및 영문 혼입 완벽 방지 ✅

### 3.3 고급 인터랙션 (Advanced Interaction) ✅
- **끌어서 채우기(Fill Handle)**: 선택 영역 우측 하단 핸들을 드래그하여 데이터 자동 채우기
- **다중 셀 선택(Range Selection)**: 클릭 드래그 또는 Shift+클릭으로 여러 셀 동시 선택
- **스마트 클립보드**: 엑셀/구글 시트의 다중 라인 및 따옴표 포함 텍스트 완벽 지원 (TSV 파서)

### 3.4 수식 처리 (Formula Engine)
- 기본 산술 연산 (+, -, *, /)
- 주요 내장 함수 (SUM, AVG, COUNT, MIN, MAX 등)
- 셀 참조 및 범위 참조 (A1, B2:C5 등)

### 3.5 시스템 유틸리티 (System Utilities)
- **Undo/Redo**: 실행 취소 및 다시 실행 (Ctrl+Z, Ctrl+Y)
- **자동 저장**: 브라우저 로컬 스토리지를 이용한 실시간 세션 유지
- **데이터 호환**: JSON 또는 CSV 형식으로 데이터 저장 및 불러오기

### 3.6 파일 시스템 및 외부 연동 ✅
- **.vsht 전용 포맷**: JSON 기반 커스텀 포맷으로 데이터, 열 너비, 행 높이 등 레이아웃 완벽 보존
- **Excel (.xlsx) 완벽 호환**: SheetJS를 이용한 바이너리 엑셀 파일 로드 및 내보내기 지원
- **스마트 파일 핸들링**: `File System Access API`를 이용해 한 번 연 파일은 'Save' 시 즉시 덮어쓰기, 'Save As...' 시 포맷 선택 저장 지원
- **CSV/TSV 지원**: UTF-8 BOM 인코딩을 적용하여 엑셀에서의 한글 깨짐 방지

### 3.7 행/열 크기 조절 및 상태 관리 ✅
- **크기 조절**: 헤더 경계선 드래그를 통한 독립적인 너비/높이 조절
- **실시간 상태 배지**: 'Edited' (수정됨) 및 'Saved' (저장됨) 상태 표시 시스템

### 3.8 테이블 구조 제어 (Table Operations) ✅
- **행 삽입**: 현재 선택 셀 위에 빈 행 삽입 — 아래 모든 셀 데이터 및 스타일 자동 시프트
- **행 삭제**: 현재 선택 행(들) 삭제 — 위 방향으로 데이터 병합, 커스텀 확인 모달 제공
- **열 삽입**: 현재 선택 열 왼쪽에 빈 열 삽입 — 오른쪽 모든 셀 자동 시프트
- **열 삭제**: 현재 선택 열(들) 삭제 — 데이터 및 스타일 좌측 병합
- **데이터 무결성**: 삽입/삭제 시 `data` 객체 및 `cellStyles` 맵의 모든 키를 재계산하여 일관성 유지
- **리본 UI**: 직관적인 아이콘(+/- 표시)이 적용된 4개 버튼을 리본 툴바에 배치

## 4. 디자인 컨셉 (Design Concept)
- **Vibrant & Premium**: 유리 질감(Glassmorphism), 다크 모드 기반의 세련된 디자인
- **Micro-interactions**: 부드러운 상태 전환 애니메이션 및 리사이즈 가이드

## 5. 단계별 개발 계획 (Implementation Roadmap)
1. **Phase 1~5 Complete**: 그리드 최적화, 파일 핸들링, 고급 셀 인터랙션, 스타일링 시스템 완료 ✅
2. **Phase 6: 행/열 제어 (Row/Column Operations)** ✅ **완료**
    - 행/열 삽입 및 삭제 기능 구현 (리본 툴바 UI 포함)
    - 데이터 시프트 엔진 (`shiftData`) 및 커스텀 확인 모달 구현
3. **Phase 7: 데이터 포맷팅 (Data Formatting)** ⬅️ **현재 목표**
    - 통화, 퍼센트, 이진수, 날짜 형식 표시기 구현
    - 소수점 자릿수 조절 기능
4. **Phase 8: 수식 엔진 (Formula Engine)**
    - 수식 파서 (Parser) 및 셀 참조 로직 (`=A1+B1`)
    - 기본 함수 구현 (`SUM`, `AVG`, `COUNT`, `MIN`, `MAX`)
5. **Phase 9: 자동 저장 및 세션 복구**
    - `localStorage`를 이용한 비정상 종료 대비 실시간 임시 저장
    - 최근 파일 목록 (Recent Files) 관리
6. **Phase 10: GitHub 프로젝트 관리 및 배포** ✅ (진행 중)
