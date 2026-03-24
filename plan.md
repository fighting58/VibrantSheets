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

### 3.1 그리드 시스템 (Grid System)
- 무한 스크롤(가상화 렌더링) 지원
- 행(Row) 및 열(Column) 헤더 표시 (A, B, C... / 1, 2, 3...)
- 셀 선택, 다중 선택 지원

### 3.2 셀 편집 및 서식 (Cell Editing & Formatting)
- 인라인 셀 편집 및 엔터/탭 키를 이용한 셀 간 이동
- 텍스트 스타일링 (굵게, 기울임꼴, 색상, 정렬) 및 셀 배경색
- 숫자 형식 지정 (통화, 퍼센트, 날짜)

### 3.3 고급 인터랙션 (Advanced Interaction)
- **끌어서 채우기(Fill Handle)**: 선택 영역 우측 하단 핸들을 드래그하여 데이터 자동 채우기
- **다중 셀 선택(Range Selection)**: 클릭 드래그 또는 Shift+클릭으로 여러 셀 동시 선택
- **실시간 수식 미리보기**: 수식 입력 시 결과값 실시간 계산

### 3.4 수식 처리 (Formula Engine)
- 기본 산술 연산 (+, -, *, /)
- 주요 내장 함수 (SUM, AVG, COUNT, MIN, MAX 등)
- 셀 참조 및 범위 참조 (A1, B2:C5 등)

### 3.5 시스템 유틸리티 (System Utilities)
- **Undo/Redo**: 실행 취소 및 다시 실행 (Ctrl+Z, Ctrl+Y)
- **자동 저장**: 브라우저 로컬 스토리지를 이용한 실시간 세션 유지
- **데이터 호환**: JSON 또는 CSV 형식으로 데이터 저장 및 불러오기

### 3.6 CSV 불러오기 / 내보내기 (CSV Import / Export) ✅
- **CSV 불러오기**: 리얼 파서를 통한 멀티라인(줄바꿈) 셀 완벽 복원 지원
  - 구분자 자동 감지 (쉼표, 탭, 세미콜론)
- **CSV 내보내기**: `File System Access API`를 이용한 OS 표준 '다른 이름으로 저장' 지원
  - UTF-8 BOM 포함 (Excel 한글 호환)

### 3.7 행/열 크기 조절 및 상태 관리 ✅
- **크기 조절**: 헤더 경계선 드래그를 통한 독립적인 너비/높이 조절
- **실시간 상태 배지**: 'Edited' (수정됨) 및 'Saved' (저장됨) 상태 표시 시스템

## 4. 디자인 컨셉 (Design Concept)
- **Vibrant & Premium**: 유리 질감(Glassmorphism), 다크 모드 기반의 세련된 디자인
- **Micro-interactions**: 부드러운 상태 전환 애니메이션 및 리사이즈 가이드

## 5. 단계별 개발 계획 (Implementation Roadmap)
1. **Phase 1**: 프로젝트 초기화 및 기본 레이아웃 ✅
2. **Phase 2**: 그리드 시스템 및 키보드 네비게이션 ✅
3. **Phase 2.5**: 다중 셀 선택, 채우기 핸들 확장, 클립보드 ✅
4. **Phase 3**: CSV 고급 불러오기 / 내보내기 (OS 표준 저장창) ✅
5. **Phase 4**: 행/열 크기 조절 및 실시간 상태 관리 ✅
6. **Phase 5**: 스타일링 툴바 및 서식 적용 기능 (B, I, 색상 등) ⬅️ 다음
7. **Phase 6**: 수식 엔진 파서 구현 및 참조 로직 연동
8. **Phase 7**: GitHub 저장소 생성 및 코드 업로드 ✅
