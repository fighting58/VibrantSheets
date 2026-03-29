# Project Structure

## Overview
VibrantSheets는 핵심 기능을 `engines/` 아래 모듈로 분리한 구조입니다.
UI/상태는 `vs_core.js`, 검색은 `vs_find.js`, 파일 I/O는 `vs_io.js`, 수식은 `formula_engine.js`가 담당합니다.

## Top-Level
```text
VibrantSheets/
  engines/
    bootstrap.js
    formula_engine.js
    vs_core.js
    vs_find.js
    vs_io.js
  DOCUMENTS/
    plan.md
    PROJECT_STRUCTURE.md
    REGRESSION_CHECKLIST.md
    TROUBLESHOOTING.md
  index.html
  style.css
  README.md
```

## Module Responsibilities

### `engines/vs_core.js`
- 메인 클래스 `VibrantSheets`
- 그리드 렌더링/선택/편집/IME 처리
- 병합/테두리/서식 상태 관리
- 인쇄 모달/프리뷰/인쇄영역/페이지브레이크 프리뷰
- 커스텀 인쇄 페이지 생성(페이지별 테이블 클론)
- 행/열 확장 및 키보드 내비게이션

### `engines/vs_find.js`
- 찾기/바꾸기 상태 관리
- 매치 하이라이트
- 다음/이전 탐색 및 일괄 치환

### `engines/vs_io.js`
- VSHT/CSV/XLSX import/export
- ExcelJS 연동
- 병합/스타일/숫자형식 매핑
- 내부 데이터 적재 보조(`setInternalData`, delimiter 감지 등)

### `engines/formula_engine.js`
- 수식 파싱/평가
- 기본 함수 처리
- 셀 참조/범위 참조 계산

### `engines/bootstrap.js`
- 앱 초기화 및 부트스트랩

## Script Load Order (`index.html`)
1. `engines/formula_engine.js`
2. `engines/vs_find.js`
3. `engines/vs_io.js`
4. `engines/vs_core.js`
5. `engines/bootstrap.js`

## Data Model (Sheet)
각 시트는 대략 아래 상태를 가집니다.
- `rows`, `cols`
- `data`
- `cellStyles`
- `cellFormats`
- `cellFormulas`
- `cellBorders`
- `mergedRanges`
- `printSettings`
- `colWidths`, `rowHeights`

## Current Notes
- 컬럼 헤더는 `A..Z` 이후 `AA, AB...`를 지원합니다.
- 열/행은 입력/붙여넣기/키보드 이동 중 필요 시 자동 확장됩니다.
- 통화 포맷은 KRW/USD를 구분해 import/export 라운드트립을 유지합니다.
- 인쇄는 페이지별 DOM을 만들어 출력하며, `style.css`의 print 전용 규칙을 함께 사용합니다.
- 기본 셀 크기는 엑셀 기본값에 맞춰 64/22(px 기준)로 동기화됩니다.
