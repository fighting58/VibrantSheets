# Project Structure

## Overview
VibrantSheets는 모든 JavaScript를 `engines/` 아래에 배치하고, 루트에는 HTML/CSS 및 문서만 두는 구조로 정리했습니다.
이 문서는 기능 구분의 기준과 각 모듈의 구현 범위를 상세히 설명합니다.

## Tree (Top-Level)
```
VibrantSheets/
├─ engines/
│  ├─ bootstrap.js
│  ├─ formula_engine.js
│  ├─ vs_core.js
│  ├─ vs_find.js
│  └─ vs_io.js
├─ documents/
│  ├─ PROCESS.md
│  ├─ PROJECT_STRUCTURE.md
│  └─ plan.md
├─ index.html
├─ style.css
├─ README.md
└─ .gitignore
```

## 기능 구분 기준
- UI와 화면 상태를 다루는 로직은 `vs_core.js`에 둡니다. (렌더링, 선택, 편집, 이벤트 바인딩)
- 검색/치환처럼 독립적으로 재사용 가능한 로직은 별도 모듈로 분리합니다. (`vs_find.js`)
- 파일 입출력처럼 외부 I/O와 형식 변환이 집중된 로직은 전용 모듈에 둡니다. (`vs_io.js`)
- 수식 계산은 독립 실행 가능한 엔진으로 유지합니다. (`formula_engine.js`)
- 초기화 및 인스턴스 생성은 최소한의 부트스트랩 파일로 분리합니다. (`bootstrap.js`)

이 기준은 “UI/상태/로직/IO” 경계를 명확히 하여, 변경 영향 범위를 최소화하는 데 목적이 있습니다.

## engines/ 역할과 구현 내용
### `vs_core.js`
메인 클래스 `VibrantSheets`를 포함하며 UI/상태 중심 로직을 담당합니다.
- 그리드 렌더링: 테이블 생성, 셀 생성, 스타일 렌더, 오버레이(선택/범위/핸들)
- 선택/입력: 단일/범위 선택, 키보드 네비게이션, 편집 모드 전환
- 셀 상태: `data`, `cellStyles`, `cellFormats`, `cellFormulas`, `mergedRanges` 관리
- 병합 로직: 병합 범위 확장, 병합 렌더링 적용, 선택 범위 보정
- 도구바 바인딩: 서식/정렬/포맷 관련 버튼 및 입력 이벤트
- 공통 유틸: 셀 ID 파싱, 컬럼 변환, 포맷/스타일 보조 함수

### `vs_find.js` (VSFind)
찾기/바꾸기 기능을 독립 모듈로 제공하며, `VibrantSheets` 인스턴스를 컨텍스트로 받습니다.
- 검색 상태 갱신: 쿼리/옵션/매치 리스트 관리
- 매치 하이라이트: 그리드에 `match` 클래스 적용
- 탐색: 이전/다음 결과로 포커스 이동
- 치환: 단일/전체 치환

### `vs_io.js` (VSIO)
파일 입출력과 형식 변환을 담당하며, 컨텍스트(`VibrantSheets`)를 인자로 받는 유틸 모듈입니다.
- 파일 열기: File Picker/폴백 처리
- VSHT/CSV/XLSX 불러오기
- 저장/다른 이름 저장
- 파일 형식 변환: XLSX/CSV/VSHT 생성
- ExcelJS 연동: 셀 포맷 및 병합 정보 변환 포함

### `formula_engine.js`
수식 평가 로직을 독립적으로 유지합니다.
- 수식 파싱 및 계산
- 기본 함수 (`SUM`, `CONCAT`, `LEFT`, `RIGHT`, `MID`)
- 범위 참조/순환 참조 감지

### `bootstrap.js`
앱 초기화 전용 파일입니다.
- `new VibrantSheets()` 인스턴스 생성
- 초기화 실패 시 오류 로깅

## 스크립트 로딩 순서
`index.html`에서 다음 순서로 로드됩니다.
1. `engines/formula_engine.js`
2. `engines/vs_find.js`
3. `engines/vs_io.js`
4. `engines/vs_core.js`
5. `engines/bootstrap.js`

## 변경 시 참고
- UI 동작/렌더 문제가 생기면 `vs_core.js`부터 확인합니다.
- 검색/치환 문제는 `vs_find.js`에서 독립적으로 수정 가능합니다.
- 파일 입출력/포맷 이슈는 `vs_io.js`가 단일 책임을 가집니다.

## 문서 위치
- 진행 프로세스: `documents/PROCESS.md`
- 구조 설명: `documents/PROJECT_STRUCTURE.md`
- 기존 계획: `documents/plan.md`
