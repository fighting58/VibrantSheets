# VibrantSheets

VibrantSheets는 Vanilla JavaScript/CSS로 만든 웹 기반 스프레드시트 앱입니다.  
빠른 그리드 상호작용, 리본 UI, 파일 입출력, 데이터 포맷팅을 중심으로 Excel 스타일 경험을 제공합니다.

## 주요 기능

### 1) Grid & Interaction
- Ready / Edit / Enter 모드 편집
- 무한 스크롤 기반 동적 행 추가
- 범위 선택(드래그, Shift 확장), Fill Handle
- 고급 TSV 파서 기반 복사/붙여넣기
- `Esc`로 편집 취소, `F2`로 셀 편집 진입
 - Fill Handle 고도화: 시리즈/복사, Alt 복사, 빈칸 스킵, 사용자 리스트

### 2) 스타일링
- 굵게/기울임/밑줄/취소선
- 텍스트 색 / 배경 색
- 정렬(좌/중/우), 폰트/크기
- IME(한글 입력) 친화 편집 흐름

### 3) 행/열 작업
- 행/열 삽입 및 삭제
- 삭제 확인 모달
- 데이터/스타일/포맷 키 재계산으로 일관성 유지

### 4) 선택 기능
- 열/행 헤더 클릭: 해당 열/행 전체 선택
- `Shift + 헤더 클릭`: 연속 다중 선택
- 헤더 드래그: 연속 범위 확장
- 코너 헤더 클릭: 전체 선택

### 5) 데이터 포맷팅 (Phase 7)
- General / Currency(KRW) / Percentage / Date
- 소수점 자리수 증감(+/-) 및 직접 입력
- `raw value` 저장 + `display value` 렌더링 분리
- `.vsht` 저장/불러오기 시 포맷 메타데이터 유지

### 6) 다중 시트
- 시트 추가/삭제/이름 변경/이동(드래그)
- 활성 시트 전환 및 독립 데이터/스타일 유지

### 7) 찾기/바꾸기
- 전체 시트 검색/치환
- 옵션: 대소문자 구분, 정확히 일치
- 결과 하이라이트 및 순차 이동

### 8) 함수 엔진(간단 버전)
- `SUM`, `CONCAT`, `LEFT`, `RIGHT`, `MID`
- 셀 참조/범위(`A1:A5`) 지원
- 순환 참조 감지(`#CYCLE`)
- 수식은 `cellFormulas`, 값은 `data`에 분리 저장
- 에러 전파 규칙: 에러 포함 연산은 에러 반환
 - 절대/상대 참조 지원(`$A$1`, `A$1`, `$A1`)

### 9) 파일 입출력
- Save / Save As (`File System Access API`)
- `.vsht`, `.xlsx`, `.csv` 지원
- CSV는 **활성 시트만** 저장 (확인 모달)

## 테스트
- `test_formula_engine.js`: 수식 엔진 단위 테스트
- `test_formula_integration.js`: 셀 참조/순환 참조 통합 테스트
- `test_ui_smoke.html`: 브라우저 UI 스모크 테스트

## 로드맵

- [x] Find/Replace
- [x] Multi-Sheet
- [x] Simple Formula Engine
- [x] Phase 8: Formula Engine Core (확장)
- [x] Phase 9: Built-in Functions
- [ ] Phase 10: Persistence & Recovery
  - [ ] 자동 저장(localStorage)
  - [ ] 세션 복구
  - [ ] 최근 파일 목록

## 함수 동작 규격(간단 버전)
- 숫자 문자열은 자동 변환, 비숫자는 문자열 유지
- 빈값(`""`, `null`, `undefined`)은 빈 문자열 처리
- `TRUE/FALSE`, 숫자(0/1)는 논리값으로 해석
- 에러 전파: 인자에 에러 포함 시 해당 에러 반환
- 비교 연산은 숫자 우선, 불가 시 문자열 비교
- 문자열 함수: `LEFT/RIGHT/MID` 범위 밖은 빈 문자열, `CONCAT`은 문자열 결합
- 숫자 함수: `SUM/AVG/COUNT/MIN/MAX`는 숫자 변환 가능한 값만 계산

## 기술 스택
- Language: Vanilla JavaScript (ES6+)
- Styling: Vanilla CSS
- Library: ExcelJS (XLSX 입출력)
- API: File System Access API

## 실행 방법
`index.html`을 최신 브라우저에서 열면 실행됩니다.  
파일 저장/불러오기 경험은 Chromium 계열(Chrome, Edge)에서 가장 안정적입니다.

---
Developed by [fighting58](https://github.com/fighting58)
