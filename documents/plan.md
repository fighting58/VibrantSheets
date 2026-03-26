# VibrantSheets 개발 기획서

## 1. 프로젝트 개요
VibrantSheets는 브라우저에서 동작하는 Excel 스타일 스프레드시트 앱입니다.  
목표는 빠른 편집 성능, 직관적인 상호작용, 실사용 가능한 파일 호환성입니다.

## 2. 현재 구현 범위 (2026-03-25 기준)

### 2.1 편집/선택 엔진
- Ready / Edit / Enter 모드
- 키보드 내비게이션(화살표, Enter, Tab, F2, Esc)
- 범위 선택, Fill Handle, 고급 클립보드 붙여넣기
- Fill Handle 고도화: 시리즈/복사, Alt 복사, 빈칸 스킵, 사용자 리스트

### 2.2 헤더 기반 선택
- 열/행 헤더 클릭: 전체 선택
- Shift+헤더 클릭: 연속 다중 선택
- 헤더 드래그: 연속 범위 확장
- 코너 헤더 클릭: 전체 선택

### 2.3 스타일링
- Bold / Italic / Underline / Strikethrough
- 텍스트색, 배경색
- 정렬, 폰트, 폰트 크기

### 2.4 데이터 포맷팅
- General, Number, Text, Currency(KRW), Percentage, Date 지원
- 소수점 자리수 조절 (Number, Currency, Percentage 타입 대응)
- raw 값 보존 + 표시값 렌더링 분리 (Text는 원형 보존)
- XLSX import/export 시 포맷 타입 및 서식 코드(numFmt) 완벽 매핑

### 2.5 구조 작업
- 행/열 삽입/삭제
- 데이터/스타일/포맷 키 시프트 처리

### 2.6 다중 시트
- 시트 추가/삭제/이름 변경/이동
- 시트별 `data/cellStyles/cellFormats` 분리

### 2.7 찾기/바꾸기
- 전체 시트 검색/치환
- 옵션: 대소문자 구분, 정확히 일치
- 결과 하이라이트 및 순차 이동

### 2.8 함수 엔진(간단 버전)
- `SUM`, `CONCAT`, `LEFT`, `RIGHT`, `MID`
- 셀 참조/범위(`A1:A5`) 지원
- 순환 참조 감지(`#CYCLE`)
- 수식은 `cellFormulas`, 값은 `data`에 분리 저장
- 에러 전파 규칙: 에러 포함 연산은 에러 반환
- 절대/상대 참조 지원(`$A$1`, `A$1`, `$A1`)

### 2.9 파일 입출력
- `.vsht`: 데이터/스타일/포맷/레이아웃 보존
- `.xlsx`: ExcelJS 기반 import/export + 스타일 보존
- `.csv`: 활성 시트만 저장 (확인 모달)

## 3. 아키텍처 요약
- 단일 엔트리: `app.js` (`VibrantSheets` 클래스)
- 함수 엔진: `formula_engine.js`
- UI: `index.html` (리본/수식바/그리드/상태바)
- 스타일: `style.css`

## 4. 다음 단계

### 2.10 테두리 엔진 및 안정성
- 9종 아이콘 기반 테두리 배치 제어 (내부 가로/세로 포함)
- 인접 셀 동기화(Mirror Rendering) 시스템
- 행/열 다중 삭제 시 데이터 잔류 버그 수정 (shiftCoord 필터링)
- IME(한글) 첫 글자 소실 해결 (Select All 전략)

## 3. 아키텍처 요약
- 단일 엔트리: `app.js` (`VibrantSheets` 클래스)
- 함수 엔진: `formula_engine.js`
- 스타일: `style.css` (컴팩트 리본 디자인 반영)

## 4. 다음 단계

### Phase 12: Persistence & Recovery
- `localStorage` 기반 자동 데이터 저장
- 세션 복구 및 최근 파일 목록 유지
- 사용자 정의 리스트 정보 유지
- 멀티 시트 저장/불러오기(XLSX) 시 스타일 유지
- CSV 저장 시 활성 시트만 저장 + 확인 모달
- 수식 범위/순환 참조 동작
- 수식 엔진 단위 테스트
- 수식 엔진 통합 테스트

## 6. 함수 동작 규격(간단 버전)
- 입력 타입: 숫자 문자열은 자동 변환, 비숫자는 문자열로 유지
- 빈값: `""`, `null`, `undefined`는 빈 문자열 처리
- Boolean: `TRUE/FALSE`, 숫자(0/1)를 논리값으로 해석
- 에러 전파: 인자에 에러 포함 시 해당 에러 반환
- 비교 연산: 숫자로 변환 가능하면 숫자 비교, 아니면 문자열 비교
- 문자열 함수: `LEFT/RIGHT/MID`는 범위 밖이면 빈 문자열, `CONCAT`은 문자열 결합
- 숫자 함수: `SUM/AVG/COUNT/MIN/MAX`는 숫자 변환 가능한 값만 계산
