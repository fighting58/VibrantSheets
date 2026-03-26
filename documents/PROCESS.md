# Project Process

## 목적
이 문서는 **프로젝트 진행 순서**, **현재 구현 상태**, **누락 기능**을 한눈에 파악할 수 있도록 정리한 진행 프로세스 문서입니다.

## 현재 진행 순서 (로드맵 기준)
1. Phase 10: 셀 병합
2. Phase 11: 셀 테두리
3. Phase 12: 인쇄 설정
4. Phase 13: Persistence & Recovery
5. Phase 13.1: 자동 저장(localStorage)
6. Phase 13.2: 세션 복구
7. Phase 13.3: 최근 파일 목록
8. Phase 14: 이미지 삽입

## 구현 상태 요약

### 완료됨
- Grid & Interaction
- 스타일링(텍스트/배경/정렬/폰트)
- 행/열 삽입/삭제 및 데이터/스타일/포맷 재계산
- 선택 기능(헤더 선택/전체 선택/범위 확장)
- 데이터 포맷팅(General/Currency/Percentage/Date)
- 다중 시트
- 찾기/바꾸기
- 함수 엔진(기본 함수 + 참조/순환 참조)
- 파일 입출력(VSHT/XLSX/CSV)

### 진행 중
- **Phase 10: 셀 병합**
- **Phase 11: 셀 테두리**

### 미완료(예정)
- Phase 11: 셀 테두리
- Phase 12: 인쇄 설정
- Phase 13: Persistence & Recovery
- Phase 14: 이미지 삽입

## Phase 10: 셀 병합 상세

### 구현된 내용
- 병합 데이터 구조: 시트 단위 `mergedRanges`
- 병합 렌더링: 병합 영역을 하나의 셀로 표시하고 내부 셀을 숨김
- 선택/편집:
  - 병합 영역 클릭 시 좌상단 셀로 포커스 이동
  - 선택 범위가 병합 셀과 겹치면 자동 확장
- 저장/불러오기:
  - VSHT 저장/복구 시 `mergedRanges` 포함
  - XLSX import/export 시 병합 반영(ExcelJS)
- UI:
  - 리본에 `Merge` / `Unmerge` 버튼 추가
- 붙여넣기 정책:
  - 다중 셀 붙여넣기 시 병합 범위와 겹치면 자동으로 병합 해제 후 적용
  - 단일 셀 붙여넣기는 병합 유지(좌상단 셀에 값 적용)

### 현재 제한 사항 및 남은 작업
- 병합 범위와 복사/붙여넣기 충돌 정책 고도화 필요 없음(자동 해제로 결정)
- 병합 영역 내 개별 셀 편집/입력 방지: 핵심 동작 구현 완료, 추가 UX 검증 필요
- 병합 상태에 따른 테두리/인쇄 연동 고려 필요

## Phase 11: 셀 테두리 (미구현)

### 요구 사항 요약
- 테두리 타입: `all`, `outer`, `inner`, `left`, `right`, `top`, `bottom`, `none`
- 스타일: `solid`, `dashed`, `dotted`
- 색상: 기본 팔레트 + 사용자 선택
- 병합 셀: 외곽 기준 테두리 적용
- 저장/불러오기: VSHT 저장 + XLSX export 기본 매핑

### 구현 체크리스트
- 테두리 저장 구조(`cellBorders` 혹은 `cellStyles` 확장) 설계
- 렌더 레이어 추가(그리드 위에 테두리 라인)
- 선택 범위 적용 로직
- 병합 셀 외곽 처리
- XLSX export 매핑

### 현재 진행 내용
- 리본에서 테두리 타입/스타일/색상 선택 및 적용 UI 추가
- `cellBorders` 구조 추가 및 렌더 반영
- VSHT/XLSX 저장 및 불러오기 매핑 추가

## Phase 12: 인쇄 설정 (미구현)

### 요구 사항 요약
- 인쇄 범위: 현재 시트 / 선택 영역 / 사용자 지정
- 페이지 설정: 방향, 용지, 여백, 배율
- 미리보기 모달
- `@media print` 스타일 분리
- VSHT 저장 (선택적으로 XLSX page setup)

### 구현 체크리스트
- 인쇄 설정 UI 및 상태 구조 설계
- 인쇄 범위 계산
- print 전용 스타일 적용
- VSHT 저장 구조 추가

## Phase 13: Persistence & Recovery (미구현)

### 요구 사항 요약
- localStorage 자동 저장
- 세션 복구(복구 여부 선택)
- 최근 파일 목록

### 구현 체크리스트
- 자동 저장 타이밍/스냅샷 정책
- 복구 충돌 우선순위 규칙
- 최근 파일 목록 포맷/저장 위치

## 모듈 구조 참고
자바스크립트는 `engines/` 하위에 모듈화되어 있습니다.
- `vs_core.js`: 핵심 UI/상태/렌더링
- `vs_find.js`: 찾기/바꾸기
- `vs_io.js`: 파일 입출력
- `formula_engine.js`: 수식 엔진
- `bootstrap.js`: 초기화

## Phase 14: 이미지 삽입 (미구현)

### 요구 사항 요약
- 이미지 업로드: 로컬 파일 선택(`<input type="file">`) 또는 URL 입력
- 이미지 배치: 셀 위에 플로팅 레이어로 표시 (셀에 종속되지 않는 독립 객체)
- 크기/위치 조정: 드래그로 이동, 핸들로 리사이즈
- 이미지 선택/삭제: 클릭으로 선택, Delete 키로 삭제
- 저장/불러오기:
  - `.vsht`: Base64로 인코딩하여 JSON 내 `sheetImages` 배열에 저장
  - `.xlsx`: ExcelJS의 `addImage` API를 사용하여 Excel 호환 이미지 삽입

### 구현 체크리스트
- [ ] 이미지 상태 구조 설계: 시트별 `sheetImages[]` (id, src, anchorCell, width, height, top, left)
- [ ] 이미지 렌더링 레이어: 그리드 위에 절대 위치 `<div>` 컨테이너로 오버레이
- [ ] 삽입 UI: 리본에 "이미지 삽입" 버튼 추가 (파일 선택 + URL 입력 모달)
- [ ] 드래그 이동 및 리사이즈 핸들 구현
- [ ] 이미지 선택 상태 및 Delete 키 삭제 처리
- [ ] VSHT 저장: Base64 직렬화 / 복원
- [ ] XLSX export: ExcelJS `workbook.addImage` + `worksheet.addImage` 매핑

### 모듈 배치 계획
- 이미지 상태 및 렌더링 로직: `vs_core.js` 내 `ImageLayer` 클래스로 분리
- VSHT/XLSX 직렬화: `vs_io.js`에 이미지 직렬화/역직렬화 메서드 추가

## 다음 단계 제안
1. 병합 기능의 붙여넣기/편집 충돌 정책 확정
2. 셀 테두리 데이터 구조 설계 및 렌더링 레이어 추가
3. 인쇄 설정 UI 및 저장 구조 설계
4. 이미지 삽입 상태 구조 및 렌더링 레이어 설계
