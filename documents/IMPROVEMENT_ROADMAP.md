# VibrantSheets 고도화 로드맵 (Improvement Roadmap)

VibrantSheets는 기본 스프레드시트 기능을 갖추었으며, “프리미엄” 사용자 경험과 대규모 데이터 처리를 위해 다음 개선 단계를 제안합니다.

---

## 1. 성능 및 확장성 (Rendering & Scalability)

### 가상 스크롤링 (Virtual Scrolling)
- 현 상태: 모든 셀을 DOM에 렌더링
- 개선: 화면에 보이는 셀만 렌더링하는 가상화 도입
- 기대효과: 대규모 데이터에서도 부드러운 스크롤

### 이벤트 위임 (Event Delegation)
- 현 상태: 각 셀에 개별 리스너 연결
- 개선: 그리드 컨테이너에 단일 리스너 적용
- 기대효과: 메모리 절감, 초기 렌더링 속도 개선

---

## 2. 아키텍처 및 유지보수성 (Architecture & Maintainability)

### ES Modules 전환
- 현 상태: 전역 스코프 기반 스크립트 로드
- 개선: `import/export` 구조로 전환, 필요 시 번들러 도입
- 기대효과: 의존성 명확화, 코드 스플리팅 가능

### 상태 관리 고도화
- 현 상태: 데이터와 UI 상태 혼재
- 개선: Store 패턴 도입, 렌더링 트리거 구조 정리
- 기대효과: Undo/Redo 구현 용이, 버그 추적 개선

---

## 3. 사용자 경험 및 기능 고도화 (UX & Features)

### 컨텍스트 메뉴
- 우클릭 기반 복사/붙여넣기, 행/열 삽입/삭제, 메모 추가

### 수식 입력기 고도화
- 자동 완성, 범위 선택, 구문 강조

### Undo/Redo 시스템
- Command 패턴 기반 편집 이력 관리

---

## 4. 안정성 및 지속성 (Persistence & Reliability)

### 자동 저장 및 복구
- `localStorage` 또는 `IndexedDB` 기반 자동 저장
- 브라우저 재시작 후 복구
- 최근 파일 목록 유지

---

## 5. 단계별 실행 계획

1. Short-term: 자동 저장 및 세션 복구 구현
2. Mid-term: ES Modules 전환 및 구조 모듈화
3. Long-term: 가상 스크롤링 및 Undo/Redo 고도화
