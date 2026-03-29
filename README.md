# 📊 VibrantSheets

**VibrantSheets**는 웹 브라우저에서 동작하는 프리미엄 Excel 스타일의 스프레드시트 애플리케이션입니다.  
초경량 Vanilla JS 기반으로 구축되었으며, 강력한 성능과 직관적인 UI, 그리고 실무에서 활용 가능한 뛰어난 파일 호환성을 목표로 개발되었습니다.

---

## ✨ 주요 특징 (Key Features)

### 🚀 고성능 그리드 엔진
- **부드러운 상호작용**: 대규모 데이터셋에서도 끊김 없는 스크롤과 셀 선택.
- **편집 모드**: Ready, Edit, Enter 모드를 지원하여 엑셀과 동일한 편집 경험을 제공합니다.
- **스마트 채우기 (Fill Handle)**: 데이터 패턴을 인식하여 자동으로 범위를 확장하거나 복사합니다.

### 🎨 프리미엄 디자인 및 스타일링
- **다크 모드 최적화**: 시각적 피로를 줄여주는 세련된 다크 모드 인터페이스.
- **풍부한 서식**: 폰트, 색상, 정렬, 테두리(9종 스타일) 및 병합 기능을 완벽하게 지원합니다.
- **데이터 포맷팅**: 일반, 숫자, 통화(KRW/USD), 백분율, 날짜, 텍스트 등 다양한 데이터 형식을 지원합니다.

### 📁 뛰어난 파일 호환성
- **XLSX 라운드트립**: `ExcelJS`를 사용하여 엑셀 파일(.xlsx)의 스타일, 병합, 이미지를 보존하며 불러오고 저장합니다.
- **CSV & VSHT**: 표준 CSV 파일 및 VibrantSheets 전용 경량 포맷(.vsht)을 지원합니다.
- **이미지 관리**: 셀 위에 이미지를 삽입, 이동, 리사이즈하고 파일과 함께 저장할 수 있습니다. (Phase 14)

### 🧮 강력한 함수 엔진
- **수식 지원**: `SUM`, `CONCAT`, `LEFT`, `RIGHT` 등 주요 함수 지원.
- **셀 참조**: 상대 참조(A1) 및 절대 참조($A$1)를 지원하며 순환 참조 감지 기능을 포함합니다.

### 🖨️ 전문적인 인쇄 기능
- **인쇄 영역 설정**: 시트의 특정 부분만 인쇄할 수 있는 영역 지정 기능.
- **실시간 프리뷰**: 인쇄 모달에서 레이아웃, 여백, 머리글/바닥글 설정을 실시간으로 확인합니다.

---

## 🛠️ 기술 스택 (Tech Stack)

| Category | Technology |
| :--- | :--- |
| **Core** | Vanilla JavaScript (ES6+) |
| **Styling** | Vanilla CSS (Modern CSS Variables & Grid) |
| **Libraries** | [ExcelJS](https://github.com/exceljs/exceljs) (XLSX I/O), [JSZip](https://github.com/Stuk/jszip) (Pre-processing) |
| **API** | File System Access API (Native File Dialogs) |

---

## 🚀 시작하기 (Getting Started)

VibrantSheets는 별도의 빌드 과정 없이 즉시 실행 가능합니다.

1. 이 저장소를 클론하거나 다운로드합니다.
2. 최신 브라우저(Chrome, Edge 권장)에서 `index.html` 파일을 엽니다.
3. 상단 리본 메뉴를 통해 파일을 열거나 새로운 시트를 작성해 보세요!

---

## 📂 프로젝트 구조 (Project Structure)

```text
VibrantSheets/
├── documents/          # 프로젝트 계획 및 상세 문서
├── engines/            # 핵심 로직 엔진 (Core, Formula, IO, Find)
├── index.html          # 메인 애플리케이션 구조
├── style.css           # 종합 스타일 매니페스트
└── README.md           # 프로젝트 가이드 (현재 파일)
```

상세한 아키텍처는 [PROJECT_STRUCTURE.md](documents/PROJECT_STRUCTURE.md)를 참고하세요.

---

## 📝 관련 문서

- [프로젝트 통합 계획서](documents/plan.md)
- [아키텍처 가이드](documents/PROJECT_STRUCTURE.md)
- [회귀 테스트 체크리스트](documents/REGRESSION_CHECKLIST.md)
- [트러블슈팅 가이드](documents/TROUBLESHOOTING.md)

---

## ⚖️ 라이선스

이 프로젝트는 학습 및 포트폴리오 목적으로 제작되었습니다.
