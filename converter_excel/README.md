# HWP 표 -> Excel 변환기 (converter_excel)

한글(HWP) 문서의 표를 Excel 파일로 변환하는 모듈입니다.

## 개요

이 모듈은 한글 문서에서 표 데이터와 스타일을 추출하여 Excel 파일로 변환합니다.
페이지 설정, 셀 스타일, 병합 셀, 행/열 크기 등을 최대한 보존합니다.

## 주요 기능

- **페이지 정보 추출**: 용지 크기, 여백, 방향 (세로/가로)
- **셀 스타일 추출**: 배경색, 글꼴(이름/크기/굵기/기울임/색상), 정렬, 테두리
- **셀 텍스트 추출**: SelectAll 방식으로 셀 내용 가져오기
- **셀 보호**: 배경색 있는 셀은 잠금, 없는 셀은 편집 가능
- **병합 셀 처리**: rowspan/colspan 유지
- **행 높이/열 너비 변환**: HWPUNIT -> Excel 단위 (pt, 문자)

## 파일 구조

```
converter_excel/
├── __init__.py              # 패키지 초기화 및 공개 API
├── page_meta.py             # 단위 변환 상수 및 페이지/표 메타 데이터 클래스
├── page_setup.py            # 엑셀 페이지 설정 추출
├── cell_style.py            # 셀 스타일(배경색, 테두리) 추출
├── hwp_table_meta.py        # 한글 표 메타 정보 추출 (행/열 크기)
├── excel_export_data.py     # 엑셀 변환용 종합 데이터 구조
├── match_page.py            # 페이지 정보 추출 및 엑셀 시트 저장
├── match_cell.py            # 셀 스타일/텍스트 추출 및 엑셀 시트 저장
├── test_export.py           # 통합 테스트/실행 스크립트
└── README.md                # 이 문서
```

### 주요 모듈 설명

| 파일 | 설명 |
|------|------|
| `page_meta.py` | `Unit` 클래스(단위 변환), `HwpPageMeta`, `TableMeta` 데이터 클래스 |
| `match_page.py` | `extract_page_info()` - 페이지 정보 추출, `apply_page_margins_to_excel()` - 여백 적용 |
| `match_cell.py` | `extract_cell_style()` - 셀 스타일 추출, `get_cell_text()` - 텍스트 추출 |
| `cell_style.py` | `get_cell_bg_color()`, `set_cell_bg_color()` - 배경색 조회/설정 |
| `test_export.py` | 통합 실행 스크립트 (페이지 + 셀 -> 엑셀 파일) |

## 사용법

### WSL에서 실행

WSL 환경에서는 `win32com`이 작동하지 않으므로 Windows Python을 호출해야 합니다.

```bash
# 한글 프로그램에서 표 안에 커서를 두고 실행
cmd.exe /c "cd /d C:\\win32hwp && python converter_excel\\test_export.py"
```

### Python에서 직접 사용

```python
from cursor import get_hwp_instance
from converter_excel import (
    extract_page_info_v2,
    extract_cell_style_v2,
    get_cell_text,
    apply_page_margins_to_excel,
)

hwp = get_hwp_instance()

# 페이지 정보 추출
page_result = extract_page_info_v2(hwp)
if page_result.success:
    print(f"용지 크기: {page_result.page_meta.page_size.width} x {page_result.page_meta.page_size.height}")

# 셀 스타일 추출
style = extract_cell_style_v2(hwp, list_id)
text = get_cell_text(hwp, list_id)
```

## 단위 변환 (HWPUNIT <-> Excel)

### HWPUNIT 기본 변환

| 단위 | HWPUNIT 값 | 계산식 |
|------|-----------|--------|
| 1 pt (포인트) | 100 | `hwpunit / 100 = pt` |
| 1 inch (인치) | 7,200 | `hwpunit / 7200 = inch` |
| 1 cm | 2,834.6 | `hwpunit / 2834.6 = cm` |
| 1 mm | 283.46 | `hwpunit / 283.46 = mm` |

### Excel 단위

| 항목 | 단위 | 변환 공식 |
|------|------|-----------|
| 행 높이 | 포인트(pt) | `hwpunit / 100` |
| 열 너비 | 문자 수 | `hwpunit / 700` (1문자 약 7pt) |
| 여백 | 인치(inch) | `hwpunit / 7200` |

### 변환 함수 (page_meta.py)

```python
from converter_excel import Unit

# HWPUNIT -> 다른 단위
pt = Unit.hwpunit_to_pt(hwpunit)      # 포인트
cm = Unit.hwpunit_to_cm(hwpunit)      # cm
mm = Unit.hwpunit_to_mm(hwpunit)      # mm

# 다른 단위 -> HWPUNIT
hwpunit = Unit.pt_to_hwpunit(pt)
hwpunit = Unit.cm_to_hwpunit(cm)
hwpunit = Unit.mm_to_hwpunit(mm)

# 엑셀 전용
chars = Unit.excel_char_to_hwpunit(width)  # 열 너비 (문자 -> HWPUNIT)
```

## 출력 Excel 시트 구성

`test_export.py` 실행 시 생성되는 Excel 파일의 시트 구성:

| 시트 이름 | 내용 |
|-----------|------|
| **표** | 한글 표 데이터 (텍스트, 배경색, 병합, 셀 보호 적용) |
| **_page** | 페이지 설정 (용지 크기, 여백 - HWPUNIT/cm/inch/pt) |
| **_cells** | 셀 스타일 상세 정보 (위치, 크기, 배경색, 글꼴, 정렬, 테두리) |
| **_sizes** | 행 높이 / 열 너비 목록 (HWPUNIT/pt/cm/문자) |

### 셀 보호 규칙

- **배경색 있는 셀**: 잠금 (편집 불가)
- **배경색 없는 셀**: 편집 가능
- 시트 보호 활성화 시 서식 변경은 허용됨

## 알려진 제한사항

1. **세로 정렬 추출 불가**: `CellShape.VertAlign`이 항상 0을 반환하는 HWP API 문제로 인해 세로 정렬(상/중/하)을 추출할 수 없습니다. 기본값 `center`가 사용됩니다.

2. **글꼴 장식 일부 미지원**: 외곽선(outline), 그림자(shadow), 양각/음각 등은 Excel에서 지원하지 않아 변환되지 않습니다.

3. **테두리 스타일 단순화**: 한글의 다양한 테두리 스타일(점선, 이중선 등)이 Excel에서 `thin` 스타일로 단순화됩니다.

4. **이미지/개체 미지원**: 셀 내 이미지나 개체는 변환되지 않습니다.

5. **복잡한 병합 셀**: 불규칙한 병합 패턴은 완벽하게 재현되지 않을 수 있습니다.

## 의존성

- Python 3.x
- `pywin32` (win32com) - Windows에서만 작동
- `openpyxl` - Excel 파일 생성

```bash
pip install pywin32 openpyxl
```

## 관련 모듈

- `table/table_info.py` - 표 탐색 및 셀 이동
- `table/cell_position.py` - 셀 위치 계산 (xline/yline 기반 그리드 매핑)
- `cursor.py` - HWP 인스턴스 연결 (ROT)
