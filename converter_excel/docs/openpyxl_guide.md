# openpyxl 라이브러리 사용 가이드

openpyxl은 Python에서 Excel 2010+ xlsx/xlsm 파일을 읽고 쓰기 위한 라이브러리입니다.
Excel 설치 없이 서버 환경에서도 사용 가능합니다.

## 설치

```bash
pip install openpyxl
```

---

## 1. 셀 읽기/쓰기 기본 사용법

### 1.1 워크북 생성 및 저장

```python
from openpyxl import Workbook

# 새 워크북 생성
wb = Workbook()
ws = wb.active  # 기본 시트 선택

# 시트 이름 변경
ws.title = "데이터"

# 파일 저장
wb.save("example.xlsx")
```

### 1.2 기존 파일 열기

```python
from openpyxl import load_workbook

# 파일 열기
wb = load_workbook("example.xlsx")
ws = wb.active

# 특정 시트 선택
ws = wb["데이터"]

# 시트 목록 확인
print(wb.sheetnames)  # ['데이터', 'Sheet2', ...]
```

### 1.3 셀에 값 쓰기

```python
# 방법 1: A1 표기법
ws["A1"] = "이름"
ws["B1"] = "나이"
ws["C1"] = "점수"

# 방법 2: cell() 메서드 (1-based 인덱스)
ws.cell(row=2, column=1, value="홍길동")
ws.cell(row=2, column=2, value=25)
ws.cell(row=2, column=3, value=95.5)

# 방법 3: append()로 행 단위 추가
ws.append(["김철수", 30, 88.0])
ws.append(["이영희", 28, 92.3])
```

### 1.4 셀에서 값 읽기

```python
# 방법 1: A1 표기법
name = ws["A2"].value  # "홍길동"

# 방법 2: cell() 메서드
age = ws.cell(row=2, column=2).value  # 25

# 방법 3: 범위로 읽기
for row in ws.iter_rows(min_row=2, max_row=4, min_col=1, max_col=3):
    for cell in row:
        print(cell.value, end="\t")
    print()
```

### 1.5 전체 데이터 순회

```python
# 모든 행 순회
for row in ws.iter_rows(values_only=True):
    print(row)  # 튜플로 반환

# 특정 범위만 순회
for row in ws.iter_rows(min_row=1, max_row=10, min_col=1, max_col=5):
    for cell in row:
        print(f"{cell.coordinate}: {cell.value}")
```

---

## 2. 셀 서식 (폰트, 배경색, 테두리, 정렬)

### 2.1 필요한 import

```python
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
```

### 2.2 폰트 (Font)

```python
# 기본 폰트 설정
font = Font(
    name='맑은 고딕',      # 폰트 이름
    size=12,              # 크기 (포인트)
    bold=True,            # 굵게
    italic=False,         # 기울임
    underline='single',   # 밑줄: 'single', 'double', 'singleAccounting', 'doubleAccounting', None
    strike=False,         # 취소선
    color='FF0000'        # 색상 (RRGGBB 또는 AARRGGBB)
)

ws["A1"].font = font
```

**폰트 색상 예시:**
```python
# 빨강
ws["A1"].font = Font(color="FF0000")

# 파랑 (투명도 포함)
ws["A2"].font = Font(color="FF0000FF")  # AARRGGBB 형식
```

### 2.3 배경색 (PatternFill)

```python
# 단색 배경
fill = PatternFill(
    fill_type='solid',        # 채우기 타입
    start_color='FFFF00',     # 노랑
    end_color='FFFF00'
)

ws["A1"].fill = fill

# 그라데이션은 GradientFill 사용
from openpyxl.styles import GradientFill
gradient = GradientFill(stop=["FF0000", "0000FF"])
ws["B1"].fill = gradient
```

**fill_type 옵션:**
- `'solid'` - 단색 채우기 (가장 많이 사용)
- `'darkDown'`, `'darkGray'`, `'darkGrid'`, `'darkHorizontal'`
- `'darkTrellis'`, `'darkUp'`, `'darkVertical'`
- `'gray0625'`, `'gray125'`
- `'lightDown'`, `'lightGray'`, `'lightGrid'`, `'lightHorizontal'`
- `'lightTrellis'`, `'lightUp'`, `'lightVertical'`
- `'mediumGray'`

### 2.4 테두리 (Border)

```python
# Side 객체로 각 테두리 정의
thin_side = Side(style='thin', color='000000')
thick_side = Side(style='thick', color='0000FF')

border = Border(
    left=thin_side,
    right=thin_side,
    top=thick_side,
    bottom=thick_side
)

ws["A1"].border = border
```

**테두리 스타일 (style) 옵션:**
| 스타일 | 설명 |
|--------|------|
| `'thin'` | 얇은 실선 |
| `'medium'` | 중간 실선 |
| `'thick'` | 두꺼운 실선 |
| `'double'` | 이중선 |
| `'hair'` | 가는 선 |
| `'dotted'` | 점선 |
| `'dashed'` | 파선 |
| `'dashDot'` | 일점쇄선 |
| `'dashDotDot'` | 이점쇄선 |
| `'mediumDashed'` | 중간 파선 |
| `'mediumDashDot'` | 중간 일점쇄선 |
| `'mediumDashDotDot'` | 중간 이점쇄선 |
| `'slantDashDot'` | 기울어진 일점쇄선 |

### 2.5 정렬 (Alignment)

```python
alignment = Alignment(
    horizontal='center',    # 수평: 'left', 'center', 'right', 'justify', 'distributed'
    vertical='center',      # 수직: 'top', 'center', 'bottom', 'justify', 'distributed'
    wrap_text=True,         # 텍스트 줄바꿈
    shrink_to_fit=False,    # 셀에 맞춤
    text_rotation=0,        # 텍스트 회전 (0-180)
    indent=0                # 들여쓰기
)

ws["A1"].alignment = alignment
```

### 2.6 숫자 형식 (Number Format)

```python
# 숫자 형식 지정
ws["A1"] = 1234567.89
ws["A1"].number_format = '#,##0.00'  # 1,234,567.89

# 백분율
ws["A2"] = 0.756
ws["A2"].number_format = '0.0%'  # 75.6%

# 통화
ws["A3"] = 50000
ws["A3"].number_format = '₩#,##0'  # ₩50,000

# 날짜
from datetime import date
ws["A4"] = date(2024, 1, 15)
ws["A4"].number_format = 'YYYY-MM-DD'  # 2024-01-15
```

### 2.7 스타일 복합 적용 예제

```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

wb = Workbook()
ws = wb.active

# 헤더 스타일
header_font = Font(name='맑은 고딕', size=12, bold=True, color='FFFFFF')
header_fill = PatternFill(fill_type='solid', start_color='4472C4', end_color='4472C4')
header_alignment = Alignment(horizontal='center', vertical='center')
thin_border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)

# 헤더 작성
headers = ['이름', '부서', '급여', '입사일']
for col, header in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col, value=header)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = header_alignment
    cell.border = thin_border

# 데이터 작성
data = [
    ['홍길동', '개발팀', 5000000, '2020-03-15'],
    ['김철수', '기획팀', 4500000, '2021-07-01'],
]

for row_idx, row_data in enumerate(data, 2):
    for col_idx, value in enumerate(row_data, 1):
        cell = ws.cell(row=row_idx, column=col_idx, value=value)
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='center', vertical='center')

        # 급여 열은 숫자 형식 적용
        if col_idx == 3:
            cell.number_format = '#,##0'

wb.save('styled_example.xlsx')
```

---

## 3. 셀 병합/병합 해제

### 3.1 셀 병합

```python
# 방법 1: 문자열 범위
ws.merge_cells('A1:D1')

# 방법 2: 좌표로 지정
ws.merge_cells(
    start_row=2,
    start_column=1,
    end_row=3,
    end_column=4
)

# 병합된 셀에 값과 스타일 적용 (왼쪽 상단 셀에만)
ws['A1'] = '병합된 제목'
ws['A1'].font = Font(bold=True, size=14)
ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
```

### 3.2 셀 병합 해제

```python
# 방법 1: 문자열 범위
ws.unmerge_cells('A1:D1')

# 방법 2: 좌표로 지정
ws.unmerge_cells(
    start_row=2,
    start_column=1,
    end_row=3,
    end_column=4
)
```

### 3.3 병합된 셀 확인

```python
# 병합된 셀 범위 목록
print(ws.merged_cells.ranges)  # [<CellRange A1:D1>, <CellRange A2:A4>]

# 병합 영역 순회
for merged_range in ws.merged_cells.ranges:
    print(f"범위: {merged_range}")
    print(f"시작: ({merged_range.min_row}, {merged_range.min_col})")
    print(f"끝: ({merged_range.max_row}, {merged_range.max_col})")
```

### 3.4 병합 셀 전체 예제

```python
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

wb = Workbook()
ws = wb.active

# 제목 병합
ws.merge_cells('A1:E1')
ws['A1'] = '2024년 매출 현황'
ws['A1'].font = Font(bold=True, size=16)
ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
ws['A1'].fill = PatternFill(fill_type='solid', start_color='E2EFDA')

# 부제목 병합
ws.merge_cells('A2:E2')
ws['A2'] = '단위: 천원'
ws['A2'].alignment = Alignment(horizontal='right')

# 행 높이 조정
ws.row_dimensions[1].height = 30
ws.row_dimensions[2].height = 20

wb.save('merged_example.xlsx')
```

---

## 4. 행 높이, 열 너비 설정

### 4.1 행 높이 설정

```python
# 단위: 포인트 (pt)
ws.row_dimensions[1].height = 30  # 1행 높이 30pt
ws.row_dimensions[2].height = 20  # 2행 높이 20pt

# 여러 행 설정
for row in range(1, 11):
    ws.row_dimensions[row].height = 25

# 기본 행 높이 (새로 생성되는 행에 적용)
ws.sheet_format.defaultRowHeight = 15
```

### 4.2 열 너비 설정

```python
from openpyxl.utils import get_column_letter

# 단위: 문자 수 (기본 폰트 기준)
ws.column_dimensions['A'].width = 15
ws.column_dimensions['B'].width = 25
ws.column_dimensions['C'].width = 10

# 숫자 인덱스로 설정
col_letter = get_column_letter(4)  # 'D'
ws.column_dimensions[col_letter].width = 20

# 여러 열 설정
for col in range(1, 6):
    ws.column_dimensions[get_column_letter(col)].width = 12
```

### 4.3 행/열 숨기기

```python
# 행 숨기기
ws.row_dimensions[3].hidden = True

# 열 숨기기
ws.column_dimensions['C'].hidden = True
```

### 4.4 자동 너비 조정 (수동 계산)

openpyxl은 자동 너비 조정 기능이 없으므로 직접 계산해야 합니다.

```python
from openpyxl.utils import get_column_letter

def auto_adjust_column_width(ws, min_width=8, max_width=50):
    """셀 내용에 맞게 열 너비 자동 조정"""
    for col in ws.columns:
        max_length = 0
        column_letter = get_column_letter(col[0].column)

        for cell in col:
            if cell.value:
                # 한글은 2자로 계산
                cell_length = 0
                for char in str(cell.value):
                    if ord(char) > 127:  # 한글, 한자 등
                        cell_length += 2
                    else:
                        cell_length += 1
                max_length = max(max_length, cell_length)

        adjusted_width = min(max(max_length + 2, min_width), max_width)
        ws.column_dimensions[column_letter].width = adjusted_width

# 사용
auto_adjust_column_width(ws)
```

### 4.5 단위 변환 (HWPUNIT -> Excel)

```python
# HWPUNIT 상수
HWPUNIT_PER_PT = 100       # 1pt = 100 HWPUNIT
HWPUNIT_PER_INCH = 7200    # 1inch = 7200 HWPUNIT
HWPUNIT_PER_CM = 2834.6    # 1cm = 약 2834.6 HWPUNIT

# HWP 행 높이 -> Excel 행 높이 (pt)
hwp_row_height = 1500  # HWPUNIT
excel_row_height = hwp_row_height / HWPUNIT_PER_PT  # 15pt

# HWP 열 너비 -> Excel 열 너비 (문자 수)
hwp_col_width = 7000  # HWPUNIT
excel_col_width = hwp_col_width / 700  # 약 10 문자
```

---

## 5. 셀 값 타입 (문자열, 숫자, 날짜)

### 5.1 자동 타입 추론

openpyxl은 Python 타입에 따라 자동으로 Excel 타입을 결정합니다.

```python
# 문자열
ws["A1"] = "텍스트"
ws["A2"] = "123"  # 문자열 "123"

# 숫자 (정수, 실수)
ws["B1"] = 42
ws["B2"] = 3.14159

# 불리언
ws["C1"] = True   # Excel에서 TRUE로 표시
ws["C2"] = False  # Excel에서 FALSE로 표시

# None
ws["D1"] = None  # 빈 셀
```

### 5.2 날짜/시간

```python
from datetime import datetime, date, time, timedelta

# 날짜
ws["A1"] = date(2024, 1, 15)
ws["A1"].number_format = 'YYYY-MM-DD'

# 시간
ws["A2"] = time(14, 30, 0)
ws["A2"].number_format = 'HH:MM:SS'

# 날짜+시간
ws["A3"] = datetime(2024, 1, 15, 14, 30, 0)
ws["A3"].number_format = 'YYYY-MM-DD HH:MM:SS'

# 시간 간격 (timedelta)
ws["A4"] = timedelta(hours=2, minutes=30)
ws["A4"].number_format = '[h]:mm:ss'  # 2:30:00
```

### 5.3 날짜 시스템 설정

```python
import openpyxl
from openpyxl.utils.datetime import CALENDAR_WINDOWS_1900, CALENDAR_MAC_1904

wb = openpyxl.Workbook()

# 날짜 시스템 확인 (기본: 1900)
print(wb.epoch)  # CALENDAR_WINDOWS_1900

# 1904 시스템으로 변경 (Mac 호환)
wb.epoch = CALENDAR_MAC_1904

# ISO 8601 형식으로 저장
wb.iso_dates = True
```

### 5.4 셀 타입 확인

```python
cell = ws["A1"]

# 값 타입 확인
print(type(cell.value))  # <class 'str'>, <class 'int'>, etc.

# 날짜 여부 확인
print(cell.is_date)  # True/False

# 내부 데이터 타입 코드
print(cell.data_type)
# 's' = 문자열
# 'n' = 숫자
# 'b' = 불리언
# 'f' = 수식
# 'd' = 날짜 (ISO 모드)
```

### 5.5 수식 입력

```python
# 수식 입력
ws["A1"] = 10
ws["A2"] = 20
ws["A3"] = "=A1+A2"  # 수식
ws["A4"] = "=SUM(A1:A2)"

# 수식이 있는 파일 열기
wb = load_workbook("example.xlsx")  # 수식 유지
wb_values = load_workbook("example.xlsx", data_only=True)  # 계산된 값만
```

### 5.6 타입별 전체 예제

```python
from openpyxl import Workbook
from datetime import datetime, date, time, timedelta

wb = Workbook()
ws = wb.active

# 헤더
ws.append(['타입', '값', '설명'])

# 문자열
ws.append(['문자열', '안녕하세요', '한글 텍스트'])
ws.append(['문자열', 'Hello', '영문 텍스트'])

# 숫자
ws.append(['정수', 12345, '정수형'])
ws.append(['실수', 3.14159, '부동소수점'])

# 날짜/시간
ws.append(['날짜', date(2024, 1, 15), '날짜만'])
ws['B6'].number_format = 'YYYY-MM-DD'

ws.append(['시간', time(14, 30, 0), '시간만'])
ws['B7'].number_format = 'HH:MM:SS'

ws.append(['날짜시간', datetime(2024, 1, 15, 14, 30), '날짜+시간'])
ws['B8'].number_format = 'YYYY-MM-DD HH:MM'

# 불리언
ws.append(['불리언', True, 'TRUE로 표시'])
ws.append(['불리언', False, 'FALSE로 표시'])

# 수식
ws.append(['수식', '=1+1', '계산 결과 2'])

wb.save('data_types_example.xlsx')
```

---

## 참고 자료

- [openpyxl 공식 문서](https://openpyxl.readthedocs.io/en/stable/)
- [Simple Usage Guide](https://openpyxl.readthedocs.io/en/stable/usage.html)
- [Working with Styles](https://openpyxl.readthedocs.io/en/3.1/styles.html)
- [Dates and Times](https://openpyxl.readthedocs.io/en/stable/datetime.html)
- [Cell Module API](https://openpyxl.readthedocs.io/en/3.1/api/openpyxl.cell.cell.html)
- [Borders Module API](https://openpyxl.readthedocs.io/en/stable/api/openpyxl.styles.borders.html)
