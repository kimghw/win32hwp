# Word 테이블을 Excel로 변환하는 패턴

Word 문서의 테이블을 Excel로 변환할 때 필요한 기술적 패턴을 정리합니다.

---

## 목차

1. [python-docx로 테이블 읽기](#1-python-docx로-테이블-읽기)
2. [win32com으로 테이블 읽기](#2-win32com으로-테이블-읽기)
3. [셀 병합 정보 확인](#3-셀-병합-정보-확인)
4. [서식 변환 (Word → Excel)](#4-서식-변환-word--excel)
5. [단위 변환](#5-단위-변환)
6. [참고 자료](#6-참고-자료)

---

## 1. python-docx로 테이블 읽기

### 1.1 기본 테이블 접근

```python
from docx import Document

# 문서 열기
doc = Document('document.docx')

# 모든 테이블 접근
for table in doc.tables:
    print(f"행 수: {len(table.rows)}, 열 수: {len(table.columns)}")
```

### 1.2 셀 접근 방법

```python
# 방법 1: 인덱스로 직접 접근 (0-based)
cell = table.cell(0, 0)  # 첫 번째 행, 첫 번째 열
text = cell.text

# 방법 2: 행과 셀 순회
for row in table.rows:
    for cell in row.cells:
        print(cell.text)

# 방법 3: 열 단위 접근
for column in table.columns:
    for cell in column.cells:
        print(cell.text)
```

### 1.3 테이블을 2D 리스트로 변환

```python
def table_to_matrix(table):
    """테이블 내용을 2D 리스트로 변환"""
    return [[cell.text for cell in row.cells] for row in table.rows]

# 사용 예
matrix = table_to_matrix(table)
```

### 1.4 Pandas DataFrame으로 변환

```python
import pandas as pd

def table_to_dataframe(table, header=True):
    """테이블을 DataFrame으로 변환"""
    data = [[cell.text for cell in row.cells] for row in table.rows]

    if header and len(data) > 0:
        return pd.DataFrame(data[1:], columns=data[0])
    return pd.DataFrame(data)
```

### 1.5 중첩 테이블 처리

```python
def get_nested_tables(cell):
    """셀 내부의 중첩 테이블 조회"""
    return cell.tables

# 또는 iter_inner_content 사용
for item in cell.iter_inner_content():
    if hasattr(item, 'rows'):  # 테이블인 경우
        print("중첩 테이블 발견")
```

### 1.6 셀 내 텍스트 상세 접근 (단락/Run)

```python
def get_cell_runs(cell):
    """셀 내 모든 Run 객체 반환"""
    runs = []
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            runs.append({
                'text': run.text,
                'bold': run.bold,
                'italic': run.italic,
                'font_size': run.font.size,
                'font_name': run.font.name,
                'color': run.font.color.rgb if run.font.color.rgb else None
            })
    return runs
```

---

## 2. win32com으로 테이블 읽기

### 2.1 기본 설정

```python
import win32com.client

# Word 애플리케이션 시작
word = win32com.client.Dispatch("Word.Application")
word.Visible = False  # 백그라운드 실행

# 문서 열기
doc = word.Documents.Open(r"C:\path\to\document.docx")
```

### 2.2 테이블 접근

```python
# 모든 테이블 조회 (1-based 인덱싱)
table_count = doc.Tables.Count
print(f"총 테이블 수: {table_count}")

# 첫 번째 테이블 접근
table = doc.Tables(1)
print(f"행 수: {table.Rows.Count}, 열 수: {table.Columns.Count}")
```

### 2.3 셀 접근 및 텍스트 읽기

```python
# 특정 셀 접근 (1-based 인덱싱)
cell = table.Cell(1, 1)  # 첫 번째 행, 첫 번째 열

# 셀 텍스트 읽기 (주의: 끝에 제어 문자 포함)
text = cell.Range.Text
clean_text = text.rstrip('\r\x07')  # 제어 문자 제거
```

### 2.4 전체 테이블 순회

```python
def read_table_win32(table):
    """win32com으로 테이블 전체 읽기"""
    data = []

    for row_idx in range(1, table.Rows.Count + 1):
        row_data = []
        for col_idx in range(1, table.Columns.Count + 1):
            try:
                cell = table.Cell(row_idx, col_idx)
                text = cell.Range.Text.rstrip('\r\x07')
                row_data.append(text)
            except Exception as e:
                # 병합된 셀로 인한 오류 처리
                row_data.append(None)
        data.append(row_data)

    return data
```

### 2.5 셀 서식 읽기

```python
def get_cell_format_win32(cell):
    """셀의 서식 정보 반환"""
    rng = cell.Range
    font = rng.Font

    return {
        'text': rng.Text.rstrip('\r\x07'),
        'font_name': font.Name,
        'font_size': font.Size,
        'bold': font.Bold,
        'italic': font.Italic,
        'color': font.Color,  # RGB 정수값
        'underline': font.Underline,
    }
```

### 2.6 문서 종료

```python
# 문서 닫기 (저장하지 않음)
doc.Close(SaveChanges=False)

# Word 종료
word.Quit()
```

---

## 3. 셀 병합 정보 확인

### 3.1 python-docx에서 병합 셀 감지

python-docx는 병합 셀에 대한 직접적인 API가 제한적입니다. 내부 XML을 통해 확인해야 합니다.

```python
from docx.oxml.ns import qn

def get_cell_span_info(cell):
    """셀의 병합 정보 반환"""
    tc = cell._tc
    tcPr = tc.tcPr

    # 수평 병합 (gridSpan)
    grid_span = 1
    if tcPr is not None:
        grid_span_elem = tcPr.find(qn('w:gridSpan'))
        if grid_span_elem is not None:
            grid_span = int(grid_span_elem.get(qn('w:val')))

    # 수직 병합 (vMerge)
    v_merge = None
    if tcPr is not None:
        v_merge_elem = tcPr.find(qn('w:vMerge'))
        if v_merge_elem is not None:
            val = v_merge_elem.get(qn('w:val'))
            v_merge = val if val else 'continue'  # 'restart' 또는 'continue'

    return {
        'grid_span': grid_span,  # 가로 병합 수
        'v_merge': v_merge,      # 'restart'=시작, 'continue'=계속, None=병합없음
    }
```

### 3.2 병합 셀 동일성 비교

```python
def are_cells_merged(cell1, cell2):
    """두 셀이 동일한 병합 셀인지 확인"""
    return cell1._tc == cell2._tc
```

### 3.3 테이블 그리드 매핑

```python
def build_grid_map(table):
    """테이블의 실제 그리드 맵 생성"""
    num_rows = len(table.rows)
    num_cols = len(table.columns)

    # 그리드 맵 초기화 (None = 빈 셀)
    grid = [[None for _ in range(num_cols)] for _ in range(num_rows)]

    for row_idx, row in enumerate(table.rows):
        col_idx = 0
        for cell in row.cells:
            # 이미 채워진 셀 건너뛰기 (수직 병합으로 인해)
            while col_idx < num_cols and grid[row_idx][col_idx] is not None:
                col_idx += 1

            if col_idx >= num_cols:
                break

            # 병합 정보 확인
            span_info = get_cell_span_info(cell)
            grid_span = span_info['grid_span']

            # 현재 셀 위치 기록
            grid[row_idx][col_idx] = {
                'cell': cell,
                'text': cell.text,
                'is_origin': True
            }

            # 수평 병합된 셀 표시
            for i in range(1, grid_span):
                if col_idx + i < num_cols:
                    grid[row_idx][col_idx + i] = {
                        'cell': cell,
                        'text': '',
                        'is_origin': False,
                        'merged_to': (row_idx, col_idx)
                    }

            col_idx += grid_span

    return grid
```

### 3.4 win32com에서 병합 셀 확인

```python
def get_merge_info_win32(cell):
    """win32com으로 셀 병합 정보 확인"""
    # MergeCells 속성으로 병합 여부 확인
    # 주의: 일부 버전에서 다르게 동작할 수 있음
    try:
        # 병합된 셀의 범위 확인
        # 이 방식은 Word 버전에 따라 다를 수 있음
        return {
            'row_index': cell.RowIndex,
            'column_index': cell.ColumnIndex,
            # 추가 병합 정보는 Range를 통해 확인
        }
    except Exception as e:
        return None
```

---

## 4. 서식 변환 (Word → Excel)

### 4.1 폰트 속성 변환

```python
from openpyxl.styles import Font
from docx.shared import Pt, RGBColor

def convert_font_to_excel(word_font):
    """Word 폰트를 Excel 폰트로 변환"""

    # 색상 변환
    color = '000000'  # 기본 검정
    if word_font.color and word_font.color.rgb:
        rgb = word_font.color.rgb
        color = f'{rgb.red:02X}{rgb.green:02X}{rgb.blue:02X}'

    # 폰트 크기 변환 (Pt → 포인트)
    size = 11  # 기본값
    if word_font.size:
        size = word_font.size.pt

    return Font(
        name=word_font.name if word_font.name else 'Calibri',
        size=size,
        bold=word_font.bold if word_font.bold else False,
        italic=word_font.italic if word_font.italic else False,
        underline='single' if word_font.underline else None,
        color=color
    )
```

### 4.2 배경색 변환

```python
from openpyxl.styles import PatternFill
from docx.oxml.ns import qn

def get_cell_background_color(cell):
    """python-docx 셀의 배경색 읽기"""
    tc = cell._tc
    tcPr = tc.tcPr

    if tcPr is None:
        return None

    shd = tcPr.find(qn('w:shd'))
    if shd is not None:
        fill = shd.get(qn('w:fill'))
        return fill  # 예: 'FFFF00' (노란색)

    return None

def convert_fill_to_excel(word_color):
    """Word 배경색을 Excel PatternFill로 변환"""
    if word_color is None or word_color == 'auto':
        return None

    return PatternFill(
        start_color=word_color,
        end_color=word_color,
        fill_type='solid'
    )
```

### 4.3 테두리 변환

```python
from openpyxl.styles import Border, Side
from docx.oxml.ns import qn

# Word 테두리 스타일 → Excel 스타일 매핑
BORDER_STYLE_MAP = {
    'single': 'thin',
    'thick': 'medium',
    'double': 'double',
    'dotted': 'dotted',
    'dashed': 'dashed',
    'nil': None,
    None: None,
}

def get_cell_borders_xml(cell):
    """python-docx 셀의 테두리 정보 읽기 (XML에서)"""
    tc = cell._tc
    tcPr = tc.tcPr

    borders = {
        'top': None,
        'bottom': None,
        'left': None,
        'right': None,
    }

    if tcPr is None:
        return borders

    tcBorders = tcPr.find(qn('w:tcBorders'))
    if tcBorders is None:
        return borders

    for edge in ['top', 'bottom', 'left', 'right']:
        border_elem = tcBorders.find(qn(f'w:{edge}'))
        if border_elem is not None:
            borders[edge] = {
                'style': border_elem.get(qn('w:val')),
                'color': border_elem.get(qn('w:color')),
                'size': border_elem.get(qn('w:sz')),  # 8분의 1 포인트 단위
            }

    return borders

def convert_border_to_excel(word_borders):
    """Word 테두리를 Excel Border로 변환"""
    def make_side(border_info):
        if border_info is None:
            return Side()

        style = BORDER_STYLE_MAP.get(border_info.get('style'), 'thin')
        color = border_info.get('color', '000000')
        if color == 'auto':
            color = '000000'

        return Side(style=style, color=color) if style else Side()

    return Border(
        top=make_side(word_borders.get('top')),
        bottom=make_side(word_borders.get('bottom')),
        left=make_side(word_borders.get('left')),
        right=make_side(word_borders.get('right'))
    )
```

### 4.4 정렬 변환

```python
from openpyxl.styles import Alignment
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT

# Word 수평 정렬 → Excel 정렬
HORIZONTAL_ALIGN_MAP = {
    WD_ALIGN_PARAGRAPH.LEFT: 'left',
    WD_ALIGN_PARAGRAPH.CENTER: 'center',
    WD_ALIGN_PARAGRAPH.RIGHT: 'right',
    WD_ALIGN_PARAGRAPH.JUSTIFY: 'justify',
    None: 'left',  # 기본값
}

# Word 수직 정렬 → Excel 정렬
VERTICAL_ALIGN_MAP = {
    WD_CELL_VERTICAL_ALIGNMENT.TOP: 'top',
    WD_CELL_VERTICAL_ALIGNMENT.CENTER: 'center',
    WD_CELL_VERTICAL_ALIGNMENT.BOTTOM: 'bottom',
    None: 'top',  # 기본값
}

def get_cell_alignment(cell):
    """셀의 정렬 정보 반환"""
    h_align = None
    if cell.paragraphs:
        h_align = cell.paragraphs[0].alignment

    v_align = cell.vertical_alignment

    return {
        'horizontal': h_align,
        'vertical': v_align,
    }

def convert_alignment_to_excel(word_alignment):
    """Word 정렬을 Excel Alignment으로 변환"""
    return Alignment(
        horizontal=HORIZONTAL_ALIGN_MAP.get(word_alignment['horizontal'], 'left'),
        vertical=VERTICAL_ALIGN_MAP.get(word_alignment['vertical'], 'top'),
        wrap_text=True  # 기본적으로 텍스트 줄바꿈 활성화
    )
```

### 4.5 전체 변환 함수

```python
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

def word_table_to_excel(table, ws, start_row=1, start_col=1):
    """Word 테이블을 Excel 워크시트에 작성"""

    for row_idx, row in enumerate(table.rows):
        for col_idx, cell in enumerate(row.cells):
            # Excel 셀 위치
            excel_row = start_row + row_idx
            excel_col = start_col + col_idx
            excel_cell = ws.cell(row=excel_row, column=excel_col)

            # 텍스트 설정
            excel_cell.value = cell.text

            # 폰트 변환 (첫 번째 Run 기준)
            if cell.paragraphs and cell.paragraphs[0].runs:
                run = cell.paragraphs[0].runs[0]
                excel_cell.font = convert_font_to_excel(run.font)

            # 배경색 변환
            bg_color = get_cell_background_color(cell)
            if bg_color:
                excel_cell.fill = convert_fill_to_excel(bg_color)

            # 테두리 변환
            borders = get_cell_borders_xml(cell)
            excel_cell.border = convert_border_to_excel(borders)

            # 정렬 변환
            alignment = get_cell_alignment(cell)
            excel_cell.alignment = convert_alignment_to_excel(alignment)

    return ws
```

---

## 5. 단위 변환

### 5.1 주요 단위 정리

| 단위 | 설명 | 변환 |
|------|------|------|
| **Twips (DXA)** | 포인트의 1/20 | 1 inch = 1,440 twips |
| **Points (pt)** | 기본 폰트/여백 단위 | 1 inch = 72 pt |
| **Half-points** | 폰트 크기 (Word 내부) | 12pt = 24 half-points |
| **EMU** | 그래픽/이미지 좌표 | 1 inch = 914,400 EMU |
| **Excel 문자 폭** | 열 너비 | 약 7pt = 1 문자 |

### 5.2 변환 상수

```python
# Word/OOXML 단위 상수
TWIPS_PER_INCH = 1440
TWIPS_PER_PT = 20
TWIPS_PER_CM = 567

POINTS_PER_INCH = 72
POINTS_PER_CM = 28.35

EMU_PER_INCH = 914400
EMU_PER_CM = 360000
EMU_PER_PT = 12700
EMU_PER_PIXEL = 9525  # 96 DPI 기준

# Excel 단위 상수
EXCEL_CHAR_WIDTH_PT = 7  # 기본 폰트 기준
DEFAULT_ROW_HEIGHT_PT = 15
DEFAULT_COL_WIDTH_CHARS = 8.43
```

### 5.3 변환 함수

```python
def twips_to_points(twips):
    """Twips → Points"""
    return twips / 20

def points_to_twips(points):
    """Points → Twips"""
    return points * 20

def twips_to_inches(twips):
    """Twips → Inches"""
    return twips / 1440

def twips_to_cm(twips):
    """Twips → Centimeters"""
    return twips / 567

def emu_to_points(emu):
    """EMU → Points"""
    return emu / 12700

def emu_to_pixels(emu, dpi=96):
    """EMU → Pixels"""
    return emu / 9525

def points_to_excel_row_height(points):
    """Points → Excel 행 높이 (직접 사용 가능)"""
    return points  # Excel 행 높이는 포인트 단위

def points_to_excel_col_width(points):
    """Points → Excel 열 너비 (문자 수)"""
    return points / 7  # 대략적인 변환

def twips_to_excel_col_width(twips):
    """Twips → Excel 열 너비 (문자 수)"""
    points = twips / 20
    return points / 7
```

### 5.4 python-docx의 내장 단위 (docx.shared)

```python
from docx.shared import Pt, Inches, Cm, Twips, Emu

# 사용 예
font_size = Pt(12)        # 12 포인트
margin = Inches(1)        # 1 인치
width = Cm(5)             # 5 센티미터
spacing = Twips(100)      # 100 트윕

# 값 추출
print(font_size.pt)       # 12.0
print(margin.inches)      # 1.0
print(width.cm)           # 5.0
print(spacing.twips)      # 100
```

### 5.5 openpyxl 내장 변환 함수

```python
from openpyxl.utils.units import (
    inch_to_dxa,
    pixels_to_EMU,
    cm_to_EMU,
    inch_to_EMU,
    pixels_to_points,
    points_to_pixels,
)

# 사용 예
dxa = inch_to_dxa(1)           # 1440 (1 inch = 1440 dxa/twips)
emu = pixels_to_EMU(100)       # 952500 (100 pixels)
pts = pixels_to_points(96)     # 72 (96 pixels = 72 points at 96dpi)
```

### 5.6 열 너비/행 높이 변환 실전 예제

```python
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

def set_column_width_from_twips(ws, col_num, twips):
    """Twips 단위로 Excel 열 너비 설정"""
    # Twips → Points → 문자 수
    points = twips / 20
    char_width = points / 7

    col_letter = get_column_letter(col_num)
    ws.column_dimensions[col_letter].width = char_width

def set_row_height_from_twips(ws, row_num, twips):
    """Twips 단위로 Excel 행 높이 설정"""
    # Twips → Points (Excel 행 높이는 포인트 단위)
    points = twips / 20
    ws.row_dimensions[row_num].height = points
```

---

## 6. 참고 자료

### 6.1 공식 문서

- [python-docx 공식 문서 - Tables](https://python-docx.readthedocs.io/en/latest/user/tables.html)
- [python-docx API - Table Objects](https://python-docx.readthedocs.io/en/latest/api/table.html)
- [python-docx - Working with Text](https://python-docx.readthedocs.io/en/latest/user/text.html)
- [openpyxl - Working with styles](https://openpyxl.readthedocs.io/en/stable/styles.html)
- [openpyxl.utils.units](https://openpyxl.readthedocs.io/en/stable/api/openpyxl.utils.units.html)
- [Office Open XML - Table Cell Borders](http://officeopenxml.com/WPtableCellProperties-Borders.php)

### 6.2 병합 셀 관련

- [python-docx - How to detect merged cells (Issue #232)](https://github.com/python-openxml/python-docx/issues/232)
- [python-docx - Merge Cells Analysis](https://python-docx.readthedocs.io/en/latest/dev/analysis/features/table/cell-merge.html)
- [python-docx - How to detect merged cells (Issue #1312)](https://github.com/python-openxml/python-docx/issues/1312)

### 6.3 셀 서식 관련

- [python-docx - Font Color](https://python-docx.readthedocs.io/en/latest/dev/analysis/features/text/font-color.html)
- [python-docx - Cell Shading (Issue #146)](https://github.com/python-openxml/python-docx/issues/146)
- [GeeksforGeeks - Formatting Cells using openpyxl](https://www.geeksforgeeks.org/python/formatting-cells-using-openpyxl-in-python/)

### 6.4 변환 예제

- [Medium - Convert Word Documents to Excel with Python](https://medium.com/@alexaae9/convert-word-documents-to-excel-with-python-a-formatting-preserving-approach-f2c3b9ad6cca)
- [Medium - Read tables from docx file to pandas DataFrames](https://medium.com/@karthikeyan.eaganathan/read-tables-from-docx-file-to-pandas-dataframes-f7e409401370)
- [Spire.Doc - Python Convert Word to Excel](https://www.e-iceblue.com/Tutorials/Python/Spire.Doc-for-Python/Program-Guide/Conversion/Python-Convert-Word-to-Excel.html)

### 6.5 win32com 관련

- [Python Tutorials - Read Table Contents in MS Word](https://www.pythontutorials.net/blog/how-to-read-contents-of-an-table-in-ms-word-file-using-python/)
- [Medium - Handling with .doc extension with Python](https://fortes-arthur.medium.com/handling-with-doc-extension-with-python-b6491792311e)

---

## 부록: 완전한 변환 예제

```python
"""Word 테이블을 Excel로 변환하는 완전한 예제"""

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter


class WordToExcelConverter:
    """Word 테이블을 Excel로 변환하는 클래스"""

    BORDER_STYLE_MAP = {
        'single': 'thin',
        'thick': 'medium',
        'double': 'double',
        'dotted': 'dotted',
        'dashed': 'dashed',
        'nil': None,
        None: None,
    }

    def __init__(self, word_path):
        self.doc = Document(word_path)
        self.wb = Workbook()
        self.ws = self.wb.active

    def convert_all_tables(self):
        """모든 테이블을 각각의 시트로 변환"""
        for idx, table in enumerate(self.doc.tables):
            if idx == 0:
                ws = self.ws
                ws.title = f"Table_{idx + 1}"
            else:
                ws = self.wb.create_sheet(f"Table_{idx + 1}")

            self._convert_table(table, ws)

        return self.wb

    def _convert_table(self, table, ws):
        """단일 테이블을 워크시트로 변환"""
        for row_idx, row in enumerate(table.rows, start=1):
            for col_idx, cell in enumerate(row.cells, start=1):
                excel_cell = ws.cell(row=row_idx, column=col_idx)

                # 텍스트
                excel_cell.value = cell.text

                # 서식 적용
                self._apply_font(cell, excel_cell)
                self._apply_fill(cell, excel_cell)
                self._apply_border(cell, excel_cell)
                self._apply_alignment(cell, excel_cell)

    def _apply_font(self, word_cell, excel_cell):
        """폰트 서식 적용"""
        if not word_cell.paragraphs or not word_cell.paragraphs[0].runs:
            return

        run = word_cell.paragraphs[0].runs[0]
        font = run.font

        color = '000000'
        if font.color and font.color.rgb:
            rgb = font.color.rgb
            color = f'{rgb}'

        size = font.size.pt if font.size else 11

        excel_cell.font = Font(
            name=font.name if font.name else 'Calibri',
            size=size,
            bold=font.bold if font.bold else False,
            italic=font.italic if font.italic else False,
            color=color
        )

    def _apply_fill(self, word_cell, excel_cell):
        """배경색 적용"""
        tc = word_cell._tc
        tcPr = tc.tcPr

        if tcPr is None:
            return

        shd = tcPr.find(qn('w:shd'))
        if shd is not None:
            fill_color = shd.get(qn('w:fill'))
            if fill_color and fill_color != 'auto':
                excel_cell.fill = PatternFill(
                    start_color=fill_color,
                    end_color=fill_color,
                    fill_type='solid'
                )

    def _apply_border(self, word_cell, excel_cell):
        """테두리 적용"""
        tc = word_cell._tc
        tcPr = tc.tcPr

        if tcPr is None:
            return

        tcBorders = tcPr.find(qn('w:tcBorders'))
        if tcBorders is None:
            return

        def make_side(edge):
            elem = tcBorders.find(qn(f'w:{edge}'))
            if elem is None:
                return Side()

            style = elem.get(qn('w:val'))
            color = elem.get(qn('w:color')) or '000000'
            if color == 'auto':
                color = '000000'

            excel_style = self.BORDER_STYLE_MAP.get(style, 'thin')
            return Side(style=excel_style, color=color) if excel_style else Side()

        excel_cell.border = Border(
            top=make_side('top'),
            bottom=make_side('bottom'),
            left=make_side('left'),
            right=make_side('right')
        )

    def _apply_alignment(self, word_cell, excel_cell):
        """정렬 적용"""
        excel_cell.alignment = Alignment(
            horizontal='left',
            vertical='top',
            wrap_text=True
        )

    def save(self, output_path):
        """Excel 파일 저장"""
        self.wb.save(output_path)


# 사용 예
if __name__ == '__main__':
    converter = WordToExcelConverter('input.docx')
    converter.convert_all_tables()
    converter.save('output.xlsx')
```
