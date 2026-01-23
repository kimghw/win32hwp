# HWP 테이블 -> 엑셀 변환 패턴 분석

이 문서는 기존 코드베이스의 HWP 테이블 변환 로직을 분석한 결과입니다.

---

## 1. 모듈 구조 개요

```
/table/
  table_info.py          # 셀 탐색, 좌표 매핑, 크기 조회
  table_boundary.py      # 테이블 경계 분석 (xend, yend 계산)
  table_grid.py          # 그리드 좌표 생성, 엑셀 스타일 그리드 매핑
  table_excel_converter.py  # HWP -> 엑셀 변환 메인 로직

/converter_excel/
  cell_style.py          # 셀 스타일 추출 (배경색, 테두리)
  excel_export_data.py   # 데이터 구조 정의 및 단위 변환
  page_meta.py           # 페이지/여백 메타 정보
```

---

## 2. 셀 데이터 추출 흐름

### 2.1 테이블 진입 및 셀 수집

```python
# table_info.py - BFS 방식 셀 수집
class TableInfo:
    def collect_cells_bfs(self) -> Dict[int, CellInfo]:
        """
        행 우선 순회로 모든 셀 탐색
        1. 첫 셀에서 시작하여 오른쪽으로 이동 (MOVE_RIGHT_OF_CELL)
        2. 행 끝에서 아래로 이동 (MOVE_DOWN_OF_CELL)
        3. 반복
        """
```

**MovePos 상수:**
```python
MOVE_LEFT_OF_CELL = 100
MOVE_RIGHT_OF_CELL = 101
MOVE_UP_OF_CELL = 102
MOVE_DOWN_OF_CELL = 103
MOVE_START_OF_CELL = 104  # 행의 시작 셀
MOVE_END_OF_CELL = 105    # 행의 끝 셀
MOVE_TOP_OF_CELL = 106    # 열의 시작 셀
MOVE_BOTTOM_OF_CELL = 107 # 열의 끝 셀
```

### 2.2 셀 텍스트 추출

```python
# table_excel_converter.py
def _get_cell_text(self, list_id: int) -> str:
    hwp.SetPos(list_id, 0, 0)
    hwp.MovePos(4, 0, 0)  # moveTopOfList - 셀 시작
    hwp.HAction.Run("MoveSelBottomOfList")  # 셀 끝까지 선택
    text = hwp.GetTextFile("TEXT", "")
    hwp.HAction.Run("Cancel")  # 선택 해제
    return text.strip()
```

### 2.3 셀 크기 조회

```python
# table_info.py
def get_cell_dimensions(self) -> tuple:
    """CellShape 속성에서 Width, Height 조회"""
    cell_shape = hwp.CellShape
    cell = cell_shape.Item("Cell")
    width = cell.Item("Width")   # HWPUNIT
    height = cell.Item("Height")  # HWPUNIT
    return (width, height)
```

---

## 3. 서식 변환 로직

### 3.1 배경색 추출 (BGR -> RGB 변환)

```python
# cell_style.py
def bgr_to_rgb(bgr: int) -> Tuple[int, int, int]:
    """HWP는 BGR 순서로 저장 -> RGB로 변환"""
    b = (bgr >> 16) & 0xFF
    g = (bgr >> 8) & 0xFF
    r = bgr & 0xFF
    return (r, g, b)

def get_cell_style(hwp, list_id: int) -> CellStyle:
    hwp.SetPos(list_id, 0, 0)
    pset = hwp.HParameterSet.HCellBorderFill
    hwp.HAction.GetDefault("CellBorderFill", pset.HSet)

    # 배경색
    bg_color = pset.FillAttr.WinBrushFaceColor
    # 유효한 색상 확인 (0이 아니고 흰색(4294967295)이 아닌 경우)
    if bg_color and bg_color > 0 and bg_color != 4294967295:
        rgb = bgr_to_rgb(bg_color)
```

### 3.2 글자 속성 추출

```python
# table_excel_converter.py / excel_export_data.py
char_shape = hwp.CharShape
if char_shape:
    font_name = char_shape.Item("FaceNameHangul")  # 한글 글꼴
    height = char_shape.Item("Height")  # HWPUNIT
    font_size_pt = height / 100  # -> pt 변환
    font_bold = bool(char_shape.Item("Bold"))
    font_italic = bool(char_shape.Item("Italic"))
```

### 3.3 정렬 속성 추출

```python
# excel_export_data.py
para_shape = hwp.ParaShape
align_val = para_shape.Item("Align")
# 0=양쪽, 1=왼쪽, 2=오른쪽, 3=가운데, 4=배분, 5=나눔
align_map = {0: 'justify', 1: 'left', 2: 'right', 3: 'center'}
align_horizontal = align_map.get(align_val, 'left')
```

### 3.4 테두리 정보 추출

```python
# cell_style.py
pset = hwp.HParameterSet.HCellBorderFill
hwp.HAction.GetDefault("CellBorderFill", pset.HSet)

border_left = CellBorder(
    type=pset.BorderTypeLeft,
    width=pset.BorderWidthLeft,
    color=pset.BorderCorlorLeft,  # 주의: API 오타 (Corlor)
)
border_right = CellBorder(
    type=pset.BorderTypeRight,
    width=pset.BorderWidthRight,
    color=pset.BorderColorRight,  # 정상 철자
)
```

---

## 4. 병합 셀 처리 방법

### 4.1 xend 기반 행/열 분리

```python
# table_boundary.py
def _calc_xend_from_top_border_cells(self, top_border_cells: List[int]) -> int:
    """첫 번째 행 셀들의 너비 합 = 테이블 전체 가로 크기"""
    cumulative_x = 0
    for cell in top_border_cells:
        w, _ = get_cell_dimensions()
        cumulative_x += w
    return cumulative_x  # xend

def _find_lastcols_by_xend(self, start_cell, xend, tolerance=50):
    """
    xend 기준으로 right_border_cells (마지막 열) 찾기
    - cumulative_x > xend: 새 행 시작 (이전 셀이 last_col)
    - cumulative_x == xend: 현재 셀이 last_col
    """
```

### 4.2 그리드 매핑 (병합 셀 -> 그리드 좌표)

```python
# table_grid.py
@dataclass
class CellGridMapping:
    """테이블 셀과 그리드 셀 매칭 결과"""
    list_id: int
    table_cell: GridCell
    grid_cells: List[Tuple[int, int]]  # (row, col) 리스트
    row_span: Tuple[int, int]  # (start_row, end_row)
    col_span: Tuple[int, int]  # (start_col, end_col)

def map_cells_to_grid(self, grid_result, excel_grid, tolerance=30):
    """
    테이블 셀의 corners와 그리드 셀을 매칭
    - 오차 허용(tolerance)을 적용하여 매칭
    """
    for table_cell in grid_result.cells:
        # 셀 범위 + tolerance 적용
        t_x1 = table_cell.corners.top_left[0] - tolerance
        t_y1 = table_cell.corners.top_left[1] - tolerance
        t_x2 = table_cell.corners.bottom_right[0] + tolerance
        t_y2 = table_cell.corners.bottom_right[1] + tolerance

        # 범위 내 그리드 셀 수집
        for grid_cell in excel_grid.cells:
            if grid_cell.x1 >= t_x1 and grid_cell.x2 <= t_x2 ...
```

### 4.3 엑셀 병합 적용

```python
# table_excel_converter.py
def to_excel(self, filepath):
    # 병합 영역 추적 (중복 방지)
    merged_coords = set()

    for cell_data in sorted_cells:
        if cell_data.rowspan > 1 or cell_data.colspan > 1:
            # 충돌 확인
            conflict_coords = []
            for r in range(cell_data.row, cell_data.end_row + 1):
                for c in range(cell_data.col, cell_data.end_col + 1):
                    if (r, c) in merged_coords:
                        conflict_coords.append((r, c))

            if not conflict_coords:
                # 엑셀 병합 (0-based -> 1-based 변환)
                ws.merge_cells(
                    start_row=cell_data.row + 1,
                    start_column=cell_data.col + 1,
                    end_row=cell_data.end_row + 1,
                    end_column=cell_data.end_col + 1
                )
                # 병합 좌표 기록
                for r, c in merged_area:
                    merged_coords.add((r, c))
```

---

## 5. 단위 변환 (HWPUNIT -> pt/inch)

### 5.1 변환 상수

```python
# excel_export_data.py / page_meta.py
class Units:
    HWPUNIT_PER_PT = 100        # 1 pt = 100 HWPUNIT
    HWPUNIT_PER_INCH = 7200     # 1 inch = 7200 HWPUNIT
    HWPUNIT_PER_CM = 2834.6     # 1 cm = 7200/2.54 HWPUNIT
    HWPUNIT_PER_MM = 283.46     # 1 mm = 7200/25.4 HWPUNIT
    EXCEL_CHAR_WIDTH_PT = 7     # 1 문자 = 7 pt (기본 폰트)
```

### 5.2 변환 공식

```python
# HWPUNIT -> 엑셀 행 높이 (pt 단위)
excel_row_height = hwpunit / 100

# HWPUNIT -> 엑셀 열 너비 (문자 단위)
excel_col_width = hwpunit / 700  # 또는 (hwpunit/100) / 7

# HWPUNIT -> 엑셀 여백 (inch 단위)
excel_margin_inch = hwpunit / 7200

# 글자 크기 HWPUNIT -> pt
font_size_pt = char_height / 100
```

### 5.3 적용 예시

```python
# table_excel_converter.py
# 열 너비 적용
HWPUNIT_PER_CHAR = 700
for col_idx in range(len(x_levels) - 1):
    col_width_hwp = x_levels[col_idx + 1] - x_levels[col_idx]
    col_width_chars = col_width_hwp / HWPUNIT_PER_CHAR
    ws.column_dimensions[col_letter].width = max(col_width_chars, 2)

# 행 높이 적용
HWPUNIT_PER_PT = 100
for row_idx in range(len(y_levels) - 1):
    row_height_hwp = y_levels[row_idx + 1] - y_levels[row_idx]
    row_height_pt = row_height_hwp / HWPUNIT_PER_PT
    ws.row_dimensions[row_idx + 1].height = max(row_height_pt, 12)

# 페이지 여백 적용 (인치 단위)
ws.page_margins = PageMargins(
    left=hwp_margins['left'],     # 이미 inch로 변환됨
    right=hwp_margins['right'],
    top=hwp_margins['top'],
    bottom=hwp_margins['bottom'],
)
```

---

## 6. 데이터 구조

### 6.1 셀 데이터 (CellData / CellInfo)

```python
@dataclass
class CellData:
    list_id: int          # HWP 셀 식별자
    row: int              # 논리 행 (0-based)
    col: int              # 논리 열 (0-based)
    end_row: int          # 병합 끝 행
    end_col: int          # 병합 끝 열
    text: str             # 셀 내용
    rowspan: int = 1      # 세로 병합 수
    colspan: int = 1      # 가로 병합 수
    # 물리 좌표 (HWPUNIT)
    start_x: int = 0
    start_y: int = 0
    end_x: int = 0
    end_y: int = 0
    # 스타일
    bg_color_rgb: tuple = None  # (R, G, B)
    style: CellStyle = None
```

### 6.2 그리드 결과 (TableGridResult)

```python
@dataclass
class TableGridResult:
    cells: List[GridCell]  # 셀 목록
    row_count: int
    col_count: int

@dataclass
class GridCell:
    list_id: int
    row: int
    col: int
    corners: CellCorner   # 4개 모서리 좌표
    lines: CellLine       # 4개 테두리선
```

### 6.3 페이지 정보 (PageInfo)

```python
@dataclass
class PageInfo:
    paper_width: int      # HWPUNIT
    paper_height: int     # HWPUNIT
    orientation: str      # 'portrait' / 'landscape'
    margin_left: int
    margin_right: int
    margin_top: int
    margin_bottom: int
    margin_header: int
    margin_footer: int
    content_width: int    # 계산됨
    content_height: int   # 계산됨
```

---

## 7. 엑셀 변환 전체 흐름

```
1. 테이블 경계 분석 (TableBoundary)
   ├─ 첫 행/마지막 행 셀 찾기
   ├─ xend (테이블 가로 크기) 계산
   └─ yend (테이블 세로 크기) 계산

2. 그리드 구축 (TableGrid)
   ├─ build_grid(): 셀별 corners 계산
   ├─ build_grid_lines(): x_lines, y_lines 생성
   └─ map_cells_to_grid(): 셀 -> 그리드 매핑

3. 데이터 추출 (TableExcelConverter)
   ├─ extract_table_data(): 텍스트 + 스타일 추출
   └─ validate_cell_positions(): 위치 검증

4. 엑셀 생성 (to_excel)
   ├─ 워크북 생성
   ├─ 페이지 설정 (여백, 방향)
   ├─ 셀 병합 적용
   ├─ 데이터 및 스타일 적용
   ├─ 행 높이 / 열 너비 설정
   └─ 메타 시트 생성 (_meta, _page)
```

---

## 8. 주요 API 참조

### HWP API

| API | 설명 |
|-----|------|
| `hwp.SetPos(list_id, para, pos)` | 커서 위치 설정 |
| `hwp.GetPos()` | 현재 위치 (list_id, para, pos) |
| `hwp.MovePos(moveID, para, pos)` | 커서 이동 |
| `hwp.CellShape.Item("Cell")` | 셀 크기 정보 |
| `hwp.CharShape.Item("...")` | 글자 속성 |
| `hwp.ParaShape.Item("...")` | 문단 속성 |
| `hwp.HParameterSet.HCellBorderFill` | 셀 테두리/채우기 |
| `hwp.ParentCtrl` | 상위 컨트롤 (테이블 확인용) |

### openpyxl API

| API | 설명 |
|-----|------|
| `ws.merge_cells(start_row, start_column, end_row, end_column)` | 셀 병합 |
| `ws.cell(row, column).value` | 셀 값 설정 |
| `ws.cell(...).fill = PatternFill(...)` | 배경색 |
| `ws.cell(...).border = Border(...)` | 테두리 |
| `ws.cell(...).alignment = Alignment(...)` | 정렬 |
| `ws.column_dimensions[col].width` | 열 너비 (문자 단위) |
| `ws.row_dimensions[row].height` | 행 높이 (pt 단위) |
| `ws.page_margins = PageMargins(...)` | 페이지 여백 (인치) |
| `ws.page_setup.orientation` | 페이지 방향 |

---

## 9. 주의사항

1. **좌표 변환**: HWP는 0-based, 엑셀은 1-based
2. **색상 순서**: HWP는 BGR, 엑셀/일반은 RGB
3. **단위 차이**:
   - HWP: HWPUNIT (100 = 1pt)
   - 엑셀 행 높이: pt
   - 엑셀 열 너비: 문자 수
   - 엑셀 여백: inch
4. **병합 충돌**: 이미 병합된 영역에 다시 병합 시도하면 오류
5. **API 오타**: `BorderCorlorLeft` (HWP API 자체 오타)
