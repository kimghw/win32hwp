# 테이블 경계 및 그리드 모듈

## 모듈 구조

```
table/
├── table_boundary.py    # 테이블 경계 판별
├── table_grid.py        # 그리드 좌표 생성 및 셀 매핑
└── table_grid_visual.py # 그리드 시각화 (Pillow)
```

---

## table_boundary.py

테이블의 4방향 경계 셀과 좌표를 계산합니다.

### TableBoundaryResult

```python
@dataclass
class TableBoundaryResult:
    table_origin: int           # 첫 번째 셀 list_id
    table_end: int              # 마지막 셀 list_id
    table_cell_counts: int      # 총 셀 개수

    top_border_cells: List[int]     # 첫 행 셀들
    bottom_border_cells: List[int]  # 마지막 행 셀들
    left_border_cells: List[int]    # 첫 열 셀들
    right_border_cells: List[int]   # 마지막 열 셀들

    start_x, start_y: int = 0       # 테이블 시작 (항상 0)
    end_x, end_y: int               # 테이블 끝 (HWPUNIT)
```

### 핵심 메서드

| 메서드 | 설명 |
|--------|------|
| `check_boundary_table()` | 경계 분석 실행, TableBoundaryResult 반환 |
| `check_first_row_cell(list_id)` | 첫 행 여부 (위로 이동 시 테이블 밖이면 True) |
| `check_bottom_row_cell(list_id)` | 마지막 행 여부 |

---

## table_grid.py

셀 좌표(corners)를 계산하고 엑셀 스타일 그리드로 변환합니다.

### 주요 데이터 클래스

| 클래스 | 설명 |
|--------|------|
| `GridCell` | 테이블 셀 정보 (list_id, row, col, corners, lines) |
| `TableGridResult` | build_grid() 결과 (cells 리스트) |
| `ExcelStyleGrid` | 엑셀 스타일 2D 그리드 (x_lines, y_lines, cells) |
| `CellGridMapping` | **테이블 셀 ↔ 그리드 셀 매핑 결과** |

### CellGridMapping (최종 출력)

```python
@dataclass
class CellGridMapping:
    list_id: int                      # HWP 셀 식별자
    table_cell: GridCell              # 원본 셀 정보
    grid_cells: List[Tuple[int, int]] # 매핑된 그리드 좌표 [(row, col), ...]
    row_span: Tuple[int, int]         # (start_row, end_row)
    col_span: Tuple[int, int]         # (start_col, end_col)
```

### 핵심 메서드

| 메서드 | 설명 |
|--------|------|
| `build_grid(boundary)` | 셀 corners 계산 → TableGridResult |
| `build_grid_lines(grid, boundary)` | x/y 라인 추출 → ExcelStyleGrid |
| `map_cells_to_grid(grid, excel_grid, tolerance)` | 셀 매핑 → List[CellGridMapping] |

### tolerance 파라미터

`map_cells_to_grid()`의 `tolerance`는 셀 경계 매칭 허용 오차입니다.

| 값 | 효과 |
|----|------|
| 큰 값 (30+) | 틀어진 줄도 같은 줄로 인식 (관대) |
| 작은 값 (5~10) | 정확히 일치하는 줄만 매칭 (엄격) |

---

## table_grid_visual.py

그리드 매핑 결과를 이미지로 시각화합니다.

### 사용법

```python
from table_grid import TableGrid
from table_grid_visual import visualize_table_grid

grid = TableGrid(hwp)
boundary = grid._boundary.check_boundary_table()
result = grid.build_grid(boundary)
excel_grid = grid.build_grid_lines(result, boundary)
mappings = grid.map_cells_to_grid(result, excel_grid, tolerance=5)

visualize_table_grid(mappings, excel_grid, "output.jpg")
```

### 출력

- 각 테이블 셀을 다른 색으로 표시
- 셀 내부에 list_id 표시
- 셀 경계는 두꺼운 검은선

---

## 처리 흐름

```
1. TableBoundary.check_boundary_table()
   → TableBoundaryResult (경계 셀 리스트, 좌표)

2. TableGrid.build_grid(boundary)
   → TableGridResult (셀별 corners 좌표)

3. TableGrid.build_grid_lines(grid, boundary)
   → ExcelStyleGrid (x_lines, y_lines → 2D 그리드)

4. TableGrid.map_cells_to_grid(grid, excel_grid, tolerance)
   → List[CellGridMapping] (list_id ↔ 그리드 좌표 매핑)
```

---

## 좌표 계산 방식

### xend (테이블 가로 크기)
- `top_border_cells` 셀들의 width 합

### yend (테이블 세로 크기)
- `left_border_cells` 셀들의 height 합

### 셀 corners
- 테이블 원점(0,0)에서 우측으로 이동하며 너비 누적
- `right_border_cells` 도달 시 행 변경, y 좌표 누적

---

## 용어 매핑

| 변수명 | 의미 |
|--------|------|
| `top_border_cells` | 테이블 상단 경계 (첫 행 셀들) |
| `bottom_border_cells` | 테이블 하단 경계 (마지막 행 셀들) |
| `left_border_cells` | 테이블 좌측 경계 (첫 열 셀들) |
| `right_border_cells` | 테이블 우측 경계 (마지막 열 셀들) |
