# Table 모듈

HWP 테이블 분석 및 그리드 매핑 모듈

## 파일 구조

| 파일 | 용도 |
|------|------|
| `table_info.py` | 셀 BFS 순회, 크기 조회, 이동 상수 |
| `table_boundary.py` | 4방향 경계 셀 및 좌표 계산 |
| `table_grid.py` | 셀 corners 계산, 엑셀 스타일 그리드 매핑 |
| `table_grid_visual.py` | 그리드 시각화 (Pillow) |
| `table_cell_info.py` | 셀 유틸리티 (컨트롤 탐색, 서식 조회) |
| `table_field.py` | 필드 CRUD |
| `table_excel_converter.py` | 엑셀 변환 |

## 사용 흐름

```python
from table import TableGrid, TableBoundary

grid = TableGrid(hwp)
boundary = grid._boundary.check_boundary_table()
result = grid.build_grid(boundary)
excel_grid = grid.build_grid_lines(result, boundary)
mappings = grid.map_cells_to_grid(result, excel_grid, tolerance=5)

# mappings: list_id ↔ 그리드 좌표 매핑
for m in mappings:
    print(f"list_id={m.list_id} → row:{m.row_span}, col:{m.col_span}")
```

## MovePos 셀 이동 상수

```python
MOVE_LEFT_OF_CELL = 100
MOVE_RIGHT_OF_CELL = 101
MOVE_UP_OF_CELL = 102
MOVE_DOWN_OF_CELL = 103
MOVE_TOP_OF_CELL = 106
MOVE_BOTTOM_OF_CELL = 107
```
