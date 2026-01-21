# Table 모듈

HWP 테이블 분석 및 조작 모듈

## 파일 구조

| 파일 | 용도 |
|------|------|
| `table_info.py` | 셀 BFS 순회, 크기 조회 |
| `table_boundary.py` | 좌표 매핑 (v4), 인접 관계 |
| `table_cell_info.py` | 셀 순회, 컨트롤 탐색 |
| `table_field.py` | 필드 CRUD |
| `measure_cell_pos.py` | X/Y 레벨 기반 좌표 측정 |

## 작업 플로우

### 1. 기본 셀 정보 수집

```python
from cursor import get_hwp_instance
from table.table_info import TableInfo

hwp = get_hwp_instance()
table = TableInfo(hwp)
table.enter_table(0)

# BFS로 모든 셀 수집
cells = table.collect_cells_bfs()
# → {list_id: CellInfo(left, right, up, down, width, height)}
```

### 2. 좌표 매핑 (권장: v4)

```python
from table.table_boundary import TableBoundary

boundary = TableBoundary(hwp)
result = boundary.check_boundary_table()
coords = boundary.map_cell_coordinates_v4(result)
# → {list_id: CellCoordinate(row, col)}
```

### 3. 인접 관계 계산

```python
adjacency = boundary.build_cell_adjacency(coords)
# → {list_id: CellAdjacency(left, right, up, down)}
```

### 4. 필드 작업

```python
from table.table_field import TableField

field = TableField(hwp)
field.enter_table(0)

# 좌표로 필드 생성/조회
field.create_field_at_coord(0, 0, "field_name")
value = field.get_field_value_by_coord(0, 0)
```

### 5. X/Y 레벨 기반 좌표 (물리적 위치)

```python
# measure_cell_pos.py 직접 실행
# → 셀의 실제 X/Y 좌표 기반으로 row/col 계산
```

## MovePos 셀 이동 상수

```python
MOVE_LEFT_OF_CELL = 100   # 왼쪽
MOVE_RIGHT_OF_CELL = 101  # 오른쪽
MOVE_UP_OF_CELL = 102     # 위
MOVE_DOWN_OF_CELL = 103   # 아래
MOVE_TOP_OF_CELL = 106    # 열의 맨 위
MOVE_BOTTOM_OF_CELL = 107 # 열의 맨 아래
```
