# 테이블 경계 추출

## 결과 구조

```python
@dataclass
class TableBoundaryResult:
    # 셀 ID 경계
    table_origin: int        # 첫 행의 첫 번째 셀
    table_end: int           # 마지막 행의 가장 오른쪽 셀
    table_cell_counts: int   # 총 셀 개수

    # 4방향 경계 셀 리스트
    first_rows: List[int]    # 첫 번째 행 셀들
    bottom_rows: List[int]   # 마지막 행 셀들
    first_cols: List[int]    # 첫 번째 열 셀들
    last_cols: List[int]     # 마지막 열 셀들

    # 좌표 경계 (HWPUNIT)
    start_x: int = 0         # 테이블 시작 x (항상 0)
    start_y: int = 0         # 테이블 시작 y (항상 0)
    end_x: int = 0           # 테이블 끝 x (xend)
    end_y: int = 0           # 테이블 끝 y (yend)
```

## 좌표 경계 계산 방식

| 변수 | 계산 방식 |
|------|-----------|
| `start_x` | 0 (첫 열 시작) |
| `start_y` | 0 (첫 행 시작) |
| `end_x` | first_rows 셀들의 width 합 (= xend) |
| `end_y` | first_cols 따라 내려가며 height 누적 → 마지막 셀 하단 |

```
(start_x, start_y) = (0, 0)
         ↓
         ┌─────────────────────┐
         │                     │
         │      테이블         │
         │                     │
         └─────────────────────┘
                               ↑
                    (end_x, end_y)
```

## 셀 리스트 계산 방식

| 변수 | 계산 방식 |
|------|-----------|
| `table_origin` | 정렬된 list_id 중 첫 번째 |
| `table_end` | `last_cols[-1]` |
| `first_rows` | 모든 셀 순회 → `MoveUp` 시 테이블 밖이거나 같은 셀이면 첫 행 |
| `bottom_rows` | 모든 셀 순회 → `MoveDown` 시 테이블 밖이거나 같은 셀이면 마지막 행 |
| `first_cols` | 시작 셀 + xend 초과/도달 시 새 행의 첫 셀 |
| `last_cols` | xend 초과 시 이전 셀, xend 도달 시 현재 셀 |

## first_cols / last_cols 상세

```
xend = 첫 행 너비 합

우측 순회하면서 너비 누적:
- 시작: first_cols = [start_cell]
- xend 초과: 이전 셀 → last_col, 현재 셀 → first_col
- xend 도달: 현재 셀 → last_col, 다음 셀 → first_col
```

## 핵심 메서드

- `check_boundary_table()` - 경계 분석 실행
- `check_first_row_cell()` / `check_bottom_row_cell()` - 행 경계 판정
- `_calc_xend_from_first_rows()` - end_x (xend) 계산
- `_calc_yend_from_first_cols()` - end_y (yend) 계산
- `_find_lastcols_by_xend()` - first_cols/last_cols 계산

## 용어

- `cell_corners` : 테이블을 구성하는 각 셀의 네 꼭짓점 좌표 집합
- `cell_lines`   : 테이블을 구성하는 각 셀의 모서리 선분들의 집합
- `right_border_line`  : 테이블 외곽 기준 가장 우측 라인
- `left_border_line `  : 테이블 외곽 기준 가장 좌측 라인
- `top_border_line  `  : 테이블 외곽 기준 가장 상단 라인
- `bottom_border_line` : 테이블 외곽 기준 가장 하단 라인
