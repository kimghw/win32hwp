# 테이블 분석 모듈 문서

## 개요

테이블의 모든 `list_id`를 추출하고, 커서 이동을 통해 테이블 구조를 파악하는 방법을 정리한다.
주요 모듈은 `table/table_boundary.py`와 `table/table_info.py`이다.

---

## 용어 정의

### 기본 용어

| 용어 | 설명 |
|------|------|
| `list_id` | 각 셀의 고유 식별자 (본문=0, 테이블 셀=1 이상) |
| `tbl` | 테이블 컨트롤 (CtrlID="tbl") |

### 테이블 전체 위치 용어

| 용어 | 설명 |
|------|------|
| `table_origin` | 테이블 시작점 (첫 번째 행, 첫 번째 열) |
| `table_end` | 테이블 끝점 (last_cols의 마지막 셀) |
| `table_cell_counts` | 테이블의 총 셀 개수 |

### 행/열 그룹 용어

| 용어 | 설명 |
|------|------|
| `first_rows` | 첫 번째 행에 속한 셀들 (리스트) |
| `bottom_rows` | 마지막 행에 속한 셀들 (리스트) |
| `first_cols` | 첫 번째 컬럼에 속한 셀들 (리스트) |
| `last_cols` | 마지막 컬럼에 속한 셀들 (리스트) |

### 특정 위치 용어

| 용어 | 설명 |
|------|------|
| `first_col` | 특정 row의 첫 번째 컬럼 값 |
| `last_col` | 특정 row의 마지막 컬럼 값 |
| `first_row` | 특정 컬럼의 첫 번째 행 값 |
| `bottom_row` | 특정 컬럼의 마지막 행 값 |

### 셀 병합 용어

| 용어 | 설명 |
|------|------|
| `merged_cell` | 병합된 셀 (여러 행/열을 합친 셀) |
| `single_cell` | 병합되지 않은 단일 셀 |
| `row_span` | 행 방향 병합 크기 (세로로 몇 칸 차지) |
| `col_span` | 열 방향 병합 크기 (가로로 몇 칸 차지) |

---

## MovePos 셀 이동 상수

```python
MOVE_LEFT_OF_CELL = 100   # 왼쪽 셀로 이동
MOVE_RIGHT_OF_CELL = 101  # 오른쪽 셀로 이동
MOVE_UP_OF_CELL = 102     # 위쪽 셀로 이동
MOVE_DOWN_OF_CELL = 103   # 아래쪽 셀로 이동
MOVE_START_OF_CELL = 104  # 행의 시작 셀
MOVE_END_OF_CELL = 105    # 행의 끝 셀
MOVE_TOP_OF_CELL = 106    # 열의 시작 셀 (맨 위)
MOVE_BOTTOM_OF_CELL = 107 # 열의 끝 셀 (맨 아래)
```

---

## 핵심 데이터 클래스

### CellCoordinate

셀의 좌표 정보를 저장하는 데이터 클래스.

```python
@dataclass
class CellCoordinate:
    list_id: int          # 셀의 list_id
    row: int = 0          # 행 번호 (0부터 시작)
    col: int = 0          # 열 번호 (0부터 시작)
    visit_count: int = 0  # 해당 셀 방문 횟수
```

### CellAdjacency

셀의 인접 노드(상/하/좌/우) 정보를 저장하는 데이터 클래스.
병합 셀의 경우 여러 셀과 인접할 수 있으므로 리스트로 관리.

```python
@dataclass
class CellAdjacency:
    list_id: int
    row: int = 0
    col: int = 0
    left: List[int] = field(default_factory=list)   # 왼쪽 인접 노드들
    right: List[int] = field(default_factory=list)  # 오른쪽 인접 노드들
    up: List[int] = field(default_factory=list)     # 위쪽 인접 노드들
    down: List[int] = field(default_factory=list)   # 아래쪽 인접 노드들
```

### CellWidthInfo / RowWidthResult

행별 셀 너비 정보를 저장하는 데이터 클래스.

```python
@dataclass
class CellWidthInfo:
    list_id: int
    width: int = 0       # 셀 너비 (HWPUNIT)
    height: int = 0      # 셀 높이 (HWPUNIT)
    start_x: int = 0     # 셀 시작 x 좌표
    end_x: int = 0       # 셀 끝 x 좌표
    col_index: int = 0   # 행 내 열 인덱스 (0부터)

@dataclass
class RowWidthResult:
    row_index: int = 0                  # first_cols 인덱스 (행 번호)
    start_cell: int = 0                 # 행 시작 셀 (first_col)
    end_cell: int = 0                   # 행 끝 셀 (last_col)
    cells: List[CellWidthInfo] = field(default_factory=list)
    total_width: int = 0                # 행 전체 너비
    start_y: int = 0                    # 행 시작 y 좌표
    end_y: int = 0                      # 행 끝 y 좌표
    row_height: int = 0                 # 행 높이 (첫 번째 셀 기준)
```

---

## TableBoundary 클래스

### 개요

테이블의 경계(첫 번째 행, 마지막 행, 첫 번째 열, 마지막 열)를 판별하고,
셀 좌표를 매핑하는 기능을 제공한다.

```python
from table.table_boundary import TableBoundary

boundary = TableBoundary(hwp, debug=True)
result = boundary.check_boundary_table()
```

### 셀 경계 판별 메서드

#### check_first_row_cell: 첫 번째 행 여부 판별

특정 셀에서 위로 이동 시 `has_tbl`이 `False`이거나 같은 셀에 머물면 `True` 반환.

```python
def check_first_row_cell(self, target_list_id: int) -> bool:
    """위로 이동 시 tbl 밖이거나 같은 셀에 머물면 True (첫 번째 행)"""
    self.hwp.SetPos(target_list_id, 0, 0)
    self.hwp.HAction.Run("MoveUp")
    new_list_id, _, _ = self.hwp.GetPos()
    parent = self.hwp.ParentCtrl
    has_tbl = parent and parent.CtrlID == "tbl"
    return not has_tbl or new_list_id == target_list_id
```

#### check_bottom_row_cell: 마지막 행 여부 판별

특정 셀에서 아래로 이동 시 `has_tbl`이 `False`이거나 같은 셀에 머물면 `True` 반환.

```python
def check_bottom_row_cell(self, target_list_id: int) -> bool:
    """아래로 이동 시 tbl 밖이거나 같은 셀에 머물면 True (마지막 행)"""
    self.hwp.SetPos(target_list_id, 0, 0)
    self.hwp.HAction.Run("MoveDown")
    new_list_id, _, _ = self.hwp.GetPos()
    parent = self.hwp.ParentCtrl
    has_tbl = parent and parent.CtrlID == "tbl"
    return not has_tbl or new_list_id == target_list_id
```

#### move_down_left_right: 하/좌/우 각각 이동 결과

특정 셀에서 하, 좌, 우 방향으로 **셀 단위** 이동했을 때 `list_id`와 `has_tbl`을 반환.

```python
def move_down_left_right(self, target_list_id: int) -> Dict[str, Tuple[int, bool]]:
    """
    Returns:
        dict: {
            'down': (list_id, has_tbl),
            'left': (list_id, has_tbl),
            'right': (list_id, has_tbl)
        }
    """
```

#### move_up_right_down: 하->우->상 커서 이동으로 last_col 판별

특정 셀에서 **하 -> 우 -> 상** 순서로 커서 이동 후, 시작점으로 돌아오면 같은 열(last_col).

```python
def move_up_right_down(self, target_list_id: int) -> Tuple[int, bool]:
    """
    Returns:
        tuple: (down_id, is_last_col)
            - down_id: 아래로 이동한 셀의 list_id
            - is_last_col: 상으로 돌아온 셀이 시작점과 같으면 True (같은 열)
    """
```

---

## 셀 좌표 매핑 알고리즘 (v1~v4)

### map_cell_coordinates (v1)

first_col 기반으로 셀 좌표 매핑. 방문 횟수로 행 결정.

- first_cols의 첫 번째 원소부터 시작
- 오른쪽 키보드 이동으로 순회하면서 동일한 셀 방문 횟수 저장
- 다음 first_col 원소를 만나면 방문 횟수로 행 결정

### map_cell_coordinates_v2

first_col 기반으로 셀 좌표 매핑 (개선 버전).

- first_cols 순회하면서 각 행의 시작점 결정
- 각 first_col에서 오른쪽으로 순회
- 열 번호는 현재 first_col로부터의 거리

### map_cell_coordinates_v3

서브셀 기반 셀 좌표 매핑.

**알고리즘:**
1. first_col[0]에서 시작하여 오른쪽으로 순회
2. 새 셀(처음 방문)을 만나면 -> 현재 행, 열+1
3. 이미 방문한 셀(2회 이상)을 만나면 -> 행+1 (한 번만), 열은 왼쪽 열+1 유지
4. first_col을 만나면 -> 열=0
5. 좌표는 최초 방문 시에만 설정

### map_cell_coordinates_v4 (권장)

v3 + 위쪽 셀 기준 열 조정.

**알고리즘:**
1. v3로 기본 좌표 매핑
2. 2행부터 위쪽 셀 기준으로 열 번호 조정

**열 조정 로직 (`_adjust_column_by_upper_cells`):**
1. 2행부터 각 셀에 대해
2. 커서를 0행까지 위로 올리면서 모든 위쪽 셀의 열 중 가장 큰 값을 찾음
3. 열 번호 = max(위쪽 셀들의 최대 열, 왼쪽 셀의 열 + 1)
4. 0행은 그대로 유지

```python
# 사용 예시
boundary = TableBoundary(hwp, debug=True)
result = boundary.check_boundary_table()
coord_result = boundary.map_cell_coordinates_v4(result)
boundary.print_cell_coordinates(coord_result)
```

---

## 병합 셀 처리 방법

### 개념

병합 셀은 여러 행/열을 합친 셀로, 하나의 `list_id`가 여러 좌표를 차지한다.

### 처리 전략

1. **방문 횟수 기반**: 같은 셀을 여러 번 방문하면 병합 셀로 판단
2. **양방향 인접 관계**: A->B 이동이 가능하면 B->A도 인접으로 설정
3. **위쪽 셀 기준 열 조정**: 병합 셀 아래의 셀들은 위쪽 셀의 최대 열을 기준으로 열 번호 조정

### build_cell_adjacency: 인접 노드 계산

모든 셀의 인접 노드(상/하/좌/우)를 계산한다. 양방향 관계를 활용하여 병합 셀도 처리.

```python
def build_cell_adjacency(self, coord_result: CellCoordinateResult = None) -> CellAdjacencyResult:
    """
    알고리즘:
    1. 각 셀에서 상/하/좌/우로 커서 이동하여 인접 셀 1개씩 찾음
    2. 양방향 관계를 활용: A->B 이면 B->A도 성립
       예: 13의 up이 3이면, 3의 down에 13 추가
    3. 이를 통해 병합 셀(13)이 여러 셀(3~8)과 인접한 경우도 처리
    """
```

**사용 예시:**
```python
adjacency_result = boundary.build_cell_adjacency(coord_result)
boundary.print_cell_adjacency(adjacency_result)
```

---

## 행별 너비 계산 (calculate_row_widths)

### 개요

first_cols에서 last_cols까지 오른쪽으로 이동하면서 셀 너비를 누적 계산한다.

### 알고리즘

1. first_cols와 last_cols를 y좌표 기준으로 정렬하여 행 매칭
2. 각 first_col에서 시작하여 해당 행의 last_col까지만 이동
3. 이동할 때마다 셀 너비를 수집하고 누적 (이미 방문한 셀은 건너뜀)
4. 해당 행의 last_col에 도달하면 행 종료
5. 다음 first_col을 만나면 종료 (다른 행으로 넘어감)

### 셀 크기 조회

```python
def get_cell_dimensions(self) -> tuple:
    """
    현재 커서 위치의 셀 너비/높이 조회 (CellShape 속성 사용)

    Returns:
        tuple: (width, height) HWPUNIT 단위
    """
    cell_shape = self.hwp.CellShape
    cell = cell_shape.Item("Cell")
    width = cell.Item("Width")
    height = cell.Item("Height")
    return (width, height)
```

### 사용 예시

```python
boundary = TableBoundary(hwp, debug=True)
result = boundary.check_boundary_table()
width_result = boundary.calculate_row_widths(result)
boundary.print_row_widths(width_result)
```

**출력 예시:**
```
=== 행별 셀 너비 결과 ===
최대 너비: 28350 HWPUNIT

[행 0] 시작=2 -> 끝=8, 전체 너비=28350
  col=0: list_id=2, width=4050, x=[0~4050]
  col=1: list_id=3, width=4050, x=[4050~8100]
  ...
```

---

## check_boundary_table

테이블의 모든 셀을 순회하면서 경계 정보를 계산한다.

### 실행 순서

#### 1. table_origin, table_cell_counts 계산

테이블의 첫 번째 셀의 `list_id`와 총 셀 개수를 구한다.

#### 2. first_rows와 bottom_rows 계산

모든 셀을 순회하면서 첫 번째 행과 마지막 행에 속한 셀들을 리스트로 저장한다.

| 변수명 | 설명 | 사용 함수 |
|--------|------|-----------|
| `first_rows` | 첫 번째 행에 속한 셀들 (리스트) | `check_first_row_cell` |
| `bottom_rows` | 마지막 행에 속한 셀들 (리스트) | `check_bottom_row_cell` |

#### 3. first_cols 계산

`table_origin`에서 아래로 내려가면서 첫 번째 컬럼 셀들을 수집한다.

#### 4. last_cols 계산

`first_rows`의 마지막 셀에서 시작하여 `move_up_right_down`을 활용해 마지막 컬럼 셀들을 수집한다.

- 시작 셀은 무조건 `last_cols`에 포함
- `is_last_col`이 True면 `down_id`를 `last_cols`에 추가
- `is_last_col`이 False여도 마지막 행이면 `last_cols`에 추가

#### 5. table_end 계산

`last_cols`의 마지막 `list_id`가 `table_end`이다.

| 변수명 | 설명 | 계산 방법 |
|--------|------|-----------|
| `first_cols` | 첫 번째 컬럼에 속한 셀들 | table_origin에서 아래로 순회 |
| `last_cols` | 마지막 컬럼에 속한 셀들 | first_rows 마지막에서 move_up_right_down 활용 |
| `table_end` | 테이블 끝점 | last_cols의 마지막 값 |

---

## 셀 위치 측정 (measure_cell_pos.py)

### 개요

first_cols를 순회하며 모든 셀의 X, Y 좌표 레벨을 수집하고, 좌표를 계산한다.

### 알고리즘

1. first_cols 순회하며 행의 시작 Y 좌표 계산
2. 각 행에서 오른쪽으로 순회하며 셀의 X 좌표와 너비 수집
3. 셀이 분할되어 있으면 (높이가 행 높이보다 작으면) 재귀적으로 서브셀 순회
4. X, Y 레벨 수집 후 근접값 병합 (TOLERANCE 이내 값들을 하나로)
5. 각 셀의 좌표 = (x_level_index, y_level_index)

### 근접값 병합

```python
TOLERANCE = 3

def merge_close_levels(levels, tolerance):
    """근접한 레벨들을 병합하여 대표값 리스트 반환"""
    if not levels:
        return []

    merged = [levels[0]]
    for level in levels[1:]:
        if level - merged[-1] <= tolerance:
            pass  # 근접한 값은 무시
        else:
            merged.append(level)
    return merged
```

---

## 함수 요약

| 함수명 | 설명 | 반환값 |
|--------|------|--------|
| `move_down_left_right` | 하/좌/우 셀 단위 이동 결과 | `dict` {방향: (list_id, has_tbl)} |
| `check_first_row_cell` | 위로 이동 시 tbl 밖이거나 같은 셀이면 True | `bool` |
| `check_bottom_row_cell` | 아래로 이동 시 tbl 밖이거나 같은 셀이면 True | `bool` |
| `move_up_right_down` | 하->우->상 커서 이동, last_col 판별 | `tuple` (down_id, is_last_col) |
| `map_cell_coordinates_v4` | 서브셀 + 위쪽 셀 기준 열 조정 좌표 매핑 | `CellCoordinateResult` |
| `build_cell_adjacency` | 모든 셀의 인접 관계 계산 | `CellAdjacencyResult` |
| `calculate_row_widths` | 행별 셀 너비 계산 | `TableWidthResult` |

---

## 전체 사용 예시

```python
from table.table_boundary import TableBoundary
from cursor import get_hwp_instance

hwp = get_hwp_instance()
boundary = TableBoundary(hwp, debug=True)

# 테이블 내부인지 확인
if not boundary._is_in_table():
    print("커서가 테이블 내부에 있지 않습니다.")
    exit(1)

# 1. 경계 분석
result = boundary.check_boundary_table()
boundary.print_boundary_info(result)

# 2. 셀 좌표 매핑 (v4 권장)
coord_result = boundary.map_cell_coordinates_v4(result)
boundary.print_cell_coordinates(coord_result)

# 3. 셀 인접 관계 (노드 그래프)
adjacency_result = boundary.build_cell_adjacency(coord_result)
boundary.print_cell_adjacency(adjacency_result)

# 4. 행별 셀 너비 계산
width_result = boundary.calculate_row_widths(result)
boundary.print_row_widths(width_result)

# 5. JSON 파일로 저장
boundary.save_adjacency_to_json(adjacency_result)
```

---

## 파일 구조

```
table/
├── __init__.py
├── table_info.py       # TableInfo 클래스 - 기본 셀 수집, 좌표 매핑
├── table_boundary.py   # TableBoundary 클래스 - 경계 판별, 좌표 매핑 v1~v4
├── table_cell_info.py  # 셀 유틸리티 함수 - 컨트롤 탐색, 서식 조회
└── table_field.py      # 필드 관련 기능

map_coordinates_to_table.py  # 셀에 좌표 텍스트 삽입
measure_cell_pos.py          # 셀 위치 X/Y 레벨 측정
```
