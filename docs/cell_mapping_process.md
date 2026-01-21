# HWP 테이블 셀 좌표 매핑 프로세스

이 문서는 HWP 테이블의 셀을 그리드 좌표로 매핑하는 전체 프로세스를 설명합니다.

---

## 1. 전체 아키텍처 개요

### 1.1 모듈 구조

```
table/
├── table_info.py           # 테이블 기본 정보 및 셀 순회
├── table_boundary.py       # 테이블 경계 판별 및 좌표 매핑
├── cell_position.py        # XY맵 기반 셀 위치 계산
└── table_excel_converter.py # 엑셀 변환
```

### 1.2 핵심 개념

| 개념 | 설명 |
|------|------|
| `list_id` | HWP 내부에서 각 셀을 식별하는 고유 ID |
| `물리 좌표` | 셀의 실제 위치와 크기 (HWPUNIT 단위) |
| `그리드 좌표` | 엑셀 스타일의 (row, col) 좌표 |
| `X/Y 레벨` | 물리 좌표를 정규화한 그리드 경계선 |

### 1.3 처리 흐름

```
HWP 테이블
    │
    ▼
┌─────────────────────────────────┐
│  1. 셀 순회 및 경계 판별        │  (TableBoundary)
│     - first_cols, last_cols    │
│     - BFS 순회                 │
└─────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────┐
│  2. 물리 좌표 수집 (XY맵)       │  (CellPositionCalculator)
│     - 각 셀의 start_x, end_x   │
│     - 각 셀의 start_y, end_y   │
└─────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────┐
│  3. 그리드 레벨 생성            │  (CellPositionCalculator)
│     - X 레벨 병합              │
│     - Y 레벨 병합              │
└─────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────┐
│  4. list_id → 그리드 좌표 매핑  │  (CellPositionCalculator)
│     - 시작/끝 좌표 계산         │
│     - rowspan, colspan 계산    │
└─────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────┐
│  5. 엑셀 변환                   │  (TableExcelConverter)
│     - 셀 병합 적용             │
│     - 텍스트 추출              │
└─────────────────────────────────┘
```

---

## 2. XY맵 작성 과정 (물리 좌표 수집)

### 2.1 개요

XY맵은 각 셀의 물리적 위치(HWPUNIT 단위)를 수집하는 과정입니다. 이 정보는 나중에 그리드 좌표로 변환됩니다.

### 2.2 주요 데이터 구조

```python
# cell_position.py
@dataclass
class CellRange:
    list_id: int
    start_row: int      # 그리드 시작 행
    start_col: int      # 그리드 시작 열
    end_row: int        # 그리드 끝 행
    end_col: int        # 그리드 끝 열
    rowspan: int        # 행 병합 수
    colspan: int        # 열 병합 수
    # 물리적 좌표 (HWPUNIT)
    start_x: int = 0    # 셀 시작 X 좌표
    start_y: int = 0    # 셀 시작 Y 좌표
    end_x: int = 0      # 셀 끝 X 좌표
    end_y: int = 0      # 셀 끝 Y 좌표
```

### 2.3 수집 알고리즘

`CellPositionCalculator.calculate()` 메서드의 동작:

```
1. 경계 분석 (TableBoundary)
   ├── first_cols: 각 행의 첫 번째 셀 목록
   └── last_cols: 각 행의 마지막 셀 목록

2. 행 단위 순회 (first_cols 기준)
   for row_start in first_cols:
       │
       ├── 현재 행의 시작 Y 좌표 = cumulative_y
       ├── 행 높이 = get_cell_dimensions()
       │
       └── 열 단위 순회 (오른쪽으로 이동)
           while current_id not in last_cols:
               │
               ├── 셀 크기 조회: width, height
               │
               ├── 물리 좌표 기록:
               │   ├── start_x = cumulative_x
               │   ├── end_x = cumulative_x + width
               │   ├── start_y = row_start_y
               │   └── end_y = row_start_y + height
               │
               ├── X/Y 레벨에 좌표 추가
               │
               └── 분할 셀 처리 (height < row_height)
                   └── collect_split_cells() 재귀 호출

3. X/Y 레벨 집합 완성
   x_levels = {0, start_x1, start_x2, ...}
   y_levels = {0, start_y1, end_y1, ...}
```

### 2.4 분할 셀 처리

행 내에서 세로로 분할된 셀(서브셀)을 처리하는 `collect_split_cells()`:

```python
def collect_split_cells(cell_id, start_x, start_y, row_end_y, visited):
    """분할된 셀 재귀 순회"""
    # 현재 셀 크기 조회
    width, height = table_info.get_cell_dimensions()

    # 물리 좌표 기록
    cell_positions[cell_id] = {
        'start_x': start_x, 'end_x': start_x + width,
        'start_y': start_y, 'end_y': start_y + height,
    }

    # 오른쪽 순회 (같은 서브 행 내)
    if right_id != cell_id and right_id not in first_cols_set:
        collect_split_cells(right_id, end_x, start_y, row_end_y, visited)

    # 아래 순회 (서브 행이 더 있으면)
    if end_y < row_end_y:
        collect_split_cells(down_id, start_x, end_y, row_end_y, visited)
```

---

## 3. 그리드 레벨 생성 과정 (레벨 병합)

### 3.1 레벨 병합의 필요성

HWP에서는 셀 경계가 정확히 일치하지 않을 수 있습니다. 예를 들어:
- 셀 A의 end_x = 1000
- 셀 B의 start_x = 1002

이런 경우 두 값을 같은 레벨로 병합해야 합니다.

### 3.2 병합 알고리즘

```python
# cell_position.py
class CellPositionCalculator:
    TOLERANCE = 3  # 레벨 병합 허용 오차

    def _merge_close_levels(self, levels: List[int]) -> List[int]:
        """근접한 레벨들을 병합"""
        if not levels:
            return []

        sorted_levels = sorted(levels)
        merged = [sorted_levels[0]]

        for level in sorted_levels[1:]:
            # 이전 레벨과의 차이가 TOLERANCE 이하면 무시
            if level - merged[-1] > self.TOLERANCE:
                merged.append(level)

        return merged
```

### 3.3 예시

```
수집된 X 좌표: {0, 1000, 1002, 2000, 3000, 3001}
         │
         ▼ _merge_close_levels()
         │
병합된 X 레벨: [0, 1000, 2000, 3000]
              │    │     │     │
              ▼    ▼     ▼     ▼
             열0  열1   열2   열3
```

### 3.4 결과 데이터

```python
@dataclass
class CellPositionResult:
    cells: Dict[int, CellRange]  # list_id -> CellRange
    x_levels: List[int]          # X 레벨 목록 (정렬됨)
    y_levels: List[int]          # Y 레벨 목록 (정렬됨)
    max_row: int                 # 최대 행 인덱스
    max_col: int                 # 최대 열 인덱스
```

---

## 4. list_id -> 그리드 좌표 매핑 과정

### 4.1 좌표 인덱스 찾기

물리 좌표를 그리드 인덱스로 변환하는 두 가지 메서드:

```python
def _find_level_index(self, value: int, levels: List[int]) -> int:
    """시작 좌표에 해당하는 레벨 인덱스 반환"""
    for idx, level in enumerate(levels):
        if abs(value - level) <= self.TOLERANCE:
            return idx
    return -1

def _find_end_level_index(self, value: int, levels: List[int]) -> int:
    """끝 좌표에 해당하는 레벨 인덱스 반환

    셀의 끝 좌표가 다음 셀의 시작 좌표와 같은 경우가 많으므로,
    끝 좌표가 레벨과 일치하면 이전 레벨 인덱스를 반환
    """
    for idx, level in enumerate(levels):
        if abs(value - level) <= self.TOLERANCE:
            return max(0, idx - 1)  # 이전 레벨 인덱스

    # 일치하는 레벨이 없으면, value보다 작은 가장 큰 레벨 찾기
    for idx in range(len(levels) - 1, -1, -1):
        if levels[idx] < value - self.TOLERANCE:
            return idx
    return 0
```

### 4.2 매핑 과정

```python
# 각 셀에 대해 그리드 좌표 계산
for list_id, pos in cell_positions.items():
    start_row = _find_level_index(pos['start_y'], y_levels)
    start_col = _find_level_index(pos['start_x'], x_levels)
    end_row = _find_end_level_index(pos['end_y'], y_levels)
    end_col = _find_end_level_index(pos['end_x'], x_levels)

    cells[list_id] = CellRange(
        list_id=list_id,
        start_row=start_row,
        start_col=start_col,
        end_row=end_row,
        end_col=end_col,
        rowspan=end_row - start_row + 1,
        colspan=end_col - start_col + 1,
        # 물리 좌표도 저장
        start_x=pos['start_x'],
        start_y=pos['start_y'],
        end_x=pos['end_x'],
        end_y=pos['end_y'],
    )
```

### 4.3 예시: 병합 셀 처리

```
물리 좌표:
┌──────────────┬──────┐
│ start_x=0    │      │
│ end_x=2000   │      │  ← 2열 병합 셀
│              │      │
└──────────────┴──────┘

X 레벨: [0, 1000, 2000]
       열0  열1   열2

매핑 결과:
- start_col = _find_level_index(0, x_levels) = 0
- end_col = _find_end_level_index(2000, x_levels) = 1
- colspan = 1 - 0 + 1 = 2
```

---

## 5. 각 클래스와 주요 메서드 설명

### 5.1 TableInfo (table_info.py)

테이블 기본 정보 수집 및 셀 순회를 담당합니다.

| 메서드 | 설명 |
|--------|------|
| `collect_cells_bfs()` | BFS로 모든 셀을 순회하며 이웃 정보 수집 |
| `get_cell_dimensions()` | 현재 셀의 너비/높이 조회 (CellShape 사용) |
| `move_to_first_cell()` | 테이블 좌상단 셀로 이동 |
| `build_coordinate_map()` | (row, col) -> list_id 매핑 생성 |

주요 상수:
```python
MOVE_LEFT_OF_CELL = 100
MOVE_RIGHT_OF_CELL = 101
MOVE_UP_OF_CELL = 102
MOVE_DOWN_OF_CELL = 103
```

### 5.2 TableBoundary (table_boundary.py)

테이블 경계를 판별하고 다양한 좌표 매핑 방식을 제공합니다.

| 메서드 | 설명 |
|--------|------|
| `check_boundary_table()` | 테이블 경계 분석 (first_rows, last_cols 등) |
| `_sort_first_cols_by_position()` | first_cols를 Y 좌표 기준으로 정렬 |
| `calculate_row_widths()` | 행별 셀 너비 누적 계산 |
| `map_cell_coordinates_v4()` | 서브셀 + 위쪽 셀 기준 열 조정 좌표 매핑 |
| `build_cell_adjacency()` | 셀 인접 관계 그래프 생성 |

주요 데이터 구조:
```python
@dataclass
class TableBoundaryResult:
    table_origin: int       # 첫 번째 셀의 list_id
    table_end: int          # 마지막 셀의 list_id
    table_cell_counts: int  # 총 셀 개수
    first_rows: List[int]   # 첫 번째 행 셀들
    bottom_rows: List[int]  # 마지막 행 셀들
    first_cols: List[int]   # 각 행의 첫 번째 셀들
    last_cols: List[int]    # 각 행의 마지막 셀들
```

### 5.3 CellPositionCalculator (cell_position.py)

XY맵 기반으로 정확한 셀 위치를 계산합니다.

| 메서드 | 설명 |
|--------|------|
| `calculate()` | 모든 셀의 위치 및 범위 계산 (핵심 메서드) |
| `_merge_close_levels()` | 근접한 레벨 병합 |
| `_find_level_index()` | 시작 좌표의 레벨 인덱스 찾기 |
| `_find_end_level_index()` | 끝 좌표의 레벨 인덱스 찾기 |
| `get_cell_at()` | 특정 좌표의 셀 반환 |
| `build_coord_to_listid_map()` | (row, col) -> list_id 매핑 생성 |
| `insert_all_coordinates()` | 모든 셀에 좌표 텍스트 삽입 (디버깅용) |

### 5.4 TableExcelConverter (table_excel_converter.py)

HWP 테이블을 엑셀로 변환합니다.

| 메서드 | 설명 |
|--------|------|
| `extract_table_data()` | 테이블의 모든 셀 데이터 추출 |
| `validate_cell_positions()` | 셀 위치 계산 결과 검증 |
| `to_excel()` | HWP 테이블을 엑셀 파일로 저장 |
| `to_dict()` | 테이블을 딕셔너리로 변환 |
| `to_2d_array()` | 테이블을 2차원 배열로 변환 |

---

## 6. 데이터 흐름 다이어그램

### 6.1 전체 데이터 흐름

```
┌─────────────────────────────────────────────────────────────────┐
│                        HWP 문서                                  │
│  ┌─────────────────────────────────────────────────────────┐   │
│  │                     테이블                               │   │
│  │  ┌──────┬──────┬──────┐                                │   │
│  │  │ id=2 │ id=3 │ id=4 │  ← list_id (HWP 내부 식별자)   │   │
│  │  ├──────┼──────┴──────┤                                │   │
│  │  │ id=5 │    id=6     │  ← 병합 셀                     │   │
│  │  └──────┴─────────────┘                                │   │
│  └─────────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                    TableBoundary                                 │
│                                                                  │
│  check_boundary_table()                                          │
│  ├── first_cols: [2, 5]      # 각 행의 첫 셀                    │
│  ├── last_cols: [4, 6]       # 각 행의 끝 셀                    │
│  ├── first_rows: [2, 3, 4]   # 첫 행의 셀들                     │
│  └── bottom_rows: [5, 6]     # 마지막 행의 셀들                 │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                CellPositionCalculator.calculate()                │
│                                                                  │
│  1단계: 물리 좌표 수집                                           │
│  ┌────────────────────────────────────────────────────────┐     │
│  │ cell_positions = {                                      │     │
│  │   2: {start_x:0,    end_x:1000, start_y:0,   end_y:500}│     │
│  │   3: {start_x:1000, end_x:2000, start_y:0,   end_y:500}│     │
│  │   4: {start_x:2000, end_x:3000, start_y:0,   end_y:500}│     │
│  │   5: {start_x:0,    end_x:1000, start_y:500, end_y:1000}│    │
│  │   6: {start_x:1000, end_x:3000, start_y:500, end_y:1000}│    │
│  │ }                                                       │     │
│  └────────────────────────────────────────────────────────┘     │
│                              │                                   │
│                              ▼                                   │
│  2단계: 레벨 수집 및 병합                                        │
│  ┌────────────────────────────────────────────────────────┐     │
│  │ x_levels (raw): {0, 1000, 2000, 3000}                   │     │
│  │ y_levels (raw): {0, 500, 1000}                          │     │
│  │                                                         │     │
│  │ x_levels (merged): [0, 1000, 2000, 3000]               │     │
│  │ y_levels (merged): [0, 500, 1000]                       │     │
│  └────────────────────────────────────────────────────────┘     │
│                              │                                   │
│                              ▼                                   │
│  3단계: 그리드 좌표 매핑                                         │
│  ┌────────────────────────────────────────────────────────┐     │
│  │ cells = {                                               │     │
│  │   2: CellRange(row=0, col=0, end_row=0, end_col=0)     │     │
│  │   3: CellRange(row=0, col=1, end_row=0, end_col=1)     │     │
│  │   4: CellRange(row=0, col=2, end_row=0, end_col=2)     │     │
│  │   5: CellRange(row=1, col=0, end_row=1, end_col=0)     │     │
│  │   6: CellRange(row=1, col=1, end_row=1, end_col=2,     │     │
│  │                colspan=2)                               │     │
│  │ }                                                       │     │
│  └────────────────────────────────────────────────────────┘     │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                    TableExcelConverter                           │
│                                                                  │
│  extract_table_data()                                            │
│  ┌────────────────────────────────────────────────────────┐     │
│  │ cells_data = [                                          │     │
│  │   CellData(list_id=2, row=0, col=0, text="A1")         │     │
│  │   CellData(list_id=3, row=0, col=1, text="B1")         │     │
│  │   CellData(list_id=4, row=0, col=2, text="C1")         │     │
│  │   CellData(list_id=5, row=1, col=0, text="A2")         │     │
│  │   CellData(list_id=6, row=1, col=1, colspan=2,         │     │
│  │            text="B2-C2")                                │     │
│  │ ]                                                       │     │
│  └────────────────────────────────────────────────────────┘     │
│                              │                                   │
│                              ▼                                   │
│  to_excel("output.xlsx")                                         │
│  ┌────────────────────────────────────────────────────────┐     │
│  │ Excel 파일:                                             │     │
│  │ ┌────┬────┬────┐                                       │     │
│  │ │ A1 │ B1 │ C1 │                                       │     │
│  │ ├────┼────┴────┤                                       │     │
│  │ │ A2 │ B2-C2   │  ← B2:C2 병합                         │     │
│  │ └────┴─────────┘                                       │     │
│  └────────────────────────────────────────────────────────┘     │
└─────────────────────────────────────────────────────────────────┘
```

### 6.2 좌표 변환 상세 흐름

```
물리 좌표 (HWPUNIT)          레벨 인덱스             그리드 좌표
─────────────────────────────────────────────────────────────────

start_x = 1000  ──────────►  _find_level_index()  ───►  col = 1
                             x_levels[1] = 1000

end_x = 3000    ──────────►  _find_end_level_index() ►  end_col = 2
                             x_levels[3] = 3000
                             → 반환: 3 - 1 = 2

colspan = end_col - col + 1 = 2 - 1 + 1 = 2
```

### 6.3 검증 흐름

```
CellPositionResult
       │
       ▼
validate_cell_positions()
       │
       ├──► 중복 좌표 검사
       │    for each cell:
       │        for (r,c) in covered_coords:
       │            if (r,c) in coord_to_cell:
       │                → 중복 발견!
       │
       ├──► 빈 좌표 검사
       │    for r in range(max_row + 1):
       │        for c in range(max_col + 1):
       │            if (r,c) not in coord_to_cell:
       │                → 빈 좌표 발견!
       │
       └──► 커버리지 계산
            coverage = covered_coords / total_coords
```

---

## 7. 사용 예시

### 7.1 셀 위치 계산

```python
from table.cell_position import CellPositionCalculator

hwp = get_hwp_instance()
calc = CellPositionCalculator(hwp, debug=True)

# 셀 위치 계산
result = calc.calculate()

# 결과 출력
calc.print_summary(result)

# 특정 좌표의 셀 찾기
cell = calc.get_cell_at(result, row=1, col=2)
print(f"list_id: {cell.list_id}, span: {cell.rowspan}x{cell.colspan}")

# 좌표 -> list_id 매핑
coord_map = calc.build_coord_to_listid_map(result)
list_id = coord_map[(0, 0)]  # (0,0) 좌표의 셀 ID
```

### 7.2 엑셀 변환

```python
from table.table_excel_converter import TableExcelConverter

converter = TableExcelConverter(hwp, debug=True)

# 검증
validation = converter.validate_cell_positions()
if validation['valid']:
    # 엑셀로 저장
    converter.to_excel("output.xlsx", with_text=True)
```

### 7.3 경계 분석

```python
from table.table_boundary import TableBoundary

boundary = TableBoundary(hwp, debug=True)

# 경계 정보 수집
result = boundary.check_boundary_table()

print(f"첫 번째 열 셀들: {result.first_cols}")
print(f"마지막 열 셀들: {result.last_cols}")
print(f"총 셀 수: {result.table_cell_counts}")
```

---

## 8. 주의사항

### 8.1 TOLERANCE 값

`CellPositionCalculator.TOLERANCE = 3`은 레벨 병합 시 허용 오차입니다. HWP 문서에 따라 이 값을 조정해야 할 수 있습니다.

### 8.2 분할 셀 처리

행 내에서 세로로 분할된 셀(서브셀)은 `collect_split_cells()` 재귀 함수로 처리됩니다. 복잡한 테이블의 경우 이 로직이 중요합니다.

### 8.3 성능 고려사항

- `max_cells` 파라미터로 처리할 최대 셀 수를 제한할 수 있습니다
- 대형 테이블의 경우 `debug=False`로 설정하여 로그 출력을 줄이세요
