# HWP 테이블 셀 좌표 매핑 프로세스

이 문서는 HWP 테이블의 셀을 그리드 좌표로 매핑하는 전체 프로세스를 설명합니다.

---

## 1. 전체 아키텍처 개요

### 1.1 모듈 구조

```
table/
├── table_info.py           # 테이블 기본 정보 및 셀 순회
├── table_boundary.py       # 테이블 경계 판별 및 좌표 매핑
├── cell_position.py        # xline/yline 기반 셀 위치 계산 (핵심)
└── table_excel_converter.py # 엑셀 변환
```

### 1.2 핵심 알고리즘: xline/yline 기반 그리드 생성

**권장 방식**: `calculate_grid()` (= `calculate()`)

```
1. 모든 셀 순회 → 물리적 좌표 수집 (start_x, end_x, start_y, end_y)
2. 수집된 x 좌표들에서 unique 값 추출 → xline (열 경계선)
3. 수집된 y 좌표들에서 unique 값 추출 → yline (행 경계선)
4. xline × yline 으로 그리드 생성 → 각 셀을 그리드에 매핑
```

**장점**:
- 순회 순서에 의존하지 않음
- 복잡한 병합 셀도 정확하게 처리
- 물리 좌표 기반이므로 오차 없음

### 1.3 핵심 개념

| 개념 | 설명 |
|------|------|
| `list_id` | HWP 내부에서 각 셀을 식별하는 고유 ID |
| `물리 좌표` | 셀의 실제 위치와 크기 (HWPUNIT 단위) |
| `그리드 좌표` | 엑셀 스타일의 (row, col) 좌표 |
| `X/Y 레벨` | 물리 좌표를 정규화한 그리드 경계선 |

### 1.4 처리 흐름

```
HWP 테이블
    │
    ▼
┌─────────────────────────────────┐
│  1. 경계 분석                    │  (TableBoundary)
│     - first_cols, last_cols     │
│     - 행 단위 순회 준비          │
└─────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────┐
│  2. 물리 좌표 수집               │  (CellPositionCalculator.calculate_grid)
│     - 모든 셀의 start/end x,y   │
│     - 좌표 집합 생성            │
└─────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────┐
│  3. xline/yline 생성            │  (CellPositionCalculator)
│     - x 좌표 → xline (열 경계)  │
│     - y 좌표 → yline (행 경계)  │
│     - 근접 레벨 병합            │
└─────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────┐
│  4. 그리드 매핑                  │  (CellPositionCalculator)
│     - 물리 좌표 → 그리드 인덱스  │
│     - rowspan, colspan 계산     │
└─────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────┐
│  5. 빈 위치 보완                 │  (_fix_empty_positions)
│     - 그리드 누락 좌표 탐지      │
│     - 인접 셀 커서 이동으로 보완 │
└─────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────┐
│  6. 중복 위치 해결               │  (_fix_overlaps)
│     - 좌표 중복 탐지             │
│     - 커서 이동 기반 범위 조정   │
└─────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────┐
│  7. 엑셀 변환                    │  (TableExcelConverter)
│     - 셀 병합 적용              │
│     - 텍스트 추출               │
└─────────────────────────────────┘
```

---

## 2. calculate_grid() 알고리즘 (xline/yline 방식)

### 2.1 개요

`calculate_grid()`는 모든 셀의 물리적 좌표를 수집한 후, xline(열 경계선)과 yline(행 경계선)을 추출하여 그리드를 생성합니다.

**이 방식의 핵심**: 순회 순서에 의존하지 않고, 물리 좌표만으로 정확한 그리드 생성

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

### 2.3 알고리즘 상세

`CellPositionCalculator.calculate_grid()` 메서드:

```
1단계: 물리 좌표 수집
───────────────────────────────────────────────────
for row_start in first_cols:
    │
    ├── row_start_y = cumulative_y
    ├── row_height = 행 시작 셀의 높이
    │
    └── BFS로 행 내 모든 셀 순회:
        for each cell in row:
            │
            ├── 셀 크기 조회: width, height
            │
            ├── 물리 좌표 계산:
            │   ├── start_x, end_x = start_x + width
            │   └── start_y, end_y = start_y + height
            │
            └── 좌표 집합에 추가:
                ├── all_x_coords.add(start_x)
                ├── all_x_coords.add(end_x)
                ├── all_y_coords.add(start_y)
                └── all_y_coords.add(end_y)

2단계: xline/yline 생성
───────────────────────────────────────────────────
x_levels = _merge_close_levels(all_x_coords)  # 열 경계선
y_levels = _merge_close_levels(all_y_coords)  # 행 경계선

3단계: 그리드 매핑
───────────────────────────────────────────────────
for each cell:
    start_col = find_level_index(start_x, x_levels)
    end_col = find_end_level_index(end_x, x_levels)
    start_row = find_level_index(start_y, y_levels)
    end_row = find_end_level_index(end_y, y_levels)

    colspan = end_col - start_col + 1
    rowspan = end_row - start_row + 1

4단계: 빈 위치 보완 (_fix_empty_positions)
───────────────────────────────────────────────────
그리드에 누락된 좌표가 있으면 인접 셀 커서 이동으로 보완:

for each empty (row, col):
    - 인접 셀 찾기 (왼쪽/오른쪽/위/아래)
    - 인접 셀에서 반대 방향으로 커서 이동
    - 이동한 셀의 범위를 빈 위치까지 확장

5단계: 중복 위치 해결 (_fix_overlaps)
───────────────────────────────────────────────────
같은 좌표를 여러 셀이 점유하면 커서 이동으로 해결:

for each overlap (row, col) with [id1, id2]:
    - id1 → 아래 이동 → id2 이면 id1이 위에 있음
    - id1의 end_row를 id2.start_row - 1로 조정
```

### 2.4 행 내 셀 순회 (BFS)

행 내에서 오른쪽/아래 방향으로 BFS 순회:

```python
queue = deque([(row_start, 0, row_start_y)])

while queue:
    current_id, start_x, start_y = queue.popleft()

    # 오른쪽 이동
    if right_id not in visited and right_id not in first_cols_set:
        queue.append((right_id, end_x, start_y))

    # 아래 이동 (행 내 분할 셀 처리)
    if end_y < row_end_y and down_id not in first_cols_set:
        queue.append((down_id, start_x, end_y))
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

## 4. 그리드 매핑 보정 (빈 위치/중복 해결)

### 4.1 보정이 필요한 이유

xline/yline 기반 매핑은 물리 좌표에 의존하므로, 복잡한 병합 셀에서 다음 문제가 발생할 수 있습니다:

| 문제 | 원인 | 예시 |
|------|------|------|
| **빈 위치** | 병합 셀의 물리 좌표가 일부 그리드 영역을 커버하지 못함 | 3행 병합 셀인데 2행만 매핑됨 |
| **중복 위치** | 두 셀의 물리 좌표가 같은 그리드 영역을 점유 | (5,0)에 셀 A, B 모두 매핑됨 |

### 4.2 빈 위치 보완 (`_fix_empty_positions`)

**알고리즘:**

```
1. 그리드 점유 맵 생성: (row, col) → list_id
2. 빈 위치 탐색: 점유되지 않은 (row, col) 수집
3. 각 빈 위치에서:
   ├── 인접 셀 찾기 (왼쪽/오른쪽/위/아래)
   ├── 인접 셀에서 반대 방향으로 커서 이동
   │   예: 왼쪽 인접 셀 → 오른쪽으로 이동 (MOVE_RIGHT_OF_CELL)
   └── 이동한 셀이 존재하면 해당 셀의 범위를 빈 위치까지 확장
```

**예시:**

```
초기 매핑 결과:
     col 0   col 1   col 2
row 0:  [2]     [3]     [4]
row 1:  [5]     [6]     [6]
row 2:  [5]      .      [7]    ← (2,1)이 빈 위치

보정 과정:
1. (2,1)의 왼쪽 인접 셀 = [5]
2. [5]에서 오른쪽으로 커서 이동 → [6] 발견
3. [6]의 범위를 row 2까지 확장

보정 후:
     col 0   col 1   col 2
row 0:  [2]     [3]     [4]
row 1:  [5]     [6]     [6]
row 2:  [5]     [6]     [7]    ← [6]이 row 1~2로 확장됨
```

### 4.3 중복 위치 해결 (`_fix_overlaps`)

**알고리즘:**

```
1. 그리드 점유 맵 생성: (row, col) → [list_ids]
2. 중복 위치 탐색: 2개 이상의 셀이 점유한 좌표 수집
3. 각 중복 셀 쌍에서:
   ├── 커서 이동으로 인접 관계 확인
   │   id1 → 아래 이동 → id2 이면 id1이 위, id2가 아래
   └── 위 셀의 end_row를 아래 셀의 start_row - 1로 조정
```

**예시:**

```
초기 매핑 결과:
     col 0   col 1
row 0:  [2]     [3]
row 1:  [4]     [3]    ← (1,1)에 [3]과 [5] 중복
row 2:  [4]     [5]

중복 확인:
- (1,1)에 셀 [3]과 [5]가 모두 매핑됨
- [3] → 아래 이동 → [5] 도착 (인접 관계 확인)
- [3]의 end_row를 [5].start_row - 1 = 1로 조정

보정 후:
     col 0   col 1
row 0:  [2]     [3]
row 1:  [4]     [5]    ← [3]은 row 0만, [5]는 row 1~2
row 2:  [4]     [5]
```

---

## 5. list_id -> 그리드 좌표 매핑 과정

### 5.1 좌표 인덱스 찾기

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

### 5.2 매핑 과정

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

### 5.3 예시: 병합 셀 처리

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

## 6. 각 클래스와 주요 메서드 설명

### 6.1 TableInfo (table_info.py)

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

### 6.2 TableBoundary (table_boundary.py)

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

### 6.3 CellPositionCalculator (cell_position.py)

xline/yline 기반으로 정확한 셀 위치를 계산합니다.

| 메서드 | 설명 |
|--------|------|
| `calculate()` | **메인 메서드** - calculate_grid() 호출 |
| `calculate_grid()` | xline/yline 기반 그리드 생성 (권장) |
| `_calculate_bfs()` | [DEPRECATED] BFS 방식 (사용하지 않음) |
| `_calculate_legacy()` | [DEPRECATED] 레거시 방식 (사용하지 않음) |
| `_merge_close_levels()` | 근접한 레벨 병합 |
| `_find_level_index()` | 시작 좌표의 레벨 인덱스 찾기 |
| `_find_end_level_index()` | 끝 좌표의 레벨 인덱스 찾기 |
| `_fix_empty_positions()` | 빈 위치 보완 (인접 셀 커서 이동 기반) |
| `_fix_overlaps()` | 중복 위치 해결 (커서 이동 기반 인접 관계 확인) |
| `get_cell_at()` | 특정 좌표의 셀 반환 |
| `build_coord_to_listid_map()` | (row, col) -> list_id 매핑 생성 |
| `insert_all_coordinates()` | 모든 셀에 좌표 텍스트 삽입 (디버깅용) |

### 6.4 TableExcelConverter (table_excel_converter.py)

HWP 테이블을 엑셀로 변환합니다.

| 메서드 | 설명 |
|--------|------|
| `extract_table_data()` | 테이블의 모든 셀 데이터 추출 |
| `validate_cell_positions()` | 셀 위치 계산 결과 검증 |
| `to_excel()` | HWP 테이블을 엑셀 파일로 저장 |
| `to_dict()` | 테이블을 딕셔너리로 변환 |
| `to_2d_array()` | 테이블을 2차원 배열로 변환 |

---

## 7. 데이터 흐름 다이어그램

### 7.1 전체 데이터 흐름

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

### 7.2 좌표 변환 상세 흐름

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

### 7.3 검증 흐름

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

## 8. 사용 예시

### 8.1 셀 위치 계산

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

### 8.2 엑셀 변환

```python
from table.table_excel_converter import TableExcelConverter

converter = TableExcelConverter(hwp, debug=True)

# 검증
validation = converter.validate_cell_positions()
if validation['valid']:
    # 엑셀로 저장
    converter.to_excel("output.xlsx", with_text=True)
```

### 8.3 경계 분석

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

## 9. 주의사항

### 9.1 TOLERANCE 값

`CellPositionCalculator.TOLERANCE = 3`은 레벨 병합 시 허용 오차입니다. HWP 문서에 따라 이 값을 조정해야 할 수 있습니다.

### 9.2 분할 셀 처리

행 내에서 세로로 분할된 셀(서브셀)은 `collect_split_cells()` 재귀 함수로 처리됩니다. 복잡한 테이블의 경우 이 로직이 중요합니다.

### 9.3 성능 고려사항

- `max_cells` 파라미터로 처리할 최대 셀 수를 제한할 수 있습니다
- 대형 테이블의 경우 `debug=False`로 설정하여 로그 출력을 줄이세요

