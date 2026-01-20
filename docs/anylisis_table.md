# 테이블 커서 이동 분석 및 list_id 매핑

## 개요

테이블의 모든 `list_id`를 추출하고, 커서 이동을 통해 테이블 구조를 파악하는 방법을 정리한다.

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
| `table_end` | 테이블 끝점 (마지막 행, 마지막 열) |

### 행/열 그룹 용어

| 용어 | 설명 |
|------|------|
| `first_rows` | 첫 번째 행에 속한 셀들 (리스트) |
| `last_rows` | 마지막 행에 속한 셀들 (리스트) |
| `first_cols` | 첫 번째 컬럼에 속한 셀들 (리스트) |
| `last_cols` | 마지막 컬럼에 속한 셀들 (리스트) |

### 특정 위치 용어

| 용어 | 설명 |
|------|------|
| `first_col` | 특정 row의 첫 번째 컬럼 값 |
| `last_col` | 특정 row의 마지막 컬럼 값 |
| `first_row` | 특정 컬럼의 첫 번째 행 값 |
| `last_row` | 특정 컬럼의 마지막 행 값 |

### 셀 병합 용어

| 용어 | 설명 |
|------|------|
| `merged_cell` | 병합된 셀 (여러 행/열을 합친 셀) |
| `single_cell` | 병합되지 않은 단일 셀 |
| `row_span` | 행 방향 병합 크기 (세로로 몇 칸 차지) |
| `col_span` | 열 방향 병합 크기 (가로로 몇 칸 차지) |

---

## 1단계: 모든 셀 수집 및 first_rows, last_rows 계산

### 알고리즘

모든 셀을 순회하면서 `MOVE_TOP_OF_CELL`, `MOVE_BOTTOM_OF_CELL`을 사용하여 경계 셀을 추출한다.

```python
all_cells = collect_all_cells()  # 테이블의 모든 list_id 수집

first_rows = set()
last_rows = set()

for cell_id in all_cells:
    hwp.SetPos(cell_id, 0, 0)

    # 열의 맨 위 셀 → first_rows
    hwp.MovePos(MOVE_TOP_OF_CELL, 0, 0)
    first_rows.add(hwp.GetPos()[0])

    # 열의 맨 아래 셀 → last_rows
    hwp.SetPos(cell_id, 0, 0)
    hwp.MovePos(MOVE_BOTTOM_OF_CELL, 0, 0)
    last_rows.add(hwp.GetPos()[0])
```

### 결과

- `first_rows`: 첫 번째 행에 속한 모든 셀의 list_id
- `last_rows`: 마지막 행에 속한 모든 셀의 list_id

---

## 2단계: first_cols, last_cols 계산

### first_cols 알고리즘

`table_origin`부터 아래로 내려가면서 하→좌→상 이동을 통해 첫 번째 컬럼 셀을 판별한다.

**핵심 로직:**
1. `table_origin`에서 시작하여 아래로 내려감
2. 각 셀에서 하→좌→상 이동 후 시작 위치와 비교
3. 시작 위치와 종착 위치가 같으면, 좌 이동 시점의 셀이 `first_col`
4. 시작 위치와 종착 위치가 다르면, 하 이동 시점의 셀이 `first_col`

```python
first_cols = set()

# table_origin에서 시작
current_cell = table_origin
first_cols.add(current_cell)  # table_origin은 항상 first_cols

while True:
    hwp.SetPos(current_cell, 0, 0)

    # 하 이동 시도
    hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
    lower_cell = hwp.GetPos()[0]

    # 더 이상 아래로 갈 수 없으면 종료
    if lower_cell == current_cell:
        break

    # 아랫셀에서 시작하여 하→좌→상 이동으로 first_col 판별
    start_cell = lower_cell

    while True:
        hwp.SetPos(start_cell, 0, 0)

        # 좌 이동
        hwp.MovePos(MOVE_LEFT_OF_CELL, 0, 0)
        left_cell = hwp.GetPos()[0]

        # 좌 이동 불가 (이미 첫 번째 컬럼)
        if left_cell == start_cell:
            first_cols.add(start_cell)
            break

        # 좌 이동 후 상 이동
        hwp.MovePos(MOVE_UP_OF_CELL, 0, 0)
        up_cell = hwp.GetPos()[0]

        # 다시 하 이동하여 원래 위치로 돌아오는지 확인
        hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
        back_cell = hwp.GetPos()[0]

        # 시작 위치와 종착 위치가 같으면 → left_cell이 first_col
        if back_cell == start_cell:
            first_cols.add(left_cell)
            break

        # 다르면 → start_cell(하 이동 시점)이 first_col
        first_cols.add(start_cell)
        break

    # 다음 행으로 이동
    current_cell = lower_cell
```

### last_cols 알고리즘

`table_end`부터 위로 올라가면서 상/우/하 이동을 통해 마지막 컬럼 셀을 판별한다.

**핵심 로직:**
1. `table_end`에서 시작하여 위로 올라감
2. 각 셀에서 상→우→하 이동 후 시작 위치와 비교
3. 시작 위치와 종착 위치가 다르면, 우 이동 시점의 셀이 해당 행의 마지막 컬럼이 아님
4. 시작 위치와 종착 위치가 같으면, 현재 셀이 해당 행의 마지막 컬럼

```python
last_cols = set()

# table_end에서 시작
current_cell = table_end
last_cols.add(current_cell)  # table_end는 항상 last_cols

while True:
    hwp.SetPos(current_cell, 0, 0)

    # 상 이동 시도
    hwp.MovePos(MOVE_UP_OF_CELL, 0, 0)
    upper_cell = hwp.GetPos()[0]

    # 더 이상 위로 갈 수 없으면 종료
    if upper_cell == current_cell:
        break

    # 윗셀에서 시작하여 상→우→하 이동으로 last_col 판별
    start_cell = upper_cell

    while True:
        hwp.SetPos(start_cell, 0, 0)

        # 우 이동
        hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
        right_cell = hwp.GetPos()[0]

        # 우 이동 불가 (이미 마지막 컬럼)
        if right_cell == start_cell:
            last_cols.add(start_cell)
            break

        # 우 이동 후 하 이동
        hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
        down_cell = hwp.GetPos()[0]

        # 다시 상 이동하여 원래 위치로 돌아오는지 확인
        hwp.MovePos(MOVE_UP_OF_CELL, 0, 0)
        back_cell = hwp.GetPos()[0]

        # 시작 위치와 종착 위치가 같으면 → start_cell이 last_col
        if back_cell == start_cell:
            last_cols.add(start_cell)
            break

        # 다르면 → 우측 셀로 이동하여 계속 탐색
        start_cell = right_cell

    # 다음 행으로 이동
    current_cell = upper_cell
```

### 결과

- `first_cols`: 각 행의 첫 번째 컬럼에 해당하는 셀들
- `last_cols`: 각 행의 마지막 컬럼에 해당하는 셀들 (상→우→하 이동 판별)

---

## MovePos 상수

| 상수 | 값 | 설명 |
|------|-----|------|
| `MOVE_LEFT_OF_CELL` | 100 | 좌측 셀로 이동 |
| `MOVE_RIGHT_OF_CELL` | 101 | 우측 셀로 이동 |
| `MOVE_UP_OF_CELL` | 102 | 상단 셀로 이동 |
| `MOVE_DOWN_OF_CELL` | 103 | 하단 셀로 이동 |
| `MOVE_START_OF_CELL` | 104 | 행의 시작 셀로 이동 |
| `MOVE_END_OF_CELL` | 105 | 행의 끝 셀로 이동 |
| `MOVE_TOP_OF_CELL` | 106 | 열의 시작(맨 위) 셀로 이동 |
| `MOVE_BOTTOM_OF_CELL` | 107 | 열의 끝(맨 아래) 셀로 이동 |

---

## table_origin, table_end 계산

```python
# table_origin: 행의 시작 → 열의 시작
hwp.MovePos(MOVE_START_OF_CELL, 0, 0)
hwp.MovePos(MOVE_TOP_OF_CELL, 0, 0)
table_origin = hwp.GetPos()[0]

# table_end: 행의 끝 → 열의 끝
hwp.MovePos(MOVE_END_OF_CELL, 0, 0)
hwp.MovePos(MOVE_BOTTOM_OF_CELL, 0, 0)
table_end = hwp.GetPos()[0]
```

---

## 전체 순서

1. **모든 셀 수집** (`_collect_all_cells`)
2. **first_rows, last_rows 계산** (MOVE_TOP_OF_CELL, MOVE_BOTTOM_OF_CELL)
3. **table_origin, table_end 계산**
4. **first_cols, last_cols 계산** (MOVE_START_OF_CELL, MOVE_END_OF_CELL)

.
