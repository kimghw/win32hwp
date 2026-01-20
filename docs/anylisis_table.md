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

## 셀 경계 판별 함수

### table_origin : 테이블의 첫번째 셀의 list_id
### table_end : last_cols의 마지막 셀의 list_id
### table_cell_counts : 테이블의 총 셀 개수

---

### move_down_left_right: 하/좌/우 각각 이동 결과

특정 셀에서 하, 좌, 우 방향으로 **셀 단위** 이동했을 때 `list_id`와 `has_tbl`을 반환한다.

```python
def move_down_left_right(hwp, target_list_id):
    """
    특정 셀에서 하/좌/우 각각 이동 시 list_id와 has_tbl 반환

    Returns:
        dict: {
            'down': (list_id, has_tbl),
            'left': (list_id, has_tbl),
            'right': (list_id, has_tbl)
        }
    """
    MOVE_DOWN_OF_CELL = 103
    MOVE_LEFT_OF_CELL = 100
    MOVE_RIGHT_OF_CELL = 101

    def move_and_check(move_const):
        hwp.SetPos(target_list_id, 0, 0)
        hwp.MovePos(move_const, 0, 0)
        list_id, _, _ = hwp.GetPos()
        parent = hwp.ParentCtrl
        has_tbl = parent and parent.CtrlID == "tbl"
        return (list_id, has_tbl)

    return {
        'down': move_and_check(MOVE_DOWN_OF_CELL),
        'left': move_and_check(MOVE_LEFT_OF_CELL),
        'right': move_and_check(MOVE_RIGHT_OF_CELL)
    }
```

---

### check_first_row_cell: 첫 번째 행 여부 판별

특정 셀에서 위로 이동 시 `has_tbl`이 `False`이거나 같은 셀에 머물면 `True` 반환.

```python
def check_first_row_cell(hwp, target_list_id):
    """위로 이동 시 tbl 밖이거나 같은 셀에 머물면 True (첫 번째 행)"""
    hwp.SetPos(target_list_id, 0, 0)
    hwp.HAction.Run("MoveUp")
    new_list_id, _, _ = hwp.GetPos()
    parent = hwp.ParentCtrl
    has_tbl = parent and parent.CtrlID == "tbl"
    return not has_tbl or new_list_id == target_list_id
```

---

### check_bottom_row_cell: 마지막 행 여부 판별

특정 셀에서 아래로 이동 시 `has_tbl`이 `False`이거나 같은 셀에 머물면 `True` 반환.

```python
def check_bottom_row_cell(hwp, target_list_id):
    """아래로 이동 시 tbl 밖이거나 같은 셀에 머물면 True (마지막 행)"""
    hwp.SetPos(target_list_id, 0, 0)
    hwp.HAction.Run("MoveDown")
    new_list_id, _, _ = hwp.GetPos()
    parent = hwp.ParentCtrl
    has_tbl = parent and parent.CtrlID == "tbl"
    return not has_tbl or new_list_id == target_list_id
```

---

### move_up_right_down: 하→우→상 커서 이동으로 last_col 판별

특정 셀에서 **하 → 우 → 상** 순서로 커서 이동 후, 시작점으로 돌아오면 같은 열(last_col).

```python
def move_up_right_down(hwp, target_list_id):
    """
    특정 셀에서 하 → 우 → 상 순서로 커서 이동 후 last_col 여부 판별

    Returns:
        tuple: (down_id, is_last_col)
            - down_id: 아래로 이동한 셀의 list_id
            - is_last_col: 상으로 돌아온 셀이 시작점과 같으면 True (같은 열)
    """
    # 시작점에서 하로 이동
    hwp.SetPos(target_list_id, 0, 0)
    hwp.HAction.Run("MoveDown")
    down_id, _, _ = hwp.GetPos()

    # 우로 이동
    hwp.HAction.Run("MoveRight")

    # 상으로 이동
    hwp.HAction.Run("MoveUp")
    up_id, _, _ = hwp.GetPos()

    # 시작점과 같으면 같은 열 -> last_col
    is_last_col = (up_id == target_list_id)

    return (down_id, is_last_col)
```

---

## 함수 요약

| 함수명 | 설명 | 반환값 |
|--------|------|--------|
| `move_down_left_right` | 하/좌/우 셀 단위 이동 결과 | `dict` {방향: (list_id, has_tbl)} |
| `check_first_row_cell` | 위로 이동 시 tbl 밖이거나 같은 셀이면 True | `bool` |
| `check_bottom_row_cell` | 아래로 이동 시 tbl 밖이거나 같은 셀이면 True | `bool` |
| `move_up_right_down` | 하→우→상 커서 이동, last_col 판별 | `tuple` (down_id, is_last_col) |

---

## check_boundary_table

테이블의 모든 셀을 순회하면서 다음의 순서대로 실행한다.

### 1. table_origin, table_cell_counts 계산

테이블의 첫 번째 셀의 `list_id`와 총 셀 개수를 구한다.

### 2. first_rows와 bottom_rows 계산

모든 셀을 순회하면서 첫 번째 행과 마지막 행에 속한 셀들을 리스트로 저장한다.

| 변수명 | 설명 | 사용 함수 |
|--------|------|-----------|
| `first_rows` | 첫 번째 행에 속한 셀들 (리스트) | `check_first_row_cell` |
| `bottom_rows` | 마지막 행에 속한 셀들 (리스트) | `check_bottom_row_cell` |

### 3. first_cols 계산

`table_origin`에서 아래로 내려가면서 첫 번째 컬럼 셀들을 수집한다.

### 4. last_cols 계산

`first_rows`의 마지막 셀에서 시작하여 `move_up_right_down`을 활용해 마지막 컬럼 셀들을 수집한다.

- 시작 셀은 무조건 `last_cols`에 포함
- `is_last_col`이 True면 `down_id`를 `last_cols`에 추가
- `is_last_col`이 False여도 마지막 행이면 `last_cols`에 추가

### 5. table_end 계산

`last_cols`의 마지막 `list_id`가 `table_end`이다.

| 변수명 | 설명 | 계산 방법 |
|--------|------|-----------|
| `first_cols` | 첫 번째 컬럼에 속한 셀들 | table_origin에서 아래로 순회 |
| `last_cols` | 마지막 컬럼에 속한 셀들 | first_rows 마지막에서 move_up_right_down 활용 |
| `table_end` | 테이블 끝점 | last_cols의 마지막 값 |
