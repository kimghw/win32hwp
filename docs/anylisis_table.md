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

## 셀 경계 판별 함수

### move_down_left_right: 하/좌/우 각각 이동 결과

특정 셀에서 하, 좌, 우 방향으로 **각각** 이동했을 때 `list_id`와 `has_tbl`을 반환한다.

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
    def move_and_check(action):
        hwp.SetPos(target_list_id, 0, 0)
        hwp.HAction.Run(action)
        list_id, _, _ = hwp.GetPos()
        parent = hwp.ParentCtrl
        has_tbl = parent and parent.CtrlID == "tbl"
        return (list_id, has_tbl)

    return {
        'down': move_and_check("MoveDown"),
        'left': move_and_check("MoveLeft"),
        'right': move_and_check("MoveRight")
    }
```

---

### check_first_row_cell: 첫 번째 행 여부 판별

특정 셀에서 위로 이동 시 `has_tbl`이 `False`면 `True` 반환.

```python
def check_first_row_cell(hwp, target_list_id):
    """위로 이동 시 tbl 밖이면 True (첫 번째 행)"""
    hwp.SetPos(target_list_id, 0, 0)
    hwp.HAction.Run("MoveUp")
    parent = hwp.ParentCtrl
    has_tbl = parent and parent.CtrlID == "tbl"
    return not has_tbl
```

---

### check_bottom_row_cell: 마지막 행 여부 판별

특정 셀에서 아래로 이동 시 `has_tbl`이 `False`면 `True` 반환.

```python
def check_bottom_row_cell(hwp, target_list_id):
    """아래로 이동 시 tbl 밖이면 True (마지막 행)"""
    hwp.SetPos(target_list_id, 0, 0)
    hwp.HAction.Run("MoveDown")
    parent = hwp.ParentCtrl
    has_tbl = parent and parent.CtrlID == "tbl"
    return not has_tbl
```

---

### move_up_right_down: 상/우/하 각각 이동 결과

특정 셀에서 상, 우, 하 방향으로 **각각** 이동했을 때 `list_id`와 `has_tbl`을 반환한다.

```python
def move_up_right_down(hwp, target_list_id):
    """
    특정 셀에서 상/우/하 각각 이동 시 list_id와 has_tbl 반환

    Returns:
        dict: {
            'up': (list_id, has_tbl),
            'right': (list_id, has_tbl),
            'down': (list_id, has_tbl)
        }
    """
    def move_and_check(action):
        hwp.SetPos(target_list_id, 0, 0)
        hwp.HAction.Run(action)
        list_id, _, _ = hwp.GetPos()
        parent = hwp.ParentCtrl
        has_tbl = parent and parent.CtrlID == "tbl"
        return (list_id, has_tbl)

    return {
        'up': move_and_check("MoveUp"),
        'right': move_and_check("MoveRight"),
        'down': move_and_check("MoveDown")
    }
```

---

## 함수 요약

| 함수명 | 설명 | 반환값 |
|--------|------|--------|
| `check_first_cols` | 하/좌/우 각각 이동 결과 | `dict` {방향: (list_id, has_tbl)} |
| `check_first_rows` | 위로 이동 시 tbl 밖이면 True | `bool` |
| `check_bottom_rows` | 아래로 이동 시 tbl 밖이면 True | `bool` |
| `move_vertical_directions` | 상/우/하 각각 이동 결과 | `dict` {방향: (list_id, has_tbl)} |
