# table_info.py - 테이블 좌표 매핑 모듈

## 개요

HWP 테이블의 (row, col) 좌표를 list_id에 매핑하는 모듈.
병합된 셀(colspan/rowspan)을 감지하고 좌표 매핑을 생성한다.

## 주요 기능

| 메서드 | 설명 |
|--------|------|
| `build_coordinate_map()` | **(메인)** 테이블 좌표 → list_id 매핑 반환 |
| `collect_cells_bfs()` | 행 우선 순회로 모든 셀 정보 수집 |
| `get_table_size()` | 테이블 행/열 수 계산 (병합 고려) |
| `is_in_table()` | 커서가 테이블 내부인지 확인 |
| `move_to_first_cell()` | 테이블 첫 번째 셀(좌상단)로 이동 |
| `find_all_tables()` | 문서 내 모든 테이블 찾기 |
| `select_table(index)` | 특정 번호 테이블 선택 |
| `enter_table(index)` | 특정 번호 테이블 첫 셀로 진입 |
| `has_caption(index)` | 테이블에 캡션이 있는지 확인 |
| `get_table_caption(index)` | 테이블 캡션 텍스트 가져오기 |
| `get_all_table_captions()` | 모든 테이블 캡션 가져오기 |

## 사용 예시

```python
from table_info import TableInfo
from cursor_utils import get_hwp_instance

hwp = get_hwp_instance()
table = TableInfo(hwp)

# 문서 내 모든 테이블 찾기
tables = table.find_all_tables()
for t in tables:
    print(f"테이블 {t['num']}: 첫 셀 list_id={t['first_cell_list_id']}")

# 특정 테이블로 진입 (0번 테이블)
table.enter_table(0)

# 좌표 매핑 가져오기
coord_map = table.build_coordinate_map()
# {(0,0): 2, (0,1): 3, (0,2): 4, (0,3): 4, ...}
# colspan 셀은 여러 좌표가 같은 list_id

# 테이블 크기
size = table.get_table_size()
print(f"{size['rows']}행 x {size['cols']}열")
```

## 사용된 HWP API

### 1. MovePos (셀 이동)

```python
hwp.MovePos(move_type, para, pos)
```

| 상수 | 값 | 설명 |
|------|-----|------|
| `MOVE_LEFT_OF_CELL` | 100 | 왼쪽 셀로 이동 |
| `MOVE_RIGHT_OF_CELL` | 101 | 오른쪽 셀로 이동 |
| `MOVE_UP_OF_CELL` | 102 | 위쪽 셀로 이동 |
| `MOVE_DOWN_OF_CELL` | 103 | 아래쪽 셀로 이동 |
| `MOVE_START_OF_CELL` | 104 | 행의 시작 셀로 이동 |
| `MOVE_END_OF_CELL` | 105 | 행의 끝 셀로 이동 |
| `MOVE_TOP_OF_CELL` | 106 | 열의 시작(맨 위) 셀로 이동 |
| `MOVE_BOTTOM_OF_CELL` | 107 | 열의 끝(맨 아래) 셀로 이동 |

**반환값:** 이동 성공 시 True, 실패 시 False

### 2. GetPos / SetPos (위치 조회/설정)

```python
# 현재 위치 조회
pos = hwp.GetPos()
# 반환: (list_id, para_id, char_pos)

# 위치 이동
hwp.SetPos(list_id, para_id, char_pos)
```

**list_id 의미:**
- 0 = 본문
- 1+ = 테이블 셀, 캡션 등 (각 텍스트 컨테이너마다 고유)

### 3. CellShape (셀 크기 조회)

```python
cell_shape = hwp.CellShape
cell = cell_shape.Item("Cell")
width = cell.Item("Width")   # HWPUNIT
height = cell.Item("Height") # HWPUNIT
```

**HWPUNIT:** 7200 = 1인치 = 25.4mm

### 4. CreateAction (테이블 내부 확인)

```python
act = hwp.CreateAction("TableCellBlock")
pset = act.CreateSet()
result = act.Execute(pset)  # True면 테이블 내부
```

### 5. 컨트롤 순회 (테이블 찾기)

```python
# 문서 내 모든 컨트롤 순회
ctrl = hwp.HeadCtrl
while ctrl:
    if ctrl.CtrlID == "tbl":  # 테이블 발견
        # 테이블 처리
        pass
    ctrl = ctrl.Next
```

| 속성/메서드 | 설명 |
|-------------|------|
| `hwp.HeadCtrl` | 문서 첫 번째 컨트롤 |
| `hwp.LastCtrl` | 문서 마지막 컨트롤 |
| `ctrl.Next` | 다음 컨트롤 |
| `ctrl.Prev` | 이전 컨트롤 |
| `ctrl.CtrlID` | 컨트롤 종류 ("tbl", "fn", "gso" 등) |
| `ctrl.GetAnchorPos(0)` | 컨트롤 앵커 위치 |

**주요 CtrlID:**

| CtrlID | 설명 |
|--------|------|
| `tbl` | 표 |
| `fn` | 각주 |
| `en` | 미주 |
| `gso` | 그리기 개체 |
| `eqed` | 수식 |

### 6. 테이블 선택/진입 액션

```python
# 테이블 위치로 이동
hwp.SetPosBySet(ctrl.GetAnchorPos(0))

# 테이블 선택 (컨트롤 선택 상태)
hwp.HAction.Run("SelectCtrlFront")
# 또는
hwp.HAction.Run("SelectCtrlReverse")

# 선택된 테이블의 첫 번째 셀로 진입
hwp.HAction.Run("ShapeObjTableSelCell")

# 테이블 밖으로 나가기
hwp.HAction.Run("MoveParentList")

# 선택 해제
hwp.HAction.Run("Cancel")
```

| 액션 | 설명 |
|------|------|
| `SelectCtrlFront` | 앞방향으로 컨트롤 선택 |
| `SelectCtrlReverse` | 뒷방향으로 컨트롤 선택 |
| `ShapeObjTableSelCell` | 테이블 선택 상태에서 첫 번째 셀 선택 |
| `MoveParentList` | 상위 리스트(본문)로 이동 |
| `Cancel` | 선택 해제 |

## 병합 셀 감지 원리

셀 너비/높이를 비교하여 colspan/rowspan 계산:

```
기준값 = 가장 작은 너비/높이 (병합되지 않은 셀)
colspan = round(셀 너비 / 기준 너비)
rowspan = round(셀 높이 / 기준 높이)
```

예시:
- 기준 너비 = 2000 HWPUNIT
- 셀 A 너비 = 4000 → colspan = 2
- 셀 B 너비 = 6000 → colspan = 3

## 좌표 매핑 알고리즘

1. 첫 셀(좌상단)에서 시작
2. 오른쪽으로만 이동하며 좌표 할당
3. colspan 셀은 여러 좌표에 같은 list_id 매핑
4. 열 수 채우면 다음 행으로 넘김
5. 마지막 list_id 도달 시 종료

```
테이블:
┌───┬───┬───────┐
│ A │ B │   C   │  ← C는 colspan=2
├───┼───┼───┬───┤
│ D │ E │ F │ G │
└───┴───┴───┴───┘

좌표 매핑:
(0,0)→A  (0,1)→B  (0,2)→C  (0,3)→C
(1,0)→D  (1,1)→E  (1,2)→F  (1,3)→G
```

## 데이터 구조

### CellInfo

```python
@dataclass
class CellInfo:
    list_id: int      # 셀의 고유 ID
    left: int = 0     # 좌측 셀 list_id (없으면 0)
    right: int = 0    # 우측 셀 list_id
    up: int = 0       # 상단 셀 list_id
    down: int = 0     # 하단 셀 list_id
    width: int = 0    # 셀 너비 (HWPUNIT)
    height: int = 0   # 셀 높이 (HWPUNIT)
```

### 반환값

```python
# build_coordinate_map()
Dict[tuple, int]  # {(row, col): list_id}

# get_table_size()
Dict[str, int]    # {'rows': 행수, 'cols': 열수}

# collect_cells_bfs()
Dict[int, CellInfo]  # {list_id: CellInfo}
```

## 주의사항

- 커서가 테이블 내부에 있어야 함
- rowspan은 좌표 매핑에서 고려하지 않음 (colspan만 처리)
- 캡션은 테이블 셀 다음 list_id를 가짐 (수집 대상 아님)
