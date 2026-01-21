# hwp_query - HWP 데이터 조회 API

HWP 문서에서 데이터를 조회하는 함수들을 모아놓은 패키지입니다.

## 설치

별도 설치 필요 없음. 프로젝트 내에서 바로 import 가능.

## 모듈 구조

```
hwp_query/
├── __init__.py      # 모든 함수 re-export
├── position.py      # 위치 정보 조회 (커서, 문단, 줄, 문장)
├── list_id.py       # list_id 관련 조회 및 좌표 변환
├── control.py       # 컨트롤 탐색 (표, 그림, 수식 등)
├── table_query.py   # 테이블 정보 조회
├── field_query.py   # 필드 조회
├── cell_query.py    # 셀 좌표/범위 조회
└── shape_query.py   # 서식 정보 조회
```

## 사용 예시

### 1. 패키지에서 직접 import (권장)

```python
from hwp_query import (
    get_current_pos,      # 현재 커서 위치
    get_list_id,          # 현재 list_id
    is_in_table,          # 테이블 내부 여부
    get_all_fields,       # 모든 필드 조회
    get_char_shape_info,  # 글자 모양 조회
)

hwp = get_hwp_instance()

# 현재 위치 조회
pos = get_current_pos(hwp)
print(f"페이지: {pos['page']}, 줄: {pos['line']}")

# 테이블 내부인지 확인
if is_in_table(hwp):
    print("테이블 안에 있습니다")
```

### 2. 개별 모듈에서 import

```python
from hwp_query.position import get_sentences, get_para_range
from hwp_query.table_query import find_all_tables, get_table_caption
from hwp_query.field_query import get_json_fields, get_field_text

# 문단 내 문장 분석
sentences = get_sentences(hwp)
for s in sentences:
    print(f"문장 {s['index']}: {s['start']}~{s['end']}")

# 모든 테이블 찾기
tables = find_all_tables(hwp)
for t in tables:
    caption = get_table_caption(hwp, t['num'])
    print(f"표 {t['num']}: {caption}")
```

### 3. 셀 좌표 계산

```python
from hwp_query.cell_query import (
    calculate_cell_positions,
    get_cell_at,
    build_coord_to_listid_map,
)

# 셀 위치 계산
result = calculate_cell_positions(hwp)
print(f"테이블 크기: {result.max_row+1}행 x {result.max_col+1}열")

# 특정 좌표의 셀 찾기
cell = get_cell_at(result, row=1, col=2)
if cell:
    print(f"list_id: {cell.list_id}, 병합: {cell.is_merged()}")

# 좌표 → list_id 맵 생성
coord_map = build_coord_to_listid_map(result)
list_id = coord_map.get((0, 0))  # (0,0) 셀의 list_id
```

## 주요 함수 목록

### position.py - 위치 정보

| 함수 | 설명 |
|------|------|
| `get_current_pos(hwp)` | 현재 커서 위치 (list_id, para_id, char_pos, page, line 등) |
| `get_para_range(hwp)` | 문단 시작/끝 pos |
| `get_line_range(hwp)` | 현재 줄의 시작/끝 pos |
| `get_sentences(hwp)` | 문단 내 문장 경계 리스트 |
| `get_cursor_index(hwp)` | 현재 커서가 몇 번째 문장/단어인지 |

### list_id.py - list_id 관련

| 함수 | 설명 |
|------|------|
| `get_list_id(hwp)` | 현재 커서의 list_id |
| `get_list_id_from_coord(hwp, row, col)` | (row, col) → list_id 변환 |
| `get_coord_from_list_id(hwp, list_id)` | list_id → (row, col) 변환 |

### control.py - 컨트롤 탐색

| 함수 | 설명 |
|------|------|
| `find_ctrl(hwp)` | 현재 위치의 컨트롤 ID |
| `get_ctrls_in_cell(hwp, list_id)` | 특정 셀 내 컨트롤 목록 |
| `CTRL_NAMES` | 컨트롤 ID → 이름 매핑 |

### table_query.py - 테이블 정보

| 함수 | 설명 |
|------|------|
| `is_in_table(hwp)` | 테이블 내부 여부 |
| `get_table_size(hwp)` | 테이블 크기 (rows, cols) |
| `get_cell_dimensions(hwp)` | 현재 셀 너비/높이 |
| `find_all_tables(hwp)` | 문서 내 모든 테이블 목록 |
| `get_table_caption(hwp, index)` | 테이블 캡션 텍스트 |

### field_query.py - 필드 조회

| 함수 | 설명 |
|------|------|
| `get_all_fields(hwp)` | 모든 필드 조회 |
| `get_field_by_name(hwp, name)` | 이름으로 필드 조회 |
| `get_field_text(hwp, name)` | 필드 텍스트 조회 |
| `field_exists(hwp, name)` | 필드 존재 여부 |
| `get_json_fields(hwp)` | JSON 형식 필드 파싱 |

### cell_query.py - 셀 좌표/범위

| 함수 | 설명 |
|------|------|
| `calculate_cell_positions(hwp)` | 모든 셀 위치 계산 |
| `get_cell_at(result, row, col)` | 좌표로 셀 찾기 |
| `get_merged_cells(result)` | 병합 셀 목록 |
| `build_coord_to_listid_map(result)` | (row, col) → list_id 맵 |

### shape_query.py - 서식 정보

| 함수 | 설명 |
|------|------|
| `get_char_shape_info(hwp)` | 글자 모양 (글꼴, 크기, 자간) |
| `get_para_shape_info(hwp)` | 문단 모양 (정렬, 줄간격) |
| `get_char_shape_detail(hwp)` | 상세 글자 모양 |
| `get_cell_shape_info(hwp)` | 셀 모양 정보 |

## 기존 모듈과의 관계

이 패키지는 기존 모듈들(`cursor.py`, `table/table_info.py` 등)의 **조회 기능**을 독립적으로 분리한 것입니다.

- 기존 모듈: 조회 + 수정/생성 기능 포함
- `hwp_query`: 조회 기능만 (읽기 전용)

새 코드에서는 `hwp_query` 사용을 권장하며, 기존 코드와도 100% 호환됩니다.
