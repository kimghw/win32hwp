# -*- coding: utf-8 -*-
"""
HWP 데이터 조회 API 모듈

이 패키지는 HWP 문서에서 데이터를 조회하는 기능들을 제공합니다.
- position: 위치 정보 조회 (커서, 문단, 줄, 문장)
- list_id: list_id 관련 조회 및 좌표 변환
- control: 컨트롤 탐색 (표, 그림, 수식 등)
- table_query: 테이블 정보 조회
- field_query: 필드 조회
- cell_query: 셀 좌표/범위 조회
- shape_query: 서식 정보 조회

사용 예시:
    from hwp_query import get_current_pos, get_list_id, is_in_table
    from hwp_query.position import get_sentences
    from hwp_query.table_query import find_all_tables
"""

# position 모듈
from hwp_query.position import (
    get_current_pos,
    get_para_range,
    get_line_range,
    get_sentences,
    get_cursor_index,
    KeyIndicatorInfo,
    PosInfo,
)

# list_id 모듈
from hwp_query.list_id import (
    get_list_id,
    get_list_id_from_coord,
    get_coord_from_list_id,
)

# control 모듈
from hwp_query.control import (
    CtrlInfo,
    get_ctrls_in_cell,
    find_ctrl,
    CTRL_NAMES,
)

# table_query 모듈
from hwp_query.table_query import (
    is_in_table,
    get_table_size,
    get_cell_dimensions,
    find_all_tables,
    get_table_caption,
    get_all_table_captions,
)

# field_query 모듈
from hwp_query.field_query import (
    get_all_fields,
    get_cell_fields,
    get_field_by_name,
    get_fields_at_coord,
    get_field_text,
    field_exists,
    get_json_fields,
    get_fields_by_table,
    find_field_by_coord,
    get_field_value_by_coord,
    get_current_field_name,
    get_current_field_state,
    get_field_summary,
    FieldInfo,
)

# cell_query 모듈
from hwp_query.cell_query import (
    CellRange,
    CellPositionResult,
    calculate_cell_positions,
    get_cell_at,
    get_cell_by_coord,
    build_coord_to_listid_map,
    get_merged_cells,
    get_merge_info,
)

# shape_query 모듈
from hwp_query.shape_query import (
    get_char_shape_info,
    get_para_shape_info,
)

__all__ = [
    # position
    'get_current_pos', 'get_para_range', 'get_line_range',
    'get_sentences', 'get_cursor_index',
    'KeyIndicatorInfo', 'PosInfo',
    # list_id
    'get_list_id', 'get_list_id_from_coord', 'get_coord_from_list_id',
    # control
    'CtrlInfo', 'get_ctrls_in_cell', 'find_ctrl', 'CTRL_NAMES',
    # table_query
    'is_in_table', 'get_table_size', 'get_cell_dimensions',
    'find_all_tables', 'get_table_caption', 'get_all_table_captions',
    # field_query
    'get_all_fields', 'get_cell_fields', 'get_field_by_name',
    'get_fields_at_coord', 'get_field_text', 'field_exists',
    'get_json_fields', 'get_fields_by_table', 'find_field_by_coord',
    'get_field_value_by_coord', 'get_current_field_name',
    'get_current_field_state', 'get_field_summary', 'FieldInfo',
    # cell_query
    'CellRange', 'CellPositionResult', 'calculate_cell_positions',
    'get_cell_at', 'get_cell_by_coord', 'build_coord_to_listid_map',
    'get_merged_cells', 'get_merge_info',
    # shape_query
    'get_char_shape_info', 'get_para_shape_info',
]
