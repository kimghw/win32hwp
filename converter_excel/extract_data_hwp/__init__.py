# -*- coding: utf-8 -*-
"""HWP 데이터 추출 패키지

한글 문서에서 데이터를 추출하는 모든 함수를 제공합니다.
- 페이지 정보 추출
- 셀 스타일 추출
- 필드 이름 생성

이 패키지는 Excel/openpyxl 관련 코드를 포함하지 않습니다.
"""

from .extract import (
    HwpExtractedData,
    extract_hwp_data,
    extract_cells_only,
    extract_fields_only,
)
from .page import PageMatchResult, extract_page_info
from .cell import CellStyleData, CellMatchResult, extract_cell_style, get_cell_text
from .field import (
    FieldInfo,
    generate_field_names,
    get_cell_bookmark,
    set_cell_field_names,
    find_left_header_cell,
    find_top_header_cell,
    clean_text_for_field_name,
    generate_random_field_name,
)

__all__ = [
    # 메인 데이터 클래스 및 함수
    'HwpExtractedData',
    'extract_hwp_data',
    'extract_cells_only',
    'extract_fields_only',

    # 페이지 관련
    'PageMatchResult',
    'extract_page_info',

    # 셀 관련
    'CellStyleData',
    'CellMatchResult',
    'extract_cell_style',
    'get_cell_text',

    # 필드 관련
    'FieldInfo',
    'generate_field_names',
    'get_cell_bookmark',
    'set_cell_field_names',
    'find_left_header_cell',
    'find_top_header_cell',
    'clean_text_for_field_name',
    'generate_random_field_name',
]
