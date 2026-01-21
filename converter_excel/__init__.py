# -*- coding: utf-8 -*-
"""Excel 변환 관련 모듈

폴더 구조:
- extract_data_hwp/: HWP 데이터 추출 (페이지, 셀, 필드)
- apply_excel/: 추출된 데이터를 엑셀에 적용/저장
- 공통 모듈: page_meta, cell_style, config 등
"""

# =============================================================================
# 공통 모듈
# =============================================================================
from .page_setup import get_page_settings, is_fit_to_one_page_wide
from .page_meta import (
    Unit,
    PageMargin,
    PageSize,
    HwpPageMeta,
    TableMeta,
    get_hwp_page_meta,
    get_hwp_table_meta,
)
from .cell_style import (
    CellBorder,
    CellStyle,
    get_cell_style,
    get_cell_bg_color,
    set_cell_bg_color,
    bgr_to_rgb,
    rgb_to_bgr,
)
from .config import (
    ExportConfig,
    load_config,
    get_default_config,
    save_default_config,
)

# =============================================================================
# HWP 데이터 추출 (extract_data_hwp/)
# =============================================================================
from .extract_data_hwp import (
    # 메인 함수
    HwpExtractedData,
    extract_hwp_data,
    extract_cells_only,
    extract_fields_only,
    # 페이지
    PageMatchResult,
    extract_page_info,
    # 셀
    CellStyleData,
    CellMatchResult,
    extract_cell_style,
    get_cell_text,
    # 필드
    FieldInfo,
    generate_field_names,
    get_cell_bookmark,
    set_cell_field_names,
    generate_random_field_name,
)

# =============================================================================
# 엑셀 적용/저장 (apply_excel/)
# =============================================================================
from .apply_excel import (
    # 메인 함수
    apply_to_excel,
    create_main_sheet,
    # 페이지
    write_page_info_to_sheet,
    apply_page_margins_to_excel,
    # 셀
    write_cell_styles_to_sheet,
    write_row_col_sizes_to_sheet,
    apply_cell_style_to_excel_cell,
    # 필드
    write_field_info_to_sheet,
)

# =============================================================================
# 통합 변환기
# =============================================================================
from .export import HwpToExcelExporter

# =============================================================================
# 레거시 호환 (excel_export_data.py)
# =============================================================================
from .excel_export_data import (
    Units,
    PageInfo,
    CellStyleInfo,
    CellInfo,
    TableInfo,
    ExcelExportData,
    extract_page_info as extract_page_info_legacy,
    extract_cell_style as extract_cell_style_legacy,
)
