# -*- coding: utf-8 -*-
"""Excel 변환 관련 모듈"""

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
from .excel_export_data import (
    Units,
    PageInfo,
    CellStyleInfo,
    CellInfo,
    TableInfo,
    ExcelExportData,
    extract_page_info,
    extract_cell_style,
)
from .match_page import (
    PageMatchResult,
    extract_page_info as extract_page_info_v2,
    apply_page_margins_to_excel,
    write_page_info_to_sheet,
    save_page_info_to_excel,
)
from .match_cell import (
    CellStyleData,
    CellMatchResult,
    extract_cell_style as extract_cell_style_v2,
    get_cell_text,
    write_cell_styles_to_sheet,
    write_row_col_sizes_to_sheet,
    apply_cell_style_to_excel_cell,
)
