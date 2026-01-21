# -*- coding: utf-8 -*-
"""한글 -> 엑셀 변환 시 필요한 데이터 구조 정의

이 모듈은 한글 문서에서 엑셀로 변환할 때 필요한 모든 정보를 정리합니다.
"""

from dataclasses import dataclass, field
from typing import List, Optional, Tuple, Dict


# ============================================================
# 단위 변환 상수
# ============================================================

class Units:
    """단위 변환 상수 및 함수"""
    # HWP 단위
    HWPUNIT_PER_PT = 100        # 1 pt = 100 HWPUNIT
    HWPUNIT_PER_INCH = 7200     # 1 inch = 7200 HWPUNIT
    HWPUNIT_PER_CM = 7200 / 2.54  # 1 cm ≈ 2834.6 HWPUNIT

    # 엑셀 단위
    EXCEL_CHAR_WIDTH_PT = 7     # 1 문자 ≈ 7 pt (기본 폰트 기준)

    @staticmethod
    def hwpunit_to_pt(hwpunit: int) -> float:
        """HWPUNIT -> 포인트"""
        return hwpunit / Units.HWPUNIT_PER_PT

    @staticmethod
    def hwpunit_to_inch(hwpunit: int) -> float:
        """HWPUNIT -> 인치"""
        return hwpunit / Units.HWPUNIT_PER_INCH

    @staticmethod
    def hwpunit_to_cm(hwpunit: int) -> float:
        """HWPUNIT -> cm"""
        return hwpunit / Units.HWPUNIT_PER_CM

    @staticmethod
    def hwpunit_to_excel_row_height(hwpunit: int) -> float:
        """HWPUNIT -> 엑셀 행 높이 (pt 단위)"""
        return hwpunit / Units.HWPUNIT_PER_PT

    @staticmethod
    def hwpunit_to_excel_col_width(hwpunit: int) -> float:
        """HWPUNIT -> 엑셀 열 너비 (문자 단위)"""
        pt = hwpunit / Units.HWPUNIT_PER_PT
        return pt / Units.EXCEL_CHAR_WIDTH_PT


# ============================================================
# 페이지 정보
# ============================================================

@dataclass
class PageInfo:
    """페이지 설정 정보 (한글 -> 엑셀)"""
    # 용지 크기 (HWPUNIT)
    paper_width: int = 0
    paper_height: int = 0
    orientation: str = 'portrait'  # portrait / landscape

    # 여백 (HWPUNIT)
    margin_left: int = 0
    margin_right: int = 0
    margin_top: int = 0
    margin_bottom: int = 0
    margin_header: int = 0
    margin_footer: int = 0
    margin_gutter: int = 0

    # 본문 영역 (계산됨)
    content_width: int = 0
    content_height: int = 0

    def calculate_content_area(self):
        """본문 영역 계산"""
        self.content_width = (self.paper_width
                              - self.margin_left
                              - self.margin_right
                              - self.margin_gutter)
        self.content_height = (self.paper_height
                               - self.margin_top
                               - self.margin_bottom
                               - self.margin_header
                               - self.margin_footer)

    def to_excel_margins_inch(self) -> Dict[str, float]:
        """엑셀 PageMargins용 인치 단위 여백"""
        return {
            'left': Units.hwpunit_to_inch(self.margin_left),
            'right': Units.hwpunit_to_inch(self.margin_right),
            'top': Units.hwpunit_to_inch(self.margin_top),
            'bottom': Units.hwpunit_to_inch(self.margin_bottom),
            'header': Units.hwpunit_to_inch(self.margin_header),
            'footer': Units.hwpunit_to_inch(self.margin_footer),
        }

    def to_dict(self) -> Dict:
        """딕셔너리 변환 (다양한 단위 포함)"""
        return {
            'paper_width_hwpunit': self.paper_width,
            'paper_height_hwpunit': self.paper_height,
            'paper_width_cm': Units.hwpunit_to_cm(self.paper_width),
            'paper_height_cm': Units.hwpunit_to_cm(self.paper_height),
            'paper_width_inch': Units.hwpunit_to_inch(self.paper_width),
            'paper_height_inch': Units.hwpunit_to_inch(self.paper_height),
            'orientation': self.orientation,
            'margin_left_cm': Units.hwpunit_to_cm(self.margin_left),
            'margin_right_cm': Units.hwpunit_to_cm(self.margin_right),
            'margin_top_cm': Units.hwpunit_to_cm(self.margin_top),
            'margin_bottom_cm': Units.hwpunit_to_cm(self.margin_bottom),
            'margin_header_cm': Units.hwpunit_to_cm(self.margin_header),
            'margin_footer_cm': Units.hwpunit_to_cm(self.margin_footer),
            'content_width_cm': Units.hwpunit_to_cm(self.content_width),
            'content_height_cm': Units.hwpunit_to_cm(self.content_height),
        }


# ============================================================
# 셀 스타일 정보
# ============================================================

@dataclass
class CellStyleInfo:
    """셀 스타일 정보"""
    # 배경색
    bg_color_rgb: Optional[Tuple[int, int, int]] = None  # (R, G, B)

    # 셀 내부 여백 (HWPUNIT)
    padding_left: int = 0
    padding_right: int = 0
    padding_top: int = 0
    padding_bottom: int = 0

    # 글자 속성
    font_name: Optional[str] = None
    font_size_pt: float = 0
    font_bold: bool = False
    font_italic: bool = False
    font_underline: bool = False
    font_color_rgb: Optional[Tuple[int, int, int]] = None

    # 정렬
    align_horizontal: str = 'left'   # left, center, right, justify
    align_vertical: str = 'center'   # top, center, bottom

    # 테두리
    border_left: bool = True
    border_right: bool = True
    border_top: bool = True
    border_bottom: bool = True

    def bg_color_hex(self) -> Optional[str]:
        """배경색을 HEX 문자열로 반환"""
        if self.bg_color_rgb:
            r, g, b = self.bg_color_rgb
            return f"{r:02X}{g:02X}{b:02X}"
        return None

    def to_dict(self) -> Dict:
        return {
            'bg_color': self.bg_color_hex(),
            'padding_left_pt': Units.hwpunit_to_pt(self.padding_left),
            'padding_right_pt': Units.hwpunit_to_pt(self.padding_right),
            'padding_top_pt': Units.hwpunit_to_pt(self.padding_top),
            'padding_bottom_pt': Units.hwpunit_to_pt(self.padding_bottom),
            'font_name': self.font_name,
            'font_size_pt': self.font_size_pt,
            'font_bold': self.font_bold,
            'font_italic': self.font_italic,
            'align_horizontal': self.align_horizontal,
            'align_vertical': self.align_vertical,
        }


# ============================================================
# 셀 데이터
# ============================================================

@dataclass
class CellInfo:
    """셀 정보 (위치 + 내용 + 스타일)"""
    # 식별자
    list_id: int = 0

    # 논리 좌표 (0-based)
    row: int = 0
    col: int = 0
    end_row: int = 0
    end_col: int = 0
    rowspan: int = 1
    colspan: int = 1

    # 물리 좌표 (HWPUNIT)
    x: int = 0
    y: int = 0
    width: int = 0
    height: int = 0

    # 내용
    text: str = ""

    # 스타일
    style: CellStyleInfo = None

    def __post_init__(self):
        if self.style is None:
            self.style = CellStyleInfo()

    def is_merged(self) -> bool:
        """병합 셀 여부"""
        return self.rowspan > 1 or self.colspan > 1

    def width_pt(self) -> float:
        """셀 너비 (pt)"""
        return Units.hwpunit_to_pt(self.width)

    def height_pt(self) -> float:
        """셀 높이 (pt)"""
        return Units.hwpunit_to_pt(self.height)

    def to_dict(self) -> Dict:
        return {
            'list_id': self.list_id,
            'row': self.row,
            'col': self.col,
            'end_row': self.end_row,
            'end_col': self.end_col,
            'rowspan': self.rowspan,
            'colspan': self.colspan,
            'x': self.x,
            'y': self.y,
            'width': self.width,
            'height': self.height,
            'width_pt': self.width_pt(),
            'height_pt': self.height_pt(),
            'text': self.text,
            'is_merged': self.is_merged(),
            'style': self.style.to_dict() if self.style else None,
        }


# ============================================================
# 테이블 전체 정보
# ============================================================

@dataclass
class TableInfo:
    """테이블 전체 정보"""
    # 크기
    row_count: int = 0
    col_count: int = 0
    cell_count: int = 0

    # 물리 크기 (HWPUNIT)
    width: int = 0
    height: int = 0

    # 행 높이 / 열 너비 목록 (HWPUNIT)
    row_heights: List[int] = field(default_factory=list)
    col_widths: List[int] = field(default_factory=list)

    # 셀 목록
    cells: List[CellInfo] = field(default_factory=list)

    def width_cm(self) -> float:
        return Units.hwpunit_to_cm(self.width)

    def height_cm(self) -> float:
        return Units.hwpunit_to_cm(self.height)

    def to_dict(self) -> Dict:
        return {
            'row_count': self.row_count,
            'col_count': self.col_count,
            'cell_count': self.cell_count,
            'width_hwpunit': self.width,
            'height_hwpunit': self.height,
            'width_cm': self.width_cm(),
            'height_cm': self.height_cm(),
            'row_heights_pt': [Units.hwpunit_to_pt(h) for h in self.row_heights],
            'col_widths_pt': [Units.hwpunit_to_pt(w) for w in self.col_widths],
        }


# ============================================================
# 엑셀 변환용 종합 데이터
# ============================================================

@dataclass
class ExcelExportData:
    """엑셀 변환에 필요한 모든 데이터"""
    page: PageInfo = None
    table: TableInfo = None

    def __post_init__(self):
        if self.page is None:
            self.page = PageInfo()
        if self.table is None:
            self.table = TableInfo()

    def to_dict(self) -> Dict:
        return {
            'page': self.page.to_dict(),
            'table': self.table.to_dict(),
        }


# ============================================================
# 한글에서 데이터 추출
# ============================================================

def extract_page_info(hwp) -> PageInfo:
    """한글 문서에서 페이지 정보 추출"""
    page = PageInfo()

    try:
        act = hwp.CreateAction("PageSetup")
        pset = act.CreateSet()
        act.GetDefault(pset)
        page_def = pset.Item("PageDef")

        page.paper_width = page_def.Item("PaperWidth")
        page.paper_height = page_def.Item("PaperHeight")
        page.orientation = 'landscape' if page_def.Item("Landscape") else 'portrait'

        page.margin_left = page_def.Item("LeftMargin")
        page.margin_right = page_def.Item("RightMargin")
        page.margin_top = page_def.Item("TopMargin")
        page.margin_bottom = page_def.Item("BottomMargin")
        page.margin_header = page_def.Item("HeaderLen")
        page.margin_footer = page_def.Item("FooterLen")
        page.margin_gutter = page_def.Item("GutterLen")

        page.calculate_content_area()

    except Exception as e:
        print(f"[오류] 페이지 정보 추출 실패: {e}")

    return page


def extract_cell_style(hwp, list_id: int) -> CellStyleInfo:
    """한글 셀에서 스타일 정보 추출"""
    style = CellStyleInfo()

    try:
        hwp.SetPos(list_id, 0, 0)

        # 배경색 및 여백
        pset = hwp.HParameterSet.HCellBorderFill
        hwp.HAction.GetDefault("CellBorderFill", pset.HSet)

        bg_color = pset.FillAttr.WinBrushFaceColor
        if bg_color and bg_color > 0 and bg_color != 4294967295:
            b = (bg_color >> 16) & 0xFF
            g = (bg_color >> 8) & 0xFF
            r = bg_color & 0xFF
            style.bg_color_rgb = (r, g, b)

        fill_attr = pset.FillAttr
        if fill_attr:
            style.padding_left = getattr(fill_attr, 'InsideMarginLeft', 0) or 0
            style.padding_right = getattr(fill_attr, 'InsideMarginRight', 0) or 0
            style.padding_top = getattr(fill_attr, 'InsideMarginTop', 0) or 0
            style.padding_bottom = getattr(fill_attr, 'InsideMarginBottom', 0) or 0

        # 글자 속성
        try:
            char_shape = hwp.CharShape
            if char_shape:
                style.font_name = char_shape.Item("FaceNameHangul")
                height = char_shape.Item("Height")
                if height:
                    style.font_size_pt = height / 100
                style.font_bold = bool(char_shape.Item("Bold"))
                style.font_italic = bool(char_shape.Item("Italic"))
        except:
            pass

        # 정렬
        try:
            para_shape = hwp.ParaShape
            if para_shape:
                align_val = para_shape.Item("Align")
                align_map = {0: 'justify', 1: 'left', 2: 'right', 3: 'center'}
                style.align_horizontal = align_map.get(align_val, 'left')
        except:
            pass

    except Exception as e:
        pass

    return style


# ============================================================
# 테스트
# ============================================================

if __name__ == "__main__":
    import sys
    import os
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

    from cursor import get_hwp_instance

    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        exit(1)

    print("=== 페이지 정보 ===")
    page = extract_page_info(hwp)
    import json
    print(json.dumps(page.to_dict(), indent=2, ensure_ascii=False))

    print("\n=== 엑셀 여백 (인치) ===")
    print(page.to_excel_margins_inch())
