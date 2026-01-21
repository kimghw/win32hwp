# -*- coding: utf-8 -*-
"""한글/엑셀 페이지 메타 정보 및 단위 변환 모듈"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from dataclasses import dataclass, field
from typing import Optional, Dict, List


# ============================================================
# 단위 변환 상수
# ============================================================

class Unit:
    """단위 변환 상수"""
    # 1 inch = 72 pt = 2.54 cm = 7200 HWPUNIT
    # 1 pt = 100 HWPUNIT
    # 1 cm = 2834.6 HWPUNIT (7200 / 2.54)

    HWPUNIT_PER_PT = 100
    HWPUNIT_PER_INCH = 7200
    HWPUNIT_PER_CM = 7200 / 2.54  # ≈ 2834.6
    HWPUNIT_PER_MM = 7200 / 25.4  # ≈ 283.46

    PT_PER_INCH = 72
    PT_PER_CM = 72 / 2.54  # ≈ 28.35

    # 엑셀 열 너비: 문자 단위 -> 포인트 (대략적 변환)
    # 1 문자 ≈ 7 pt (기본 폰트 기준)
    EXCEL_CHAR_TO_PT = 7

    @staticmethod
    def hwpunit_to_pt(hwpunit: int) -> float:
        """HWPUNIT -> 포인트"""
        return hwpunit / Unit.HWPUNIT_PER_PT

    @staticmethod
    def pt_to_hwpunit(pt: float) -> int:
        """포인트 -> HWPUNIT"""
        return int(pt * Unit.HWPUNIT_PER_PT)

    @staticmethod
    def hwpunit_to_cm(hwpunit: int) -> float:
        """HWPUNIT -> cm"""
        return hwpunit / Unit.HWPUNIT_PER_CM

    @staticmethod
    def cm_to_hwpunit(cm: float) -> int:
        """cm -> HWPUNIT"""
        return int(cm * Unit.HWPUNIT_PER_CM)

    @staticmethod
    def hwpunit_to_mm(hwpunit: int) -> float:
        """HWPUNIT -> mm"""
        return hwpunit / Unit.HWPUNIT_PER_MM

    @staticmethod
    def mm_to_hwpunit(mm: float) -> int:
        """mm -> HWPUNIT"""
        return int(mm * Unit.HWPUNIT_PER_MM)

    @staticmethod
    def excel_pt_to_hwpunit(pt: float) -> int:
        """엑셀 포인트 -> HWPUNIT"""
        return int(pt * Unit.HWPUNIT_PER_PT)

    @staticmethod
    def excel_char_to_hwpunit(chars: float) -> int:
        """엑셀 문자 단위(열 너비) -> HWPUNIT"""
        pt = chars * Unit.EXCEL_CHAR_TO_PT
        return int(pt * Unit.HWPUNIT_PER_PT)


# ============================================================
# 페이지 메타 데이터 클래스
# ============================================================

@dataclass
class PageMargin:
    """페이지 여백 (HWPUNIT)"""
    left: int = 0
    right: int = 0
    top: int = 0
    bottom: int = 0
    header: int = 0
    footer: int = 0
    gutter: int = 0  # 제본 여백

    def to_dict(self) -> Dict:
        return {
            'left': self.left,
            'right': self.right,
            'top': self.top,
            'bottom': self.bottom,
            'header': self.header,
            'footer': self.footer,
            'gutter': self.gutter,
            # cm 변환
            'left_cm': Unit.hwpunit_to_cm(self.left),
            'right_cm': Unit.hwpunit_to_cm(self.right),
            'top_cm': Unit.hwpunit_to_cm(self.top),
            'bottom_cm': Unit.hwpunit_to_cm(self.bottom),
        }


@dataclass
class PageSize:
    """페이지 크기 (HWPUNIT)"""
    width: int = 0
    height: int = 0
    orientation: str = 'portrait'  # portrait / landscape

    # 표준 용지 크기 (HWPUNIT)
    PAPER_SIZES = {
        'A4': (59528, 84188),      # 210mm x 297mm
        'A3': (84188, 119055),     # 297mm x 420mm
        'Letter': (61200, 79200),  # 8.5" x 11"
        'Legal': (61200, 100800),  # 8.5" x 14"
    }

    def to_dict(self) -> Dict:
        return {
            'width': self.width,
            'height': self.height,
            'orientation': self.orientation,
            'width_cm': Unit.hwpunit_to_cm(self.width),
            'height_cm': Unit.hwpunit_to_cm(self.height),
            'width_mm': Unit.hwpunit_to_mm(self.width),
            'height_mm': Unit.hwpunit_to_mm(self.height),
        }


@dataclass
class HwpPageMeta:
    """한글 페이지 메타 정보"""
    page_size: PageSize = field(default_factory=PageSize)
    margin: PageMargin = field(default_factory=PageMargin)

    # 본문 영역 (페이지 - 여백)
    content_width: int = 0
    content_height: int = 0

    # 단 설정
    columns: int = 1
    column_gap: int = 0

    def calculate_content_area(self):
        """본문 영역 계산"""
        self.content_width = (self.page_size.width
                              - self.margin.left
                              - self.margin.right
                              - self.margin.gutter)
        self.content_height = (self.page_size.height
                               - self.margin.top
                               - self.margin.bottom
                               - self.margin.header
                               - self.margin.footer)

    def to_dict(self) -> Dict:
        return {
            'page_size': self.page_size.to_dict(),
            'margin': self.margin.to_dict(),
            'content_width': self.content_width,
            'content_height': self.content_height,
            'content_width_cm': Unit.hwpunit_to_cm(self.content_width),
            'content_height_cm': Unit.hwpunit_to_cm(self.content_height),
            'columns': self.columns,
            'column_gap': self.column_gap,
        }


@dataclass
class TableMeta:
    """표 메타 정보"""
    width: int = 0
    height: int = 0
    row_count: int = 0
    col_count: int = 0
    row_heights: List[int] = field(default_factory=list)
    col_widths: List[int] = field(default_factory=list)

    def to_dict(self) -> Dict:
        return {
            'width': self.width,
            'height': self.height,
            'width_cm': Unit.hwpunit_to_cm(self.width),
            'height_cm': Unit.hwpunit_to_cm(self.height),
            'row_count': self.row_count,
            'col_count': self.col_count,
            'row_heights': self.row_heights,
            'col_widths': self.col_widths,
        }


# ============================================================
# 한글에서 메타 정보 추출
# ============================================================

def get_hwp_page_meta(hwp=None) -> Optional[HwpPageMeta]:
    """
    한글 문서의 페이지 메타 정보 추출

    Returns:
        HwpPageMeta 객체
    """
    if hwp is None:
        from cursor import get_hwp_instance
        hwp = get_hwp_instance()
        if hwp is None:
            return None

    meta = HwpPageMeta()

    try:
        # 방법 1: 현재 섹션 정보 직접 가져오기
        act = hwp.CreateAction("PageSetup")
        pset = act.CreateSet()
        act.GetDefault(pset)

        # PageDef 항목 접근
        page_def = pset.Item("PageDef")

        # 용지 크기
        meta.page_size.width = page_def.Item("PaperWidth")
        meta.page_size.height = page_def.Item("PaperHeight")

        # 용지 방향
        landscape = page_def.Item("Landscape")
        meta.page_size.orientation = 'landscape' if landscape else 'portrait'

        # 여백
        meta.margin.left = page_def.Item("LeftMargin")
        meta.margin.right = page_def.Item("RightMargin")
        meta.margin.top = page_def.Item("TopMargin")
        meta.margin.bottom = page_def.Item("BottomMargin")
        meta.margin.header = page_def.Item("HeaderLen")
        meta.margin.footer = page_def.Item("FooterLen")
        meta.margin.gutter = page_def.Item("GutterLen")

        # 본문 영역 계산
        meta.calculate_content_area()

    except Exception as e:
        print(f"[오류] 페이지 정보 추출 실패: {e}")
        import traceback
        traceback.print_exc()
        return None

    return meta


def get_hwp_table_meta(hwp=None) -> Optional[TableMeta]:
    """
    한글 표의 메타 정보 추출 (커서가 표 안에 있어야 함)

    Returns:
        TableMeta 객체
    """
    if hwp is None:
        from cursor import get_hwp_instance
        hwp = get_hwp_instance()
        if hwp is None:
            return None

    from table.table_info import TableInfo, MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL

    table_info = TableInfo(hwp, debug=False)

    if not table_info.is_in_table():
        return None

    meta = TableMeta()

    # 표 전체 크기
    ctrl = hwp.ParentCtrl
    if ctrl and ctrl.CtrlID == 'tbl':
        try:
            props = ctrl.Properties
            meta.width = props.Item('Width')
            meta.height = props.Item('Height')
        except:
            pass

    # 첫 셀로 이동
    table_info.move_to_first_cell()

    # 첫 번째 열 순회하며 행 높이 수집
    row_heights = []
    while True:
        list_id = hwp.GetPos()[0]
        _, height = table_info.get_cell_dimensions()
        row_heights.append(height)

        hwp.SetPos(list_id, 0, 0)
        hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
        next_id = hwp.GetPos()[0]

        if next_id == list_id:
            break

    meta.row_heights = row_heights
    meta.row_count = len(row_heights)

    # 첫 번째 행 순회하며 열 너비 수집
    table_info.move_to_first_cell()
    col_widths = []
    while True:
        list_id = hwp.GetPos()[0]
        width, _ = table_info.get_cell_dimensions()
        col_widths.append(width)

        hwp.SetPos(list_id, 0, 0)
        hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
        next_id = hwp.GetPos()[0]

        if next_id == list_id:
            break

    meta.col_widths = col_widths
    meta.col_count = len(col_widths)

    return meta


# ============================================================
# 테스트
# ============================================================

if __name__ == "__main__":
    print("=== 단위 변환 테스트 ===")
    print(f"100 HWPUNIT = {Unit.hwpunit_to_pt(100)} pt")
    print(f"100 HWPUNIT = {Unit.hwpunit_to_cm(100):.4f} cm")
    print(f"1 cm = {Unit.cm_to_hwpunit(1)} HWPUNIT")
    print(f"1 mm = {Unit.mm_to_hwpunit(1)} HWPUNIT")
    print(f"엑셀 24.52 pt = {Unit.excel_pt_to_hwpunit(24.52)} HWPUNIT")

    print("\n=== 한글 페이지 메타 정보 ===")
    page_meta = get_hwp_page_meta()
    if page_meta:
        import json
        print(json.dumps(page_meta.to_dict(), indent=2, ensure_ascii=False))

    print("\n=== 한글 표 메타 정보 ===")
    table_meta = get_hwp_table_meta()
    if table_meta:
        print(f"표 크기: {table_meta.width} x {table_meta.height} HWPUNIT")
        print(f"행 수: {table_meta.row_count}, 열 수: {table_meta.col_count}")
        print(f"행 높이 (처음 10개): {table_meta.row_heights[:10]}")
        print(f"열 너비 (처음 10개): {table_meta.col_widths[:10]}")
