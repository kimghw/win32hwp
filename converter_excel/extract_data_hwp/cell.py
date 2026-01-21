# -*- coding: utf-8 -*-
"""한글 셀 스타일 추출 모듈

셀 배경색, 테두리, 여백, 글꼴 등의 스타일 정보를 추출합니다.
Excel/openpyxl 관련 코드는 포함하지 않습니다.
"""

from dataclasses import dataclass, field
from typing import Optional, List, Tuple


@dataclass
class CellStyleData:
    """셀 스타일 데이터"""
    list_id: int = 0

    # 위치 정보
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

    # 배경색 (RGB)
    bg_color_rgb: Optional[Tuple[int, int, int]] = None

    # 셀 내부 여백 (HWPUNIT)
    margin_left: int = 0
    margin_right: int = 0
    margin_top: int = 0
    margin_bottom: int = 0

    # 글꼴 기본
    font_name: Optional[str] = None
    font_size_pt: float = 0
    font_bold: bool = False
    font_italic: bool = False
    font_color_rgb: Optional[Tuple[int, int, int]] = None

    # 글꼴 장식
    font_underline: int = 0        # 밑줄 타입 (0=없음, 1=실선, ...)
    font_underline_color: Optional[Tuple[int, int, int]] = None
    font_strikeout: int = 0        # 취소선 타입
    font_strikeout_color: Optional[Tuple[int, int, int]] = None
    font_outline: int = 0          # 외곽선 타입
    font_shadow: int = 0           # 그림자 타입
    font_emboss: bool = False      # 양각
    font_engrave: bool = False     # 음각
    font_superscript: bool = False # 위첨자
    font_subscript: bool = False   # 아래첨자

    # 자간/장평
    char_spacing: int = 0      # 자간 (%)
    char_ratio: int = 100      # 장평 (%)

    # 정렬
    align_horizontal: str = 'left'
    align_vertical: str = 'center'

    # 줄간격
    line_spacing: int = 0      # 줄간격 (%)

    # 테두리 (type, width)
    border_left_type: int = 0
    border_left_width: int = 0
    border_right_type: int = 0
    border_right_width: int = 0
    border_top_type: int = 0
    border_top_width: int = 0
    border_bottom_type: int = 0
    border_bottom_width: int = 0

    # 텍스트
    text: str = ""

    # 필드 이름 (배경색 없는 셀에만 설정)
    field_name: str = ""
    field_source: str = ""  # bookmark, A_B, A, B, random


@dataclass
class CellMatchResult:
    """셀 매칭 결과"""
    cells: List[CellStyleData] = field(default_factory=list)
    row_heights: List[int] = field(default_factory=list)  # HWPUNIT
    col_widths: List[int] = field(default_factory=list)   # HWPUNIT
    success: bool = False
    error: str = None


def extract_cell_style(hwp, list_id: int) -> CellStyleData:
    """한글 셀에서 스타일 정보 추출

    Args:
        hwp: HWP 객체
        list_id: 셀의 list_id

    Returns:
        CellStyleData 객체
    """
    style = CellStyleData(list_id=list_id)

    try:
        hwp.SetPos(list_id, 0, 0)

        # 1. 배경색 및 셀 내부 여백 (CellBorderFill)
        pset = hwp.HParameterSet.HCellBorderFill
        hwp.HAction.GetDefault("CellBorderFill", pset.HSet)

        # 배경색
        fill_attr = pset.FillAttr
        if fill_attr:
            bg_color = fill_attr.WinBrushFaceColor
            if bg_color and bg_color > 0 and bg_color != 4294967295:
                b = (bg_color >> 16) & 0xFF
                g = (bg_color >> 8) & 0xFF
                r = bg_color & 0xFF
                style.bg_color_rgb = (r, g, b)

            # 셀 내부 여백
            style.margin_left = getattr(fill_attr, 'InsideMarginLeft', 0) or 0
            style.margin_right = getattr(fill_attr, 'InsideMarginRight', 0) or 0
            style.margin_top = getattr(fill_attr, 'InsideMarginTop', 0) or 0
            style.margin_bottom = getattr(fill_attr, 'InsideMarginBottom', 0) or 0

        # 테두리
        style.border_left_type = pset.BorderTypeLeft
        style.border_left_width = pset.BorderWidthLeft
        style.border_right_type = pset.BorderTypeRight
        style.border_right_width = pset.BorderWidthRight
        style.border_top_type = pset.BorderTypeTop
        style.border_top_width = pset.BorderWidthTop
        style.border_bottom_type = pset.BorderTypeBottom
        style.border_bottom_width = pset.BorderWidthBottom

        # 2. 글자 속성 (HCharShape 파라미터셋)
        try:
            char_pset = hwp.HParameterSet.HCharShape
            hwp.HAction.GetDefault("CharShape", char_pset.HSet)

            # 기본 속성
            style.font_name = char_pset.FaceNameHangul
            style.font_size_pt = char_pset.Height / 100  # HWPUNIT -> pt
            style.font_bold = bool(char_pset.Bold)
            style.font_italic = bool(char_pset.Italic)

            # 글자색 (BGR -> RGB)
            text_color = char_pset.TextColor
            if text_color and text_color > 0:
                b = (text_color >> 16) & 0xFF
                g = (text_color >> 8) & 0xFF
                r = text_color & 0xFF
                style.font_color_rgb = (r, g, b)

            # 밑줄
            style.font_underline = char_pset.UnderlineType
            underline_color = char_pset.UnderlineColor
            if underline_color and underline_color > 0:
                b = (underline_color >> 16) & 0xFF
                g = (underline_color >> 8) & 0xFF
                r = underline_color & 0xFF
                style.font_underline_color = (r, g, b)

            # 취소선
            style.font_strikeout = char_pset.StrikeOutType
            strikeout_color = char_pset.StrikeOutColor
            if strikeout_color and strikeout_color > 0:
                b = (strikeout_color >> 16) & 0xFF
                g = (strikeout_color >> 8) & 0xFF
                r = strikeout_color & 0xFF
                style.font_strikeout_color = (r, g, b)

            # 외곽선, 그림자, 양각/음각
            style.font_outline = char_pset.OutLineType
            style.font_shadow = char_pset.ShadowType
            style.font_emboss = bool(char_pset.Emboss)
            style.font_engrave = bool(char_pset.Engrave)

            # 위첨자/아래첨자
            style.font_superscript = bool(char_pset.SuperScript)
            style.font_subscript = bool(char_pset.SubScript)

            # 자간/장평 (한글 기준)
            style.char_spacing = char_pset.SpacingHangul
            style.char_ratio = char_pset.RatioHangul
        except:
            pass

        # 3. 정렬/줄간격 (텍스트 선택 후 ParaShape)
        try:
            # 위치 재설정 후 텍스트 선택
            hwp.SetPos(list_id, 0, 0)
            hwp.HAction.Run("MoveSelParaEnd")
            ps = hwp.ParaShape
            if ps:
                align_val = ps.Item("AlignType")
                if align_val is not None:
                    # 0=양쪽, 1=왼쪽, 2=오른쪽, 3=가운데, 4=배분, 5=나눔
                    align_map = {0: 'justify', 1: 'left', 2: 'right', 3: 'center', 4: 'distribute', 5: 'divide'}
                    style.align_horizontal = align_map.get(align_val, 'left')

                line_sp = ps.Item("LineSpacing")
                if line_sp is not None:
                    style.line_spacing = line_sp
            hwp.HAction.Run("Cancel")
        except:
            pass

        # 4. 셀 세로 정렬 - CellShape.VertAlign이 항상 0 반환하는 문제 있음
        # 기본값 'center' 사용 (CellStyleData에서 설정됨)

    except Exception as e:
        pass

    return style


def get_cell_text(hwp, list_id: int) -> str:
    """셀 텍스트 추출 (SelectAll 사용)"""
    try:
        hwp.SetPos(list_id, 0, 0)

        # SelectAll로 셀 전체 선택
        hwp.HAction.Run("SelectAll")

        # 선택된 텍스트 가져오기 (saveblock 옵션 사용)
        text = hwp.GetTextFile("TEXT", "saveblock")

        # 선택 해제
        hwp.HAction.Run("Cancel")

        # 텍스트 정리
        if text:
            text = text.strip().replace('\r\n', ' ').replace('\r', ' ').replace('\n', ' ')
            return text
        return ""
    except:
        return ""


# ============================================================
# 테스트
# ============================================================

if __name__ == "__main__":
    import sys
    import os
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

    from cursor import get_hwp_instance
    from table.table_info import TableInfo, MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL

    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        exit(1)

    table_info = TableInfo(hwp, debug=False)
    if not table_info.is_in_table():
        print("[오류] 커서가 표 안에 있지 않습니다.")
        exit(1)

    print("=== 셀 스타일 추출 테스트 ===\n")

    # 첫 셀로 이동
    table_info.move_to_first_cell()
    first_id = hwp.GetPos()[0]

    # BFS로 모든 셀 순회
    visited = set()
    queue = [first_id]
    visited.add(first_id)
    cells = []

    while queue and len(cells) < 50:
        cur_id = queue.pop(0)

        style = extract_cell_style(hwp, cur_id)
        style.text = get_cell_text(hwp, cur_id)
        cells.append(style)

        if style.bg_color_rgb:
            print(f"셀 {len(cells)}: bg={style.bg_color_rgb}, font={style.font_name}, size={style.font_size_pt}pt")

        # 다음 셀 탐색
        for move_cmd in [MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL]:
            hwp.SetPos(cur_id, 0, 0)
            hwp.MovePos(move_cmd, 0, 0)
            next_id = hwp.GetPos()[0]
            if next_id != cur_id and next_id not in visited:
                visited.add(next_id)
                queue.append(next_id)

    print(f"\n총 {len(cells)}개 셀 탐색 완료")
