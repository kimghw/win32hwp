# -*- coding: utf-8 -*-
"""한글 셀 스타일(배경색, 테두리) 추출 모듈"""

from dataclasses import dataclass
from typing import Optional, Tuple, Dict


@dataclass
class CellBorder:
    """셀 테두리 정보"""
    type: int = 0      # 테두리 타입 (0=없음, 1=실선 등)
    width: int = 0     # 테두리 두께
    color: int = 0     # 테두리 색상 (BGR)
    color_rgb: Optional[Tuple[int, int, int]] = None


@dataclass
class CellStyle:
    """셀 스타일 정보"""
    bg_color: int = 0                    # 배경색 (BGR)
    bg_color_rgb: Optional[Tuple[int, int, int]] = None  # RGB 튜플
    border_left: CellBorder = None
    border_right: CellBorder = None
    border_top: CellBorder = None
    border_bottom: CellBorder = None

    def __post_init__(self):
        if self.border_left is None:
            self.border_left = CellBorder()
        if self.border_right is None:
            self.border_right = CellBorder()
        if self.border_top is None:
            self.border_top = CellBorder()
        if self.border_bottom is None:
            self.border_bottom = CellBorder()

    def has_bg_color(self) -> bool:
        """배경색이 설정되어 있는지 확인"""
        return self.bg_color > 0 and self.bg_color != 4294967295

    def to_dict(self) -> Dict:
        return {
            'bg_color': self.bg_color,
            'bg_color_rgb': self.bg_color_rgb,
            'borders': {
                'left': {'type': self.border_left.type, 'width': self.border_left.width},
                'right': {'type': self.border_right.type, 'width': self.border_right.width},
                'top': {'type': self.border_top.type, 'width': self.border_top.width},
                'bottom': {'type': self.border_bottom.type, 'width': self.border_bottom.width},
            }
        }


def bgr_to_rgb(bgr: int) -> Tuple[int, int, int]:
    """BGR 정수를 RGB 튜플로 변환"""
    b = (bgr >> 16) & 0xFF
    g = (bgr >> 8) & 0xFF
    r = bgr & 0xFF
    return (r, g, b)


def rgb_to_bgr(r: int, g: int, b: int) -> int:
    """RGB를 BGR 정수로 변환"""
    return (b << 16) | (g << 8) | r


def get_cell_style(hwp, list_id: int) -> CellStyle:
    """
    특정 셀의 스타일 정보 추출

    Args:
        hwp: HWP 객체
        list_id: 셀의 list_id

    Returns:
        CellStyle 객체
    """
    hwp.SetPos(list_id, 0, 0)

    style = CellStyle()

    try:
        pset = hwp.HParameterSet.HCellBorderFill
        hwp.HAction.GetDefault("CellBorderFill", pset.HSet)

        # 배경색 추출
        fill_attr = pset.FillAttr
        if fill_attr:
            bg_color = fill_attr.WinBrushFaceColor
            if bg_color and bg_color > 0 and bg_color != 4294967295:
                style.bg_color = bg_color
                style.bg_color_rgb = bgr_to_rgb(bg_color)

        # 테두리 정보 추출
        style.border_left = CellBorder(
            type=pset.BorderTypeLeft,
            width=pset.BorderWidthLeft,
            color=pset.BorderCorlorLeft,  # HWP API 오타
        )
        if style.border_left.color > 0:
            style.border_left.color_rgb = bgr_to_rgb(style.border_left.color)

        style.border_right = CellBorder(
            type=pset.BorderTypeRight,
            width=pset.BorderWidthRight,
            color=pset.BorderColorRight,
        )
        if style.border_right.color > 0:
            style.border_right.color_rgb = bgr_to_rgb(style.border_right.color)

        style.border_top = CellBorder(
            type=pset.BorderTypeTop,
            width=pset.BorderWidthTop,
            color=pset.BorderColorTop,
        )
        if style.border_top.color > 0:
            style.border_top.color_rgb = bgr_to_rgb(style.border_top.color)

        style.border_bottom = CellBorder(
            type=pset.BorderTypeBottom,
            width=pset.BorderWidthBottom,
            color=pset.BorderColorBottom,
        )
        if style.border_bottom.color > 0:
            style.border_bottom.color_rgb = bgr_to_rgb(style.border_bottom.color)

    except Exception as e:
        print(f"[오류] 셀 스타일 추출 실패 (list_id={list_id}): {e}")

    return style


def get_cell_bg_color(hwp, list_id: int) -> Optional[Tuple[int, int, int]]:
    """
    셀 배경색만 추출 (RGB 튜플 반환)

    Args:
        hwp: HWP 객체
        list_id: 셀의 list_id

    Returns:
        (R, G, B) 튜플 또는 None (배경색 없음)
    """
    hwp.SetPos(list_id, 0, 0)

    try:
        pset = hwp.HParameterSet.HCellBorderFill
        hwp.HAction.GetDefault("CellBorderFill", pset.HSet)

        bg_color = pset.FillAttr.WinBrushFaceColor
        if bg_color and bg_color > 0 and bg_color != 4294967295:
            return bgr_to_rgb(bg_color)

    except Exception as e:
        pass

    return None


def set_cell_bg_color(hwp, color_rgb: Tuple[int, int, int]) -> bool:
    """
    현재 셀에 배경색 설정

    Args:
        hwp: HWP 객체
        color_rgb: (R, G, B) 튜플

    Returns:
        성공 여부
    """
    try:
        r, g, b = color_rgb
        color_bgr = rgb_to_bgr(r, g, b)

        pset = hwp.HParameterSet.HCellBorderFill
        hwp.HAction.GetDefault("CellBorderFill", pset.HSet)

        pset.FillAttr.type = 1  # 단색 채우기
        pset.FillAttr.WindowsBrush = 1
        pset.FillAttr.WinBrushFaceColor = color_bgr
        pset.FillAttr.WinBrushFaceStyle = 0  # 솔리드

        hwp.HAction.Execute("CellBorderFill", pset.HSet)
        return True

    except Exception as e:
        print(f"[오류] 셀 배경색 설정 실패: {e}")
        return False


# 테스트 코드
if __name__ == "__main__":
    import sys
    import os
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

    from cursor import get_hwp_instance
    from table.table_info import TableInfo, MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL

    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        exit(1)

    table_info = TableInfo(hwp, debug=False)
    if not table_info.is_in_table():
        print("[오류] 커서가 테이블 안에 있지 않습니다.")
        exit(1)

    print("=== 셀 스타일 모듈 테스트 ===\n")

    table_info.move_to_first_cell()
    first_id = hwp.GetPos()[0]

    visited = set()
    queue = [first_id]
    visited.add(first_id)

    count = 0
    colored_cells = []

    while queue and count < 30:
        cur_id = queue.pop(0)
        count += 1

        style = get_cell_style(hwp, cur_id)

        if style.has_bg_color():
            colored_cells.append((cur_id, style))
            print(f"셀 {count} (id={cur_id}): BGR={hex(style.bg_color)}, RGB={style.bg_color_rgb}")

        # BFS 탐색
        for move_cmd in [MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL]:
            hwp.SetPos(cur_id, 0, 0)
            hwp.MovePos(move_cmd, 0, 0)
            next_id = hwp.GetPos()[0]
            if next_id != cur_id and next_id not in visited:
                visited.add(next_id)
                queue.append(next_id)

    print(f"\n총 {count}개 셀 탐색, 배경색 있는 셀: {len(colored_cells)}개")
