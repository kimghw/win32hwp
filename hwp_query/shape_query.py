# -*- coding: utf-8 -*-
"""
서식 정보 조회 모듈

HWP 문서의 글자 모양, 문단 모양 등 서식 정보를 조회합니다.
"""

from typing import Tuple, Dict, Any


def get_char_shape_info(hwp) -> Tuple[str, int, int]:
    """
    현재 위치의 글자 모양 정보 조회

    Args:
        hwp: HWP COM 객체

    Returns:
        tuple: (font_name, font_size, char_spacing)
            - font_name: 글꼴 이름 (한글 글꼴)
            - font_size: 글자 크기 (pt)
            - char_spacing: 자간 (%)
    """
    try:
        cs = hwp.CharShape
        if cs:
            font_name = cs.Item("FaceNameHangul") or ""
            height = cs.Item("Height") or 0
            font_size = height // 100 if height else 0
            char_spacing = cs.Item("CharSpacing") or 0
            return font_name, font_size, char_spacing
    except:
        pass
    return "", 0, 0


def get_para_shape_info(hwp) -> Tuple[str, int]:
    """
    현재 위치의 문단 모양 정보 조회

    Args:
        hwp: HWP COM 객체

    Returns:
        tuple: (align, line_spacing)
            - align: 정렬 방식 ('양쪽', '왼쪽', '오른쪽', '가운데', '배분', '나눔')
            - line_spacing: 줄 간격 (%)
    """
    align_map = {0: "양쪽", 1: "왼쪽", 2: "오른쪽", 3: "가운데", 4: "배분", 5: "나눔"}
    try:
        ps = hwp.ParaShape
        if ps:
            align_val = ps.Item("Align")
            align = align_map.get(align_val, str(align_val))
            line_spacing = ps.Item("LineSpacing") or 0
            return align, line_spacing
    except:
        pass
    return "", 0


def get_char_shape_detail(hwp) -> Dict[str, Any]:
    """
    현재 위치의 상세 글자 모양 정보 조회

    Args:
        hwp: HWP COM 객체

    Returns:
        dict: {
            'font_name_hangul': 한글 글꼴,
            'font_name_latin': 영문 글꼴,
            'font_size': 글자 크기 (pt),
            'bold': 굵게 여부,
            'italic': 기울임 여부,
            'underline': 밑줄 여부,
            'strikeout': 취소선 여부,
            'text_color': 글자 색상 (BGR),
            'char_spacing': 자간 (%),
            'ratio': 장평 (%)
        }
    """
    result = {
        'font_name_hangul': '',
        'font_name_latin': '',
        'font_size': 0,
        'bold': False,
        'italic': False,
        'underline': False,
        'strikeout': False,
        'text_color': 0,
        'char_spacing': 0,
        'ratio': 100
    }

    try:
        cs = hwp.CharShape
        if not cs:
            return result

        result['font_name_hangul'] = cs.Item("FaceNameHangul") or ""
        result['font_name_latin'] = cs.Item("FaceNameLatin") or ""

        height = cs.Item("Height") or 0
        result['font_size'] = height // 100 if height else 0

        result['bold'] = bool(cs.Item("Bold"))
        result['italic'] = bool(cs.Item("Italic"))
        result['underline'] = bool(cs.Item("UnderlineType"))
        result['strikeout'] = bool(cs.Item("StrikeOutType"))
        result['text_color'] = cs.Item("TextColor") or 0
        result['char_spacing'] = cs.Item("CharSpacing") or 0
        result['ratio'] = cs.Item("Ratio") or 100

    except:
        pass

    return result


def get_para_shape_detail(hwp) -> Dict[str, Any]:
    """
    현재 위치의 상세 문단 모양 정보 조회

    Args:
        hwp: HWP COM 객체

    Returns:
        dict: {
            'align': 정렬 방식,
            'line_spacing': 줄 간격 (%),
            'space_before': 문단 위 간격,
            'space_after': 문단 아래 간격,
            'indent_left': 왼쪽 여백,
            'indent_right': 오른쪽 여백,
            'indent_first': 첫 줄 들여쓰기
        }
    """
    align_map = {0: "양쪽", 1: "왼쪽", 2: "오른쪽", 3: "가운데", 4: "배분", 5: "나눔"}

    result = {
        'align': '',
        'line_spacing': 0,
        'space_before': 0,
        'space_after': 0,
        'indent_left': 0,
        'indent_right': 0,
        'indent_first': 0
    }

    try:
        ps = hwp.ParaShape
        if not ps:
            return result

        align_val = ps.Item("Align")
        result['align'] = align_map.get(align_val, str(align_val))
        result['line_spacing'] = ps.Item("LineSpacing") or 0
        result['space_before'] = ps.Item("SpaceBeforePara") or 0
        result['space_after'] = ps.Item("SpaceAfterPara") or 0
        result['indent_left'] = ps.Item("LeftMargin") or 0
        result['indent_right'] = ps.Item("RightMargin") or 0
        result['indent_first'] = ps.Item("Indent") or 0

    except:
        pass

    return result


def get_cell_shape_info(hwp) -> Dict[str, Any]:
    """
    현재 위치의 셀 모양 정보 조회

    Args:
        hwp: HWP COM 객체

    Returns:
        dict: {
            'width': 셀 너비 (HWPUNIT),
            'height': 셀 높이 (HWPUNIT),
            'margin_left': 왼쪽 여백,
            'margin_right': 오른쪽 여백,
            'margin_top': 위쪽 여백,
            'margin_bottom': 아래쪽 여백,
            'border_fill': 테두리/배경 정보
        }
    """
    result = {
        'width': 0,
        'height': 0,
        'margin_left': 0,
        'margin_right': 0,
        'margin_top': 0,
        'margin_bottom': 0,
        'border_fill': None
    }

    try:
        cell_shape = hwp.CellShape
        if not cell_shape:
            return result

        cell = cell_shape.Item("Cell")
        if cell:
            result['width'] = cell.Item("Width") or 0
            result['height'] = cell.Item("Height") or 0
            result['margin_left'] = cell.Item("LeftMargin") or 0
            result['margin_right'] = cell.Item("RightMargin") or 0
            result['margin_top'] = cell.Item("TopMargin") or 0
            result['margin_bottom'] = cell.Item("BottomMargin") or 0

        border_fill = cell_shape.Item("BorderFill")
        if border_fill:
            result['border_fill'] = border_fill

    except:
        pass

    return result
