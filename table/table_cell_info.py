# -*- coding: utf-8 -*-
"""
테이블 셀 유틸리티 모듈

테이블 셀 순회, 컨트롤 탐색, 서식 정보 조회 등의 유틸리티 함수 제공
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    from .table_info import MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL
except ImportError:
    from table_info import MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL


def get_ctrls_in_cell(hwp, target_list_id, target_para_id=None):
    """
    특정 셀(list_id)에 속한 컨트롤 찾기

    Args:
        hwp: HWP 인스턴스
        target_list_id: 대상 list_id
        target_para_id: 특정 문단만 필터링 (None=전체)

    Returns:
        list: 컨트롤 정보 리스트
    """
    ctrls = []
    ctrl = hwp.HeadCtrl

    while ctrl:
        try:
            anchor = ctrl.GetAnchorPos(0)
            ctrl_list_id = anchor.Item("List")
            ctrl_para_id = anchor.Item("Para")

            if ctrl_list_id == target_list_id:
                if target_para_id is None or ctrl_para_id == target_para_id:
                    ctrl_id = ctrl.CtrlID
                    if ctrl_id and ctrl_id not in ("secd", "cold"):  # 섹션/열 정의 제외
                        ctrls.append({
                            'ctrl': ctrl,
                            'id': ctrl_id,
                            'desc': getattr(ctrl, 'UserDesc', ctrl_id),
                            'para': ctrl_para_id
                        })
        except:
            pass
        ctrl = ctrl.Next

    return ctrls


def get_char_shape_info(hwp):
    """
    현재 위치의 글자 모양 정보 조회

    Returns:
        tuple: (font_name, font_size, char_spacing)
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


def get_para_shape_info(hwp):
    """
    현재 위치의 문단 모양 정보 조회

    Returns:
        tuple: (align, line_spacing)
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


def insert_colored_text(hwp, text, color_bgr):
    """
    색상이 적용된 텍스트 삽입

    Args:
        hwp: HWP 인스턴스
        text: 삽입할 텍스트
        color_bgr: BGR 형식 색상 (예: 0x0000FF = 빨강, 0xFF0000 = 파랑)
    """
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

    # 삽입한 텍스트 선택
    for _ in range(len(text)):
        hwp.HAction.Run("MoveSelLeft")

    # 색상 적용
    hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
    hwp.HParameterSet.HCharShape.TextColor = color_bgr
    hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)

    hwp.HAction.Run("Cancel")
