# -*- coding: utf-8 -*-
"""
테이블 정보 조회 모듈

HWP 문서 내 테이블의 구조, 크기, 캡션 등을 조회합니다.
"""

from typing import Dict, List, Tuple, Optional, Any


# MovePos 셀 이동 상수
MOVE_LEFT_OF_CELL = 100
MOVE_RIGHT_OF_CELL = 101
MOVE_UP_OF_CELL = 102
MOVE_DOWN_OF_CELL = 103
MOVE_START_OF_CELL = 104
MOVE_END_OF_CELL = 105
MOVE_TOP_OF_CELL = 106
MOVE_BOTTOM_OF_CELL = 107


def is_in_table(hwp) -> bool:
    """
    현재 커서가 테이블 내부에 있는지 확인

    Args:
        hwp: HWP COM 객체

    Returns:
        bool: 테이블 내부이면 True
    """
    try:
        act = hwp.CreateAction("TableCellBlock")
        pset = act.CreateSet()
        return act.Execute(pset)
    except:
        return False


def get_cell_dimensions(hwp) -> Tuple[int, int]:
    """
    현재 커서 위치의 셀 너비/높이 조회

    CellShape 속성을 사용하여 현재 셀의 크기 정보를 가져옵니다.

    Args:
        hwp: HWP COM 객체

    Returns:
        tuple: (width, height) HWPUNIT 단위, 실패 시 (0, 0)
    """
    try:
        cell_shape = hwp.CellShape
        if cell_shape is None:
            return (0, 0)

        cell = cell_shape.Item("Cell")
        if cell is None:
            return (0, 0)

        width = cell.Item("Width")
        height = cell.Item("Height")
        return (width or 0, height or 0)
    except Exception:
        return (0, 0)


def get_table_size(hwp, cells: Dict = None) -> Dict[str, int]:
    """
    테이블 크기 계산

    현재 테이블의 행/열 수를 반환합니다.

    Args:
        hwp: HWP COM 객체
        cells: 셀 정보 딕셔너리 (없으면 내부에서 수집)

    Returns:
        Dict: {'rows': 행수, 'cols': 열수}
    """
    try:
        from table.table_info import TableInfo
        table_info = TableInfo(hwp, debug=False)

        if cells:
            table_info.cells = cells
        else:
            table_info.collect_cells_bfs()

        return table_info.get_table_size()
    except Exception:
        return {'rows': 0, 'cols': 0}


def find_all_tables(hwp) -> List[Dict]:
    """
    문서 내 모든 테이블을 찾아 정보 반환

    HeadCtrl에서 시작하여 Next로 순회하며 CtrlID="tbl"인 컨트롤 수집

    Args:
        hwp: HWP COM 객체

    Returns:
        list: [{
            'num': 테이블 번호 (0부터),
            'ctrl': 컨트롤 객체,
            'first_cell_list_id': 첫 번째 셀의 list_id
        }, ...]
    """
    tables = []
    ctrl = hwp.HeadCtrl
    table_num = 0

    while ctrl:
        if ctrl.CtrlID == "tbl":
            # 테이블 위치로 이동 후 선택
            hwp.SetPosBySet(ctrl.GetAnchorPos(0))
            hwp.HAction.Run("SelectCtrlFront")

            # 첫 번째 셀 선택
            hwp.HAction.Run("ShapeObjTableSelCell")

            # 첫 셀의 list_id
            pos = hwp.GetPos()
            first_cell_list_id = pos[0]

            tables.append({
                'num': table_num,
                'ctrl': ctrl,
                'first_cell_list_id': first_cell_list_id
            })

            # 선택 해제
            hwp.HAction.Run("Cancel")
            hwp.HAction.Run("MoveParentList")

            table_num += 1

        ctrl = ctrl.Next

    return tables


def select_table(hwp, table_index: int) -> Optional[Any]:
    """
    특정 번호의 테이블 선택

    Args:
        hwp: HWP COM 객체
        table_index: 테이블 번호 (0부터 시작)

    Returns:
        ctrl: 선택된 테이블 컨트롤 (없으면 None)
    """
    ctrl = hwp.HeadCtrl
    current_index = 0

    while ctrl:
        if ctrl.CtrlID == "tbl":
            if current_index == table_index:
                hwp.SetPosBySet(ctrl.GetAnchorPos(0))
                if not hwp.HAction.Run("SelectCtrlFront"):
                    hwp.HAction.Run("SelectCtrlReverse")
                return ctrl
            current_index += 1
        ctrl = ctrl.Next

    return None


def enter_table(hwp, table_index: int) -> bool:
    """
    특정 번호의 테이블 첫 번째 셀로 진입

    Args:
        hwp: HWP COM 객체
        table_index: 테이블 번호 (0부터 시작)

    Returns:
        bool: 성공 여부
    """
    ctrl = select_table(hwp, table_index)
    if not ctrl:
        return False

    hwp.HAction.Run("ShapeObjTableSelCell")
    return True


def get_table_caption(hwp, table_index: int = None) -> str:
    """
    테이블 캡션 텍스트 가져오기

    캡션은 테이블 마지막 셀 다음 list_id를 가집니다.

    Args:
        hwp: HWP COM 객체
        table_index: 테이블 번호 (None이면 현재 테이블)

    Returns:
        str: 캡션 텍스트 (없으면 빈 문자열)
    """
    if table_index is not None:
        ctrl = select_table(hwp, table_index)
        if not ctrl:
            return ""
        hwp.HAction.Run("ShapeObjTableSelCell")

    try:
        # 테이블의 마지막 셀로 이동
        hwp.MovePos(MOVE_BOTTOM_OF_CELL, 0, 0)
        hwp.MovePos(MOVE_END_OF_CELL, 0, 0)

        last_cell_list_id = hwp.GetPos()[0]

        # 캡션 list_id = 마지막 셀 + 1
        caption_list_id = last_cell_list_id + 1

        # 캡션 위치로 이동 시도
        hwp.SetPos(caption_list_id, 0, 0)
        current_list_id = hwp.GetPos()[0]

        if current_list_id != caption_list_id:
            return ""

        # 캡션 텍스트 읽기
        hwp.HAction.Run("SelectAll")
        text = hwp.GetTextFile("TEXT", "")
        hwp.HAction.Run("Cancel")

        if text:
            return text.strip()

        return ""
    except Exception:
        return ""


def get_all_table_captions(hwp) -> List[Dict]:
    """
    문서 내 모든 테이블의 캡션 가져오기

    Args:
        hwp: HWP COM 객체

    Returns:
        list: [{
            'num': 테이블 번호,
            'first_cell_list_id': 첫 셀 ID,
            'caption': 캡션 텍스트
        }, ...]
    """
    captions = []
    tables = find_all_tables(hwp)

    for t in tables:
        caption = get_table_caption(hwp, t['num'])
        captions.append({
            'num': t['num'],
            'first_cell_list_id': t['first_cell_list_id'],
            'caption': caption
        })

    return captions


def has_caption(hwp, table_index: int = None) -> bool:
    """
    테이블에 캡션이 있는지 확인

    Args:
        hwp: HWP COM 객체
        table_index: 테이블 번호 (None이면 현재 테이블)

    Returns:
        bool: 캡션 존재 여부
    """
    caption = get_table_caption(hwp, table_index)
    return len(caption) > 0
