# -*- coding: utf-8 -*-
"""
list_id 조회 모듈

HWP 문서의 list_id 관련 조회 및 좌표 변환 기능을 제공합니다.

list_id 설명:
- 0: 본문
- 1~9: 머리말/꼬리말/각주/미주
- 10 이상: 테이블 셀
"""

from typing import Optional, Tuple, Dict


def get_list_id(hwp) -> int:
    """
    현재 커서 위치의 list_id 반환

    Args:
        hwp: HWP COM 객체

    Returns:
        int: list_id (0=본문, 10+=셀)
    """
    pos = hwp.GetPos()
    return pos[0]


class ListIdMapper:
    """
    list_id와 (row, col) 좌표 간 변환을 위한 매퍼 클래스

    TableInfo와 연동하여 좌표 맵을 생성합니다.
    """

    def __init__(self, hwp):
        self.hwp = hwp
        self._coord_map: Dict[Tuple[int, int], int] = {}  # (row, col) -> list_id
        self._list_id_to_coord: Dict[int, Tuple[int, int]] = {}  # list_id -> (row, col)
        self._initialized = False

    def _ensure_initialized(self):
        """좌표 맵이 초기화되어 있는지 확인하고, 안 되어있으면 초기화"""
        if self._initialized:
            return

        try:
            from table.table_info import TableInfo
            table_info = TableInfo(self.hwp, debug=False)

            if table_info.is_in_table():
                table_info.collect_cells_bfs()
                self._coord_map = table_info.build_coordinate_map()
                self._list_id_to_coord = dict(table_info._representative_coords)
                self._initialized = True
        except Exception:
            pass

    def get_coord_from_list_id(self, list_id: int) -> Optional[Tuple[int, int]]:
        """
        list_id로부터 대표 좌표 반환

        병합 셀의 경우 가장 위-왼쪽 좌표를 반환합니다.

        Args:
            list_id: 셀의 list_id

        Returns:
            (row, col) 좌표 또는 None
        """
        self._ensure_initialized()
        return self._list_id_to_coord.get(list_id)

    def get_list_id_from_coord(self, row: int, col: int) -> Optional[int]:
        """
        (row, col) 좌표로부터 list_id 반환

        병합 셀의 경우 해당 좌표를 포함하는 셀의 list_id를 반환합니다.

        Args:
            row: 행 좌표 (0-based)
            col: 열 좌표 (0-based)

        Returns:
            list_id 또는 None
        """
        self._ensure_initialized()
        return self._coord_map.get((row, col))

    def refresh(self):
        """좌표 맵 새로고침"""
        self._initialized = False
        self._coord_map.clear()
        self._list_id_to_coord.clear()
        self._ensure_initialized()


# 편의 함수 (TableField와 호환)
def get_list_id_from_coord(hwp, row: int, col: int, coord_map: Dict = None) -> Optional[int]:
    """
    (row, col) 좌표로부터 list_id 반환

    Args:
        hwp: HWP COM 객체
        row: 행 좌표 (0-based)
        col: 열 좌표 (0-based)
        coord_map: 좌표 맵 (없으면 내부에서 생성)

    Returns:
        list_id 또는 None
    """
    if coord_map:
        return coord_map.get((row, col))

    mapper = ListIdMapper(hwp)
    return mapper.get_list_id_from_coord(row, col)


def get_coord_from_list_id(hwp, list_id: int, list_id_to_coord: Dict = None) -> Optional[Tuple[int, int]]:
    """
    list_id로부터 대표 좌표 반환

    Args:
        hwp: HWP COM 객체
        list_id: 셀의 list_id
        list_id_to_coord: list_id → 좌표 맵 (없으면 내부에서 생성)

    Returns:
        (row, col) 좌표 또는 None
    """
    if list_id_to_coord:
        return list_id_to_coord.get(list_id)

    mapper = ListIdMapper(hwp)
    return mapper.get_coord_from_list_id(list_id)
