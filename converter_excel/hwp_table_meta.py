# -*- coding: utf-8 -*-
"""한글 표 메타 정보 추출 모듈"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor import get_hwp_instance
from table.table_info import TableInfo, MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL


def get_table_meta(hwp=None):
    """
    한글 표의 메타 정보 추출
    - 행 높이
    - 열 너비
    - 셀 정보
    """
    if hwp is None:
        hwp = get_hwp_instance()
        if hwp is None:
            print("[오류] 한글이 실행 중이지 않습니다.")
            return None

    table_info = TableInfo(hwp, debug=False)

    if not table_info.is_in_table():
        print("[오류] 커서가 테이블 내부에 있지 않습니다.")
        return None

    # 첫 셀로 이동
    table_info.move_to_first_cell()

    # 표 전체 정보
    ctrl = hwp.ParentCtrl
    if ctrl and ctrl.CtrlID == 'tbl':
        props = ctrl.Properties
        print("=== 표 전체 속성 ===")
        print(f"  CtrlID: {ctrl.CtrlID}")
        print(f"  UserDesc: {ctrl.UserDesc}")

        # Properties 항목 확인
        try:
            print(f"  Width: {props.Item('Width')}")
            print(f"  Height: {props.Item('Height')}")
            print(f"  RowCount: {props.Item('RowCount')}")
            print(f"  ColCount: {props.Item('ColCount')}")
        except:
            pass

    # 첫 번째 열 순회하며 행 높이 수집
    print("\n=== 행 높이 (HWPUNIT) ===")
    table_info.move_to_first_cell()
    row_heights = []
    row_idx = 0

    while True:
        list_id = hwp.GetPos()[0]
        width, height = table_info.get_cell_dimensions()
        row_heights.append(height)

        # 포인트 변환 (1pt = 100 HWPUNIT)
        height_pt = height / 100
        print(f"  행 {row_idx}: {height} HWPUNIT ({height_pt:.2f} pt)")

        row_idx += 1

        # 아래로 이동
        hwp.SetPos(list_id, 0, 0)
        hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
        next_id = hwp.GetPos()[0]

        if next_id == list_id:
            break

    # 첫 번째 행 순회하며 열 너비 수집
    print("\n=== 열 너비 (HWPUNIT) ===")
    table_info.move_to_first_cell()
    col_widths = []
    col_idx = 0

    while True:
        list_id = hwp.GetPos()[0]
        width, height = table_info.get_cell_dimensions()
        col_widths.append(width)

        # 포인트 변환
        width_pt = width / 100
        print(f"  열 {col_idx}: {width} HWPUNIT ({width_pt:.2f} pt)")

        col_idx += 1

        # 우측으로 이동
        hwp.SetPos(list_id, 0, 0)
        hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
        next_id = hwp.GetPos()[0]

        if next_id == list_id:
            break

    return {
        'row_heights': row_heights,
        'col_widths': col_widths,
        'row_count': len(row_heights),
        'col_count': len(col_widths),
    }


def get_all_cells_meta(hwp=None, max_cells=1000):
    """모든 셀의 메타 정보 추출"""
    if hwp is None:
        hwp = get_hwp_instance()
        if hwp is None:
            return None

    table_info = TableInfo(hwp, debug=False)

    if not table_info.is_in_table():
        return None

    table_info.move_to_first_cell()
    first_id = hwp.GetPos()[0]

    cells = {}
    visited = set()

    # BFS로 모든 셀 순회
    from collections import deque
    queue = deque([first_id])
    visited.add(first_id)

    while queue and len(visited) < max_cells:
        cur_id = queue.popleft()

        hwp.SetPos(cur_id, 0, 0)
        width, height = table_info.get_cell_dimensions()

        cells[cur_id] = {
            'width': width,
            'height': height,
            'width_pt': width / 100,
            'height_pt': height / 100,
        }

        # 4방향 탐색
        for move_cmd in [MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL]:
            hwp.SetPos(cur_id, 0, 0)
            hwp.MovePos(move_cmd, 0, 0)
            next_id = hwp.GetPos()[0]

            if next_id != cur_id and next_id not in visited:
                visited.add(next_id)
                queue.append(next_id)

    return cells


if __name__ == "__main__":
    meta = get_table_meta()

    if meta:
        print(f"\n=== 요약 ===")
        print(f"행 수: {meta['row_count']}")
        print(f"열 수: {meta['col_count']}")
        print(f"총 행 높이: {sum(meta['row_heights'])} HWPUNIT")
        print(f"총 열 너비: {sum(meta['col_widths'])} HWPUNIT")
