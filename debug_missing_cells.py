# -*- coding: utf-8 -*-
"""누락된 셀 확인 - 38, 39가 왜 빠졌는지"""

from table.table_boundary import TableBoundary
from table.table_info import TableInfo, MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL
from cursor import get_hwp_instance

hwp = get_hwp_instance()
if not hwp:
    print('[오류] 한글이 실행 중이지 않습니다.')
    exit(1)

table_info = TableInfo(hwp, debug=False)

# 셀 37에서 아래로 이동 테스트
print("=== 셀 37에서 아래로 이동 테스트 ===")
hwp.SetPos(37, 0, 0)
width, height = table_info.get_cell_dimensions()
print(f"셀 37: width={width}, height={height}")

hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
next_id, _, _ = hwp.GetPos()
print(f"37에서 아래로 이동 -> {next_id}")

# 셀 38, 39 직접 확인
print("\n=== 셀 38, 39 직접 확인 ===")
for cell_id in [38, 39, 40]:
    try:
        hwp.SetPos(cell_id, 0, 0)
        current, _, _ = hwp.GetPos()
        if current == cell_id:
            width, height = table_info.get_cell_dimensions()
            print(f"셀 {cell_id}: width={width}, height={height}")
        else:
            print(f"셀 {cell_id}: SetPos 후 위치가 {current}로 변경됨")
    except Exception as e:
        print(f"셀 {cell_id}: 오류 - {e}")

# 셀 40에서 위로 이동 테스트
print("\n=== 셀 40에서 위로 이동 테스트 ===")
from table.table_info import MOVE_UP_OF_CELL
hwp.SetPos(40, 0, 0)
hwp.MovePos(MOVE_UP_OF_CELL, 0, 0)
prev_id, _, _ = hwp.GetPos()
print(f"40에서 위로 이동 -> {prev_id}")
