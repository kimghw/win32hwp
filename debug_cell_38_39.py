# -*- coding: utf-8 -*-
"""셀 38, 39 위치 확인"""

from table.table_boundary import TableBoundary
from table.table_info import TableInfo, MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL, MOVE_LEFT_OF_CELL, MOVE_UP_OF_CELL
from cursor import get_hwp_instance

hwp = get_hwp_instance()
table_info = TableInfo(hwp, debug=False)

# 셀 38, 39의 인접 셀 확인
for cell_id in [38, 39]:
    print(f"\n=== 셀 {cell_id} 인접 확인 ===")
    hwp.SetPos(cell_id, 0, 0)
    width, height = table_info.get_cell_dimensions()
    print(f"크기: width={width}, height={height}")

    # 왼쪽
    hwp.SetPos(cell_id, 0, 0)
    hwp.MovePos(MOVE_LEFT_OF_CELL, 0, 0)
    left_id, _, _ = hwp.GetPos()
    print(f"왼쪽: {left_id}")

    # 오른쪽
    hwp.SetPos(cell_id, 0, 0)
    hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
    right_id, _, _ = hwp.GetPos()
    print(f"오른쪽: {right_id}")

    # 위쪽
    hwp.SetPos(cell_id, 0, 0)
    hwp.MovePos(MOVE_UP_OF_CELL, 0, 0)
    up_id, _, _ = hwp.GetPos()
    print(f"위쪽: {up_id}")

    # 아래쪽
    hwp.SetPos(cell_id, 0, 0)
    hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
    down_id, _, _ = hwp.GetPos()
    print(f"아래쪽: {down_id}")
