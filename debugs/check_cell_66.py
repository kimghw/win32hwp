# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, '.')
from cursor import get_hwp_instance
from table.table_info import TableInfo, MOVE_RIGHT_OF_CELL

hwp = get_hwp_instance()
table_info = TableInfo(hwp, debug=False)

# 48에서 오른쪽 이동 확인
hwp.SetPos(48, 0, 0)
print(f'48 위치 설정')
hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
next_id = hwp.GetPos()[0]
print(f'48 -> MoveRight -> {next_id}')

# 66 위치 확인
hwp.SetPos(66, 0, 0)
w, h = table_info.get_cell_dimensions()
print(f'66: width={w}, height={h}')

# 66에서 오른쪽 이동
hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
next_id = hwp.GetPos()[0]
print(f'66 -> MoveRight -> {next_id}')

# 65에서 오른쪽 이동
hwp.SetPos(65, 0, 0)
hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
next_id = hwp.GetPos()[0]
print(f'65 -> MoveRight -> {next_id}')
