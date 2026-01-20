# -*- coding: utf-8 -*-
"""행 3 순회 디버그"""

from table.table_boundary import TableBoundary
from table.table_info import TableInfo, MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL
from cursor import get_hwp_instance

hwp = get_hwp_instance()
boundary = TableBoundary(hwp, debug=False)
table_info = TableInfo(hwp, debug=False)

result = boundary.check_boundary_table()
first_cols = boundary._sort_first_cols_by_position(result.first_cols)
last_cols_set = set(result.last_cols)
first_cols_set = set(first_cols)

print(f"first_cols: {first_cols}")
print(f"last_cols: {result.last_cols}")

# 행 3 (first_cols[2] = 15)
row_start = first_cols[2]  # 15
print(f"\n=== 행 3 순회 (시작: {row_start}) ===")

hwp.SetPos(row_start, 0, 0)
_, row_height = table_info.get_cell_dimensions()
print(f"행 높이: {row_height}")

current_id = row_start
visited = []

for i in range(50):  # 최대 50번
    if current_id in visited:
        print(f"이미 방문: {current_id}")
        break
    visited.append(current_id)

    hwp.SetPos(current_id, 0, 0)
    width, height = table_info.get_cell_dimensions()
    is_split = "분할됨" if height < row_height else ""
    print(f"  셀 {current_id}: width={width}, height={height} {is_split}")

    if current_id in last_cols_set:
        print(f"  -> last_col 도달, 종료")
        break

    hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
    next_id, _, _ = hwp.GetPos()

    if next_id == current_id:
        print(f"  -> 이동 불가, 종료")
        break

    if next_id in first_cols_set and next_id != row_start:
        print(f"  -> first_col {next_id} 만남, 종료")
        break

    current_id = next_id

print(f"\n방문 순서: {visited}")
print(f"38, 39 포함 여부: 38 in visited = {38 in visited}, 39 in visited = {39 in visited}")
