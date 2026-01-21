# -*- coding: utf-8 -*-
"""빈 셀 위치에서 인접 셀 확인 및 커서 이동으로 매핑"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from cursor import get_hwp_instance
from table.cell_position import CellPositionCalculator
from table.table_info import TableInfo, MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL, MOVE_LEFT_OF_CELL, MOVE_UP_OF_CELL


def find_missing_cells():
    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return

    print("=== 빈 셀 인접 셀 확인 및 커서 이동 매핑 ===\n")

    calc = CellPositionCalculator(hwp, debug=False)
    table_info = TableInfo(hwp, debug=False)

    try:
        result = calc.calculate(max_cells=2000)

        # 빈 위치 찾기
        grid = {}
        for list_id, cell in result.cells.items():
            for r in range(cell.start_row, cell.end_row + 1):
                for c in range(cell.start_col, cell.end_col + 1):
                    key = (r, c)
                    if key not in grid:
                        grid[key] = list_id

        empty_positions = []
        for r in range(result.max_row + 1):
            for c in range(result.max_col + 1):
                if (r, c) not in grid:
                    empty_positions.append((r, c))

        print(f"빈 위치: {empty_positions}")

        # 각 빈 위치에 대해 인접 셀 확인
        for empty_row, empty_col in empty_positions:
            print(f"\n=== 빈 위치 ({empty_row}, {empty_col}) 분석 ===")

            # 인접 셀 찾기 (왼쪽, 오른쪽, 위, 아래)
            adjacent = {}

            # 왼쪽 인접 셀
            for c in range(empty_col - 1, -1, -1):
                if (empty_row, c) in grid:
                    adjacent['left'] = grid[(empty_row, c)]
                    break

            # 오른쪽 인접 셀
            for c in range(empty_col + 1, result.max_col + 1):
                if (empty_row, c) in grid:
                    adjacent['right'] = grid[(empty_row, c)]
                    break

            # 위쪽 인접 셀
            for r in range(empty_row - 1, -1, -1):
                if (r, empty_col) in grid:
                    adjacent['up'] = grid[(r, empty_col)]
                    break

            # 아래쪽 인접 셀
            for r in range(empty_row + 1, result.max_row + 1):
                if (r, empty_col) in grid:
                    adjacent['down'] = grid[(r, empty_col)]
                    break

            print(f"인접 셀: {adjacent}")

            # 인접 셀 상세 정보
            for direction, list_id in adjacent.items():
                cell = result.cells[list_id]
                print(f"  {direction}: list_id={list_id}, "
                      f"grid=({cell.start_row},{cell.start_col})~({cell.end_row},{cell.end_col})")

            # 커서 이동으로 빈 위치의 실제 셀 찾기
            print(f"\n커서 이동으로 빈 위치 셀 탐색:")

            # 왼쪽 셀에서 오른쪽으로 이동
            if 'left' in adjacent:
                left_id = adjacent['left']
                hwp.SetPos(left_id, 0, 0)
                hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
                found_id = hwp.GetPos()[0]
                print(f"  왼쪽({left_id}) → 오른쪽 이동 → list_id={found_id}")

                if found_id != left_id and found_id not in result.cells:
                    print(f"    *** 새로운 셀 발견! list_id={found_id} ***")
                    # 새 셀 크기 확인
                    hwp.SetPos(found_id, 0, 0)
                    width, height = table_info.get_cell_dimensions()
                    print(f"    크기: {width} x {height}")

            # 위쪽 셀에서 아래로 이동
            if 'up' in adjacent:
                up_id = adjacent['up']
                hwp.SetPos(up_id, 0, 0)
                hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
                found_id = hwp.GetPos()[0]
                print(f"  위({up_id}) → 아래 이동 → list_id={found_id}")

                if found_id != up_id and found_id not in result.cells:
                    print(f"    *** 새로운 셀 발견! list_id={found_id} ***")
                    hwp.SetPos(found_id, 0, 0)
                    width, height = table_info.get_cell_dimensions()
                    print(f"    크기: {width} x {height}")

            # 오른쪽 셀에서 왼쪽으로 이동
            if 'right' in adjacent:
                right_id = adjacent['right']
                hwp.SetPos(right_id, 0, 0)
                hwp.MovePos(MOVE_LEFT_OF_CELL, 0, 0)
                found_id = hwp.GetPos()[0]
                print(f"  오른쪽({right_id}) → 왼쪽 이동 → list_id={found_id}")

                if found_id != right_id and found_id not in result.cells:
                    print(f"    *** 새로운 셀 발견! list_id={found_id} ***")
                    hwp.SetPos(found_id, 0, 0)
                    width, height = table_info.get_cell_dimensions()
                    print(f"    크기: {width} x {height}")

            # 아래 셀에서 위로 이동
            if 'down' in adjacent:
                down_id = adjacent['down']
                hwp.SetPos(down_id, 0, 0)
                hwp.MovePos(MOVE_UP_OF_CELL, 0, 0)
                found_id = hwp.GetPos()[0]
                print(f"  아래({down_id}) → 위 이동 → list_id={found_id}")

                if found_id != down_id and found_id not in result.cells:
                    print(f"    *** 새로운 셀 발견! list_id={found_id} ***")
                    hwp.SetPos(found_id, 0, 0)
                    width, height = table_info.get_cell_dimensions()
                    print(f"    크기: {width} x {height}")

    except ValueError as e:
        print(f"[오류] {e}")


if __name__ == "__main__":
    find_missing_cells()
