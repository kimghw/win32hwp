# -*- coding: utf-8 -*-
"""중복 위치의 실제 소유 셀 확인"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from cursor import get_hwp_instance
from table.cell_position import CellPositionCalculator
from table.table_info import TableInfo, MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL, MOVE_LEFT_OF_CELL, MOVE_UP_OF_CELL


def check_overlap_owner():
    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return

    print("=== 중복 위치 실제 소유 셀 확인 ===\n")

    calc = CellPositionCalculator(hwp, debug=False)
    table_info = TableInfo(hwp, debug=False)

    try:
        result = calc.calculate(max_cells=2000)

        # 중복 위치: (4, 17), (4, 18), (4, 19)
        # 셀 7: row=2~4, col=17~19
        # 셀 16: row=4~5, col=17~26

        cell7 = result.cells[7]
        cell16 = result.cells[16]

        print(f"셀 7: grid=({cell7.start_row},{cell7.start_col})~({cell7.end_row},{cell7.end_col})")
        print(f"       y={cell7.start_y}~{cell7.end_y}")
        print(f"셀 16: grid=({cell16.start_row},{cell16.start_col})~({cell16.end_row},{cell16.end_col})")
        print(f"        y={cell16.start_y}~{cell16.end_y}")

        # row=4의 Y 범위 확인
        row4_y = result.y_levels[4] if len(result.y_levels) > 4 else None
        row5_y = result.y_levels[5] if len(result.y_levels) > 5 else None
        print(f"\nrow=4 Y 범위: {row4_y} ~ {row5_y}")

        # 중복 위치에서 커서 이동으로 실제 셀 확인
        print("\n=== 중복 위치 커서 이동 분석 ===")

        # 셀 7에서 오른쪽으로 이동하면?
        hwp.SetPos(7, 0, 0)
        hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
        right_of_7 = hwp.GetPos()[0]
        print(f"셀 7 → 오른쪽 = list_id={right_of_7}")

        # 셀 7에서 아래로 이동하면?
        hwp.SetPos(7, 0, 0)
        hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
        down_of_7 = hwp.GetPos()[0]
        print(f"셀 7 → 아래 = list_id={down_of_7}")

        # 셀 16에서 위로 이동하면?
        hwp.SetPos(16, 0, 0)
        hwp.MovePos(MOVE_UP_OF_CELL, 0, 0)
        up_of_16 = hwp.GetPos()[0]
        print(f"셀 16 → 위 = list_id={up_of_16}")

        # 셀 16에서 왼쪽으로 이동하면?
        hwp.SetPos(16, 0, 0)
        hwp.MovePos(MOVE_LEFT_OF_CELL, 0, 0)
        left_of_16 = hwp.GetPos()[0]
        print(f"셀 16 → 왼쪽 = list_id={left_of_16}")

        # row=3에서 col=17~19 셀 찾기
        print("\n=== row=3 영역 분석 ===")
        # 셀 13 (row=3, col=5~16)에서 오른쪽으로 이동
        if 13 in result.cells:
            hwp.SetPos(13, 0, 0)
            hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
            right_of_13 = hwp.GetPos()[0]
            print(f"셀 13 (row=3, col=5~16) → 오른쪽 = list_id={right_of_13}")

        # 셀 10 (row=3, col=20~26)에서 왼쪽으로 이동
        if 10 in result.cells:
            hwp.SetPos(10, 0, 0)
            hwp.MovePos(MOVE_LEFT_OF_CELL, 0, 0)
            left_of_10 = hwp.GetPos()[0]
            print(f"셀 10 (row=3, col=20~26) → 왼쪽 = list_id={left_of_10}")

        # 결론
        print("\n=== 결론 ===")
        print(f"셀 7은 row={cell7.start_row}~{cell7.end_row}를 차지")
        print(f"셀 16은 row={cell16.start_row}~{cell16.end_row}를 차지")

        # 실제 row=4에서 col=17~19를 차지하는 것은 누구?
        # 셀 16 위로 이동하면 어디로 가는지로 판단
        if up_of_16 == 7:
            print("\n→ 셀 16 위에 셀 7이 있음 (인접)")
            print("→ 셀 7은 row=2~3, 셀 16은 row=4~5가 되어야 함")
        elif up_of_16 == 16:
            print("\n→ 셀 16 위로 이동해도 셀 16 (병합 셀)")

    except Exception as e:
        print(f"[오류] {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    check_overlap_owner()
