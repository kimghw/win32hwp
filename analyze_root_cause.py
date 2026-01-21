# -*- coding: utf-8 -*-
"""중복/빈칸 근본 원인 분석"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from cursor import get_hwp_instance
from table.cell_position import CellPositionCalculator


def analyze_root_cause():
    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return

    print("=== 중복/빈칸 근본 원인 분석 ===\n")

    calc = CellPositionCalculator(hwp, debug=False)

    try:
        result = calc.calculate(max_cells=2000)

        # Y 레벨 출력 (row 0~10)
        print("=== Y 레벨 (행 경계선) - row 0~10 ===")
        for i, y in enumerate(result.y_levels[:11]):
            print(f"  row[{i}] = {y}")

        # 문제의 셀들 분석
        print("\n=== 문제의 셀 분석 ===")
        problem_cells = [7, 16]
        for lid in problem_cells:
            if lid in result.cells:
                cell = result.cells[lid]
                print(f"\nlist_id={lid}:")
                print(f"  그리드: ({cell.start_row},{cell.start_col}) ~ ({cell.end_row},{cell.end_col})")
                print(f"  물리Y: {cell.start_y} ~ {cell.end_y}")
                # Y 레벨에서 해당 좌표 찾기
                for i, y in enumerate(result.y_levels):
                    if abs(cell.start_y - y) <= 3:
                        print(f"  start_y({cell.start_y}) -> row[{i}] = {y}")
                    if abs(cell.end_y - y) <= 3:
                        print(f"  end_y({cell.end_y}) -> row[{i}] = {y}")

        # row=2 분석 - 빈 칸 원인
        print("\n=== row=2 분석 (빈 칸 원인) ===")
        row2_y_start = result.y_levels[2] if len(result.y_levels) > 2 else None
        row2_y_end = result.y_levels[3] if len(result.y_levels) > 3 else None
        print(f"row=2 Y 범위: {row2_y_start} ~ {row2_y_end}")

        row2_cells = [(lid, c) for lid, c in result.cells.items()
                      if c.start_row == 2]
        print(f"row=2 시작 셀: {len(row2_cells)}개")
        for lid, cell in sorted(row2_cells, key=lambda x: x[1].start_col):
            print(f"  list_id={lid}: col={cell.start_col}~{cell.end_col}, "
                  f"물리X={cell.start_x}~{cell.end_x}")

        # col=17~19 범위의 X 좌표 확인
        print("\n=== col=17~19 X 좌표 확인 ===")
        for c in [17, 18, 19, 20]:
            if c < len(result.x_levels):
                print(f"  col[{c}] = {result.x_levels[c]}")

        # row=2에서 col=17~19를 커버하는 셀 찾기
        col17_x = result.x_levels[17] if len(result.x_levels) > 17 else None
        col19_x = result.x_levels[19] if len(result.x_levels) > 19 else None
        print(f"\ncol=17~19 X 범위: {col17_x} ~ {col19_x}")

        print("\nrow=2에서 col=17~19 범위를 커버하는 셀:")
        for lid, cell in result.cells.items():
            if cell.start_row <= 2 <= cell.end_row:
                if cell.start_x <= col17_x and cell.end_x >= col19_x:
                    print(f"  list_id={lid}: col={cell.start_col}~{cell.end_col}, "
                          f"row={cell.start_row}~{cell.end_row}")

        # row=4에서 중복 원인 분석
        print("\n=== row=4에서 중복 원인 분석 ===")
        row4_y_start = result.y_levels[4] if len(result.y_levels) > 4 else None
        row4_y_end = result.y_levels[5] if len(result.y_levels) > 5 else None
        print(f"row=4 Y 범위: {row4_y_start} ~ {row4_y_end}")

        # 셀 7과 16의 Y 좌표가 row=4와 어떻게 매핑되는지 확인
        if 7 in result.cells and 16 in result.cells:
            cell7 = result.cells[7]
            cell16 = result.cells[16]
            print(f"\n셀 7: y={cell7.start_y}~{cell7.end_y}, grid_row={cell7.start_row}~{cell7.end_row}")
            print(f"셀 16: y={cell16.start_y}~{cell16.end_y}, grid_row={cell16.start_row}~{cell16.end_row}")

            # 겹치는 Y 범위 확인
            overlap_start = max(cell7.start_y, cell16.start_y)
            overlap_end = min(cell7.end_y, cell16.end_y)
            if overlap_start < overlap_end:
                print(f"\n물리적 Y 겹침: {overlap_start} ~ {overlap_end} (높이: {overlap_end - overlap_start})")
            else:
                print(f"\n물리적 Y 겹침 없음")

    except ValueError as e:
        print(f"[오류] {e}")


if __name__ == "__main__":
    analyze_root_cause()
