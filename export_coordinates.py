# -*- coding: utf-8 -*-
"""그리드 좌표와 테이블 좌표를 txt 파일로 출력"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from cursor import get_hwp_instance
from table.cell_position import CellPositionCalculator


def export_coordinates():
    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return

    print("=== 좌표 데이터 txt 파일 출력 ===\n")

    calc = CellPositionCalculator(hwp, debug=False)

    try:
        result = calc.calculate(max_cells=2000)

        # 1. 그리드 레벨 좌표 출력 (xline, yline)
        grid_file = "C:\\win32hwp\\grid_levels.txt"
        with open(grid_file, "w", encoding="utf-8") as f:
            f.write("=== 그리드 레벨 좌표 ===\n\n")

            f.write(f"테이블 크기: {result.max_row + 1}행 x {result.max_col + 1}열\n")
            f.write(f"X 레벨 개수: {len(result.x_levels)}\n")
            f.write(f"Y 레벨 개수: {len(result.y_levels)}\n\n")

            f.write("--- X 레벨 (열 경계선, HWPUNIT) ---\n")
            for i, x in enumerate(result.x_levels):
                f.write(f"col[{i:2d}] = {x:6d}\n")

            f.write("\n--- Y 레벨 (행 경계선, HWPUNIT) ---\n")
            for i, y in enumerate(result.y_levels):
                f.write(f"row[{i:2d}] = {y:6d}\n")

        print(f"[저장] {grid_file}")

        # 2. 셀 좌표 출력 (그리드 좌표 + 물리 좌표)
        cells_file = "C:\\win32hwp\\cell_coordinates.txt"
        with open(cells_file, "w", encoding="utf-8") as f:
            f.write("=== 셀 좌표 데이터 ===\n\n")

            f.write(f"총 셀 수: {len(result.cells)}\n")
            f.write(f"테이블 크기: {result.max_row + 1}행 x {result.max_col + 1}열\n\n")

            f.write("list_id | 그리드(row,col) | rowspan | colspan | 물리X (start~end) | 물리Y (start~end)\n")
            f.write("-" * 100 + "\n")

            for list_id in sorted(result.cells.keys()):
                cell = result.cells[list_id]
                f.write(f"{list_id:7d} | ({cell.start_row:2d},{cell.start_col:2d})~({cell.end_row:2d},{cell.end_col:2d}) | "
                        f"{cell.rowspan:7d} | {cell.colspan:7d} | "
                        f"{cell.start_x:6d}~{cell.end_x:6d} | "
                        f"{cell.start_y:6d}~{cell.end_y:6d}\n")

        print(f"[저장] {cells_file}")

        # 3. 그리드 맵 출력 (2D 배열 형태)
        gridmap_file = "C:\\win32hwp\\grid_map.txt"
        with open(gridmap_file, "w", encoding="utf-8") as f:
            f.write("=== 그리드 맵 (각 좌표의 list_id) ===\n\n")

            # 좌표 → list_id 맵 생성
            coord_map = calc.build_coord_to_listid_map(result)

            # 헤더
            f.write("     ")
            for col in range(result.max_col + 1):
                f.write(f"{col:4d}")
            f.write("\n")
            f.write("     " + "----" * (result.max_col + 1) + "\n")

            # 각 행
            for row in range(result.max_row + 1):
                f.write(f"{row:3d} |")
                for col in range(result.max_col + 1):
                    list_id = coord_map.get((row, col), 0)
                    if list_id:
                        f.write(f"{list_id:4d}")
                    else:
                        f.write("   .")
                f.write("\n")

        print(f"[저장] {gridmap_file}")

        print("\n[완료] 3개 파일 생성됨")

    except Exception as e:
        print(f"[오류] {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    export_coordinates()
