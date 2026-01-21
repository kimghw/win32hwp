# -*- coding: utf-8 -*-
"""중복 셀 상세 분석"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from cursor import get_hwp_instance
from table.cell_position import CellPositionCalculator


def analyze_duplicates():
    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return

    print("=== 중복 셀 상세 분석 ===\n")

    calc = CellPositionCalculator(hwp, debug=False)

    try:
        result = calc.calculate(max_cells=2000)

        # 중복 점유 위치 찾기
        grid = {}
        for list_id, cell in result.cells.items():
            for r in range(cell.start_row, cell.end_row + 1):
                for c in range(cell.start_col, cell.end_col + 1):
                    key = (r, c)
                    if key not in grid:
                        grid[key] = []
                    grid[key].append(list_id)

        duplicates = {k: v for k, v in grid.items() if len(v) > 1}

        # 중복된 셀 ID들 수집
        dup_cell_ids = set()
        for ids in duplicates.values():
            dup_cell_ids.update(ids)

        print(f"중복에 관련된 고유 셀 수: {len(dup_cell_ids)}")
        print(f"중복 위치 수: {len(duplicates)}")

        # 각 중복 셀의 상세 정보
        print("\n=== 중복 셀 상세 정보 ===")
        for list_id in sorted(dup_cell_ids):
            cell = result.cells[list_id]
            print(f"\nlist_id={list_id}:")
            print(f"  그리드: ({cell.start_row},{cell.start_col}) ~ ({cell.end_row},{cell.end_col})")
            print(f"  span: {cell.rowspan}x{cell.colspan}")
            print(f"  물리좌표: x={cell.start_x}~{cell.end_x}, y={cell.start_y}~{cell.end_y}")

        # 중복 쌍 분석
        print("\n\n=== 중복 쌍 상세 분석 (처음 10개) ===")
        analyzed_pairs = set()
        count = 0
        for pos, ids in sorted(duplicates.items()):
            pair = tuple(sorted(ids))
            if pair in analyzed_pairs:
                continue
            analyzed_pairs.add(pair)
            count += 1
            if count > 10:
                break

            print(f"\n위치 {pos}: {ids}")
            for lid in ids:
                cell = result.cells[lid]
                print(f"  list_id={lid}:")
                print(f"    그리드: ({cell.start_row},{cell.start_col}) ~ ({cell.end_row},{cell.end_col})")
                print(f"    물리X: {cell.start_x} ~ {cell.end_x} (너비: {cell.end_x - cell.start_x})")
                print(f"    물리Y: {cell.start_y} ~ {cell.end_y} (높이: {cell.end_y - cell.start_y})")

        # X 레벨 분석 - 문제가 되는 영역
        print("\n\n=== X 레벨 상세 (col 9~20 영역) ===")
        for i, x in enumerate(result.x_levels):
            if 9 <= i <= 20:
                print(f"  col[{i}] = {x}")

        # Y 레벨 분석 - 문제가 되는 영역
        print("\n=== Y 레벨 상세 (row 26~35 영역) ===")
        for i, y in enumerate(result.y_levels):
            if 26 <= i <= 35:
                print(f"  row[{i}] = {y}")

    except ValueError as e:
        print(f"[오류] {e}")


if __name__ == "__main__":
    analyze_duplicates()
