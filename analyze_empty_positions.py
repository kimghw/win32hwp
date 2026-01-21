# -*- coding: utf-8 -*-
"""빈 위치 원인 분석"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from cursor import get_hwp_instance
from table.cell_position import CellPositionCalculator


def analyze_empty():
    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return

    print("=== 빈 위치 원인 분석 ===\n")

    calc = CellPositionCalculator(hwp, debug=False)

    try:
        result = calc.calculate(max_cells=2000)

        # 그리드 점유 상태 구축
        grid = {}
        for list_id, cell in result.cells.items():
            for r in range(cell.start_row, cell.end_row + 1):
                for c in range(cell.start_col, cell.end_col + 1):
                    key = (r, c)
                    if key not in grid:
                        grid[key] = []
                    grid[key].append(list_id)

        # 빈 위치 수집
        empty_positions = []
        for r in range(result.max_row + 1):
            for c in range(result.max_col + 1):
                if (r, c) not in grid:
                    empty_positions.append((r, c))

        print(f"빈 위치 수: {len(empty_positions)}")
        print(f"테이블 크기: {result.max_row + 1}행 x {result.max_col + 1}열")

        # col=35가 비어있는 패턴 분석
        print("\n=== col=35 분석 ===")
        col35_empty = [(r, c) for r, c in empty_positions if c == 35]
        print(f"col=35에서 빈 행: {len(col35_empty)}개")

        # col=35에 있는 셀 확인
        col35_cells = [(lid, c) for lid, c in result.cells.items() if c.start_col <= 35 <= c.end_col]
        print(f"col=35을 포함하는 셀: {len(col35_cells)}개")
        for lid, cell in sorted(col35_cells, key=lambda x: x[1].start_row)[:10]:
            print(f"  list_id={lid}: row={cell.start_row}~{cell.end_row}, col={cell.start_col}~{cell.end_col}")

        # X 레벨 최대값 확인
        print(f"\n최대 X 레벨: {result.x_levels[-1]}")
        print(f"col=35 X 레벨: {result.x_levels[35] if len(result.x_levels) > 35 else 'N/A'}")

        # row=0~5, col=35 분석
        print("\n=== row=0~5, col=35 주변 분석 ===")
        for r in range(6):
            row_cells = [(lid, c) for lid, c in result.cells.items()
                         if c.start_row <= r <= c.end_row]
            max_col_in_row = max([c.end_col for _, c in row_cells]) if row_cells else -1
            print(f"  row={r}: 최대 col={max_col_in_row}, 빈칸={(r, 35) in empty_positions}")

        # 행 0의 셀 상세
        print("\n=== row=0 셀 상세 ===")
        row0_cells = [(lid, c) for lid, c in result.cells.items() if c.start_row == 0]
        for lid, cell in sorted(row0_cells, key=lambda x: x[1].start_col):
            print(f"  list_id={lid}: col={cell.start_col}~{cell.end_col} (colspan={cell.colspan})")
            print(f"    물리X: {cell.start_x}~{cell.end_x}")

        # col=15, 16 분석 (row=6,7에서 빈 칸)
        print("\n=== col=15,16 분석 (row 6,7) ===")
        for c in [15, 16]:
            col_cells = [(lid, cell) for lid, cell in result.cells.items()
                         if cell.start_col <= c <= cell.end_col]
            print(f"\ncol={c}을 포함하는 셀:")
            for lid, cell in sorted(col_cells, key=lambda x: x[1].start_row)[:15]:
                print(f"  list_id={lid}: row={cell.start_row}~{cell.end_row}")

    except ValueError as e:
        print(f"[오류] {e}")


if __name__ == "__main__":
    analyze_empty()
