# -*- coding: utf-8 -*-
"""
calculate_grid() 테스트 스크립트

xline/yline 기반 그리드 레벨 생성 방식 테스트
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from cursor import get_hwp_instance
from table.cell_position import CellPositionCalculator


def print_grid_visual(result, calc):
    """그리드를 시각적으로 출력"""
    print(f"\n=== 그리드 시각화 ({result.max_row + 1}행 x {result.max_col + 1}열) ===")

    # 좌표 -> list_id 맵 생성
    coord_map = calc.build_coord_to_listid_map(result)

    # 헤더 출력
    header = "     "
    for col in range(result.max_col + 1):
        header += f"  {col:2d}  "
    print(header)
    print("     " + "-----" * (result.max_col + 1))

    # 각 행 출력
    for row in range(result.max_row + 1):
        row_str = f" {row:2d} |"
        for col in range(result.max_col + 1):
            list_id = coord_map.get((row, col))
            if list_id:
                # 해당 셀의 대표 좌표인지 확인
                cell = result.cells.get(list_id)
                if cell and cell.start_row == row and cell.start_col == col:
                    row_str += f" {list_id:3d} "
                else:
                    row_str += "  ·  "  # 병합 셀의 일부
            else:
                row_str += "  -  "  # 빈 셀
        print(row_str)


def compare_methods(hwp, debug=False):
    """calculate_grid()와 calculate_bfs() 결과 비교"""
    calc = CellPositionCalculator(hwp, debug=debug)

    print("=" * 60)
    print("calculate_grid() 실행 (xline/yline 기반)")
    print("=" * 60)

    result_grid = calc.calculate_grid()
    calc.print_summary(result_grid)
    print_grid_visual(result_grid, calc)

    print("\n" + "=" * 60)
    print("calculate_bfs() 실행 (BFS 기반 - 레거시)")
    print("=" * 60)

    result_bfs = calc.calculate_bfs()
    calc.print_summary(result_bfs)
    print_grid_visual(result_bfs, calc)

    # 비교
    print("\n" + "=" * 60)
    print("비교 결과")
    print("=" * 60)

    grid_cells = set(result_grid.cells.keys())
    bfs_cells = set(result_bfs.cells.keys())

    only_in_grid = grid_cells - bfs_cells
    only_in_bfs = bfs_cells - grid_cells

    print(f"calculate_grid 셀 수: {len(grid_cells)}")
    print(f"calculate_bfs 셀 수: {len(bfs_cells)}")

    if only_in_grid:
        print(f"\ncalculate_grid에만 있는 셀: {sorted(only_in_grid)}")
    if only_in_bfs:
        print(f"\ncalculate_bfs에만 있는 셀: {sorted(only_in_bfs)}")

    if not only_in_grid and not only_in_bfs:
        print("\n두 방법의 셀 집합이 동일합니다.")

    # 좌표 비교
    print("\n좌표 차이:")
    common_cells = grid_cells & bfs_cells
    diff_count = 0
    for list_id in sorted(common_cells):
        cell_grid = result_grid.cells[list_id]
        cell_bfs = result_bfs.cells[list_id]

        if (cell_grid.start_row != cell_bfs.start_row or
            cell_grid.start_col != cell_bfs.start_col or
            cell_grid.end_row != cell_bfs.end_row or
            cell_grid.end_col != cell_bfs.end_col):
            diff_count += 1
            print(f"  셀 {list_id}:")
            print(f"    grid: ({cell_grid.start_row},{cell_grid.start_col})~({cell_grid.end_row},{cell_grid.end_col})")
            print(f"    bfs:  ({cell_bfs.start_row},{cell_bfs.start_col})~({cell_bfs.end_row},{cell_bfs.end_col})")

    if diff_count == 0:
        print("  좌표 차이 없음")


def main():
    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return

    calc = CellPositionCalculator(hwp, debug=False)

    try:
        # 테이블 내부인지 확인
        from table.table_info import TableInfo
        table_info = TableInfo(hwp)
        if not table_info.is_in_table():
            print("[오류] 커서가 테이블 내부에 있지 않습니다.")
            return

        # 두 방법 비교
        compare_methods(hwp, debug=False)

    except Exception as e:
        print(f"[오류] {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
