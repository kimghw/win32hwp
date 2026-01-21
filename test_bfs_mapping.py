# -*- coding: utf-8 -*-
"""BFS 방식 셀 좌표 매핑 테스트 - CellPositionCalculator.calculate_bfs() 사용"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from cursor import get_hwp_instance
from table.cell_position import CellPositionCalculator


def main():
    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return

    print("=== BFS 방식 셀 좌표 매핑 테스트 ===\n")

    calc = CellPositionCalculator(hwp, debug=True)

    try:
        result = calc.calculate_bfs()

        print(f"\n=== 결과 ===")
        print(f"총 셀: {len(result.cells)}개")
        print(f"X 레벨 ({len(result.x_levels)}개): {result.x_levels[:10]}{'...' if len(result.x_levels) > 10 else ''}")
        print(f"Y 레벨 ({len(result.y_levels)}개): {result.y_levels[:10]}{'...' if len(result.y_levels) > 10 else ''}")
        print(f"테이블 크기: {result.max_row + 1}행 x {result.max_col + 1}열")

        # 그리드 좌표 출력
        print(f"\n=== 그리드 좌표 (처음 30개) ===")
        for i, (list_id, cell) in enumerate(sorted(result.cells.items())):
            if i >= 30:
                print("...")
                break
            span_info = ""
            if cell.rowspan > 1 or cell.colspan > 1:
                span_info = f" [span: {cell.rowspan}x{cell.colspan}]"
            print(f"list_id={list_id}: (row={cell.start_row}, col={cell.start_col}){span_info}")

        # 병합 셀 정보
        merged = calc.get_merged_cells(result)
        if merged:
            print(f"\n=== 병합 셀 ({len(merged)}개) ===")
            for cell in sorted(merged, key=lambda c: (c.start_row, c.start_col))[:10]:
                print(f"  list_id={cell.list_id}: ({cell.start_row},{cell.start_col})~({cell.end_row},{cell.end_col}) [{cell.rowspan}x{cell.colspan}]")
            if len(merged) > 10:
                print(f"  ... 외 {len(merged) - 10}개")

    except ValueError as e:
        print(f"[오류] {e}")


if __name__ == "__main__":
    main()
