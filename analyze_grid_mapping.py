# -*- coding: utf-8 -*-
"""그리드 매핑 정상 여부 확인"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from cursor import get_hwp_instance
from table.cell_position import CellPositionCalculator


def analyze_grid():
    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return

    print("=== 그리드 매핑 검증 ===\n")

    calc = CellPositionCalculator(hwp, debug=False)

    try:
        result = calc.calculate(max_cells=2000)

        print(f"총 셀: {len(result.cells)}개")
        print(f"테이블 크기: {result.max_row + 1}행 x {result.max_col + 1}열")
        print(f"예상 총 그리드 칸: {(result.max_row + 1) * (result.max_col + 1)}개")

        # 그리드 점유 상태 구축
        grid = {}  # (row, col) -> list_id
        overlap_count = 0

        for list_id, cell in result.cells.items():
            for r in range(cell.start_row, cell.end_row + 1):
                for c in range(cell.start_col, cell.end_col + 1):
                    key = (r, c)
                    if key in grid:
                        overlap_count += 1
                    else:
                        grid[key] = list_id

        # 점유된 칸 수
        occupied = len(grid)
        total_grid = (result.max_row + 1) * (result.max_col + 1)
        empty = total_grid - occupied

        print(f"\n점유된 그리드 칸: {occupied}개")
        print(f"빈 그리드 칸: {empty}개")
        print(f"중복 발생 횟수: {overlap_count}개")

        # 그리드 시각화 (처음 15행, 처음 20열)
        print("\n=== 그리드 시각화 (row 0~14, col 0~19) ===")
        print("    ", end="")
        for c in range(min(20, result.max_col + 1)):
            print(f"{c:3}", end="")
        print()
        print("    " + "---" * min(20, result.max_col + 1))

        for r in range(min(15, result.max_row + 1)):
            print(f"{r:2} |", end="")
            for c in range(min(20, result.max_col + 1)):
                if (r, c) in grid:
                    # 셀이 있으면 X 표시
                    print("  X", end="")
                else:
                    # 빈 칸이면 . 표시
                    print("  .", end="")
            print()

        # row=6,7 상세 분석
        print("\n=== row 6,7 상세 (col 0~35) ===")
        for r in [6, 7]:
            print(f"row={r}: ", end="")
            for c in range(result.max_col + 1):
                if (r, c) in grid:
                    print("X", end="")
                else:
                    print(".", end="")
            print()

        # col=35 상세 분석
        print("\n=== col=35 점유 상태 (row 0~72) ===")
        col35_status = []
        for r in range(result.max_row + 1):
            if (r, 35) in grid:
                col35_status.append("X")
            else:
                col35_status.append(".")

        # 10개씩 출력
        for i in range(0, len(col35_status), 10):
            chunk = col35_status[i:i+10]
            print(f"  row {i:2}~{i+len(chunk)-1:2}: {''.join(chunk)}")

        # 정상 여부 판단
        print("\n=== 결론 ===")
        if empty == 0 and overlap_count == 0:
            print("[OK] 그리드 매핑 정상 - 모든 칸이 정확히 1개 셀로 점유됨")
        else:
            print(f"[문제] 그리드 매핑 비정상")
            print(f"  - 빈 칸: {empty}개")
            print(f"  - 중복: {overlap_count}개")

    except ValueError as e:
        print(f"[오류] {e}")


if __name__ == "__main__":
    analyze_grid()
