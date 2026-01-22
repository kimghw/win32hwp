"""xend 초과 시점으로 마지막 셀(last_col) 찾기
- 그냥 우측으로 계속 이동
- xend 초과하면: 이전 셀 = last_col, 현재 셀 = first_col, cumulative 리셋
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor import get_hwp_instance
from table.table_info import TableInfo, MOVE_RIGHT_OF_CELL

XEND = 42319
TOLERANCE = 50

def find_lastcols_by_traversal(hwp, start_cell: int, xend: int):
    table_info = TableInfo(hwp, debug=False)

    hwp.SetPos(start_cell, 0, 0)
    current_id = start_cell
    cumulative_x = 0
    last_cols = []
    first_cols = [start_cell]
    row_num = 1
    prev_id = None

    print(f"=== xend={xend} 기준 순회 ===\n")
    print(f"[Row {row_num}] 시작: {current_id}")

    iteration = 0
    max_iterations = 500

    while iteration < max_iterations:
        iteration += 1

        hwp.SetPos(current_id, 0, 0)
        w, h = table_info.get_cell_dimensions()
        cumulative_x += w

        print(f"  [{iteration}] cell={current_id}, w={w}, cumul={cumulative_x}")

        # xend 초과 → 이전 셀이 last_col, 현재 셀은 새 행의 first_col
        if cumulative_x > xend + TOLERANCE:
            if prev_id is not None:
                last_cols.append(prev_id)
                print(f"  >> xend 초과! last_col={prev_id}")

            first_cols.append(current_id)
            row_num += 1
            print(f"\n[Row {row_num}] 시작: {current_id}")

            # 리셋: 현재 셀 너비부터 다시 시작
            cumulative_x = w
            prev_id = None

        # xend 도달 (오차 범위 내)
        elif abs(cumulative_x - xend) <= TOLERANCE:
            last_cols.append(current_id)
            print(f"  >> xend 도달! last_col={current_id}")

            # 우측 이동
            hwp.SetPos(current_id, 0, 0)
            hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
            next_id = hwp.GetPos()[0]

            if next_id == current_id:
                print(f"\n[종료] 더 이상 우측 이동 불가")
                break

            first_cols.append(next_id)
            row_num += 1
            cumulative_x = 0
            prev_id = None
            current_id = next_id
            print(f"\n[Row {row_num}] 시작: {current_id}")
            continue

        # 우측 이동
        prev_id = current_id
        hwp.SetPos(current_id, 0, 0)
        hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
        next_id = hwp.GetPos()[0]

        if next_id == current_id:
            last_cols.append(current_id)
            print(f"  >> 우측 끝! last_col={current_id}")
            break

        current_id = next_id

    print(f"\n총 {iteration}회 반복")

    return {
        'last_cols': last_cols,
        'first_cols': first_cols,
    }


def main():
    hwp = get_hwp_instance()
    if not hwp:
        print("한글이 실행 중이 아닙니다")
        return

    original_pos = hwp.GetPos()

    start_cell = 2
    result = find_lastcols_by_traversal(hwp, start_cell, XEND)

    print(f"\n{'='*50}")
    print(f"=== 결과 ===")
    print(f"xend: {XEND} HWPUNIT")
    print(f"first_cols ({len(result['first_cols'])}): {result['first_cols']}")
    print(f"last_cols  ({len(result['last_cols'])}): {result['last_cols']}")

    hwp.SetPos(*original_pos)


if __name__ == "__main__":
    main()
