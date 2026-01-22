# -*- coding: utf-8 -*-
"""오른쪽 이동 + 너비 누적으로 first_cols, last_cols 동시 추출

알고리즘:
1. first_rows의 너비 합 = xend
2. 테이블 첫 셀에서 시작, 계속 오른쪽 이동
3. xend 초과 시: 이전 셀 = last_col, 현재 셀 = 새 first_col
4. xend 도달 시: 현재 셀 = last_col
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor import get_hwp_instance
from table.table_info import TableInfo, MOVE_RIGHT_OF_CELL

TOLERANCE = 50


def calc_xend_from_first_rows(hwp, first_rows: list, debug: bool = False) -> int:
    """first_rows 셀들의 너비 합 = xend"""
    table_info = TableInfo(hwp, debug=False)

    if not first_rows:
        return 0

    first_rows_set = set(first_rows)
    start_cell = first_rows[0]
    hwp.SetPos(start_cell, 0, 0)

    cumulative_x = 0
    current_id = start_cell
    path = []

    while current_id in first_rows_set:
        w, _ = table_info.get_cell_dimensions()
        cumulative_x += w
        path.append(current_id)

        hwp.SetPos(current_id, 0, 0)
        hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
        next_id = hwp.GetPos()[0]

        if next_id == current_id:
            break

        current_id = next_id

    if debug:
        print(f"[xend] first_rows: {first_rows}")
        print(f"[xend] 순회 경로: {path}")
        print(f"[xend] xend = {cumulative_x} HWPUNIT")

    return cumulative_x


def find_cols_by_xend(hwp, start_cell: int, xend: int, debug: bool = True) -> dict:
    """
    단일 순회로 first_cols, last_cols 동시 추출

    - xend 초과 시: 이전 셀 = last_col, 현재 셀 = 새 first_col
    - xend 도달 시: 현재 셀 = last_col, 다음 셀 = first_col
    """
    table_info = TableInfo(hwp, debug=False)

    hwp.SetPos(start_cell, 0, 0)
    current_id = start_cell
    cumulative_x = 0
    first_cols = [start_cell]
    last_cols = []
    prev_id = None
    path = []

    max_iterations = 1000

    for i in range(max_iterations):
        hwp.SetPos(current_id, 0, 0)
        w, _ = table_info.get_cell_dimensions()
        cumulative_x += w
        path.append((current_id, cumulative_x))

        # xend 초과 → 이전 셀이 last_col, 현재 셀은 새 행의 first_col
        if cumulative_x > xend + TOLERANCE:
            if prev_id is not None:
                last_cols.append(prev_id)
                if debug:
                    print(f"  xend 초과: {current_id} (x={cumulative_x}) -> last_col={prev_id}, new first_col={current_id}")

            first_cols.append(current_id)
            cumulative_x = w  # 현재 셀 너비부터 다시 시작
            prev_id = None

        # xend 도달 (오차 범위 내)
        elif abs(cumulative_x - xend) <= TOLERANCE:
            last_cols.append(current_id)
            if debug:
                print(f"  xend 도달: {current_id} (x={cumulative_x}) -> last_col={current_id}")

            # 우측 이동
            hwp.SetPos(current_id, 0, 0)
            hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
            next_id = hwp.GetPos()[0]

            if next_id == current_id:
                if debug:
                    print(f"  이동 불가 -> 종료")
                break

            first_cols.append(next_id)
            if debug:
                print(f"  새 행 시작: first_col={next_id}")
            cumulative_x = 0
            prev_id = None
            current_id = next_id
            continue

        # 우측 이동
        prev_id = current_id
        hwp.SetPos(current_id, 0, 0)
        hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
        next_id = hwp.GetPos()[0]

        if next_id == current_id:
            last_cols.append(current_id)
            if debug:
                print(f"  이동 불가: {current_id} -> last_col={current_id}")
            break

        current_id = next_id

    return {
        'first_cols': first_cols,
        'last_cols': last_cols,
        'path': path,
    }


def main():
    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return

    table_info = TableInfo(hwp, debug=False)
    if not table_info.is_in_table():
        print("[오류] 커서가 테이블 내부에 있지 않습니다.")
        return

    # 기존 방식으로 first_rows, first_cols 가져오기
    from table.table_boundary import TableBoundary
    boundary = TableBoundary(hwp, debug=False)
    boundary_result = boundary.check_boundary_table()

    print(f"\n=== 기존 결과 ===")
    print(f"first_cols ({len(boundary_result.first_cols)}개): {sorted(boundary_result.first_cols)}")
    print(f"last_cols ({len(boundary_result.last_cols)}개): {sorted(boundary_result.last_cols)}")

    # 1단계: first_rows에서 xend 계산
    print(f"\n=== 새 방식: xend 계산 ===")
    xend = calc_xend_from_first_rows(hwp, boundary_result.first_rows, debug=True)

    # 2단계: 단일 순회로 first_cols, last_cols 찾기
    print(f"\n=== 새 방식: 단일 순회 ===")
    start_cell = boundary_result.table_origin
    result = find_cols_by_xend(hwp, start_cell, xend, debug=True)

    print(f"\n=== 새 방식 결과 ===")
    print(f"first_cols ({len(result['first_cols'])}개): {result['first_cols']}")
    print(f"last_cols ({len(result['last_cols'])}개): {result['last_cols']}")

    # 비교
    print(f"\n=== 비교 ===")
    old_first = set(boundary_result.first_cols)
    new_first = set(result['first_cols'])
    old_last = set(boundary_result.last_cols)
    new_last = set(result['last_cols'])

    print(f"first_cols 일치: {'O' if old_first == new_first else 'X'}")
    print(f"last_cols 일치: {'O' if old_last == new_last else 'X'}")

    if old_first != new_first:
        print(f"  기존에만: {sorted(old_first - new_first)}")
        print(f"  새로에만: {sorted(new_first - old_first)}")

    if old_last != new_last:
        print(f"  기존에만: {sorted(old_last - new_last)}")
        print(f"  새로에만: {sorted(new_last - old_last)}")


if __name__ == "__main__":
    main()
