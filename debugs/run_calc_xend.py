"""현재 열린 HWP 문서에서 xend 계산"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor import get_hwp_instance
from table.table_info import TableInfo, MOVE_RIGHT_OF_CELL
from table.table_boundary import TableBoundary

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
        print(f"[xend] xend = {cumulative_x / 100:.2f} pt")
        print(f"[xend] xend = {cumulative_x / 2834.6:.2f} cm")

    return cumulative_x


def main():
    hwp = get_hwp_instance()
    if not hwp:
        print("한글이 실행 중이 아닙니다")
        return

    # 현재 위치 저장
    original_pos = hwp.GetPos()

    print("=== 테이블 경계 분석 ===")
    boundary = TableBoundary(hwp, debug=False)
    boundary_result = boundary.check_boundary_table()

    if not boundary_result or not boundary_result.first_rows:
        print("테이블이 없거나 first_rows를 찾을 수 없습니다")
        return

    print(f"first_rows: {boundary_result.first_rows}")
    print(f"first_cols: {boundary_result.first_cols}")
    print()

    # xend 계산
    xend = calc_xend_from_first_rows(hwp, boundary_result.first_rows, debug=True)

    print(f"\n=== 결과 요약 ===")
    print(f"xend = {xend} HWPUNIT")
    print(f"xend = {xend / 100:.2f} pt")
    print(f"xend = {xend / 2834.6:.2f} cm")
    print(f"xend = {xend / 7200:.4f} inch")

    # 원래 위치로 복원
    hwp.SetPos(*original_pos)


if __name__ == "__main__":
    main()
