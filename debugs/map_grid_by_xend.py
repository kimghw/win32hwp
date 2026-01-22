"""xend 기준 그리드 좌표 매핑 - 단일 순회"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor import get_hwp_instance
from table.table_info import TableInfo, MOVE_RIGHT_OF_CELL
from table.table_boundary import TableBoundary

TOLERANCE = 50


def map_grid_by_xend(hwp, start_cell: int, xend: int) -> dict:
    """
    xend 기준으로 그리드 좌표 매핑 (단일 순회)

    - cumulative_x > xend → 새 행 시작 (이전 셀이 마지막 열)
    - cumulative_x == xend → 현재 셀이 마지막 열

    Args:
        hwp: 한글 인스턴스
        start_cell: 시작 셀 list_id
        xend: 테이블 전체 가로 크기 (HWPUNIT)

    Returns:
        dict: {
            'grid': {list_id: {'row': r, 'col': c, 'width': w, 'height': h}},
            'max_row': 최대 행,
            'max_col': 최대 열
        }
    """
    table_info = TableInfo(hwp, debug=False)

    grid = {}
    row = 0
    col = 0
    cumulative_x = 0
    max_col = 0

    hwp.SetPos(start_cell, 0, 0)
    current_id = start_cell

    print(f"=== xend={xend} 기준 그리드 매핑 ===\n")
    print(f"[Row 0] 시작")

    max_iterations = 1000

    for _ in range(max_iterations):
        hwp.SetPos(current_id, 0, 0)
        w, h = table_info.get_cell_dimensions()
        cumulative_x += w

        # xend 초과 → 새 행 시작
        if cumulative_x > xend + TOLERANCE:
            # 현재 셀은 새 행의 첫 번째
            row += 1
            col = 0
            cumulative_x = w
            print(f"\n[Row {row}] 시작 (xend 초과)")

        # 이미 매핑된 셀인지 확인 (rowspan)
        if current_id not in grid:
            grid[current_id] = {
                'row': row,
                'col': col,
                'width': w,
                'height': h
            }
            print(f"  ({row}, {col}): id={current_id}, w={w}, cumul={cumulative_x}")
        else:
            # rowspan 셀
            print(f"  ({row}, {col}): id={current_id} (rowspan)")

        if col > max_col:
            max_col = col

        # xend 도달 → 다음 셀부터 새 행
        if abs(cumulative_x - xend) <= TOLERANCE:
            # 우측 이동
            hwp.SetPos(current_id, 0, 0)
            hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
            next_id = hwp.GetPos()[0]

            if next_id == current_id:
                print(f"  → 테이블 끝")
                break

            row += 1
            col = 0
            cumulative_x = 0
            print(f"\n[Row {row}] 시작 (xend 도달)")
            current_id = next_id
            continue

        # 우측 이동
        hwp.SetPos(current_id, 0, 0)
        hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
        next_id = hwp.GetPos()[0]

        if next_id == current_id:
            print(f"  → 테이블 끝")
            break

        col += 1
        current_id = next_id

    return {
        'grid': grid,
        'max_row': row,
        'max_col': max_col
    }


def main():
    hwp = get_hwp_instance()
    if not hwp:
        print("한글이 실행 중이 아닙니다")
        return

    original_pos = hwp.GetPos()

    # 경계 분석 (xend 계산용)
    boundary = TableBoundary(hwp, debug=False)
    boundary_result = boundary.check_boundary_table()

    # xend 계산
    xend = boundary._calc_xend_from_first_rows(boundary_result.first_rows)
    start_cell = boundary_result.first_rows[0]

    print(f"xend = {xend} HWPUNIT ({xend/100:.1f}pt)")
    print(f"start_cell = {start_cell}")
    print()

    # 그리드 매핑
    result = map_grid_by_xend(hwp, start_cell, xend)

    print(f"\n{'='*50}")
    print(f"=== 결과 ===")
    print(f"총 셀 수: {len(result['grid'])}")
    print(f"그리드 크기: {result['max_row'] + 1} 행 x {result['max_col'] + 1} 열")

    print(f"\n=== 상세 매핑 ===")
    for list_id, info in sorted(result['grid'].items(), key=lambda x: (x[1]['row'], x[1]['col'])):
        print(f"  id={list_id}: ({info['row']}, {info['col']}) w={info['width']}")

    hwp.SetPos(*original_pos)


if __name__ == "__main__":
    main()
