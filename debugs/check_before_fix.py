# -*- coding: utf-8 -*-
"""_fix 함수 전후 비교"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor import get_hwp_instance
from table.table_info import TableInfo, MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL
from table.table_boundary import TableBoundary
from table.cell_position import CellRange
from collections import deque


def calculate_without_fix():
    """_fix 함수 없이 순수 매핑 결과"""
    hwp = get_hwp_instance()
    if not hwp:
        return None

    table_info = TableInfo(hwp, debug=False)
    boundary = TableBoundary(hwp, debug=False)

    if not table_info.is_in_table():
        return None

    TOLERANCE = 3

    def merge_close_levels(levels):
        if not levels:
            return []
        sorted_levels = sorted(levels)
        merged = [sorted_levels[0]]
        for level in sorted_levels[1:]:
            if level - merged[-1] > TOLERANCE:
                merged.append(level)
        return merged

    def find_level_index(value, levels):
        for idx, level in enumerate(levels):
            if abs(value - level) <= TOLERANCE:
                return idx
        return -1

    def find_end_level_index(value, levels):
        for idx, level in enumerate(levels):
            if abs(value - level) <= TOLERANCE:
                return max(0, idx - 1)
        for idx in range(len(levels) - 1, -1, -1):
            if levels[idx] < value - TOLERANCE:
                return idx
        return 0

    # 경계 분석
    boundary_result = boundary.check_boundary_table()
    first_cols = boundary._sort_first_cols_by_position(boundary_result.first_cols)
    last_cols_set = set(boundary_result.last_cols)
    first_cols_set = set(first_cols)

    # 셀 좌표 수집
    cell_positions = {}
    all_x_coords = set()
    all_y_coords = set()

    cumulative_y = 0

    for row_idx, row_start in enumerate(first_cols):
        hwp.SetPos(row_start, 0, 0)
        _, row_height = table_info.get_cell_dimensions()

        row_start_y = cumulative_y
        row_end_y = cumulative_y + row_height

        visited_in_row = set()
        queue = deque()
        queue.append((row_start, 0, row_start_y))
        visited_in_row.add(row_start)

        while queue:
            current_id, start_x, start_y = queue.popleft()

            hwp.SetPos(current_id, 0, 0)
            width, height = table_info.get_cell_dimensions()

            end_x = start_x + width
            end_y = start_y + height

            cell_positions[current_id] = {
                'start_x': start_x, 'end_x': end_x,
                'start_y': start_y, 'end_y': end_y,
            }

            all_x_coords.add(start_x)
            all_x_coords.add(end_x)
            all_y_coords.add(start_y)
            all_y_coords.add(end_y)

            # 오른쪽
            hwp.SetPos(current_id, 0, 0)
            hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
            right_id = hwp.GetPos()[0]
            if right_id != current_id and right_id not in visited_in_row and right_id not in first_cols_set:
                visited_in_row.add(right_id)
                queue.append((right_id, end_x, start_y))

            # 아래
            if end_y < row_end_y:
                hwp.SetPos(current_id, 0, 0)
                hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
                down_id = hwp.GetPos()[0]
                if down_id != current_id and down_id not in visited_in_row and down_id not in first_cols_set:
                    visited_in_row.add(down_id)
                    queue.append((down_id, start_x, end_y))

        cumulative_y = row_end_y

    # 경계 필터링
    table_max_x = max(cell_positions[lid]['end_x'] for lid in last_cols_set if lid in cell_positions)
    table_max_y = cumulative_y

    all_x_coords = {x for x in all_x_coords if -TOLERANCE <= x <= table_max_x + TOLERANCE}
    all_y_coords = {y for y in all_y_coords if -TOLERANCE <= y <= table_max_y + TOLERANCE}

    x_levels = merge_close_levels(list(all_x_coords))
    y_levels = merge_close_levels(list(all_y_coords))

    # 순수 그리드 매핑 (fix 없이)
    cells = {}
    for list_id, pos in cell_positions.items():
        start_row = find_level_index(pos['start_y'], y_levels)
        start_col = find_level_index(pos['start_x'], x_levels)
        end_row = find_end_level_index(pos['end_y'], y_levels)
        end_col = find_end_level_index(pos['end_x'], x_levels)

        if start_row < 0 or start_col < 0:
            continue

        cells[list_id] = {
            'start_row': start_row, 'start_col': start_col,
            'end_row': end_row, 'end_col': end_col,
            'start_y': pos['start_y'], 'end_y': pos['end_y'],
        }

    return cells, y_levels


def main():
    result = calculate_without_fix()
    if not result:
        print("오류")
        return

    cells, y_levels = result

    print(f"y_levels: {y_levels[:10]}...")
    print()

    print("=== _fix 전 순수 매핑 (처음 10개) ===")
    for list_id in sorted(cells.keys())[:10]:
        c = cells[list_id]
        print(f"list_id={list_id}: ({c['start_row']},{c['start_col']})~({c['end_row']},{c['end_col']})")
        print(f"  물리 y: [{c['start_y']}~{c['end_y']}]")


if __name__ == "__main__":
    main()
