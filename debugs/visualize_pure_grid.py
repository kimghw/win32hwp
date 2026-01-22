# -*- coding: utf-8 -*-
"""순수 그리드 매핑 시각화 (_fix 없이)"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor import get_hwp_instance
from table.table_info import TableInfo, MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL
from table.table_boundary import TableBoundary
from collections import deque


def calculate_pure_grid():
    """_fix 함수 없이 순수 매핑"""
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

            hwp.SetPos(current_id, 0, 0)
            hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
            right_id = hwp.GetPos()[0]
            if right_id != current_id and right_id not in visited_in_row and right_id not in first_cols_set:
                visited_in_row.add(right_id)
                queue.append((right_id, end_x, start_y))

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

    # 순수 그리드 매핑
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
            'rowspan': end_row - start_row + 1,
            'colspan': end_col - start_col + 1,
        }

    max_row = max(c['end_row'] for c in cells.values()) if cells else 0
    max_col = max(c['end_col'] for c in cells.values()) if cells else 0

    return {
        'cells': cells,
        'x_levels': x_levels,
        'y_levels': y_levels,
        'max_row': max_row,
        'max_col': max_col,
    }


def visualize(data):
    import matplotlib.pyplot as plt
    import matplotlib.patches as patches
    from matplotlib import font_manager

    font_path = "C:/Windows/Fonts/malgun.ttf"
    if os.path.exists(font_path):
        font_prop = font_manager.FontProperties(fname=font_path)
        plt.rcParams['font.family'] = font_prop.get_name()
    plt.rcParams['axes.unicode_minus'] = False

    max_row = data['max_row']
    max_col = data['max_col']
    cells = data['cells']

    fig, ax = plt.subplots(figsize=(18, 22))
    colors = plt.cm.tab20.colors

    for i, (list_id, cell) in enumerate(sorted(cells.items())):
        x = cell['start_col']
        y = max_row - cell['end_row']
        w = cell['colspan']
        h = cell['rowspan']

        color = colors[i % len(colors)]

        rect = patches.Rectangle(
            (x, y), w, h,
            linewidth=1.5, edgecolor='black',
            facecolor=color, alpha=0.5
        )
        ax.add_patch(rect)

        center_x = x + w / 2
        center_y = y + h / 2

        if cell['rowspan'] > 1 or cell['colspan'] > 1:
            label = f"{list_id}\n({cell['start_row']},{cell['start_col']})\n~({cell['end_row']},{cell['end_col']})"
        else:
            label = f"{list_id}\n({cell['start_row']},{cell['start_col']})"

        fontsize = 6 if w < 1.5 or h < 1.5 else 7
        ax.text(center_x, center_y, label,
                ha='center', va='center',
                fontsize=fontsize, fontweight='bold')

    # 그리드 라인
    for col in range(max_col + 2):
        ax.axvline(x=col, color='gray', linewidth=0.5, alpha=0.3)
    for row in range(max_row + 2):
        ax.axhline(y=row, color='gray', linewidth=0.5, alpha=0.3)

    ax.set_xlim(-1, max_col + 2)
    ax.set_ylim(-1, max_row + 2)
    ax.set_aspect('equal')
    ax.set_xlabel('Column')
    ax.set_ylabel('Row (반전)')
    ax.set_title(f'순수 그리드 매핑 (_fix 없음)\n'
                f'{max_row+1}행 x {max_col+1}열, {len(cells)}개 셀',
                fontsize=13, fontweight='bold')

    plt.tight_layout()
    output_path = "C:\\win32hwp\\debugs\\pure_grid.png"
    plt.savefig(output_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"시각화 저장: {output_path}")


def main():
    data = calculate_pure_grid()
    if not data:
        print("오류")
        return

    print(f"그리드: {data['max_row']+1}행 x {data['max_col']+1}열")
    print(f"셀 수: {len(data['cells'])}개")
    visualize(data)


if __name__ == "__main__":
    main()
