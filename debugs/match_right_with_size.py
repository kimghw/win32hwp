# -*- coding: utf-8 -*-
"""오른쪽 순회 + 셀 크기 기반 그리드 매칭

알고리즘:
1. x_levels, y_levels 먼저 수집
2. first_cols[0]에서 시작
3. 오른쪽 이동하면서:
   - 방문 안한 셀: width/height로 colspan/rowspan 계산 → 그리드 매칭
   - 방문한 셀: 스킵하고 계속 오른쪽
4. last_cols(xend) 만나면 → 다음 first_cols(xstart)로
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor import get_hwp_instance
from table.table_info import TableInfo, MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL
from table.table_boundary import TableBoundary
from collections import deque

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


def collect_levels(hwp, table_info, boundary):
    """1단계: x_levels, y_levels 수집"""
    boundary_result = boundary.check_boundary_table()
    first_cols = boundary._sort_first_cols_by_position(boundary_result.first_cols)
    last_cols_set = set(boundary_result.last_cols)
    first_cols_set = set(first_cols)

    cell_positions = {}
    all_x = set()
    all_y = set()

    cumulative_y = 0

    for row_start in first_cols:
        hwp.SetPos(row_start, 0, 0)
        _, row_height = table_info.get_cell_dimensions()

        row_start_y = cumulative_y
        row_end_y = cumulative_y + row_height

        visited = set()
        queue = deque([(row_start, 0, row_start_y)])
        visited.add(row_start)

        while queue:
            cid, sx, sy = queue.popleft()
            hwp.SetPos(cid, 0, 0)
            w, h = table_info.get_cell_dimensions()
            ex, ey = sx + w, sy + h

            cell_positions[cid] = {'sx': sx, 'ex': ex, 'sy': sy, 'ey': ey}
            all_x.update([sx, ex])
            all_y.update([sy, ey])

            # 오른쪽
            hwp.SetPos(cid, 0, 0)
            hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
            rid = hwp.GetPos()[0]
            if rid != cid and rid not in visited and rid not in first_cols_set:
                visited.add(rid)
                queue.append((rid, ex, sy))

            # 아래
            if ey < row_end_y:
                hwp.SetPos(cid, 0, 0)
                hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
                did = hwp.GetPos()[0]
                if did != cid and did not in visited and did not in first_cols_set:
                    visited.add(did)
                    queue.append((did, sx, ey))

        cumulative_y = row_end_y

    # 경계 필터링
    table_max_x = max(cell_positions[lid]['ex'] for lid in last_cols_set if lid in cell_positions)
    table_max_y = cumulative_y

    all_x = {x for x in all_x if -TOLERANCE <= x <= table_max_x + TOLERANCE}
    all_y = {y for y in all_y if -TOLERANCE <= y <= table_max_y + TOLERANCE}

    x_levels = merge_close_levels(list(all_x))
    y_levels = merge_close_levels(list(all_y))

    return x_levels, y_levels, cell_positions, first_cols, last_cols_set


def match_right_with_size():
    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return None

    table_info = TableInfo(hwp, debug=False)
    boundary = TableBoundary(hwp, debug=False)

    if not table_info.is_in_table():
        print("[오류] 커서가 테이블 내부에 있지 않습니다.")
        return None

    # 1단계: levels 수집
    x_levels, y_levels, cell_positions, first_cols, last_cols_set = \
        collect_levels(hwp, table_info, boundary)

    print(f"x_levels: {len(x_levels)}개")
    print(f"y_levels: {len(y_levels)}개")
    print(f"first_cols: {len(first_cols)}개")
    print()

    # 2단계: 오른쪽 순회하면서 그리드 매칭
    matched = []
    visited = set()

    for row_idx, row_start_id in enumerate(first_cols):
        print(f"=== Row {row_idx} (first_col: {row_start_id}) ===")

        current_id = row_start_id

        while True:
            if current_id in visited:
                print(f"  [{current_id}] 이미 방문 → 스킵")
            else:
                visited.add(current_id)

                # 셀 위치 가져오기
                pos = cell_positions.get(current_id)
                if pos:
                    # 그리드 좌표 계산
                    start_col = find_level_index(pos['sx'], x_levels)
                    end_col = find_end_level_index(pos['ex'], x_levels)
                    start_row = find_level_index(pos['sy'], y_levels)
                    end_row = find_end_level_index(pos['ey'], y_levels)

                    colspan = end_col - start_col + 1
                    rowspan = end_row - start_row + 1

                    matched.append({
                        'list_id': current_id,
                        'start_row': start_row,
                        'start_col': start_col,
                        'end_row': end_row,
                        'end_col': end_col,
                        'rowspan': rowspan,
                        'colspan': colspan,
                    })

                    if rowspan > 1 or colspan > 1:
                        print(f"  [{current_id}] → ({start_row},{start_col})~({end_row},{end_col}) [병합 {rowspan}x{colspan}]")
                    else:
                        print(f"  [{current_id}] → ({start_row},{start_col})")

            # last_cols면 다음 행으로
            if current_id in last_cols_set:
                print(f"  [{current_id}] last_col → 다음 행")
                break

            # 오른쪽으로 이동
            hwp.SetPos(current_id, 0, 0)
            hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
            next_id = hwp.GetPos()[0]

            if next_id == current_id:
                print(f"  [{current_id}] 오른쪽 이동 불가 → 끝")
                break

            current_id = next_id

        print()

    max_row = max(m['end_row'] for m in matched) if matched else 0
    max_col = max(m['end_col'] for m in matched) if matched else 0

    print(f"=== 매칭 완료: {len(matched)}개 셀, {max_row+1}행 x {max_col+1}열 ===")

    return {
        'cells': matched,
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

    for i, m in enumerate(cells):
        x = m['start_col']
        y = max_row - m['end_row']  # Y 반전
        w = m['colspan']
        h = m['rowspan']

        color = colors[i % len(colors)]

        rect = patches.Rectangle(
            (x, y), w, h,
            linewidth=1.5, edgecolor='black',
            facecolor=color, alpha=0.5
        )
        ax.add_patch(rect)

        center_x = x + w / 2
        center_y = y + h / 2

        if m['rowspan'] > 1 or m['colspan'] > 1:
            label = f"{m['list_id']}\n({m['start_row']},{m['start_col']})\n~({m['end_row']},{m['end_col']})"
        else:
            label = f"{m['list_id']}\n({m['start_row']},{m['start_col']})"

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
    ax.set_title(f'오른쪽 순회 + 셀 크기 매칭\n{max_row+1}행 x {max_col+1}열, {len(cells)}개 셀',
                fontsize=13, fontweight='bold')

    plt.tight_layout()
    output_path = "C:\\win32hwp\\debugs\\right_traverse_with_size.png"
    plt.savefig(output_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"\n시각화 저장: {output_path}")


def main():
    data = match_right_with_size()
    if not data:
        return

    print("\n=== 매칭 결과 (처음 15개) ===")
    for m in data['cells'][:15]:
        if m['rowspan'] > 1 or m['colspan'] > 1:
            print(f"list_id={m['list_id']}: ({m['start_row']},{m['start_col']})~({m['end_row']},{m['end_col']})")
        else:
            print(f"list_id={m['list_id']}: ({m['start_row']},{m['start_col']})")

    visualize(data)


if __name__ == "__main__":
    main()
