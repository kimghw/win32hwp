# -*- coding: utf-8 -*-
"""순수 오른쪽 순회 매칭

알고리즘:
1. first_cols[0]에서 시작
2. 무조건 오른쪽으로만 이동
3. last_cols 만나면 → 다음 first_cols로 이동
4. 방문한 셀 → skip하고 계속 오른쪽
5. 방문 안한 셀 → 매칭
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor import get_hwp_instance
from table.table_info import TableInfo, MOVE_RIGHT_OF_CELL
from table.table_boundary import TableBoundary

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


def match_pure_right():
    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return None

    table_info = TableInfo(hwp, debug=False)
    boundary = TableBoundary(hwp, debug=False)

    if not table_info.is_in_table():
        print("[오류] 커서가 테이블 내부에 있지 않습니다.")
        return None

    # 경계 분석
    boundary_result = boundary.check_boundary_table()
    first_cols = boundary._sort_first_cols_by_position(boundary_result.first_cols)
    last_cols_set = set(boundary_result.last_cols)

    print(f"=== TableBoundary 원본 데이터 ===")
    print(f"first_rows: {boundary_result.first_rows}")
    print(f"first_cols: {boundary_result.first_cols}")
    print(f"last_cols: {boundary_result.last_cols}")
    print(f"table_origin: {boundary_result.table_origin}")
    print(f"table_end: {boundary_result.table_end}")
    print()
    print(f"first_cols (sorted by y): {first_cols}")
    print()

    # 1단계: 순수 오른쪽 순회로 셀 좌표 수집 + last_cols 직접 추출
    # last_cols 판단: 다음 셀이 first_cols면 현재 셀이 last_col
    first_cols_set = set(first_cols)
    cell_positions = {}
    all_x = set()
    all_y = set()
    visited = set()
    extracted_last_cols = []  # 직접 추출한 last_cols

    cumulative_y = 0

    for row_idx, row_start_id in enumerate(first_cols):
        # 행 시작 셀의 높이로 row_height 결정
        hwp.SetPos(row_start_id, 0, 0)
        _, row_height = table_info.get_cell_dimensions()

        row_start_y = cumulative_y
        row_end_y = cumulative_y + row_height

        current_x = 0
        current_id = row_start_id

        while True:
            if current_id not in visited:
                visited.add(current_id)

                hwp.SetPos(current_id, 0, 0)
                width, height = table_info.get_cell_dimensions()

                sx, sy = current_x, row_start_y
                ex, ey = current_x + width, row_start_y + height

                cell_positions[current_id] = {
                    'sx': sx, 'ex': ex,
                    'sy': sy, 'ey': ey,
                }

                all_x.update([sx, ex])
                all_y.update([sy, ey])

                current_x = ex  # 다음 셀 시작 x
            else:
                # 이미 방문한 셀이면 해당 셀의 너비만큼 x 이동
                if current_id in cell_positions:
                    current_x = cell_positions[current_id]['ex']

            # 오른쪽으로 이동
            hwp.SetPos(current_id, 0, 0)
            hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
            next_id = hwp.GetPos()[0]

            # 이동 불가 → 현재 셀이 last_col
            if next_id == current_id:
                extracted_last_cols.append(current_id)
                break

            # 다음 셀이 first_cols → 현재 셀이 last_col
            if next_id in first_cols_set:
                print(f"  Row {row_idx}: {current_id} → next={next_id} (first_col) → last_col={current_id}")
                extracted_last_cols.append(current_id)
                break

            current_id = next_id

        cumulative_y = row_end_y

    print(f"last_cols (extracted): {extracted_last_cols}")
    print(f"last_cols 개수: {len(extracted_last_cols)} (first_cols: {len(first_cols)})")

    # 경계 필터링
    table_max_x = max(pos['ex'] for pos in cell_positions.values()) if cell_positions else 0
    table_max_y = cumulative_y

    all_x = {x for x in all_x if -TOLERANCE <= x <= table_max_x + TOLERANCE}
    all_y = {y for y in all_y if -TOLERANCE <= y <= table_max_y + TOLERANCE}

    x_levels = merge_close_levels(list(all_x))
    y_levels = merge_close_levels(list(all_y))

    print(f"x_levels: {len(x_levels)}개")
    print(f"y_levels: {len(y_levels)}개")
    print()

    # 디버그: 처음 20개 셀의 물리 좌표 출력
    print("=== 물리 좌표 (처음 20개) ===")
    for lid in sorted(cell_positions.keys())[:20]:
        pos = cell_positions[lid]
        print(f"list_id={lid}: sx={pos['sx']}, ex={pos['ex']}, sy={pos['sy']}, ey={pos['ey']}")
    print()
    print(f"y_levels: {y_levels[:10]}...")
    print()

    # 2단계: 그리드 매칭
    matched = []
    for list_id, pos in cell_positions.items():
        start_col = find_level_index(pos['sx'], x_levels)
        end_col = find_end_level_index(pos['ex'], x_levels)
        start_row = find_level_index(pos['sy'], y_levels)
        end_row = find_end_level_index(pos['ey'], y_levels)

        colspan = end_col - start_col + 1
        rowspan = end_row - start_row + 1

        matched.append({
            'list_id': list_id,
            'start_row': start_row,
            'start_col': start_col,
            'end_row': end_row,
            'end_col': end_col,
            'rowspan': rowspan,
            'colspan': colspan,
        })

    # list_id 순으로 정렬
    matched.sort(key=lambda x: x['list_id'])

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
    ax.set_title(f'순수 오른쪽 순회 매칭\n{max_row+1}행 x {max_col+1}열, {len(cells)}개 셀',
                fontsize=13, fontweight='bold')

    plt.tight_layout()
    output_path = "C:\\win32hwp\\debugs\\pure_right_match.png"
    plt.savefig(output_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"\n시각화 저장: {output_path}")


def main():
    data = match_pure_right()
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
