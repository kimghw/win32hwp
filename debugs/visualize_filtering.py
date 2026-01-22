# -*- coding: utf-8 -*-
"""1단계 셀 모서리 좌표 추출 및 필터링 시각화"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor import get_hwp_instance
from table.table_info import TableInfo, MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL
from table.table_boundary import TableBoundary
from collections import deque


def extract_and_visualize():
    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return

    table_info = TableInfo(hwp, debug=False)
    boundary = TableBoundary(hwp, debug=False)

    if not table_info.is_in_table():
        print("[오류] 커서가 테이블 내부에 있지 않습니다.")
        return

    # 경계 분석
    boundary_result = boundary.check_boundary_table()
    first_cols = boundary._sort_first_cols_by_position(boundary_result.first_cols)
    last_cols_set = set(boundary_result.last_cols)
    first_cols_set = set(first_cols)

    print(f"first_cols: {len(first_cols)}개")
    print(f"last_cols: {len(last_cols_set)}개")

    # 1단계: 모든 셀의 물리적 좌표 수집
    cell_positions = {}
    all_x_coords = set()
    all_y_coords = set()

    cumulative_y = 0

    for row_idx, row_start in enumerate(first_cols):
        hwp.SetPos(row_start, 0, 0)
        _, row_height = table_info.get_cell_dimensions()

        row_start_y = cumulative_y
        row_end_y = cumulative_y + row_height

        # 행 내 셀들 BFS
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

            # 모든 좌표 수집 (필터링 전)
            all_x_coords.add(start_x)
            all_x_coords.add(end_x)
            all_y_coords.add(start_y)
            all_y_coords.add(end_y)

            # 오른쪽 이동
            hwp.SetPos(current_id, 0, 0)
            hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
            right_id = hwp.GetPos()[0]

            if (right_id != current_id and
                right_id not in visited_in_row and
                right_id not in first_cols_set):
                visited_in_row.add(right_id)
                queue.append((right_id, end_x, start_y))

            # 아래 이동
            if end_y < row_end_y:
                hwp.SetPos(current_id, 0, 0)
                hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
                down_id = hwp.GetPos()[0]

                if (down_id != current_id and
                    down_id not in visited_in_row and
                    down_id not in first_cols_set):
                    visited_in_row.add(down_id)
                    queue.append((down_id, start_x, end_y))

        cumulative_y = row_end_y

    # 필터링 전 좌표 저장
    x_before = sorted(all_x_coords)
    y_before = sorted(all_y_coords)

    print(f"\n=== 필터링 전 ===")
    print(f"X 좌표: {len(x_before)}개")
    print(f"Y 좌표: {len(y_before)}개")

    # 경계 계산
    table_min_x = 0
    table_max_x = 0
    table_min_y = 0
    table_max_y = cumulative_y

    for last_col_id in last_cols_set:
        if last_col_id in cell_positions:
            table_max_x = max(table_max_x, cell_positions[last_col_id]['end_x'])

    print(f"\n=== 테이블 경계 ===")
    print(f"X: [{table_min_x} ~ {table_max_x}]")
    print(f"Y: [{table_min_y} ~ {table_max_y}]")

    # 필터링
    TOLERANCE = 3
    x_after = sorted([x for x in all_x_coords
                      if table_min_x - TOLERANCE <= x <= table_max_x + TOLERANCE])
    y_after = sorted([y for y in all_y_coords
                      if table_min_y - TOLERANCE <= y <= table_max_y + TOLERANCE])

    print(f"\n=== 필터링 후 ===")
    print(f"X 좌표: {len(x_after)}개")
    print(f"Y 좌표: {len(y_after)}개")

    # 제거된 좌표
    x_removed = set(x_before) - set(x_after)
    y_removed = set(y_before) - set(y_after)

    print(f"\n=== 제거된 좌표 ===")
    print(f"X: {sorted(x_removed) if x_removed else '없음'}")
    print(f"Y: {sorted(y_removed) if y_removed else '없음'}")

    # 시각화
    visualize_filtering(
        cell_positions,
        x_before, y_before,
        x_after, y_after,
        table_min_x, table_max_x, table_min_y, table_max_y,
        x_removed, y_removed
    )


def visualize_filtering(cells, x_before, y_before, x_after, y_after,
                        min_x, max_x, min_y, max_y, x_removed, y_removed):
    """필터링 전/후 시각화"""
    try:
        import matplotlib.pyplot as plt
        import matplotlib.patches as patches
        from matplotlib import font_manager
    except ImportError:
        print("[경고] matplotlib이 설치되어 있지 않습니다.")
        return

    # 한글 폰트
    font_path = "C:/Windows/Fonts/malgun.ttf"
    if os.path.exists(font_path):
        font_prop = font_manager.FontProperties(fname=font_path)
        plt.rcParams['font.family'] = font_prop.get_name()
    plt.rcParams['axes.unicode_minus'] = False

    fig, axes = plt.subplots(1, 2, figsize=(20, 12))

    for ax_idx, (ax, title, x_lines, y_lines) in enumerate([
        (axes[0], "필터링 전", x_before, y_before),
        (axes[1], "필터링 후", x_after, y_after),
    ]):
        # 셀 그리기
        colors = plt.cm.tab20.colors
        for i, (lid, pos) in enumerate(cells.items()):
            x = pos['start_x']
            y = max_y - pos['end_y']  # Y 반전
            w = pos['end_x'] - pos['start_x']
            h = pos['end_y'] - pos['start_y']

            rect = patches.Rectangle(
                (x, y), w, h,
                linewidth=0.5, edgecolor='gray',
                facecolor=colors[i % len(colors)], alpha=0.3
            )
            ax.add_patch(rect)

        # X 라인 그리기
        for x in x_lines:
            is_removed = x in x_removed
            color = 'red' if is_removed else 'blue'
            lw = 2 if is_removed else 0.8
            alpha = 1.0 if is_removed else 0.5
            ax.axvline(x=x, color=color, linewidth=lw, alpha=alpha)

        # Y 라인 그리기
        for y in y_lines:
            y_flipped = max_y - y
            is_removed = y in y_removed
            color = 'red' if is_removed else 'green'
            lw = 2 if is_removed else 0.8
            alpha = 1.0 if is_removed else 0.5
            ax.axhline(y=y_flipped, color=color, linewidth=lw, alpha=alpha)

        # 경계 표시
        ax.axvline(x=min_x, color='black', linewidth=2, linestyle='--', label='경계')
        ax.axvline(x=max_x, color='black', linewidth=2, linestyle='--')
        ax.axhline(y=0, color='black', linewidth=2, linestyle='--')
        ax.axhline(y=max_y, color='black', linewidth=2, linestyle='--')

        ax.set_xlim(min_x - 2000, max(x_before) + 2000)
        ax.set_ylim(-2000, max_y + 2000)
        ax.set_aspect('equal')
        ax.set_xlabel('X (HWPUNIT)')
        ax.set_ylabel('Y (HWPUNIT)')
        ax.set_title(f'{title}\nX: {len(x_lines)}개, Y: {len(y_lines)}개')
        ax.grid(True, alpha=0.2)

    # 범례
    from matplotlib.lines import Line2D
    legend_elements = [
        Line2D([0], [0], color='blue', linewidth=1, label='X 라인 (유지)'),
        Line2D([0], [0], color='green', linewidth=1, label='Y 라인 (유지)'),
        Line2D([0], [0], color='red', linewidth=2, label='제거된 라인'),
        Line2D([0], [0], color='black', linewidth=2, linestyle='--', label='테이블 경계'),
    ]
    fig.legend(handles=legend_elements, loc='upper center', ncol=4, fontsize=10)

    plt.suptitle('1단계: 셀 모서리 좌표 추출 및 경계 필터링', fontsize=14, fontweight='bold', y=0.98)
    plt.tight_layout(rect=[0, 0, 1, 0.95])

    output_path = "C:\\win32hwp\\debugs\\filtering_comparison.png"
    plt.savefig(output_path, dpi=150, bbox_inches='tight')
    plt.close()

    print(f"\n시각화 저장: {output_path}")


if __name__ == "__main__":
    extract_and_visualize()
