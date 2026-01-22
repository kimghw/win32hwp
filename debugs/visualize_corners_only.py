# -*- coding: utf-8 -*-
"""1단계: 순수 좌표점(corners)만 시각화 - 셀 매칭 없이"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor import get_hwp_instance
from table.table_info import TableInfo, MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL
from table.table_boundary import TableBoundary
from collections import deque


def extract_corners_only():
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
    first_cols_set = set(first_cols)

    # 1단계: 모든 셀의 모서리 좌표 수집
    all_corners = []  # (x, y) 튜플 리스트
    cell_positions = {}

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

            # 4개 꼭지점 수집
            all_corners.append((start_x, start_y))  # 좌상단
            all_corners.append((end_x, start_y))    # 우상단
            all_corners.append((start_x, end_y))    # 좌하단
            all_corners.append((end_x, end_y))      # 우하단

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

    # X, Y 좌표 분리
    all_x = sorted(set(c[0] for c in all_corners))
    all_y = sorted(set(c[1] for c in all_corners))

    # 경계 계산
    table_max_x = 0
    for last_col_id in last_cols_set:
        if last_col_id in cell_positions:
            table_max_x = max(table_max_x, cell_positions[last_col_id]['end_x'])

    table_max_y = cumulative_y

    # 필터링
    TOLERANCE = 3
    filtered_x = sorted([x for x in all_x if 0 - TOLERANCE <= x <= table_max_x + TOLERANCE])
    filtered_y = sorted([y for y in all_y if 0 - TOLERANCE <= y <= table_max_y + TOLERANCE])

    # 제거된 좌표
    removed_x = set(all_x) - set(filtered_x)
    removed_y = set(all_y) - set(filtered_y)

    return {
        'corners': all_corners,
        'x_before': all_x,
        'y_before': all_y,
        'x_after': filtered_x,
        'y_after': filtered_y,
        'removed_x': removed_x,
        'removed_y': removed_y,
        'table_max_x': table_max_x,
        'table_max_y': table_max_y,
    }


def visualize_corners(data):
    """순수 좌표점만 시각화"""
    try:
        import matplotlib.pyplot as plt
        from matplotlib import font_manager
    except ImportError:
        print("[경고] matplotlib이 설치되어 있지 않습니다.")
        return

    font_path = "C:/Windows/Fonts/malgun.ttf"
    if os.path.exists(font_path):
        font_prop = font_manager.FontProperties(fname=font_path)
        plt.rcParams['font.family'] = font_prop.get_name()
    plt.rcParams['axes.unicode_minus'] = False

    max_y = data['table_max_y']

    fig, axes = plt.subplots(1, 2, figsize=(18, 10))

    for ax_idx, (ax, title, x_list, y_list) in enumerate([
        (axes[0], f"필터링 전\nX: {len(data['x_before'])}개, Y: {len(data['y_before'])}개",
         data['x_before'], data['y_before']),
        (axes[1], f"필터링 후\nX: {len(data['x_after'])}개, Y: {len(data['y_after'])}개",
         data['x_after'], data['y_after']),
    ]):
        # 꼭지점 표시 (작은 점)
        corners_x = [c[0] for c in data['corners']]
        corners_y = [max_y - c[1] for c in data['corners']]  # Y 반전
        ax.scatter(corners_x, corners_y, s=3, c='gray', alpha=0.3, label='셀 꼭지점')

        # X 라인 (세로선)
        for x in x_list:
            is_removed = x in data['removed_x']
            if ax_idx == 0 and is_removed:
                ax.axvline(x=x, color='red', linewidth=2, alpha=0.8)
            else:
                ax.axvline(x=x, color='blue', linewidth=0.8, alpha=0.5)

        # Y 라인 (가로선)
        for y in y_list:
            y_flipped = max_y - y
            is_removed = y in data['removed_y']
            if ax_idx == 0 and is_removed:
                ax.axhline(y=y_flipped, color='red', linewidth=2, alpha=0.8)
            else:
                ax.axhline(y=y_flipped, color='green', linewidth=0.8, alpha=0.5)

        # 테이블 경계 (검은 점선)
        ax.axvline(x=0, color='black', linewidth=2, linestyle='--')
        ax.axvline(x=data['table_max_x'], color='black', linewidth=2, linestyle='--')
        ax.axhline(y=0, color='black', linewidth=2, linestyle='--')
        ax.axhline(y=max_y, color='black', linewidth=2, linestyle='--')

        # 제거된 좌표 표시 (필터링 전에서만)
        if ax_idx == 0:
            for x in data['removed_x']:
                ax.annotate(f'제거: {x}', xy=(x, max_y * 0.95),
                           fontsize=9, color='red', fontweight='bold',
                           ha='center', rotation=90)

        ax.set_xlim(-3000, max(data['x_before']) + 5000)
        ax.set_ylim(-5000, max_y + 5000)
        ax.set_aspect('equal')
        ax.set_xlabel('X (HWPUNIT)')
        ax.set_ylabel('Y (HWPUNIT, 반전)')
        ax.set_title(title, fontsize=12)
        ax.grid(True, alpha=0.2)

    # 범례
    from matplotlib.lines import Line2D
    from matplotlib.patches import Patch
    legend_elements = [
        Line2D([0], [0], marker='o', color='w', markerfacecolor='gray',
               markersize=6, label='셀 꼭지점 (corners)'),
        Line2D([0], [0], color='blue', linewidth=1, label='X 라인 (xline)'),
        Line2D([0], [0], color='green', linewidth=1, label='Y 라인 (yline)'),
        Line2D([0], [0], color='red', linewidth=2, label='제거된 라인'),
        Line2D([0], [0], color='black', linewidth=2, linestyle='--', label='테이블 경계'),
    ]
    fig.legend(handles=legend_elements, loc='upper center', ncol=5, fontsize=10)

    plt.suptitle('1단계: 셀 꼭지점(corners) 수집 → X/Y 라인 추출 → 경계 필터링',
                fontsize=14, fontweight='bold', y=0.98)
    plt.tight_layout(rect=[0, 0, 1, 0.93])

    output_path = "C:\\win32hwp\\debugs\\corners_only.png"
    plt.savefig(output_path, dpi=150, bbox_inches='tight')
    plt.close()

    print(f"\n시각화 저장: {output_path}")


def main():
    data = extract_corners_only()
    if not data:
        return

    print("=== 1단계: 셀 꼭지점 좌표 수집 ===")
    print(f"총 꼭지점: {len(data['corners'])}개")
    print(f"\n필터링 전: X {len(data['x_before'])}개, Y {len(data['y_before'])}개")
    print(f"필터링 후: X {len(data['x_after'])}개, Y {len(data['y_after'])}개")
    print(f"\n제거된 X: {sorted(data['removed_x']) if data['removed_x'] else '없음'}")
    print(f"제거된 Y: {sorted(data['removed_y']) if data['removed_y'] else '없음'}")
    print(f"\n테이블 경계: X=[0~{data['table_max_x']}], Y=[0~{data['table_max_y']}]")

    visualize_corners(data)


if __name__ == "__main__":
    main()
