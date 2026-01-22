# -*- coding: utf-8 -*-
"""2단계: 매칭 결과 정규 간격 시각화"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor import get_hwp_instance
from table.cell_position import CellPositionCalculator


def visualize_matched_normalized():
    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return

    calc = CellPositionCalculator(hwp, debug=False)

    try:
        result = calc.calculate()
    except ValueError as e:
        print(f"[오류] {e}")
        return

    print("=== 2단계: 매칭 결과 ===")
    print(f"x_levels: {len(result.x_levels)}개")
    print(f"y_levels: {len(result.y_levels)}개")
    print(f"총 셀: {len(result.cells)}개")
    print(f"그리드: {result.max_row + 1}행 x {result.max_col + 1}열")

    # 시각화
    try:
        import matplotlib.pyplot as plt
        import matplotlib.patches as patches
        from matplotlib import font_manager
    except ImportError:
        print("[경고] matplotlib이 설치되어 있지 않습니다.")
        return

    font_path = "C:/Windows/Fonts/malgun.ttf"
    if os.path.exists(font_path):
        font_prop = font_manager.FontProperties(fname=font_path)
        plt.rcParams['font.family'] = font_prop.get_name()
    plt.rcParams['axes.unicode_minus'] = False

    max_row = result.max_row
    max_col = result.max_col

    fig, ax = plt.subplots(figsize=(18, 22))

    colors = plt.cm.tab20.colors

    # 각 셀 그리기 (정규화된 그리드)
    for i, (list_id, cell) in enumerate(sorted(result.cells.items())):
        # 정규화된 좌표
        x = cell.start_col
        y = max_row - cell.end_row  # Y 반전
        w = cell.colspan
        h = cell.rowspan

        color = colors[i % len(colors)]

        rect = patches.Rectangle(
            (x, y), w, h,
            linewidth=1.5,
            edgecolor='black',
            facecolor=color,
            alpha=0.5
        )
        ax.add_patch(rect)

        # 라벨
        center_x = x + w / 2
        center_y = y + h / 2

        if cell.rowspan > 1 or cell.colspan > 1:
            label = f"{list_id}\n({cell.start_row},{cell.start_col})\n~({cell.end_row},{cell.end_col})"
        else:
            label = f"{list_id}\n({cell.start_row},{cell.start_col})"

        fontsize = 6 if w < 1.5 or h < 1.5 else 7
        ax.text(center_x, center_y, label,
                ha='center', va='center',
                fontsize=fontsize, fontweight='bold')

    # 그리드 라인
    for col in range(max_col + 2):
        ax.axvline(x=col, color='gray', linewidth=0.5, alpha=0.3)
    for row in range(max_row + 2):
        ax.axhline(y=row, color='gray', linewidth=0.5, alpha=0.3)

    # X 라벨 (열 인덱스)
    for col in range(max_col + 1):
        ax.text(col + 0.5, max_row + 1.3, str(col),
                ha='center', va='center', fontsize=7, color='blue')

    # Y 라벨 (행 인덱스)
    for row in range(max_row + 1):
        ax.text(-0.5, max_row - row + 0.5, str(row),
                ha='center', va='center', fontsize=7, color='green')

    ax.set_xlim(-1, max_col + 2)
    ax.set_ylim(-1, max_row + 2)
    ax.set_aspect('equal')
    ax.set_xlabel('Column (정규화)', fontsize=11)
    ax.set_ylabel('Row (정규화, 반전)', fontsize=11)
    ax.set_title(f'2단계: 매칭 결과 (정규 간격)\n'
                f'{max_row+1}행 x {max_col+1}열, {len(result.cells)}개 셀\n'
                f'x_levels: {len(result.x_levels)}개, y_levels: {len(result.y_levels)}개',
                fontsize=13, fontweight='bold')

    plt.tight_layout()

    output_path = "C:\\win32hwp\\debugs\\matched_normalized.png"
    plt.savefig(output_path, dpi=150, bbox_inches='tight')
    plt.close()

    print(f"\n시각화 저장: {output_path}")


if __name__ == "__main__":
    visualize_matched_normalized()
