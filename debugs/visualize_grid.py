"""그리드 매핑 시각화 - TableBoundary.map_grid_by_xend 사용"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib import font_manager
from cursor import get_hwp_instance
from table.table_boundary import TableBoundary

# 한글 폰트 설정
font_path = "C:/Windows/Fonts/malgun.ttf"
if os.path.exists(font_path):
    font_manager.fontManager.addfont(font_path)
    plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False


def visualize_grid(result: dict, save_path: str = None):
    """그리드를 시각화"""
    grid = result['grid']
    max_row = result['max_row']
    max_col = result['max_col']
    xend = result.get('xend', 0)

    # rowspan_positions가 있으면 사용, 없으면 기본 위치만
    rowspan_positions = result.get('rowspan_positions', {})

    # 그리드 배열 생성 (list_id, is_primary, colspan)
    grid_array = [[None for _ in range(max_col + 1)] for _ in range(max_row + 1)]

    # 각 셀의 모든 위치 추적 (colspan 포함)
    for list_id, info in grid.items():
        r, c = info['row'], info['col']
        colspan = info.get('colspan', 1)
        for dc in range(colspan):
            if r <= max_row and c + dc <= max_col:
                grid_array[r][c + dc] = (list_id, dc == 0)  # (list_id, is_start)

    # rowspan 위치 추가 (colspan도 적용)
    for list_id, positions in rowspan_positions.items():
        colspan = grid[list_id].get('colspan', 1) if list_id in grid else 1
        for (r, c) in positions:
            for dc in range(colspan):
                if r <= max_row and c + dc <= max_col:
                    grid_array[r][c + dc] = (list_id, False)  # rowspan은 is_start=False

    # 시각화
    fig, ax = plt.subplots(figsize=(14, 18))

    cell_width = 1.0
    cell_height = 0.5

    # 색상 맵
    unique_ids = list(grid.keys())
    colors = plt.cm.tab20(range(len(unique_ids)))
    color_map = {lid: colors[i % 20] for i, lid in enumerate(unique_ids)}

    for r in range(max_row + 1):
        for c in range(max_col + 1):
            x = c * cell_width
            y = (max_row - r) * cell_height

            cell_data = grid_array[r][c]

            if cell_data is not None:
                list_id, is_start = cell_data
                color = color_map[list_id]
                info = grid[list_id]
                is_primary = (r == info['row'] and c == info['col'])

                rect = patches.FancyBboxPatch(
                    (x + 0.02, y + 0.02),
                    cell_width - 0.04,
                    cell_height - 0.04,
                    boxstyle="round,pad=0.02",
                    facecolor=color,
                    edgecolor='black',
                    linewidth=0.5
                )
                ax.add_patch(rect)

                # 셀 시작점에만 라벨 표시
                if is_start or is_primary:
                    if is_primary:
                        label = f"{list_id}"
                        fontweight = 'bold'
                        fontsize = 6
                    else:
                        label = f"({list_id})"
                        fontweight = 'normal'
                        fontsize = 5

                    ax.text(
                        x + cell_width / 2,
                        y + cell_height / 2,
                        label,
                        ha='center',
                        va='center',
                        fontsize=fontsize,
                        fontweight=fontweight,
                        color='black' if sum(color[:3]) > 1.5 else 'white'
                    )
            else:
                rect = patches.Rectangle(
                    (x + 0.02, y + 0.02),
                    cell_width - 0.04,
                    cell_height - 0.04,
                    facecolor='#f0f0f0',
                    edgecolor='lightgray',
                    linewidth=0.5,
                    linestyle='--'
                )
                ax.add_patch(rect)

    # 행/열 번호
    for r in range(max_row + 1):
        y = (max_row - r) * cell_height
        ax.text(-0.3, y + cell_height / 2, f"R{r}", ha='right', va='center', fontsize=6)

    for c in range(max_col + 1):
        x = c * cell_width
        ax.text(x + cell_width / 2, (max_row + 1) * cell_height + 0.1, f"C{c}",
                ha='center', va='bottom', fontsize=6)

    ax.set_xlim(-0.5, (max_col + 1) * cell_width + 0.5)
    ax.set_ylim(-0.5, (max_row + 2) * cell_height)
    ax.set_aspect('equal')
    ax.axis('off')
    ax.set_title(f'HWP 테이블 그리드 매핑 (xend={xend})\n'
                 f'{max_row + 1}행 x {max_col + 1}열, {len(grid)}셀',
                 fontsize=11, fontweight='bold')

    plt.tight_layout()

    if save_path:
        plt.savefig(save_path, dpi=150, bbox_inches='tight', facecolor='white')
        print(f"저장됨: {save_path}")

    plt.close()
    return save_path


def main():
    hwp = get_hwp_instance()
    if not hwp:
        print("한글이 실행 중이 아닙니다")
        return

    original_pos = hwp.GetPos()

    # TableBoundary의 map_grid_by_xend 사용
    boundary = TableBoundary(hwp, debug=False)
    result = boundary.map_grid_by_xend()

    print(f"xend = {result['xend']} HWPUNIT")
    print(f"총 셀 수: {len(result['grid'])}")
    print(f"그리드 크기: {result['max_row'] + 1} 행 x {result['max_col'] + 1} 열")

    # 시각화
    save_path = os.path.join(os.path.dirname(__file__), "grid_mapping.png")
    visualize_grid(result, save_path)

    hwp.SetPos(*original_pos)


if __name__ == "__main__":
    main()
