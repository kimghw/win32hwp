# -*- coding: utf-8 -*-
"""오른쪽 순회 방식 셀 매칭 (BFS 아님)

알고리즘:
1. first_cols[0]에서 시작
2. 오른쪽으로 이동하며 매칭
3. last_cols 만나면 → 다음 first_cols로 이동
4. 방문한 셀 → 매칭 없이 오른쪽 계속
5. 방문 안한 셀 → 매칭
6. 끝까지 진행
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor import get_hwp_instance
from table.table_info import TableInfo, MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL
from table.table_boundary import TableBoundary


def match_right_traverse():
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

    print(f"first_cols: {len(first_cols)}개")
    print(f"last_cols: {len(last_cols_set)}개")
    print()

    # 매칭 결과
    matched = []
    visited = set()

    # 그리드 좌표
    row = 0
    col = 0

    for row_idx, row_start_id in enumerate(first_cols):
        print(f"=== Row {row_idx} (first_col: {row_start_id}) ===")

        col = 0
        current_id = row_start_id

        while True:
            # 이미 방문했으면 매칭 없이 오른쪽으로
            if current_id in visited:
                print(f"  [{current_id}] 이미 방문 → 스킵")
            else:
                # 방문 안 했으면 매칭
                visited.add(current_id)

                # 셀 크기 가져오기
                hwp.SetPos(current_id, 0, 0)
                width, height = table_info.get_cell_dimensions()

                matched.append({
                    'list_id': current_id,
                    'row': row_idx,
                    'col': col,
                    'width': width,
                    'height': height,
                })

                print(f"  [{current_id}] → ({row_idx}, {col}) 매칭")

            # last_cols면 다음 행으로
            if current_id in last_cols_set:
                print(f"  [{current_id}] last_col → 다음 행")
                break

            # 오른쪽으로 이동
            hwp.SetPos(current_id, 0, 0)
            hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
            next_id = hwp.GetPos()[0]

            # 이동 불가 (같은 셀)
            if next_id == current_id:
                print(f"  [{current_id}] 오른쪽 이동 불가 → 끝")
                break

            col += 1
            current_id = next_id

        print()

    print(f"=== 매칭 완료: {len(matched)}개 셀 ===")
    return matched


def visualize(matched):
    """매칭 결과 시각화"""
    import matplotlib.pyplot as plt
    import matplotlib.patches as patches
    from matplotlib import font_manager

    font_path = "C:/Windows/Fonts/malgun.ttf"
    if os.path.exists(font_path):
        font_prop = font_manager.FontProperties(fname=font_path)
        plt.rcParams['font.family'] = font_prop.get_name()
    plt.rcParams['axes.unicode_minus'] = False

    max_row = max(m['row'] for m in matched)
    max_col = max(m['col'] for m in matched)

    fig, ax = plt.subplots(figsize=(18, 14))
    colors = plt.cm.tab20.colors

    for i, m in enumerate(matched):
        x = m['col']
        y = max_row - m['row']  # Y 반전
        w = 1
        h = 1

        color = colors[i % len(colors)]

        rect = patches.Rectangle(
            (x, y), w, h,
            linewidth=1.5, edgecolor='black',
            facecolor=color, alpha=0.5
        )
        ax.add_patch(rect)

        label = f"{m['list_id']}\n({m['row']},{m['col']})"
        ax.text(x + 0.5, y + 0.5, label,
                ha='center', va='center',
                fontsize=6, fontweight='bold')

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
    ax.set_title(f'오른쪽 순회 매칭 결과\n{max_row+1}행 x {max_col+1}열, {len(matched)}개 셀',
                fontsize=13, fontweight='bold')

    plt.tight_layout()
    output_path = "C:\\win32hwp\\debugs\\right_traverse_matched.png"
    plt.savefig(output_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"\n시각화 저장: {output_path}")


def main():
    matched = match_right_traverse()
    if not matched:
        return

    print("\n=== 매칭 결과 ===")
    for m in matched[:20]:
        print(f"list_id={m['list_id']}: ({m['row']}, {m['col']})")

    if len(matched) > 20:
        print(f"... 외 {len(matched) - 20}개")

    visualize(matched)


if __name__ == "__main__":
    main()
