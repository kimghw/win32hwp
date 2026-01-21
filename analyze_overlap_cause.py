# -*- coding: utf-8 -*-
"""중복 점유 근본 원인 분석 - BFS 경로 추적"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from cursor import get_hwp_instance
from table.table_info import TableInfo, MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL, MOVE_UP_OF_CELL


def trace_cell_path():
    """특정 셀들이 어떤 경로로 도달했는지 추적"""
    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return

    table_info = TableInfo(hwp, debug=False)

    if not table_info.is_in_table():
        print("[오류] 커서가 테이블 내부에 있지 않습니다.")
        return

    print("=== 중복 원인 분석: 셀 137, 138 경로 추적 ===\n")

    # 타겟 셀들
    target_cells = {137, 138, 158, 159, 160, 161}

    # 첫 셀로 이동
    table_info.move_to_first_cell()
    first_id = hwp.GetPos()[0]
    print(f"첫 셀: list_id={first_id}")

    # BFS로 탐색하며 타겟 셀에 도달하는 경로 기록
    from collections import deque

    visited = set()
    parent = {}  # child -> (parent, direction, x, y)
    cell_data = {}  # list_id -> (start_x, start_y, width, height)

    hwp.SetPos(first_id, 0, 0)
    first_width, first_height = table_info.get_cell_dimensions()
    cell_data[first_id] = (0, 0, first_width, first_height)

    queue = deque([(first_id, 0, 0)])  # (list_id, x, y)
    visited.add(first_id)
    parent[first_id] = None

    while queue:
        current_id, cur_x, cur_y = queue.popleft()
        cur_w = cell_data[current_id][2]
        cur_h = cell_data[current_id][3]

        directions = [
            (MOVE_RIGHT_OF_CELL, 'right'),
            (MOVE_DOWN_OF_CELL, 'down'),
            (MOVE_UP_OF_CELL, 'up'),
        ]

        for move_cmd, direction in directions:
            hwp.SetPos(current_id, 0, 0)
            hwp.MovePos(move_cmd, 0, 0)
            next_id = hwp.GetPos()[0]

            if next_id == current_id or next_id in visited:
                continue

            visited.add(next_id)

            hwp.SetPos(next_id, 0, 0)
            next_w, next_h = table_info.get_cell_dimensions()

            # 좌표 계산
            if direction == 'right':
                next_x = cur_x + cur_w
                next_y = cur_y
            elif direction == 'down':
                next_x = cur_x
                next_y = cur_y + cur_h
            elif direction == 'up':
                next_x = cur_x
                next_y = cur_y - next_h

            cell_data[next_id] = (next_x, next_y, next_w, next_h)
            parent[next_id] = (current_id, direction, next_x, next_y)

            queue.append((next_id, next_x, next_y))

            # 타겟 셀이면 경로 출력
            if next_id in target_cells:
                print(f"\n=== 셀 {next_id} 도달 ===")
                print(f"  좌표: x={next_x}, y={next_y}")
                print(f"  크기: w={next_w}, h={next_h}")
                print(f"  end: x={next_x + next_w}, y={next_y + next_h}")

                # 경로 역추적
                path = []
                node = next_id
                while parent[node] is not None:
                    p_id, p_dir, p_x, p_y = parent[node]
                    path.append((node, p_dir, p_x, p_y))
                    node = p_id
                path.append((first_id, 'start', 0, 0))
                path.reverse()

                print(f"  경로:")
                for i, (n_id, n_dir, n_x, n_y) in enumerate(path):
                    print(f"    [{i}] {n_dir} -> list_id={n_id} (x={n_x})")

    # 문제 분석
    print("\n\n=== 중복 원인 분석 ===")
    print("\n셀 137 vs 138 비교:")
    if 137 in cell_data and 138 in cell_data:
        c137 = cell_data[137]
        c138 = cell_data[138]
        print(f"  셀 137: x={c137[0]}~{c137[0]+c137[2]}, y={c137[1]}~{c137[1]+c137[3]}")
        print(f"  셀 138: x={c138[0]}~{c138[0]+c138[2]}, y={c138[1]}~{c138[1]+c138[3]}")
        print(f"  137 end_x ({c137[0]+c137[2]}) vs 138 start_x ({c138[0]})")
        if c137[0] + c137[2] > c138[0]:
            print(f"  --> 물리적으로 겹침! 차이: {c137[0]+c137[2] - c138[0]}")

    print("\n셀 158, 159, 160, 161 비교:")
    for lid in [158, 159, 160, 161]:
        if lid in cell_data:
            c = cell_data[lid]
            print(f"  셀 {lid}: x={c[0]}~{c[0]+c[2]} (w={c[2]})")


if __name__ == "__main__":
    trace_cell_path()
