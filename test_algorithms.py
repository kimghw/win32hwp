# -*- coding: utf-8 -*-
"""다양한 셀 매핑 알고리즘 비교 테스트"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from collections import deque
from cursor import get_hwp_instance
from table.table_info import (
    TableInfo, MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL,
    MOVE_LEFT_OF_CELL, MOVE_UP_OF_CELL
)


def get_cell_info(hwp, table_info, list_id):
    """셀의 크기 정보 반환"""
    hwp.SetPos(list_id, 0, 0)
    width, height = table_info.get_cell_dimensions()
    return width, height


def algorithm_1_bfs_4dir(hwp, table_info, first_id, max_cells=2000):
    """알고리즘 1: 일반 BFS (4방향) - 모든 방향에서 좌표 누적"""
    print("\n=== 알고리즘 1: 일반 BFS (4방향 모두 좌표 누적) ===")

    visited = set()
    cell_positions = {}

    first_w, first_h = get_cell_info(hwp, table_info, first_id)
    cell_positions[first_id] = {'x': 0, 'y': 0, 'w': first_w, 'h': first_h}
    visited.add(first_id)

    queue = deque([(first_id, 0, 0)])

    directions = [
        (MOVE_RIGHT_OF_CELL, 'right', 1, 0),
        (MOVE_DOWN_OF_CELL, 'down', 0, 1),
        (MOVE_LEFT_OF_CELL, 'left', -1, 0),
        (MOVE_UP_OF_CELL, 'up', 0, -1),
    ]

    while queue and len(visited) < max_cells:
        cur_id, cur_x, cur_y = queue.popleft()
        cur_info = cell_positions[cur_id]

        for move_cmd, dir_name, dx, dy in directions:
            hwp.SetPos(cur_id, 0, 0)
            hwp.MovePos(move_cmd, 0, 0)
            next_id = hwp.GetPos()[0]

            if next_id == cur_id or next_id in visited:
                continue

            visited.add(next_id)
            next_w, next_h = get_cell_info(hwp, table_info, next_id)

            # 좌표 계산
            if dir_name == 'right':
                next_x = cur_x + cur_info['w']
                next_y = cur_y
            elif dir_name == 'left':
                next_x = cur_x - next_w
                next_y = cur_y
            elif dir_name == 'down':
                next_x = cur_x
                next_y = cur_y + cur_info['h']
            elif dir_name == 'up':
                next_x = cur_x
                next_y = cur_y - next_h

            cell_positions[next_id] = {'x': next_x, 'y': next_y, 'w': next_w, 'h': next_h}
            queue.append((next_id, next_x, next_y))

    return cell_positions


def algorithm_2_row_first(hwp, table_info, first_id, max_cells=2000):
    """알고리즘 2: 행 우선 순회 - first_cols 기반"""
    print("\n=== 알고리즘 2: 행 우선 순회 (first_cols 기반) ===")

    visited = set()
    cell_positions = {}

    # 첫 번째 열의 모든 셀 찾기 (first_cols)
    first_cols = []
    hwp.SetPos(first_id, 0, 0)
    current_y = 0

    while len(first_cols) < max_cells:
        cur_id = hwp.GetPos()[0]
        if cur_id in visited:
            break

        cur_w, cur_h = get_cell_info(hwp, table_info, cur_id)
        first_cols.append((cur_id, current_y, cur_h))
        visited.add(cur_id)
        cell_positions[cur_id] = {'x': 0, 'y': current_y, 'w': cur_w, 'h': cur_h}

        current_y += cur_h

        # 아래로 이동
        hwp.SetPos(cur_id, 0, 0)
        hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
        next_id = hwp.GetPos()[0]
        if next_id == cur_id:
            break

    # 각 first_col에서 우측으로 순회
    for start_id, start_y, _ in first_cols:
        hwp.SetPos(start_id, 0, 0)
        cur_x = cell_positions[start_id]['w']

        while len(visited) < max_cells:
            hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
            next_id = hwp.GetPos()[0]

            if next_id in visited:
                # 이미 방문한 셀이면 해당 셀의 end_x로 이동
                cur_x = cell_positions[next_id]['x'] + cell_positions[next_id]['w']
                hwp.SetPos(next_id, 0, 0)
                continue

            if next_id == start_id:
                break

            visited.add(next_id)
            next_w, next_h = get_cell_info(hwp, table_info, next_id)
            cell_positions[next_id] = {'x': cur_x, 'y': start_y, 'w': next_w, 'h': next_h}
            cur_x += next_w

            hwp.SetPos(next_id, 0, 0)

    return cell_positions


def algorithm_3_dfs(hwp, table_info, first_id, max_cells=2000):
    """알고리즘 3: DFS (깊이 우선 탐색)"""
    print("\n=== 알고리즘 3: DFS (깊이 우선 탐색) ===")

    visited = set()
    cell_positions = {}

    first_w, first_h = get_cell_info(hwp, table_info, first_id)
    cell_positions[first_id] = {'x': 0, 'y': 0, 'w': first_w, 'h': first_h}
    visited.add(first_id)

    stack = [(first_id, 0, 0)]

    directions = [
        (MOVE_RIGHT_OF_CELL, 'right'),
        (MOVE_DOWN_OF_CELL, 'down'),
        (MOVE_LEFT_OF_CELL, 'left'),
        (MOVE_UP_OF_CELL, 'up'),
    ]

    while stack and len(visited) < max_cells:
        cur_id, cur_x, cur_y = stack.pop()
        cur_info = cell_positions[cur_id]

        for move_cmd, dir_name in directions:
            hwp.SetPos(cur_id, 0, 0)
            hwp.MovePos(move_cmd, 0, 0)
            next_id = hwp.GetPos()[0]

            if next_id == cur_id or next_id in visited:
                continue

            visited.add(next_id)
            next_w, next_h = get_cell_info(hwp, table_info, next_id)

            if dir_name == 'right':
                next_x = cur_x + cur_info['w']
                next_y = cur_y
            elif dir_name == 'left':
                next_x = cur_x - next_w
                next_y = cur_y
            elif dir_name == 'down':
                next_x = cur_x
                next_y = cur_y + cur_info['h']
            elif dir_name == 'up':
                next_x = cur_x
                next_y = cur_y - next_h

            cell_positions[next_id] = {'x': next_x, 'y': next_y, 'w': next_w, 'h': next_h}
            stack.append((next_id, next_x, next_y))

    return cell_positions


def algorithm_4_bfs_right_down_only(hwp, table_info, first_id, max_cells=2000):
    """알고리즘 4: BFS (우측+하단만) - 역방향 없이"""
    print("\n=== 알고리즘 4: BFS (우측+하단만) ===")

    visited = set()
    cell_positions = {}

    first_w, first_h = get_cell_info(hwp, table_info, first_id)
    cell_positions[first_id] = {'x': 0, 'y': 0, 'w': first_w, 'h': first_h}
    visited.add(first_id)

    queue = deque([(first_id, 0, 0)])

    while queue and len(visited) < max_cells:
        cur_id, cur_x, cur_y = queue.popleft()
        cur_info = cell_positions[cur_id]

        # 우측 이동
        hwp.SetPos(cur_id, 0, 0)
        hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
        right_id = hwp.GetPos()[0]

        if right_id != cur_id and right_id not in visited:
            visited.add(right_id)
            r_w, r_h = get_cell_info(hwp, table_info, right_id)
            next_x = cur_x + cur_info['w']
            cell_positions[right_id] = {'x': next_x, 'y': cur_y, 'w': r_w, 'h': r_h}
            queue.append((right_id, next_x, cur_y))

        # 하단 이동
        hwp.SetPos(cur_id, 0, 0)
        hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
        down_id = hwp.GetPos()[0]

        if down_id != cur_id and down_id not in visited:
            visited.add(down_id)
            d_w, d_h = get_cell_info(hwp, table_info, down_id)
            next_y = cur_y + cur_info['h']
            cell_positions[down_id] = {'x': cur_x, 'y': next_y, 'w': d_w, 'h': d_h}
            queue.append((down_id, cur_x, next_y))

    return cell_positions


def analyze_result(name, cell_positions, tolerance=100):
    """결과 분석"""
    if not cell_positions:
        print(f"  셀 수: 0")
        return

    # X, Y 레벨 수집
    x_levels = set()
    y_levels = set()

    for info in cell_positions.values():
        x_levels.add(info['x'])
        y_levels.add(info['y'])

    # 레벨 병합
    def merge_levels(levels, tol):
        sorted_levels = sorted(levels)
        merged = [sorted_levels[0]]
        for lv in sorted_levels[1:]:
            if lv - merged[-1] > tol:
                merged.append(lv)
        return merged

    x_merged = merge_levels(x_levels, tolerance)
    y_merged = merge_levels(y_levels, tolerance)

    # 그리드 매핑
    def find_index(val, levels, tol):
        for i, lv in enumerate(levels):
            if abs(val - lv) <= tol:
                return i
        return -1

    grid = {}
    for list_id, info in cell_positions.items():
        col = find_index(info['x'], x_merged, tolerance)
        row = find_index(info['y'], y_merged, tolerance)

        # end 위치 계산
        end_x = info['x'] + info['w']
        end_y = info['y'] + info['h']

        # end_col, end_row 찾기
        end_col = col
        end_row = row
        for i, lv in enumerate(x_merged):
            if lv < end_x - tolerance:
                end_col = i
        for i, lv in enumerate(y_merged):
            if lv < end_y - tolerance:
                end_row = i

        for r in range(row, end_row + 1):
            for c in range(col, end_col + 1):
                if (r, c) not in grid:
                    grid[(r, c)] = []
                grid[(r, c)].append(list_id)

    # 분석
    max_row = max(r for r, c in grid.keys()) if grid else 0
    max_col = max(c for r, c in grid.keys()) if grid else 0
    total_cells = (max_row + 1) * (max_col + 1)

    duplicates = sum(1 for v in grid.values() if len(v) > 1)
    empty = total_cells - len(grid)

    # 음수 좌표
    negative = sum(1 for info in cell_positions.values() if info['x'] < 0 or info['y'] < 0)

    print(f"  셀 수: {len(cell_positions)}")
    print(f"  그리드 크기: {max_row + 1}행 x {max_col + 1}열")
    print(f"  X 레벨: {len(x_merged)}개, Y 레벨: {len(y_merged)}개")
    print(f"  중복: {duplicates}개, 빈칸: {empty}개, 음수좌표: {negative}개")

    if duplicates == 0 and empty == 0 and negative == 0:
        print(f"  --> [OK] 정상")
    else:
        print(f"  --> [문제] 비정상")

    return {
        'cells': len(cell_positions),
        'rows': max_row + 1,
        'cols': max_col + 1,
        'duplicates': duplicates,
        'empty': empty,
        'negative': negative
    }


def main():
    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return

    table_info = TableInfo(hwp, debug=False)

    if not table_info.is_in_table():
        print("[오류] 커서가 테이블 내부에 있지 않습니다.")
        return

    # 첫 셀로 이동
    table_info.move_to_first_cell()
    first_id = hwp.GetPos()[0]
    print(f"첫 번째 셀: list_id={first_id}")

    results = {}

    # 알고리즘 1: 일반 BFS (4방향)
    pos1 = algorithm_1_bfs_4dir(hwp, table_info, first_id)
    results['BFS 4방향'] = analyze_result('BFS 4방향', pos1)

    # 알고리즘 2: 행 우선 순회
    table_info.move_to_first_cell()
    pos2 = algorithm_2_row_first(hwp, table_info, first_id)
    results['행 우선'] = analyze_result('행 우선', pos2)

    # 알고리즘 3: DFS
    table_info.move_to_first_cell()
    pos3 = algorithm_3_dfs(hwp, table_info, first_id)
    results['DFS'] = analyze_result('DFS', pos3)

    # 알고리즘 4: BFS (우측+하단만)
    table_info.move_to_first_cell()
    pos4 = algorithm_4_bfs_right_down_only(hwp, table_info, first_id)
    results['BFS 우하'] = analyze_result('BFS 우하', pos4)

    # 비교 요약
    print("\n" + "=" * 60)
    print("=== 알고리즘 비교 요약 ===")
    print("=" * 60)
    print(f"{'알고리즘':<15} {'셀수':<8} {'크기':<12} {'중복':<8} {'빈칸':<8} {'음수':<8}")
    print("-" * 60)
    for name, r in results.items():
        size = f"{r['rows']}x{r['cols']}"
        status = "OK" if r['duplicates'] == 0 and r['empty'] == 0 and r['negative'] == 0 else "문제"
        print(f"{name:<15} {r['cells']:<8} {size:<12} {r['duplicates']:<8} {r['empty']:<8} {r['negative']:<8} {status}")


if __name__ == "__main__":
    main()
