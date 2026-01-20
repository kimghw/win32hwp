# -*- coding: utf-8 -*-
"""셀 위치 좌표 측정 - first_cols 순회하며 X, Y 레벨 수집"""

from table.table_boundary import TableBoundary
from table.table_info import TableInfo, MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL
from cursor import get_hwp_instance

hwp = get_hwp_instance()
if not hwp:
    print('[오류] 한글이 실행 중이지 않습니다.')
    exit(1)

boundary = TableBoundary(hwp, debug=False)
table_info = TableInfo(hwp, debug=False)

if not boundary._is_in_table():
    print('[오류] 커서가 테이블 내부에 있지 않습니다.')
    exit(1)

# 경계 분석
result = boundary.check_boundary_table()
first_cols = boundary._sort_first_cols_by_position(result.first_cols)
last_cols_set = set(result.last_cols)

print(f"table_origin: {result.table_origin}")
print(f"first_cols (행 시작): {first_cols}")

# === first_cols 순회하며 모든 셀의 좌표와 레벨 수집 ===
x_levels = set()
y_levels = set()
x_levels.add(0)
y_levels.add(0)

cell_positions = {}
first_cols_set = set(first_cols)


def collect_split_cells(cell_id, start_x, start_y, row_end_y, visited_global):
    """분할된 셀을 재귀적으로 순회하며 Y 레벨과 위치 수집 (아래 + 오른쪽)"""
    hwp.SetPos(cell_id, 0, 0)
    width, height = table_info.get_cell_dimensions()
    end_x = start_x + width
    end_y = start_y + height

    x_levels.add(end_x)
    y_levels.add(end_y)

    # 셀 위치 저장
    if cell_id not in cell_positions:
        cell_positions[cell_id] = {
            'start_x': start_x, 'end_x': end_x,
            'start_y': start_y, 'end_y': end_y,
            'width': width, 'height': height
        }

    # 오른쪽으로 순회 (같은 Y 레벨에 있는 셀들)
    hwp.SetPos(cell_id, 0, 0)
    hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
    right_id, _, _ = hwp.GetPos()

    if right_id != cell_id and right_id not in visited_global and right_id not in first_cols_set:
        visited_global.add(right_id)
        collect_split_cells(right_id, end_x, start_y, row_end_y, visited_global)

    # 아래로 순회 (행 끝에 도달하지 않았으면)
    if end_y < row_end_y:
        hwp.SetPos(cell_id, 0, 0)
        hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
        down_id, _, _ = hwp.GetPos()

        if down_id != cell_id and down_id not in visited_global and down_id not in first_cols_set:
            visited_global.add(down_id)
            collect_split_cells(down_id, start_x, end_y, row_end_y, visited_global)


cumulative_y = 0  # 행의 시작 Y 좌표

for row_idx, row_start in enumerate(first_cols):
    # 행의 시작 셀에서 행 높이 가져오기
    hwp.SetPos(row_start, 0, 0)
    _, row_height = table_info.get_cell_dimensions()

    row_start_y = cumulative_y
    row_end_y = cumulative_y + row_height
    y_levels.add(row_end_y)

    # 이 행에서 오른쪽으로 순회
    cumulative_x = 0
    current_id = row_start
    visited = set()

    while True:
        if current_id in visited:
            break
        visited.add(current_id)

        hwp.SetPos(current_id, 0, 0)
        width, height = table_info.get_cell_dimensions()

        start_x = cumulative_x
        end_x = cumulative_x + width
        start_y = row_start_y
        end_y = row_start_y + height

        # X, Y 레벨 추가
        x_levels.add(end_x)
        y_levels.add(end_y)

        # 셀 위치 저장
        cell_positions[current_id] = {
            'start_x': start_x, 'end_x': end_x,
            'start_y': start_y, 'end_y': end_y,
            'width': width, 'height': height
        }

        # 셀이 분할되어 있으면 (높이가 행 높이보다 작으면) 재귀적으로 서브셀 순회
        if height < row_height:
            hwp.SetPos(current_id, 0, 0)
            hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
            next_sub, _, _ = hwp.GetPos()

            if next_sub != current_id and next_sub not in visited and next_sub not in first_cols_set:
                visited.add(next_sub)
                collect_split_cells(next_sub, start_x, end_y, row_end_y, visited)

        cumulative_x = end_x

        # last_col 도달시 종료
        if current_id in last_cols_set:
            break

        # 오른쪽으로 이동
        hwp.SetPos(current_id, 0, 0)
        hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
        next_id, _, _ = hwp.GetPos()

        if next_id == current_id:
            break
        # 다음 first_col 만나면 종료
        if next_id in first_cols_set and next_id != row_start:
            break

        current_id = next_id

    cumulative_y = row_end_y

x_levels = sorted(x_levels)
y_levels = sorted(y_levels)

print(f"\n=== X 레벨 ===")
print(f"총 {len(x_levels)}개: {x_levels}")

print(f"\n=== Y 레벨 ===")
print(f"총 {len(y_levels)}개: {y_levels}")


# === 좌표를 레벨 인덱스로 변환 (±10 허용) ===
TOLERANCE = 10


def find_level_index(value, levels):
    """값이 레벨 ±TOLERANCE 범위 내에 있으면 해당 레벨 인덱스 반환"""
    for idx, level in enumerate(levels):
        if abs(value - level) <= TOLERANCE:
            return idx
    return -1  # 매칭 안됨


# 각 셀에 레벨 인덱스 할당 (왼쪽 x, 위쪽 y 기준)
for list_id, pos in cell_positions.items():
    pos['col'] = find_level_index(pos['start_x'], x_levels)  # 왼쪽 x 기준
    pos['row'] = find_level_index(pos['start_y'], y_levels)  # 위쪽 y 기준


print(f"\n=== 셀 위치 정보 ({len(cell_positions)}개) ===")
for list_id, pos in sorted(cell_positions.items()):
    print(f"list_id={list_id}: x={pos['col']}, y={pos['row']}")


# === 각 셀에 좌표 텍스트 삽입 ===
def insert_coordinate_text(hwp, list_id: int, col: int, row: int):
    """셀에 좌표 텍스트를 파란색으로 삽입"""
    hwp.SetPos(list_id, 0, 0)
    hwp.MovePos(5, 0, 0)  # moveBottomOfList - 셀 끝으로 이동

    coord_text = f"\r({row}, {col})"

    # 텍스트 삽입
    act = hwp.CreateAction("InsertText")
    pset = act.CreateSet()
    act.GetDefault(pset)
    pset.SetItem("Text", coord_text)
    act.Execute(pset)

    # 삽입한 텍스트 선택 (뒤에서부터 좌표 길이만큼)
    coord_len = len(f"({row}, {col})")
    for _ in range(coord_len):
        hwp.HAction.Run("MoveSelPrevChar")

    # 파란색 글자 모양 적용
    act = hwp.CreateAction("CharShape")
    pset = act.CreateSet()
    act.GetDefault(pset)
    pset.SetItem("TextColor", 0xFF0000)  # BGR: 파란색
    act.Execute(pset)

    # 선택 해제
    hwp.HAction.Run("Cancel")


print(f"\n총 {len(cell_positions)}개 셀에 좌표 삽입 중...")

for list_id, pos in cell_positions.items():
    insert_coordinate_text(hwp, list_id, pos['col'], pos['row'])
    print(f"  list_id={list_id} → ({pos['row']}, {pos['col']})")

print("\n완료!")
