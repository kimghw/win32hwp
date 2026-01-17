"""
테이블 셀 순회 및 좌/우/위/아래 위치 정보 수집
"""
import win32com.client as win32
from cursor_utils import get_hwp_instance


def collect_table_cells():
    """
    커서 위치의 테이블을 순회하며 각 셀의 위치 정보 수집

    Returns:
        list: [
            {
                'index': int,
                'current': (list_id, para_id, char_pos),
                'left': (list_id, para_id, char_pos) or 0,
                'right': (list_id, para_id, char_pos) or 0,
                'up': (list_id, para_id, char_pos) or 0,
                'down': (list_id, para_id, char_pos) or 0
            },
            ...
        ]
    """
    hwp = get_hwp_instance()
    if not hwp:
        print("한글이 실행되지 않았습니다")
        return []

    # 보안 모듈 등록
    hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')

    # 원래 위치 저장
    original_pos = hwp.GetPos()

    # 첫 번째 셀로 이동
    hwp.MovePos(106, 0, 0)  # MOVE_TOP_OF_CELL
    hwp.MovePos(104, 0, 0)  # MOVE_START_OF_CELL

    cells = []
    index = 0

    # 행별로 순회
    row = 0
    while True:
        # 각 행의 시작으로 이동
        hwp.MovePos(106, 0, 0)  # MOVE_TOP_OF_CELL (첫 번째 행)
        hwp.MovePos(104, 0, 0)  # MOVE_START_OF_CELL (첫 번째 열)

        # 현재 행으로 이동
        for _ in range(row):
            prev_pos = hwp.GetPos()
            hwp.MovePos(103, 0, 0)  # MOVE_DOWN_OF_CELL
            if hwp.GetPos() == prev_pos:
                # 더 이상 아래로 못 감 = 마지막 행 도달
                hwp.SetPos(original_pos[0], original_pos[1], original_pos[2])
                print(f"\n총 {len(cells)}개 셀 수집 완료 ({row}행)")
                return cells

        # 현재 행의 모든 열 순회
        col = 0
        while True:
            current_pos = hwp.GetPos()

            # 좌 이동
            hwp.MovePos(100, 0, 0)  # MOVE_LEFT_OF_CELL
            left_pos = hwp.GetPos()
            left = left_pos if left_pos != current_pos else 0
            hwp.SetPos(current_pos[0], current_pos[1], current_pos[2])

            # 우 이동
            hwp.MovePos(101, 0, 0)  # MOVE_RIGHT_OF_CELL
            right_pos = hwp.GetPos()
            right = right_pos if right_pos != current_pos else 0
            hwp.SetPos(current_pos[0], current_pos[1], current_pos[2])

            # 위 이동
            hwp.MovePos(102, 0, 0)  # MOVE_UP_OF_CELL
            up_pos = hwp.GetPos()
            up = up_pos if up_pos != current_pos else 0
            hwp.SetPos(current_pos[0], current_pos[1], current_pos[2])

            # 아래 이동
            hwp.MovePos(103, 0, 0)  # MOVE_DOWN_OF_CELL
            down_pos = hwp.GetPos()
            down = down_pos if down_pos != current_pos else 0
            hwp.SetPos(current_pos[0], current_pos[1], current_pos[2])

            # 셀 정보 저장
            cell_info = {
                'index': index,
                'row': row,
                'col': col,
                'current': current_pos,
                'left': left,
                'right': right,
                'up': up,
                'down': down
            }
            cells.append(cell_info)

            print(f"[{row},{col}] 현재={current_pos}, 좌={left}, 우={right}, 위={up}, 아래={down}")

            # 오른쪽으로 이동
            hwp.MovePos(101, 0, 0)  # MOVE_RIGHT_OF_CELL
            next_pos = hwp.GetPos()

            # 더 이상 오른쪽으로 못 가면 행 종료
            if next_pos == current_pos:
                break

            col += 1
            index += 1

        row += 1
        index += 1

    # 원래 위치로 복원
    hwp.SetPos(original_pos[0], original_pos[1], original_pos[2])

    return cells


if __name__ == "__main__":
    print("=== 테이블 셀 정보 수집 시작 ===\n")

    cells = collect_table_cells()

    if cells:
        print(f"\n수집된 셀 개수: {len(cells)}")

        # 행별 통계
        rows_dict = {}
        for cell in cells:
            row = cell['row']
            if row not in rows_dict:
                rows_dict[row] = []
            rows_dict[row].append(cell['col'])

        print(f"\n행별 셀 개수:")
        for row in sorted(rows_dict.keys()):
            cols = sorted(rows_dict[row])
            print(f"  행 {row}: {len(cols)}개 셀 (열: {cols})")

        # 최대 행/열 확인
        max_row = max(cell['row'] for cell in cells)
        max_col = max(cell['col'] for cell in cells)
        print(f"\n최대 행: {max_row}, 최대 열: {max_col}")
        print(f"예상 셀 개수: {(max_row + 1) * (max_col + 1)}")

        # 누락된 셀 찾기
        collected_positions = {(cell['row'], cell['col']) for cell in cells}
        missing = []
        for row in range(max_row + 1):
            for col in range(max_col + 1):
                if (row, col) not in collected_positions:
                    missing.append((row, col))

        if missing:
            print(f"\n누락된 셀: {len(missing)}개")
            for pos in missing:
                print(f"  ({pos[0]}, {pos[1]})")

        print("\n처음 10개 셀:")
        for cell in cells[:10]:
            print(f"  [{cell['row']},{cell['col']}]")
            print(f"    현재: {cell['current']}")
            print(f"    좌: {cell['left']}")
            print(f"    우: {cell['right']}")
            print(f"    위: {cell['up']}")
            print(f"    아래: {cell['down']}")
    else:
        print("셀 정보 수집 실패")
