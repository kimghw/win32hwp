# BFS로 테이블 셀 구조를 탐지하는 모듈
#
# === 목표 ===
# 테이블의 모든 셀을 BFS로 탐색하고 각 셀의 이웃 정보(상하좌우 list_id)를 수집한다.
#
# === 알고리즘 ===
# 1. 테이블의 첫 번째 셀(좌상단)로 이동
#    - 좌측/상단으로 더 이상 이동 불가할 때까지 이동하여 시작점 확보
#
# 2. BFS 탐색으로 모든 접근 가능한 셀 수집
#    - 각 셀에서 상하좌우로 커서 이동 시도
#    - 이동 성공 시 해당 방향의 list_id 기록
#    - 이동 실패(반환값 0 또는 False) 시 해당 방향 list_id = 0 (경계 또는 병합 영역)
#    - 방문한 list_id는 중복 탐색 방지
#
# 3. 각 셀의 이웃 정보 저장
#    - {list_id: 27, left: 0, right: 28, up: 0, down: 0}

from collections import deque
from dataclasses import dataclass
from typing import Dict, List
from cursor_utils import get_hwp_instance


# MovePos 셀 이동 상수
MOVE_LEFT_OF_CELL = 100
MOVE_RIGHT_OF_CELL = 101
MOVE_UP_OF_CELL = 102
MOVE_DOWN_OF_CELL = 103
MOVE_START_OF_CELL = 104  # 행의 시작 셀
MOVE_END_OF_CELL = 105    # 행의 끝 셀
MOVE_TOP_OF_CELL = 106    # 열의 시작 셀
MOVE_BOTTOM_OF_CELL = 107 # 열의 끝 셀


@dataclass
class CellInfo:
    """셀 정보를 저장하는 데이터 클래스"""
    list_id: int
    left: int = 0    # 좌측 셀의 list_id (없으면 0)
    right: int = 0   # 우측 셀의 list_id
    up: int = 0      # 상단 셀의 list_id
    down: int = 0    # 하단 셀의 list_id


class TableInfo:
    """BFS로 테이블 구조를 탐지하는 클래스"""

    def __init__(self, hwp=None, debug: bool = False):
        self.hwp = hwp or get_hwp_instance()
        self.debug = debug
        self.cells: Dict[int, CellInfo] = {}  # list_id -> CellInfo

    def _log(self, msg: str):
        """디버그 메시지 출력"""
        if self.debug:
            print(f"[TableInfo] {msg}")

    def _get_list_id(self) -> int:
        """현재 커서 위치의 list_id 반환"""
        pos = self.hwp.GetPos()
        return pos[0]

    def _move_to_first_cell_simple(self) -> bool:
        """테이블의 첫 번째 셀(좌상단)로 이동 (로그 없음)"""
        # 행의 시작으로 이동
        self.hwp.MovePos(MOVE_START_OF_CELL, 0, 0)
        # 열의 시작(맨 위)으로 이동
        self.hwp.MovePos(MOVE_TOP_OF_CELL, 0, 0)
        return True

    def is_in_table(self) -> bool:
        """현재 커서가 테이블 내부에 있는지 확인"""
        try:
            act = self.hwp.CreateAction("TableCellBlock")
            pset = act.CreateSet()
            return act.Execute(pset)
        except:
            return False

    def move_to_first_cell(self) -> bool:
        """테이블의 첫 번째 셀(좌상단)로 이동"""
        if not self.is_in_table():
            self._log("테이블 내부가 아닙니다")
            return False

        # 먼저 행의 시작으로 이동 (MOVE_START_OF_CELL)
        self.hwp.MovePos(MOVE_START_OF_CELL, 0, 0)
        self._log(f"행 시작 이동 후: {self._get_list_id()}")

        # 그 다음 열의 시작(맨 위)으로 이동 (MOVE_TOP_OF_CELL)
        self.hwp.MovePos(MOVE_TOP_OF_CELL, 0, 0)
        self._log(f"열 시작 이동 후: {self._get_list_id()}")

        # 추가: 왼쪽으로 더 이동 가능한지 확인
        before = self._get_list_id()
        result = self.hwp.MovePos(MOVE_LEFT_OF_CELL, 0, 0)
        after = self._get_list_id()
        self._log(f"왼쪽 이동 시도: result={result}, before={before}, after={after}")

        if result and after != before:
            # 왼쪽에 더 있으면 다시 행의 시작으로
            self.hwp.MovePos(MOVE_START_OF_CELL, 0, 0)
            self._log(f"추가 이동 후: {self._get_list_id()}")

        self._log(f"첫 번째 셀 list_id: {self._get_list_id()}")
        return True

    def collect_cells_bfs(self) -> Dict[int, CellInfo]:
        """
        행 우선 순회로 모든 셀을 탐색하고 이웃 정보 수집

        첫 셀에서 시작하여:
        1. 오른쪽으로 이동하며 행의 모든 셀 수집
        2. 행 끝에서 첫 셀로 돌아가 아래로 이동
        3. 다음 행 반복

        Returns:
            Dict[int, CellInfo]: list_id -> CellInfo 매핑
        """
        if not self.move_to_first_cell():
            return {}

        self.cells.clear()
        row_start_positions = []  # 각 행의 시작 위치 저장

        row = 0
        while True:
            # 현재 행의 시작 위치 저장
            row_start_pos = self.hwp.GetPos()
            row_start_positions.append(row_start_pos)
            self._log(f"행 {row} 시작: list_id={row_start_pos[0]}")

            # 현재 행을 오른쪽으로 순회
            col = 0
            while True:
                current_id = self._get_list_id()

                # 이미 수집한 셀이면 건너뛰기
                if current_id in self.cells:
                    self._log(f"  ({row},{col}) 이미 수집됨: {current_id}")
                else:
                    # 이웃 정보 수집
                    cell_info = self._collect_neighbors(current_id)
                    self.cells[current_id] = cell_info
                    self._log(f"  ({row},{col}) 셀 {current_id}: L={cell_info.left}, R={cell_info.right}, U={cell_info.up}, D={cell_info.down}")

                # 오른쪽으로 이동 시도
                before = current_id
                result = self.hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
                after = self._get_list_id()

                if not result or after == before:
                    # 행 끝 도달
                    self._log(f"  행 {row} 끝 (총 {col+1}개 셀)")
                    break

                col += 1

            # 다음 행으로 이동: 행 시작으로 돌아간 후 아래로
            self.hwp.SetPos(row_start_pos[0], row_start_pos[1], row_start_pos[2])
            before = self._get_list_id()
            result = self.hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
            after = self._get_list_id()

            if not result or after == before:
                # 마지막 행
                self._log(f"마지막 행 도달 (총 {row+1}개 행)")
                break

            row += 1

        self._log(f"총 {len(self.cells)}개 셀 수집 완료")
        return self.cells

    def _collect_neighbors(self, current_id: int) -> CellInfo:
        """현재 셀의 이웃 정보 수집 (이동 후 복귀)"""
        cell_info = CellInfo(list_id=current_id)

        # 현재 위치 저장
        saved_pos = self.hwp.GetPos()

        directions = [
            (MOVE_LEFT_OF_CELL, 'left'),
            (MOVE_RIGHT_OF_CELL, 'right'),
            (MOVE_UP_OF_CELL, 'up'),
            (MOVE_DOWN_OF_CELL, 'down')
        ]

        for direction, attr in directions:
            # 이웃으로 이동 시도
            result = self.hwp.MovePos(direction, 0, 0)
            after = self._get_list_id()

            if result and after != current_id:
                setattr(cell_info, attr, after)
            else:
                setattr(cell_info, attr, 0)

            # 원래 위치로 복귀
            self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])

        return cell_info

    def find_duplicate_neighbors(self) -> Dict[str, Dict[int, List[int]]]:
        """
        각 방향별로 같은 이웃을 공유하는 셀들 찾기

        예: 셀 A와 셀 B가 같은 up을 가리키면, 그 up 셀은 가로 병합된 것

        Returns:
            Dict[str, Dict[int, List[int]]]: {
                'left': {이웃_list_id: [참조하는_셀들]},
                'right': {이웃_list_id: [참조하는_셀들]},
                'up': {이웃_list_id: [참조하는_셀들]},
                'down': {이웃_list_id: [참조하는_셀들]}
            }
        """
        if not self.cells:
            self.collect_cells_bfs()

        if not self.cells:
            return {}

        # 각 방향별로 그룹화
        groups = {
            'left': {},
            'right': {},
            'up': {},
            'down': {}
        }

        for cell_id, cell in self.cells.items():
            # left 그룹
            if cell.left != 0:
                if cell.left not in groups['left']:
                    groups['left'][cell.left] = []
                groups['left'][cell.left].append(cell_id)

            # right 그룹
            if cell.right != 0:
                if cell.right not in groups['right']:
                    groups['right'][cell.right] = []
                groups['right'][cell.right].append(cell_id)

            # up 그룹
            if cell.up != 0:
                if cell.up not in groups['up']:
                    groups['up'][cell.up] = []
                groups['up'][cell.up].append(cell_id)

            # down 그룹
            if cell.down != 0:
                if cell.down not in groups['down']:
                    groups['down'][cell.down] = []
                groups['down'][cell.down].append(cell_id)

        # 중복된 것만 필터링 (2개 이상 참조하는 경우)
        duplicates = {
            'left': {},
            'right': {},
            'up': {},
            'down': {}
        }

        for direction in ['left', 'right', 'up', 'down']:
            for neighbor_id, referring_cells in groups[direction].items():
                if len(referring_cells) >= 2:
                    duplicates[direction][neighbor_id] = referring_cells

        return duplicates

    def print_duplicate_neighbors(self):
        """중복된 이웃 정보 출력"""
        duplicates = self.find_duplicate_neighbors()

        print("\n=== 중복된 이웃 분석 ===")

        for direction, data in duplicates.items():
            if data:
                print(f"\n[{direction}] 중복:")
                for neighbor_id, referring_cells in data.items():
                    print(f"  셀 {neighbor_id} <- {referring_cells} ({len(referring_cells)}개 셀이 참조)")
            else:
                print(f"\n[{direction}] 중복 없음")

    def get_table_size(self) -> Dict[str, int]:
        """
        테이블 크기 계산 (병합 셀 고려)

        열 수: 첫 행을 오른쪽으로 이동하며 셀 수 + rowspan 병합 보정
        행 수: 첫 열을 아래로 이동하며 셀 수 + colspan 병합 보정

        Returns:
            Dict: {'rows': 행수, 'cols': 열수, 'rowspan_count': rowspan수, 'colspan_count': colspan수}
        """
        if not self.cells:
            self.collect_cells_bfs()

        if not self.cells:
            return {'rows': 0, 'cols': 0, 'rowspan_count': 0, 'colspan_count': 0}

        # 중복 이웃 분석
        duplicates = self.find_duplicate_neighbors()

        # rowspan: up 또는 down이 중복된 경우 (여러 셀이 같은 위/아래 셀을 가리킴)
        # → 가로로 병합된 셀이 세로 방향에 영향
        rowspan_cells = set()
        for neighbor_id in duplicates['up'].keys():
            rowspan_cells.add(neighbor_id)
        for neighbor_id in duplicates['down'].keys():
            rowspan_cells.add(neighbor_id)

        # colspan: left 또는 right가 중복된 경우 (여러 셀이 같은 좌/우 셀을 가리킴)
        # → 세로로 병합된 셀이 가로 방향에 영향
        colspan_cells = set()
        for neighbor_id in duplicates['left'].keys():
            colspan_cells.add(neighbor_id)
        for neighbor_id in duplicates['right'].keys():
            colspan_cells.add(neighbor_id)

        # 첫 번째 셀로 이동
        if not self.move_to_first_cell():
            return {'rows': 0, 'cols': 0, 'rowspan_count': 0, 'colspan_count': 0}

        # 열 수 계산: 첫 행을 오른쪽으로 순회
        col_count = 0
        rowspan_in_first_row = 0
        while True:
            current_id = self._get_list_id()
            col_count += 1

            # 현재 셀이 rowspan 병합 셀인지 확인
            if current_id in rowspan_cells:
                # 이 셀을 up으로 가리키는 셀 수 - 1 = 추가 열
                up_refs = len(duplicates['up'].get(current_id, []))
                down_refs = len(duplicates['down'].get(current_id, []))
                extra = max(up_refs, down_refs) - 1
                if extra > 0:
                    rowspan_in_first_row += extra
                    self._log(f"  열 {col_count}: 셀 {current_id}는 rowspan, 추가 열 +{extra}")

            # 오른쪽으로 이동
            before = current_id
            result = self.hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
            after = self._get_list_id()

            if not result or after == before:
                break

        total_cols = col_count + rowspan_in_first_row

        # 첫 번째 셀로 다시 이동
        self.move_to_first_cell()

        # 행 수 계산: 첫 열을 아래로 순회
        row_count = 0
        colspan_in_first_col = 0
        while True:
            current_id = self._get_list_id()
            row_count += 1

            # 현재 셀이 colspan 병합 셀인지 확인
            if current_id in colspan_cells:
                # 이 셀을 left로 가리키는 셀 수 - 1 = 추가 행
                left_refs = len(duplicates['left'].get(current_id, []))
                right_refs = len(duplicates['right'].get(current_id, []))
                extra = max(left_refs, right_refs) - 1
                if extra > 0:
                    colspan_in_first_col += extra
                    self._log(f"  행 {row_count}: 셀 {current_id}는 colspan, 추가 행 +{extra}")

            # 아래로 이동
            before = current_id
            result = self.hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
            after = self._get_list_id()

            if not result or after == before:
                break

        total_rows = row_count + colspan_in_first_col

        self._log(f"테이블 크기: {total_rows}행 x {total_cols}열 (rowspan={rowspan_in_first_row}, colspan={colspan_in_first_col})")

        return {
            'rows': total_rows,
            'cols': total_cols,
            'rowspan_count': rowspan_in_first_row,
            'colspan_count': colspan_in_first_col
        }


if __name__ == "__main__":
    print("=== TableInfo 테스트 ===\n")

    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        exit(1)

    table = TableInfo(hwp, debug=True)

    if not table.is_in_table():
        print("[오류] 커서가 테이블 내부에 있지 않습니다.")
        print("테이블 안에 커서를 위치시킨 후 다시 실행하세요.")
        exit(1)

    print("\n1. BFS로 셀 수집...")
    cells = table.collect_cells_bfs()
    print(f"\n수집된 셀: {len(cells)}개")

    print("\n2. 셀 정보:")
    for cell_id, cell in cells.items():
        print(f"   셀 {cell_id}: L={cell.left}, R={cell.right}, U={cell.up}, D={cell.down}")

    print("\n3. 중복 이웃 분석...")
    table.print_duplicate_neighbors()

    print("\n4. 테이블 크기 계산...")
    size = table.get_table_size()
    print(f"   행 수: {size['rows']} (colspan 보정: +{size['colspan_count']})")
    print(f"   열 수: {size['cols']} (rowspan 보정: +{size['rowspan_count']})")

    print("\n=== 테스트 완료 ===")
