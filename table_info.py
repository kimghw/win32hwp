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
# 4. 병합 셀 탐지
#    - up, down 이 중복된 경우 rowspan
#   - left, right 이 중복된 경우 colspan
#
# 4.1 추가 병합셀 탐지 - 너비/높이 비교 방식
#   - 가로병합된 셀에서 위, 아래로 너비를 비교해서 같다면 병합 수준이 같은 것으로 간주
#   - 예로) 31번이 3개 병합된 것인데, 아래로 내려가 34번의 셀 넓이를 비교하니 같으면 31번과 동일한 수로 병합된것
#   - 세로 병합된 셀에서 좌우로 연결된 셀들의 높이를 비교해서 같다면 동일한 수로 병합된 것으로 간주
#
# 5. 테이블 크기 계산
#  열 수 계산
#  1. 수집된 셀 중 up=0인 셀만 필터링 (첫 행)
#  2. 물리적 셀 개수 + colspan 추가
#     - 예: 셀 6개, 하나가 colspan=2 → 6 + 1 = 7열

#  행 수 계산
#  1. 물리적 셀 수 + 병합 추가 셀 수 = 논리적 총 셀 수
#     - 병합 추가 = (colspan × rowspan - 1) 합산
#  2. 논리적 총 셀 수 ÷ 열 수 = 행 수
# 
# 6. 테이블 좌표 매핑(병합되기전 테이블과 list_id 매핑  )
# - .테이블 시작 셀부터 우측으로 커서를 이동한다. (0,0)부터 시작하고 colspan이 보유한 셀을 만나면 열만 증가하는 좌표를 추가한다.
# - 예로) (0,0)  (0,1)  (0,2)  (0,3)-이게 colspan=2인 셀 이면 커서 없이 (0,4) 추가
# - 행의 마지막 열에서 우측이로 이동한 경우 다음 행의 시작 셀로 이동한다.
# - rowspann은 고려하지 않는다. 
# - 테이블이랑  list_id 매칭하고 list_id는 중복가능하다. 

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
    width: int = 0   # 셀 너비 (HWPUNIT)
    height: int = 0  # 셀 높이 (HWPUNIT)


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
        """현재 셀의 이웃 정보 및 크기 수집 (이동 후 복귀)"""
        cell_info = CellInfo(list_id=current_id)

        # 현재 위치 저장
        saved_pos = self.hwp.GetPos()

        # 셀 크기 정보 수집
        width, height = self.get_cell_dimensions()
        cell_info.width = width
        cell_info.height = height

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

    def get_cell_dimensions(self) -> tuple:
        """
        현재 커서 위치의 셀 너비/높이 조회

        CellShape 속성을 사용하여 현재 셀의 크기 정보를 가져옴

        Returns:
            tuple: (width, height) HWPUNIT 단위, 실패 시 (0, 0)
        """
        try:
            cell_shape = self.hwp.CellShape
            if cell_shape is None:
                return (0, 0)

            # Cell 아이템에서 Width, Height 조회
            cell = cell_shape.Item("Cell")
            if cell is None:
                return (0, 0)

            width = cell.Item("Width")
            height = cell.Item("Height")
            return (width or 0, height or 0)
        except Exception as e:
            self._log(f"셀 크기 조회 실패: {e}")
            return (0, 0)

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

        열 수: 첫 행을 순회하며 셀 개수 + (colspan-1) 합산
        행 수: (총 list_id 수 + 병합 셀 수) / 열 수

        Returns:
            Dict: {'rows': 행수, 'cols': 열수}
        """
        if not self.cells:
            self.collect_cells_bfs()

        if not self.cells:
            return {'rows': 0, 'cols': 0}

        # 병합 정보 가져오기
        merge_info = self.find_merged_by_dimensions()
        cell_merge = merge_info.get('cell_merge_info', {})

        # 첫 번째 셀로 이동
        if not self.move_to_first_cell():
            return {'rows': 0, 'cols': 0}

        # 열 수 계산: 첫 행(up=0) 셀만 카운트 + 병합 추가
        # 세로 병합 셀을 지나면 다른 행으로 넘어가므로 up=0인 셀만 첫 행으로 인정
        first_row_cells = []

        for cell_id, cell in self.cells.items():
            if cell.up == 0:  # 위에 셀이 없음 = 첫 행
                first_row_cells.append(cell_id)

        first_row_cells.sort()  # list_id 순으로 정렬
        self._log(f"첫 행 셀 (up=0): {first_row_cells}")

        # 물리적 셀 개수
        cols = len(first_row_cells)

        # 병합 추가: colspan-1 (2개 병합이면 +1, 3개 병합이면 +2)
        colspan_extra = 0
        for cell_id in first_row_cells:
            if cell_id in cell_merge:
                colspan = cell_merge[cell_id].get('colspan', 1)
                if colspan > 1:
                    colspan_extra += (colspan - 1)
                    self._log(f"  셀 {cell_id}: colspan={colspan}, +{colspan-1}")

        cols += colspan_extra
        self._log(f"열 수: {len(first_row_cells)} + {colspan_extra} = {cols}")

        # 행 수 계산: (총 list_id 수 + 병합으로 인한 추가 셀 수) / 열 수
        # 병합 셀 수 = 각 셀의 (colspan-1) + (rowspan-1) 합
        merge_cell_count = 0
        for cell_id, info in cell_merge.items():
            colspan = info.get('colspan', 1)
            rowspan = info.get('rowspan', 1)
            # 병합된 셀이 차지하는 추가 공간
            merge_cell_count += (colspan * rowspan) - 1

        total_logical_cells = len(self.cells) + merge_cell_count
        rows = total_logical_cells // cols if cols > 0 else 0

        self._log(f"테이블 크기: {rows}행 x {cols}열")
        self._log(f"  - 물리적 셀: {len(self.cells)}개")
        self._log(f"  - 병합 추가: {merge_cell_count}개")
        self._log(f"  - 논리적 총: {total_logical_cells}개")

        return {
            'rows': rows,
            'cols': cols
        }

    def find_merged_by_dimensions(self) -> Dict[str, List[Dict]]:
        """
        4.1 추가 병합셀 탐지 - 너비/높이 비교 방식

        가로 병합된 셀에서 위, 아래로 너비를 비교해서 같다면 병합 수준이 같은 것으로 간주
        - 예: 31번이 3개 병합된 것인데, 아래로 내려가 34번의 셀 너비를 비교하니 같으면 31번과 동일한 수로 병합된 것

        세로 병합된 셀에서 좌우로 연결된 셀들의 높이를 비교해서 같다면 동일한 수로 병합된 것으로 간주

        Returns:
            Dict: {
                'horizontal_groups': [{'width': w, 'cells': [list_id, ...], 'colspan': n}],
                'vertical_groups': [{'height': h, 'cells': [list_id, ...], 'rowspan': n}],
                'base_width': 기준 너비,
                'base_height': 기준 높이,
                'cell_merge_info': {list_id: {'colspan': n, 'rowspan': n}}
            }
        """
        if not self.cells:
            self.collect_cells_bfs()

        if not self.cells:
            return {'horizontal_groups': [], 'vertical_groups': [], 'cell_merge_info': {}}

        # 1. 너비별 셀 그룹화 (가로 병합 탐지용)
        width_groups: Dict[int, List[int]] = {}
        for cell_id, cell in self.cells.items():
            if cell.width > 0:
                if cell.width not in width_groups:
                    width_groups[cell.width] = []
                width_groups[cell.width].append(cell_id)

        # 2. 높이별 셀 그룹화 (세로 병합 탐지용)
        height_groups: Dict[int, List[int]] = {}
        for cell_id, cell in self.cells.items():
            if cell.height > 0:
                if cell.height not in height_groups:
                    height_groups[cell.height] = []
                height_groups[cell.height].append(cell_id)

        # 3. 상/하 연결된 셀 중 너비가 같은 것 찾기 (가로 병합 수준 동일)
        horizontal_merged = []
        for cell_id, cell in self.cells.items():
            # 위쪽 셀과 비교
            if cell.up != 0 and cell.up in self.cells:
                up_cell = self.cells[cell.up]
                if cell.width == up_cell.width and cell.width > 0:
                    horizontal_merged.append({
                        'type': 'same_width_vertical',
                        'cell': cell_id,
                        'neighbor': cell.up,
                        'width': cell.width,
                        'direction': 'up'
                    })

            # 아래쪽 셀과 비교
            if cell.down != 0 and cell.down in self.cells:
                down_cell = self.cells[cell.down]
                if cell.width == down_cell.width and cell.width > 0:
                    horizontal_merged.append({
                        'type': 'same_width_vertical',
                        'cell': cell_id,
                        'neighbor': cell.down,
                        'width': cell.width,
                        'direction': 'down'
                    })

        # 4. 좌/우 연결된 셀 중 높이가 같은 것 찾기 (세로 병합 수준 동일)
        vertical_merged = []
        for cell_id, cell in self.cells.items():
            # 왼쪽 셀과 비교
            if cell.left != 0 and cell.left in self.cells:
                left_cell = self.cells[cell.left]
                if cell.height == left_cell.height and cell.height > 0:
                    vertical_merged.append({
                        'type': 'same_height_horizontal',
                        'cell': cell_id,
                        'neighbor': cell.left,
                        'height': cell.height,
                        'direction': 'left'
                    })

            # 오른쪽 셀과 비교
            if cell.right != 0 and cell.right in self.cells:
                right_cell = self.cells[cell.right]
                if cell.height == right_cell.height and cell.height > 0:
                    vertical_merged.append({
                        'type': 'same_height_horizontal',
                        'cell': cell_id,
                        'neighbor': cell.right,
                        'height': cell.height,
                        'direction': 'right'
                    })

        # 5. 기준 너비/높이 찾기 (가장 작은 값 = 병합되지 않은 셀)
        all_widths = [c.width for c in self.cells.values() if c.width > 0]
        all_heights = [c.height for c in self.cells.values() if c.height > 0]

        base_width = min(all_widths) if all_widths else 0
        base_height = min(all_heights) if all_heights else 0

        self._log(f"기준 너비: {base_width} ({base_width / 7200 * 25.4:.1f}mm)")
        self._log(f"기준 높이: {base_height} ({base_height / 7200 * 25.4:.1f}mm)")

        # 6. 각 셀의 병합 개수 계산
        cell_merge_info: Dict[int, Dict] = {}
        for cell_id, cell in self.cells.items():
            colspan = 1
            rowspan = 1

            if base_width > 0 and cell.width > 0:
                # 너비를 기준 너비로 나누어 가로 병합 개수 계산
                colspan = round(cell.width / base_width)
                if colspan < 1:
                    colspan = 1

            if base_height > 0 and cell.height > 0:
                # 높이를 기준 높이로 나누어 세로 병합 개수 계산
                rowspan = round(cell.height / base_height)
                if rowspan < 1:
                    rowspan = 1

            cell_merge_info[cell_id] = {
                'colspan': colspan,
                'rowspan': rowspan,
                'width': cell.width,
                'height': cell.height
            }

            if colspan > 1 or rowspan > 1:
                self._log(f"셀 {cell_id}: colspan={colspan}, rowspan={rowspan}")

        # 7. 결과 정리 - 너비별 그룹 (2개 이상 셀이 같은 너비를 가진 경우만)
        horizontal_groups = []
        for w, cells in width_groups.items():
            if len(cells) >= 2:
                colspan = round(w / base_width) if base_width > 0 else 1
                horizontal_groups.append({
                    'width': w,
                    'cells': cells,
                    'colspan': colspan
                })

        vertical_groups = []
        for h, cells in height_groups.items():
            if len(cells) >= 2:
                rowspan = round(h / base_height) if base_height > 0 else 1
                vertical_groups.append({
                    'height': h,
                    'cells': cells,
                    'rowspan': rowspan
                })

        self._log(f"너비 기반 병합 그룹: {len(horizontal_groups)}개")
        self._log(f"높이 기반 병합 그룹: {len(vertical_groups)}개")

        return {
            'horizontal_groups': horizontal_groups,
            'vertical_groups': vertical_groups,
            'horizontal_merged': horizontal_merged,  # 상하 이웃 중 같은 너비
            'vertical_merged': vertical_merged,      # 좌우 이웃 중 같은 높이
            'base_width': base_width,
            'base_height': base_height,
            'cell_merge_info': cell_merge_info       # 각 셀의 병합 개수
        }

    def print_merged_by_dimensions(self):
        """너비/높이 기반 병합 정보 출력"""
        result = self.find_merged_by_dimensions()

        print("\n=== 4.1 너비/높이 기반 병합 분석 ===")

        # 기준 값 출력
        base_w_mm = result['base_width'] / 7200 * 25.4 if result['base_width'] else 0
        base_h_mm = result['base_height'] / 7200 * 25.4 if result['base_height'] else 0
        print(f"\n[기준 값] 너비: {result['base_width']} ({base_w_mm:.1f}mm), 높이: {result['base_height']} ({base_h_mm:.1f}mm)")

        print("\n[가로 병합 그룹] (같은 너비를 가진 셀들)")
        if result['horizontal_groups']:
            for group in result['horizontal_groups']:
                width_mm = group['width'] / 7200 * 25.4  # HWPUNIT → mm
                print(f"  너비 {group['width']} ({width_mm:.1f}mm) → colspan={group['colspan']}: 셀 {group['cells']}")
        else:
            print("  없음")

        print("\n[세로 병합 그룹] (같은 높이를 가진 셀들)")
        if result['vertical_groups']:
            for group in result['vertical_groups']:
                height_mm = group['height'] / 7200 * 25.4  # HWPUNIT → mm
                print(f"  높이 {group['height']} ({height_mm:.1f}mm) → rowspan={group['rowspan']}: 셀 {group['cells']}")
        else:
            print("  없음")

        print("\n[각 셀의 병합 정보]")
        if result['cell_merge_info']:
            merged_cells = [(cid, info) for cid, info in result['cell_merge_info'].items()
                          if info['colspan'] > 1 or info['rowspan'] > 1]
            if merged_cells:
                for cell_id, info in merged_cells:
                    print(f"  셀 {cell_id}: colspan={info['colspan']}, rowspan={info['rowspan']}")
            else:
                print("  병합된 셀 없음 (모든 셀이 1x1)")
        else:
            print("  정보 없음")

    def build_coordinate_map(self) -> Dict[tuple, int]:
        """
        6. 테이블 좌표 매핑 (병합되기 전 테이블과 list_id 매핑)

        첫 셀에서 시작하여 오른쪽으로만 이동하며 좌표 매핑
        - 열 수 채우면 다음 행으로 넘김
        - colspan이 있는 셀은 여러 좌표가 같은 list_id
        - 마지막 list_id 도달 시 종료

        Returns:
            Dict[tuple, int]: {(row, col): list_id} 매핑
        """
        if not self.cells:
            self.collect_cells_bfs()

        if not self.cells:
            return {}

        # 병합 정보 가져오기
        merge_info = self.find_merged_by_dimensions()
        cell_merge = merge_info.get('cell_merge_info', {})

        # 테이블 크기 가져오기
        size = self.get_table_size()
        total_cols = size['cols']

        # 마지막 list_id
        last_list_id = max(self.cells.keys())
        self._log(f"마지막 list_id: {last_list_id}, 열 수: {total_cols}")

        # 첫 번째 셀로 이동
        if not self.move_to_first_cell():
            return {}

        coord_map = {}  # (row, col) → list_id
        row = 0
        col = 0

        while True:
            current_id = self._get_list_id()

            # 현재 셀의 colspan 가져오기
            colspan = 1
            if current_id in cell_merge:
                colspan = cell_merge[current_id].get('colspan', 1)

            # colspan만큼 좌표 매핑 (같은 list_id)
            for i in range(colspan):
                coord_map[(row, col)] = current_id
                self._log(f"좌표 ({row}, {col}) → 셀 {current_id}")
                col += 1

                # 열 수 채우면 다음 행
                if col >= total_cols:
                    row += 1
                    col = 0

            # 마지막 list_id 도달 시 종료
            if current_id == last_list_id:
                self._log(f"마지막 셀 {last_list_id} 도달")
                break

            # 오른쪽으로만 이동
            before = current_id
            self.hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
            after = self._get_list_id()

            # 이동 실패 시 종료
            if after == before:
                self._log(f"이동 불가, 종료")
                break

        self._log(f"좌표 매핑 완료: {len(coord_map)}개")
        return coord_map

    def print_coordinate_map(self):
        """좌표 매핑 정보 출력"""
        coord_map = self.build_coordinate_map()

        if not coord_map:
            print("좌표 매핑 없음")
            return

        # 테이블 크기 계산
        max_row = max(r for r, c in coord_map.keys())
        max_col = max(c for r, c in coord_map.keys())

        print(f"\n=== 6. 테이블 좌표 매핑 ({max_row+1}행 x {max_col+1}열) ===")
        print("\n[좌표 → list_id 매핑]")

        for row in range(max_row + 1):
            row_str = f"행 {row}: "
            cells = []
            for col in range(max_col + 1):
                list_id = coord_map.get((row, col), '-')
                cells.append(f"({row},{col})→{list_id}")
            print(row_str + "  ".join(cells))


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
        w_mm = cell.width / 7200 * 25.4 if cell.width > 0 else 0
        h_mm = cell.height / 7200 * 25.4 if cell.height > 0 else 0
        print(f"   셀 {cell_id}: L={cell.left}, R={cell.right}, U={cell.up}, D={cell.down}, W={w_mm:.1f}mm, H={h_mm:.1f}mm")

    print("\n3. 중복 이웃 분석...")
    table.print_duplicate_neighbors()

    print("\n4. 테이블 크기 계산...")
    size = table.get_table_size()
    print(f"   행 수: {size['rows']}")
    print(f"   열 수: {size['cols']}")

    print("\n5. 너비/높이 기반 병합 분석...")
    table.print_merged_by_dimensions()

    print("\n6. 좌표 매핑...")
    table.print_coordinate_map()

    print("\n=== 테스트 완료 ===")
