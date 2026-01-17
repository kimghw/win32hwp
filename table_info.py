# 테이블 좌표 매핑 모듈
#
# 주요 기능: build_coordinate_map()
# - 테이블의 (row, col) 좌표를 list_id에 매핑
# - colspan 셀은 여러 좌표가 같은 list_id를 가짐
# - 예: {(0,0): 2, (0,1): 3, (0,2): 4, (0,3): 4, ...} 

from dataclasses import dataclass
from typing import Dict
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

        # 행의 시작 → 열의 시작(맨 위)으로 이동
        self.hwp.MovePos(MOVE_START_OF_CELL, 0, 0)
        self.hwp.MovePos(MOVE_TOP_OF_CELL, 0, 0)

        # 왼쪽으로 더 이동 가능한지 확인
        before = self._get_list_id()
        result = self.hwp.MovePos(MOVE_LEFT_OF_CELL, 0, 0)
        after = self._get_list_id()

        if result and after != before:
            self.hwp.MovePos(MOVE_START_OF_CELL, 0, 0)

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

        row = 0
        while True:
            row_start_pos = self.hwp.GetPos()

            # 현재 행을 오른쪽으로 순회
            while True:
                current_id = self._get_list_id()

                if current_id not in self.cells:
                    cell_info = self._collect_neighbors(current_id)
                    self.cells[current_id] = cell_info

                before = current_id
                result = self.hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
                after = self._get_list_id()

                if not result or after == before:
                    break

            # 다음 행으로 이동
            self.hwp.SetPos(row_start_pos[0], row_start_pos[1], row_start_pos[2])
            before = self._get_list_id()
            result = self.hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
            after = self._get_list_id()

            if not result or after == before:
                break

            row += 1

        self._log(f"셀 수집 완료: {len(self.cells)}개")
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
        cell_merge = self._get_cell_merge_info()

        # 첫 번째 셀로 이동
        if not self.move_to_first_cell():
            return {'rows': 0, 'cols': 0}

        # 열 수 계산: 첫 행(up=0) 셀만 카운트 + 병합 추가
        # 세로 병합 셀을 지나면 다른 행으로 넘어가므로 up=0인 셀만 첫 행으로 인정
        first_row_cells = []

        for cell_id, cell in self.cells.items():
            if cell.up == 0:  # 위에 셀이 없음 = 첫 행
                first_row_cells.append(cell_id)

        first_row_cells.sort()

        # 물리적 셀 개수 + colspan 추가
        cols = len(first_row_cells)
        for cell_id in first_row_cells:
            if cell_id in cell_merge:
                colspan = cell_merge[cell_id].get('colspan', 1)
                if colspan > 1:
                    cols += (colspan - 1)

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

        return {
            'rows': rows,
            'cols': cols
        }

    def _get_cell_merge_info(self) -> Dict[int, Dict]:
        """각 셀의 colspan/rowspan 계산 (내부용)"""
        if not self.cells:
            return {}

        # 기준 너비/높이 (가장 작은 값 = 병합되지 않은 셀)
        all_widths = [c.width for c in self.cells.values() if c.width > 0]
        all_heights = [c.height for c in self.cells.values() if c.height > 0]

        base_width = min(all_widths) if all_widths else 0
        base_height = min(all_heights) if all_heights else 0

        cell_merge_info = {}
        for cell_id, cell in self.cells.items():
            colspan = round(cell.width / base_width) if base_width > 0 and cell.width > 0 else 1
            rowspan = round(cell.height / base_height) if base_height > 0 and cell.height > 0 else 1

            cell_merge_info[cell_id] = {
                'colspan': max(1, colspan),
                'rowspan': max(1, rowspan)
            }

        return cell_merge_info

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
        cell_merge = self._get_cell_merge_info()

        # 테이블 크기 가져오기
        size = self.get_table_size()
        total_cols = size['cols']

        # 마지막 list_id
        last_list_id = max(self.cells.keys())

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
                col += 1

                # 열 수 채우면 다음 행
                if col >= total_cols:
                    row += 1
                    col = 0

            # 마지막 list_id 도달 시 종료
            if current_id == last_list_id:
                break

            # 오른쪽으로만 이동
            before = current_id
            self.hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
            after = self._get_list_id()

            if after == before:
                break

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
    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        exit(1)

    table = TableInfo(hwp, debug=False)

    if not table.is_in_table():
        print("[오류] 커서가 테이블 내부에 있지 않습니다.")
        exit(1)

    # 셀 수집
    cells = table.collect_cells_bfs()
    print(f"수집된 셀: {len(cells)}개")

    # 테이블 크기
    size = table.get_table_size()
    print(f"테이블 크기: {size['rows']}행 x {size['cols']}열")

    # 좌표 매핑
    table.print_coordinate_map()
