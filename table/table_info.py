# 테이블 좌표 매핑 모듈
#
# 주요 기능: build_coordinate_map()
# - 테이블의 (row, col) 좌표를 list_id에 매핑
# - colspan 셀은 여러 좌표가 같은 list_id를 가짐
# - 예: {(0,0): 2, (0,1): 3, (0,2): 4, (0,3): 4, ...} 

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from dataclasses import dataclass, field
from typing import Dict, List, Tuple, Optional
from cursor import get_hwp_instance


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
        self._coord_map: Dict[Tuple[int, int], int] = {}  # (row, col) -> list_id
        self._representative_coords: Dict[int, Tuple[int, int]] = {}  # list_id -> 대표 좌표
        self._cell_coords: Dict[int, List[Tuple[int, int]]] = {}  # list_id -> 해당 셀의 모든 좌표
        self._table_size: Dict[str, int] = {}  # 캐시된 테이블 크기

    def _log(self, msg: str):
        """디버그 메시지 출력"""
        if self.debug:
            print(f"[TableInfo] {msg}")

    def _get_list_id(self) -> int:
        """현재 커서 위치의 list_id 반환"""
        pos = self.hwp.GetPos()
        return pos[0]

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
        # 캐시된 값이 있으면 반환
        if self._table_size:
            return self._table_size

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

        # 캐시에 저장
        self._table_size = {
            'rows': rows,
            'cols': cols
        }
        return self._table_size

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

    def build_coordinate_map(self) -> Dict[Tuple[int, int], int]:
        """
        테이블 좌표 매핑 (병합 셀 포함)

        - colspan/rowspan을 고려하여 모든 좌표를 list_id에 매핑
        - 병합 셀은 여러 좌표가 같은 list_id를 가짐
        - 대표 좌표 = 병합 영역의 (min_row, min_col)

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
        total_rows = size['rows']
        total_cols = size['cols']

        if total_rows == 0 or total_cols == 0:
            return {}

        # 2D 그리드 초기화 (None = 비어있음)
        grid = [[None for _ in range(total_cols)] for _ in range(total_rows)]

        # 첫 번째 셀로 이동
        if not self.move_to_first_cell():
            return {}

        # 초기화
        self._coord_map = {}
        self._representative_coords = {}
        self._cell_coords = {}

        # 모든 셀 순회하며 그리드 채우기 (list_id 순서대로)
        visited_cells = set()

        # list_id 순으로 정렬하여 처리 (왼쪽 위 → 오른쪽 아래 순서)
        for list_id in sorted(self.cells.keys()):
            if list_id in visited_cells:
                continue

            visited_cells.add(list_id)

            # 병합 정보
            colspan = cell_merge.get(list_id, {}).get('colspan', 1)
            rowspan = cell_merge.get(list_id, {}).get('rowspan', 1)

            # 이 셀의 시작 위치 찾기 (그리드에서 비어있는 첫 위치)
            start_row, start_col = None, None
            for r in range(total_rows):
                for c in range(total_cols):
                    if grid[r][c] is None:
                        # 이 위치에서 시작할 수 있는지 확인
                        can_place = True
                        for dr in range(rowspan):
                            for dc in range(colspan):
                                rr, cc = r + dr, c + dc
                                if rr >= total_rows or cc >= total_cols or grid[rr][cc] is not None:
                                    can_place = False
                                    break
                            if not can_place:
                                break

                        if can_place:
                            start_row, start_col = r, c
                            break
                if start_row is not None:
                    break

            if start_row is None:
                continue

            # 해당 셀이 차지하는 모든 좌표 기록
            cell_coords = []
            for dr in range(rowspan):
                for dc in range(colspan):
                    r, c = start_row + dr, start_col + dc
                    if r < total_rows and c < total_cols:
                        grid[r][c] = list_id
                        self._coord_map[(r, c)] = list_id
                        cell_coords.append((r, c))

            # 대표 좌표 = (min_row, min_col) = 가장 위, 가장 왼쪽
            if cell_coords:
                rep_coord = (min(r for r, c in cell_coords),
                             min(c for r, c in cell_coords))
                self._representative_coords[list_id] = rep_coord
                self._cell_coords[list_id] = cell_coords

        return self._coord_map

    def get_representative_coord(self, list_id: int) -> Optional[Tuple[int, int]]:
        """
        병합 셀의 대표 좌표 반환 (가장 위, 가장 왼쪽)

        Args:
            list_id: 셀의 list_id

        Returns:
            (row, col) 대표 좌표, 없으면 None
        """
        if not self._representative_coords:
            self.build_coordinate_map()
        return self._representative_coords.get(list_id)

    def get_cell_coords(self, list_id: int) -> List[Tuple[int, int]]:
        """
        셀이 차지하는 모든 좌표 반환

        Args:
            list_id: 셀의 list_id

        Returns:
            [(row, col), ...] 좌표 리스트
        """
        if not self._cell_coords:
            self.build_coordinate_map()
        return self._cell_coords.get(list_id, [])

    def get_merge_info(self, list_id: int) -> Dict:
        """
        셀의 병합 정보 반환

        Returns:
            {'colspan': int, 'rowspan': int, 'representative': (row, col), 'coords': [...]}
        """
        if not self._representative_coords:
            self.build_coordinate_map()

        cell_merge = self._get_cell_merge_info()
        merge = cell_merge.get(list_id, {'colspan': 1, 'rowspan': 1})

        return {
            'colspan': merge.get('colspan', 1),
            'rowspan': merge.get('rowspan', 1),
            'representative': self._representative_coords.get(list_id),
            'coords': self._cell_coords.get(list_id, [])
        }

    def print_coordinate_map(self):
        """좌표 매핑 정보 출력 (대표 좌표 포함)"""
        coord_map = self.build_coordinate_map()

        if not coord_map:
            print("좌표 매핑 없음")
            return

        # 테이블 크기 계산
        max_row = max(r for r, c in coord_map.keys())
        max_col = max(c for r, c in coord_map.keys())

        print(f"\n=== 테이블 좌표 매핑 ({max_row+1}행 x {max_col+1}열) ===")
        print("\n[좌표 → list_id 매핑]")

        for row in range(max_row + 1):
            row_str = f"행 {row}: "
            cells = []
            for col in range(max_col + 1):
                list_id = coord_map.get((row, col), '-')
                cells.append(f"({row},{col})→{list_id}")
            print(row_str + "  ".join(cells))

        # 병합 셀 정보 출력
        merged_cells = [(lid, coords) for lid, coords in self._cell_coords.items() if len(coords) > 1]
        if merged_cells:
            print("\n[병합 셀 정보]")
            for list_id, coords in merged_cells:
                rep = self._representative_coords.get(list_id)
                merge = self._get_cell_merge_info().get(list_id, {})
                colspan = merge.get('colspan', 1)
                rowspan = merge.get('rowspan', 1)
                print(f"  list_id={list_id}: {rowspan}x{colspan} 병합, 대표좌표={rep}, 좌표={coords}")

    def find_all_tables(self) -> List[Dict]:
        """
        문서 내 모든 테이블을 찾아 정보 반환

        HeadCtrl에서 시작하여 Next로 순회하며 CtrlID="tbl"인 컨트롤 수집

        Returns:
            list: [{'num': 번호, 'ctrl': 컨트롤, 'first_cell_list_id': 첫셀ID}, ...]
        """
        tables = []
        ctrl = self.hwp.HeadCtrl
        table_num = 0

        while ctrl:
            if ctrl.CtrlID == "tbl":
                # 테이블 위치로 이동 후 선택
                self.hwp.SetPosBySet(ctrl.GetAnchorPos(0))
                self.hwp.HAction.Run("SelectCtrlFront")

                # 첫 번째 셀 선택 (테이블 선택 상태에서)
                self.hwp.HAction.Run("ShapeObjTableSelCell")

                # 첫 셀의 list_id
                pos = self.hwp.GetPos()
                first_cell_list_id = pos[0]

                tables.append({
                    'num': table_num,
                    'ctrl': ctrl,
                    'first_cell_list_id': first_cell_list_id
                })

                self._log(f"테이블 {table_num}: 첫번째 셀 list_id={first_cell_list_id}")

                # 선택 해제
                self.hwp.HAction.Run("Cancel")
                self.hwp.HAction.Run("MoveParentList")

                table_num += 1

            ctrl = ctrl.Next

        return tables

    def select_table(self, table_index: int):
        """
        특정 번호의 테이블 선택

        Args:
            table_index: 테이블 번호 (0부터 시작)

        Returns:
            ctrl: 선택된 테이블 컨트롤 (없으면 None)
        """
        ctrl = self.hwp.HeadCtrl
        current_index = 0

        while ctrl:
            if ctrl.CtrlID == "tbl":
                if current_index == table_index:
                    # 테이블 선택
                    self.hwp.SetPosBySet(ctrl.GetAnchorPos(0))
                    if not self.hwp.HAction.Run("SelectCtrlFront"):
                        self.hwp.HAction.Run("SelectCtrlReverse")
                    self._log(f"테이블 {table_index} 선택됨")
                    return ctrl
                current_index += 1
            ctrl = ctrl.Next

        self._log(f"테이블 {table_index} 없음")
        return None

    def enter_table(self, table_index: int) -> bool:
        """
        특정 번호의 테이블 첫 번째 셀로 진입

        Args:
            table_index: 테이블 번호 (0부터 시작)

        Returns:
            bool: 성공 여부
        """
        ctrl = self.select_table(table_index)
        if not ctrl:
            return False

        # 첫 번째 셀로 진입
        self.hwp.HAction.Run("ShapeObjTableSelCell")
        return True

    def has_caption(self, table_index: int = None) -> bool:
        """
        테이블에 캡션이 있는지 확인

        캡션 텍스트를 실제로 가져와서 존재 여부 판단

        Args:
            table_index: 테이블 번호 (None이면 현재 선택된 테이블)

        Returns:
            bool: 캡션 존재 여부
        """
        caption = self.get_table_caption(table_index)
        return len(caption) > 0

    def get_table_caption(self, table_index: int = None) -> str:
        """
        테이블 캡션 텍스트 가져오기

        캡션은 테이블 마지막 셀 다음 list_id를 가짐
        해당 위치로 이동하여 텍스트 추출

        Args:
            table_index: 테이블 번호 (None이면 현재 테이블)

        Returns:
            str: 캡션 텍스트 (없으면 빈 문자열)
        """
        if table_index is not None:
            ctrl = self.select_table(table_index)
            if not ctrl:
                return ""
            # 테이블 첫 셀로 진입
            self.hwp.HAction.Run("ShapeObjTableSelCell")

        try:
            # 테이블의 마지막 셀로 이동
            self.hwp.MovePos(MOVE_BOTTOM_OF_CELL, 0, 0)  # 맨 아래 행
            self.hwp.MovePos(MOVE_END_OF_CELL, 0, 0)     # 맨 오른쪽 열

            last_cell_list_id = self._get_list_id()

            # 캡션 list_id = 마지막 셀 + 1
            caption_list_id = last_cell_list_id + 1

            # 캡션 위치로 이동 시도
            self.hwp.SetPos(caption_list_id, 0, 0)
            current_list_id = self._get_list_id()

            # 이동 성공 확인
            if current_list_id != caption_list_id:
                return ""

            # 캡션 텍스트 읽기 - 해당 리스트의 텍스트 추출
            self.hwp.HAction.Run("SelectAll")
            text = self.hwp.GetTextFile("TEXT", "")
            self.hwp.HAction.Run("Cancel")

            if text:
                return text.strip()

            return ""
        except Exception as e:
            self._log(f"캡션 조회 실패: {e}")
            return ""

    def get_all_table_captions(self) -> List[Dict]:
        """
        문서 내 모든 테이블의 캡션 가져오기

        Returns:
            list: [{'num': 번호, 'caption': 캡션텍스트}, ...]
        """
        captions = []
        tables = self.find_all_tables()

        for t in tables:
            caption = self.get_table_caption(t['num'])
            captions.append({
                'num': t['num'],
                'first_cell_list_id': t['first_cell_list_id'],
                'caption': caption
            })

        return captions
