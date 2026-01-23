# -*- coding: utf-8 -*-
"""
테이블 경계 판별 모듈

테이블의 경계(첫 번째 행, 마지막 행, 첫 번째 열, 마지막 열)를 판별하는 함수 제공
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import json
from dataclasses import dataclass, field, asdict
from typing import Dict, List, Tuple, Optional

try:
    from .table_info import (
        TableInfo, MOVE_LEFT_OF_CELL, MOVE_RIGHT_OF_CELL,
        MOVE_UP_OF_CELL, MOVE_DOWN_OF_CELL
    )
except ImportError:
    from table_info import (
        TableInfo, MOVE_LEFT_OF_CELL, MOVE_RIGHT_OF_CELL,
        MOVE_UP_OF_CELL, MOVE_DOWN_OF_CELL
    )


@dataclass
class TableBoundaryResult:
    """테이블 경계 분석 결과"""
    # 셀 ID 경계
    table_origin: int = 0       # 테이블 첫 번째 셀의 list_id
    table_end: int = 0          # last_cols의 마지막 셀의 list_id
    table_cell_counts: int = 0  # 총 셀 개수

    # 4방향 경계 셀 리스트
    first_rows: List[int] = field(default_factory=list)   # 첫 번째 행 셀들
    bottom_rows: List[int] = field(default_factory=list)  # 마지막 행 셀들
    first_cols: List[int] = field(default_factory=list)   # 첫 번째 열 셀들
    last_cols: List[int] = field(default_factory=list)    # 마지막 열 셀들

    # 좌표 경계 (HWPUNIT)
    start_x: int = 0            # 테이블 시작 x (항상 0)
    start_y: int = 0            # 테이블 시작 y (항상 0)
    end_x: int = 0              # 테이블 끝 x (xend)
    end_y: int = 0              # 테이블 끝 y (yend)


@dataclass
class SubTableResult:
    """서브 테이블 결과"""
    start: int = 0              # 서브셋 시작 셀 list_id
    end: int = 0                # 서브셋 마지막 셀 list_id
    cells: List[int] = field(default_factory=list)  # 서브셋에 포함된 셀들
    next_subset: Optional[int] = None  # 세로 분할된 다음 서브셋 시작 위치


@dataclass
class SubCellInfo:
    """서브셀 정보 (크기 기반 분할)"""
    list_id: int
    width: int = 0
    height: int = 0
    is_new_subcell: bool = False  # 크기가 달라져 새 서브셀로 정의됨


@dataclass
class RowSubsetResult:
    """행 단위 서브셋 분할 결과"""
    row_index: int = 0                  # first_cols 인덱스 (행 번호)
    start_cell: int = 0                 # 행 시작 셀 list_id
    cells: List[SubCellInfo] = field(default_factory=list)  # 행의 모든 셀
    subcell_boundaries: List[int] = field(default_factory=list)  # 서브셀 경계 list_id들


@dataclass
class CellWidthInfo:
    """셀 너비 정보"""
    list_id: int
    width: int = 0              # 셀 너비 (HWPUNIT)
    height: int = 0             # 셀 높이 (HWPUNIT)
    start_x: int = 0            # 셀 시작 x 좌표
    end_x: int = 0              # 셀 끝 x 좌표
    col_index: int = 0          # 행 내 열 인덱스 (0부터)


@dataclass
class RowWidthResult:
    """행별 셀 너비 결과"""
    row_index: int = 0                              # first_cols 인덱스 (행 번호)
    start_cell: int = 0                             # 행 시작 셀 (first_col)
    end_cell: int = 0                               # 행 끝 셀 (last_col)
    cells: List[CellWidthInfo] = field(default_factory=list)  # 행의 모든 셀 너비 정보
    total_width: int = 0                            # 행 전체 너비
    start_y: int = 0                                # 행 시작 y 좌표
    end_y: int = 0                                  # 행 끝 y 좌표
    row_height: int = 0                             # 행 높이 (첫 번째 셀 기준)


@dataclass
class TableWidthResult:
    """테이블 전체 셀 너비 결과"""
    rows: List[RowWidthResult] = field(default_factory=list)
    max_width: int = 0                              # 가장 넓은 행의 너비


class TableBoundary:
    """테이블 경계 판별 클래스"""

    def __init__(self, hwp=None, debug: bool = False):
        from cursor import get_hwp_instance
        self.hwp = hwp or get_hwp_instance()
        self.debug = debug
        self._table_info = TableInfo(self.hwp, debug)
        self._current_tbl = None  # 현재 테이블의 tbl 컨트롤

    def _log(self, msg: str):
        """디버그 메시지 출력"""
        if self.debug:
            print(f"[TableBoundary] {msg}")

    def _is_in_table(self) -> bool:
        """현재 커서 위치가 테이블 내부인지 확인"""
        parent = self.hwp.ParentCtrl
        return parent and parent.CtrlID == "tbl"

    def move_down_left_right(self, target_list_id: int) -> Dict[str, Tuple[int, bool]]:
        """
        특정 셀에서 하/좌/우 각각 이동 시 list_id와 has_tbl 반환

        Args:
            target_list_id: 대상 셀의 list_id

        Returns:
            dict: {
                'down': (list_id, has_tbl),
                'left': (list_id, has_tbl),
                'right': (list_id, has_tbl)
            }
        """
        def move_and_check(move_const: int) -> Tuple[int, bool]:
            self.hwp.SetPos(target_list_id, 0, 0)
            self.hwp.MovePos(move_const, 0, 0)
            list_id, _, _ = self.hwp.GetPos()
            parent = self.hwp.ParentCtrl
            has_tbl = parent and parent.CtrlID == "tbl"
            return (list_id, has_tbl)

        return {
            'down': move_and_check(MOVE_DOWN_OF_CELL),
            'left': move_and_check(MOVE_LEFT_OF_CELL),
            'right': move_and_check(MOVE_RIGHT_OF_CELL)
        }

    def check_first_row_cell(self, target_list_id: int) -> bool:
        """
        위로 이동 시 tbl 밖이거나 같은 셀에 머물면 True (첫 번째 행)

        Args:
            target_list_id: 대상 셀의 list_id

        Returns:
            bool: 첫 번째 행이면 True
        """
        self.hwp.SetPos(target_list_id, 0, 0)
        self.hwp.HAction.Run("MoveUp")
        new_list_id, _, _ = self.hwp.GetPos()
        parent = self.hwp.ParentCtrl
        has_tbl = parent and parent.CtrlID == "tbl"
        # 테이블 밖으로 나가거나, 같은 셀에 머물러 있으면 첫 번째 행
        return not has_tbl or new_list_id == target_list_id

    def check_bottom_row_cell(self, target_list_id: int) -> bool:
        """
        아래로 이동 시 tbl 밖이거나 같은 셀에 머물면 True (마지막 행)

        Args:
            target_list_id: 대상 셀의 list_id

        Returns:
            bool: 마지막 행이면 True
        """
        self.hwp.SetPos(target_list_id, 0, 0)
        self.hwp.HAction.Run("MoveDown")
        new_list_id, _, _ = self.hwp.GetPos()
        parent = self.hwp.ParentCtrl
        has_tbl = parent and parent.CtrlID == "tbl"
        # 테이블 밖으로 나가거나, 같은 셀에 머물러 있으면 마지막 행
        return not has_tbl or new_list_id == target_list_id

    def move_up_right_down(self, target_list_id: int) -> Tuple[int, bool]:
        """
        특정 셀에서 하 → 우 → 상 순서로 커서 이동 후 last_col 여부 판별

        Args:
            target_list_id: 대상 셀의 list_id

        Returns:
            tuple: (down_id, is_last_col)
                - down_id: 아래로 이동한 셀의 list_id
                - is_last_col: 상으로 돌아온 셀이 시작점과 같으면 True (같은 열)
        """
        # 시작점에서 하로 이동
        self.hwp.SetPos(target_list_id, 0, 0)
        self.hwp.HAction.Run("MoveDown")
        down_id, _, _ = self.hwp.GetPos()

        # 우로 이동
        self.hwp.HAction.Run("MoveRight")

        # 상으로 이동
        self.hwp.HAction.Run("MoveUp")
        up_id, _, _ = self.hwp.GetPos()

        # 시작점과 같으면 같은 열 -> last_col
        is_last_col = (up_id == target_list_id)

        return (down_id, is_last_col)

    def _check_first_col_cell(self, target_list_id: int) -> bool:
        """
        왼쪽 셀로 이동 시 tbl 밖이거나 같은 셀이면 True (첫 번째 열)

        Args:
            target_list_id: 대상 셀의 list_id

        Returns:
            bool: 첫 번째 열이면 True
        """
        self.hwp.SetPos(target_list_id, 0, 0)
        result = self.hwp.MovePos(MOVE_LEFT_OF_CELL, 0, 0)
        new_list_id, _, _ = self.hwp.GetPos()
        new_tbl = self.hwp.ParentCtrl
        has_tbl = new_tbl and new_tbl.CtrlID == "tbl"
        # 이동 실패, tbl 밖, 또는 같은 셀이면 첫 번째 열
        return not result or not has_tbl or new_list_id == target_list_id

    def _check_last_col_cell(self, target_list_id: int) -> bool:
        """
        오른쪽 셀로 이동 시 tbl 밖이거나 같은 셀이면 True (마지막 열)

        Args:
            target_list_id: 대상 셀의 list_id

        Returns:
            bool: 마지막 열이면 True
        """
        self.hwp.SetPos(target_list_id, 0, 0)
        result = self.hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
        new_list_id, _, _ = self.hwp.GetPos()
        new_tbl = self.hwp.ParentCtrl
        has_tbl = new_tbl and new_tbl.CtrlID == "tbl"
        # 이동 실패, tbl 밖, 또는 같은 셀이면 마지막 열
        return not result or not has_tbl or new_list_id == target_list_id

    def _calc_xend_from_first_rows(self, first_rows: List[int]) -> int:
        """
        first_rows 셀들의 너비 합 = xend (테이블 전체 가로 크기)

        Args:
            first_rows: 첫 번째 행의 셀 list_id 리스트

        Returns:
            int: xend (HWPUNIT)
        """
        if not first_rows:
            return 0

        first_rows_set = set(first_rows)
        start_cell = first_rows[0]
        self.hwp.SetPos(start_cell, 0, 0)

        cumulative_x = 0
        current_id = start_cell

        while current_id in first_rows_set:
            w, _ = self._table_info.get_cell_dimensions()
            cumulative_x += w

            self.hwp.SetPos(current_id, 0, 0)
            self.hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
            next_id = self.hwp.GetPos()[0]

            if next_id == current_id:
                break

            current_id = next_id

        self._log(f"[xend] first_rows: {first_rows}, xend={cumulative_x}")
        return cumulative_x

    def _calc_yend_from_first_cols(self, first_cols: List[int]) -> int:
        """
        first_cols 셀들의 높이 합 = yend (테이블 전체 세로 크기)

        Args:
            first_cols: 첫 번째 열의 셀 list_id 리스트

        Returns:
            int: yend (HWPUNIT)
        """
        if not first_cols:
            return 0

        # 중복 제거 (순서 유지)
        seen = set()
        unique_first_cols = []
        for cell_id in first_cols:
            if cell_id not in seen:
                seen.add(cell_id)
                unique_first_cols.append(cell_id)

        cumulative_y = 0

        for cell_id in unique_first_cols:
            self.hwp.SetPos(cell_id, 0, 0)
            _, h = self._table_info.get_cell_dimensions()
            cumulative_y += h

        self._log(f"[yend] first_cols (unique): {unique_first_cols}, yend={cumulative_y}")
        return cumulative_y

    def _find_lastcols_by_xend(self, start_cell: int, xend: int, tolerance: int = 50) -> dict:
        """
        start_cell부터 우측으로 순회하면서 xend 기준으로 last_cols 찾기

        - xend 초과 시: 이전 셀 = last_col, 현재 셀 = first_col
        - xend 도달 시: 현재 셀 = last_col

        Args:
            start_cell: 시작 셀 list_id
            xend: 테이블 전체 가로 크기 (HWPUNIT)
            tolerance: 허용 오차 (HWPUNIT)

        Returns:
            dict: {'first_cols': [...], 'last_cols': [...]}
        """
        self.hwp.SetPos(start_cell, 0, 0)
        current_id = start_cell
        cumulative_x = 0
        last_cols = []
        first_cols = [start_cell]
        prev_id = None

        max_iterations = 1000

        for _ in range(max_iterations):
            self.hwp.SetPos(current_id, 0, 0)
            w, _ = self._table_info.get_cell_dimensions()
            cumulative_x += w

            # xend 초과 → 이전 셀이 last_col, 현재 셀은 새 행의 first_col
            if cumulative_x > xend + tolerance:
                if prev_id is not None:
                    last_cols.append(prev_id)

                first_cols.append(current_id)
                cumulative_x = w
                prev_id = None

            # xend 도달 (오차 범위 내)
            elif abs(cumulative_x - xend) <= tolerance:
                last_cols.append(current_id)

                # 우측 이동
                self.hwp.SetPos(current_id, 0, 0)
                self.hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
                next_id = self.hwp.GetPos()[0]

                if next_id == current_id:
                    break

                first_cols.append(next_id)
                cumulative_x = 0
                prev_id = None
                current_id = next_id
                continue

            # 우측 이동
            prev_id = current_id
            self.hwp.SetPos(current_id, 0, 0)
            self.hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
            next_id = self.hwp.GetPos()[0]

            if next_id == current_id:
                last_cols.append(current_id)
                break

            current_id = next_id

        self._log(f"[xend 순회] first_cols: {first_cols}")
        self._log(f"[xend 순회] last_cols: {last_cols}")

        return {
            'first_cols': first_cols,
            'last_cols': last_cols,
        }

    def map_grid_by_xend(self, start_cell: int = None, xend: int = None, tolerance: int = 50) -> dict:
        """
        xend 기준으로 그리드 좌표 매핑 (x좌표 기반 컬럼 결정)

        - cumulative_x > xend → 새 행 시작
        - cumulative_x == xend → 현재 셀이 마지막 열
        - 컬럼 위치는 셀의 start_x 좌표로 결정 (rowspan 셀 일관성 유지)

        Args:
            start_cell: 시작 셀 list_id (없으면 자동 계산)
            xend: 테이블 전체 가로 크기 (없으면 자동 계산)
            tolerance: 허용 오차 (HWPUNIT)

        Returns:
            dict: {
                'grid': {list_id: {'row': r, 'col': c, 'start_x': x, 'width': w, 'height': h}},
                'rowspan_positions': {list_id: [(row, col), ...]},
                'col_positions': [x좌표들],  # 컬럼 경계 x좌표
                'max_row': 최대 행,
                'max_col': 최대 열,
                'xend': 사용된 xend 값
            }
        """
        # xend와 start_cell 자동 계산
        if xend is None or start_cell is None:
            cells = self._table_info.collect_cells_bfs()
            if not cells:
                return {'grid': {}, 'rowspan_positions': {}, 'col_positions': [],
                        'max_row': 0, 'max_col': 0, 'xend': 0}

            all_list_ids = sorted(cells.keys())

            first_rows = []
            for list_id in all_list_ids:
                if self.check_first_row_cell(list_id):
                    first_rows.append(list_id)
            first_rows.sort()

            if not first_rows:
                return {'grid': {}, 'rowspan_positions': {}, 'col_positions': [],
                        'max_row': 0, 'max_col': 0, 'xend': 0}

            if start_cell is None:
                start_cell = first_rows[0]
            if xend is None:
                xend = self._calc_xend_from_first_rows(first_rows)

        self._log(f"[map_grid] start_cell={start_cell}, xend={xend}")

        grid = {}
        rowspan_positions = {}
        col_positions = set()  # 모든 컬럼 시작 x좌표 수집
        row = 0
        start_x = 0  # 현재 셀의 시작 x좌표
        cumulative_x = 0
        max_col = 0

        self.hwp.SetPos(start_cell, 0, 0)
        current_id = start_cell

        # 1차 순회: 모든 셀의 start_x 수집
        temp_cells = []
        for _ in range(1000):
            self.hwp.SetPos(current_id, 0, 0)
            w, h = self._table_info.get_cell_dimensions()

            # xend 초과 → 새 행
            if cumulative_x + w > xend + tolerance:
                row += 1
                start_x = 0
                cumulative_x = 0

            temp_cells.append({
                'id': current_id,
                'row': row,
                'start_x': start_x,
                'width': w,
                'height': h
            })
            col_positions.add(start_x)

            cumulative_x += w
            start_x = cumulative_x

            # xend 도달
            if abs(cumulative_x - xend) <= tolerance:
                self.hwp.SetPos(current_id, 0, 0)
                self.hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
                next_id = self.hwp.GetPos()[0]
                if next_id == current_id:
                    break
                row += 1
                start_x = 0
                cumulative_x = 0
                current_id = next_id
                continue

            self.hwp.SetPos(current_id, 0, 0)
            self.hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
            next_id = self.hwp.GetPos()[0]
            if next_id == current_id:
                break
            current_id = next_id

        # end_x도 컬럼 경계에 추가 (colspan 계산용)
        for cell in temp_cells:
            end_x = cell['start_x'] + cell['width']
            col_positions.add(end_x)

        # 컬럼 위치 정렬 (x좌표 → col 인덱스 매핑)
        sorted_col_positions = sorted(col_positions)
        x_to_col = {x: i for i, x in enumerate(sorted_col_positions)}
        max_col = len(sorted_col_positions) - 1 if sorted_col_positions else 0

        self._log(f"[map_grid] col_positions: {sorted_col_positions}")

        # 2차: 셀 매핑 (start_x → col, end_x → end_col로 colspan 계산)
        for cell in temp_cells:
            list_id = cell['id']
            cell_row = cell['row']
            cell_col = x_to_col.get(cell['start_x'], 0)
            end_x = cell['start_x'] + cell['width']
            end_col = x_to_col.get(end_x, cell_col + 1)
            colspan = end_col - cell_col

            if list_id not in grid:
                grid[list_id] = {
                    'row': cell_row,
                    'col': cell_col,
                    'colspan': colspan,
                    'start_x': cell['start_x'],
                    'width': cell['width'],
                    'height': cell['height']
                }
                self._log(f"  ({cell_row}, {cell_col}-{end_col}): id={list_id}, colspan={colspan}")
            else:
                if list_id not in rowspan_positions:
                    rowspan_positions[list_id] = []
                rowspan_positions[list_id].append((cell_row, cell_col))
                self._log(f"  ({cell_row}, {cell_col}): id={list_id} (rowspan)")

        self._log(f"[map_grid] 완료: {len(grid)}셀, {row+1}행 x {max_col+1}열")

        return {
            'grid': grid,
            'rowspan_positions': rowspan_positions,
            'col_positions': sorted_col_positions,
            'max_row': row,
            'max_col': max_col,
            'xend': xend
        }

    def check_boundary_table(self) -> TableBoundaryResult:
        """
        테이블의 모든 셀을 순회하면서 경계 정보 계산

        1. table_origin 계산 - 테이블 첫 번째 셀의 list_id
        2. first_rows, bottom_rows 계산 - 첫/마지막 행에 속한 셀들
        3. first_cols, last_cols 계산 - 첫/마지막 열에 속한 셀들
        4. table_end 계산 - bottom_rows의 마지막 list_id

        Returns:
            TableBoundaryResult: 경계 분석 결과
        """
        result = TableBoundaryResult()

        # 셀 정보 수집
        cells = self._table_info.collect_cells_bfs()
        if not cells:
            self._log("셀 정보를 수집할 수 없습니다")
            return result

        all_list_ids = sorted(cells.keys())
        self._log(f"수집된 셀: {len(all_list_ids)}개")

        # 첫 번째 셀로 이동 후 현재 테이블의 tbl 컨트롤 저장
        self.hwp.SetPos(all_list_ids[0], 0, 0)
        self._current_tbl = self.hwp.ParentCtrl
        self._log(f"current_tbl: {self._current_tbl}")

        # 1. table_origin = 첫 번째 셀의 list_id
        result.table_origin = all_list_ids[0]
        result.table_cell_counts = len(all_list_ids)
        self._log(f"table_origin: {result.table_origin}")
        self._log(f"table_cell_counts: {result.table_cell_counts}")

        # 2. first_rows, bottom_rows 계산
        for list_id in all_list_ids:
            if self.check_first_row_cell(list_id):
                result.first_rows.append(list_id)
            if self.check_bottom_row_cell(list_id):
                result.bottom_rows.append(list_id)

        result.first_rows.sort()
        result.bottom_rows.sort()
        self._log(f"first_rows: {result.first_rows}")
        self._log(f"bottom_rows: {result.bottom_rows}")

        # 3. xend 계산 및 last_cols/first_cols 계산 (xend 기반 우측 순회)
        xend = self._calc_xend_from_first_rows(result.first_rows)

        if xend > 0 and result.first_rows:
            xend_result = self._find_lastcols_by_xend(result.first_rows[0], xend)
            result.first_cols = xend_result['first_cols']
            result.last_cols = xend_result['last_cols']
        else:
            # fallback: table_origin에서 아래로 내려가면서 수집
            self.hwp.SetPos(result.table_origin, 0, 0)
            while True:
                current_id, _, _ = self.hwp.GetPos()
                result.first_cols.append(current_id)
                before_id = current_id
                self.hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
                after_id, _, _ = self.hwp.GetPos()
                if after_id == before_id:
                    break

        # table_end = last_cols의 마지막 list_id
        if result.last_cols:
            result.table_end = result.last_cols[-1]

        # 4. 좌표 경계 계산
        result.start_x = 0
        result.start_y = 0
        result.end_x = xend
        result.end_y = self._calc_yend_from_first_cols(result.first_cols)

        self._log(f"first_cols: {result.first_cols}")
        self._log(f"last_cols: {result.last_cols}")
        self._log(f"table_end: {result.table_end}")
        self._log(f"좌표 경계: ({result.start_x}, {result.start_y}) ~ ({result.end_x}, {result.end_y})")

        return result

    def print_boundary_info(self, result: TableBoundaryResult = None):
        """경계 정보 출력"""
        if result is None:
            result = self.check_boundary_table()

        print("\n=== 테이블 경계 정보 ===")
        print(f"table_origin: {result.table_origin}")
        print(f"table_end: {result.table_end}")
        print(f"table_cell_counts: {result.table_cell_counts}")
        print(f"first_rows: {result.first_rows}")
        print(f"bottom_rows: {result.bottom_rows}")
        print(f"first_cols: {result.first_cols}")
        print(f"last_cols: {result.last_cols}")
        print(f"좌표 경계: ({result.start_x}, {result.start_y}) ~ ({result.end_x}, {result.end_y})")

    def move_right_down_left(self, target_list_id: int, last_cols_set: set) -> Tuple[List[int], int, bool]:
        """
        특정 셀에서 우 → 하 → 좌 순서로 커서 이동하며 셀 수집

        Args:
            target_list_id: 대상 셀의 list_id
            last_cols_set: 마지막 열 셀들의 집합 (종료 조건)

        Returns:
            tuple: (collected_cells, last_cell, is_stopped)
                - collected_cells: 이동하며 수집한 셀들
                - last_cell: 마지막으로 도달한 셀
                - is_stopped: 같은 셀에서 멈춤 (서브셋 완성)
        """
        collected = []
        current = target_list_id

        while current not in last_cols_set:
            self.hwp.SetPos(current, 0, 0)

            # 우로 이동
            self.hwp.HAction.Run("MoveRight")
            right_id, _, _ = self.hwp.GetPos()
            if right_id != current:
                collected.append(right_id)
                current = right_id
                if current in last_cols_set:
                    break
                continue

            # 하로 이동
            self.hwp.HAction.Run("MoveDown")
            down_id, _, _ = self.hwp.GetPos()
            if down_id != current:
                collected.append(down_id)
                current = down_id
                if current in last_cols_set:
                    break
                continue

            # 좌로 이동
            self.hwp.HAction.Run("MoveLeft")
            left_id, _, _ = self.hwp.GetPos()
            if left_id != current:
                collected.append(left_id)
                current = left_id
                if current in last_cols_set:
                    break
                continue

            # 세 방향 모두 같은 셀 -> 멈춤
            return (collected, current, True)

        return (collected, current, False)

    def _sort_first_cols_by_position(self, first_cols: List[int]) -> List[int]:
        """first_cols를 y좌표(위치) 기준으로 정렬"""
        # 각 셀의 위치 정보 수집
        positions = []
        for list_id in first_cols:
            self.hwp.SetPos(list_id, 0, 0)
            # KeyIndicator로 페이지/줄 정보 획득
            key = self.hwp.KeyIndicator()
            page = key[3] if len(key) > 3 else 0
            line = key[4] if len(key) > 4 else 0
            positions.append((list_id, page, line))

        # 페이지, 줄 순서로 정렬
        positions.sort(key=lambda x: (x[1], x[2]))
        return [p[0] for p in positions]

    def sub_table(self, boundary_result: TableBoundaryResult = None) -> List[SubTableResult]:
        """
        테이블을 행 단위 서브셋으로 분할

        Args:
            boundary_result: 경계 분석 결과 (없으면 새로 계산)

        Returns:
            List[SubTableResult]: 서브 테이블 리스트
        """
        if boundary_result is None:
            boundary_result = self.check_boundary_table()

        subsets = []
        last_cols_set = set(boundary_result.last_cols)

        # 1. first_cols 정렬 (y좌표 기준)
        sorted_first_cols = self._sort_first_cols_by_position(boundary_result.first_cols)
        self._log(f"sorted_first_cols: {sorted_first_cols}")

        # 2. 각 first_cols에서 시작하여 서브셋 탐색
        for start_cell in sorted_first_cols:
            subset = SubTableResult(start=start_cell, cells=[start_cell])

            self._log(f"서브셋 시작: {start_cell}")

            # 우→하→좌 순회하며 셀 수집
            collected, last_cell, is_stopped = self.move_right_down_left(start_cell, last_cols_set)
            subset.cells.extend(collected)
            subset.end = last_cell

            if is_stopped:
                # 세로 분할 셀 탐색 (새 서브셋 시작점)
                vertical_start = self._find_vertical_split(last_cell)
                if vertical_start and vertical_start != last_cell:
                    subset.next_subset = vertical_start
                    self._log(f"세로 분할 연결: {vertical_start}")

            self._log(f"서브셋 완성: {subset.cells}")
            subsets.append(subset)

        return subsets

    def _find_vertical_split(self, current_id: int) -> Optional[int]:
        """
        세로 분할된 셀 탐색 (멈춘 셀에서 우→하→좌로 다른 셀 찾기)

        Args:
            current_id: 현재 셀의 list_id

        Returns:
            Optional[int]: 세로 분할된 새 서브셋의 시작 셀, 없으면 None
        """
        self.hwp.SetPos(current_id, 0, 0)

        # 우→하→좌 반복하며 다른 list_id 찾기
        for _ in range(10):  # 최대 10번 시도
            self.hwp.HAction.Run("MoveRight")
            self.hwp.HAction.Run("MoveDown")
            self.hwp.HAction.Run("MoveLeft")
            next_id, _, _ = self.hwp.GetPos()

            if next_id != current_id:
                return next_id

        return None

    def print_sub_table(self, subsets: List[SubTableResult] = None):
        """서브 테이블 정보 출력"""
        if subsets is None:
            subsets = self.sub_table()

        print("\n=== 서브 테이블 정보 ===")
        for i, subset in enumerate(subsets):
            print(f"[서브셋 {i}]")
            print(f"  start: {subset.start}")
            print(f"  end: {subset.end}")
            print(f"  cells: {subset.cells}")
            print(f"  next_subset: {subset.next_subset}")

    def _get_cell_size(self, list_id: int) -> Tuple[int, int]:
        """
        특정 셀의 너비/높이 조회

        Args:
            list_id: 셀의 list_id

        Returns:
            tuple: (width, height) HWPUNIT 단위
        """
        self.hwp.SetPos(list_id, 0, 0)
        return self._table_info.get_cell_dimensions()

    def sub_table_by_size(self, boundary_result: TableBoundaryResult = None) -> List[RowSubsetResult]:
        """
        테이블을 행 단위로 서브셋 분할 (셀 높이 기반)

        알고리즘:
        1. first_cols의 각 원소(행 시작 셀)에서 시작
        2. 우측으로 이동하면서 list_id와 셀의 높이/넓이 수집
        3. 우측 이동 시 list_id가 동일하거나 다음 first_col을 만나면 해당 행 종료
        4. 높이가 달라지는 셀을 새로운 subcell로 정의 (같은 높이 = 같은 서브셀)
        5. 마지막 first_cols에서 우측 이동 시 list_id 동일하면 전체 종료

        Args:
            boundary_result: 경계 분석 결과 (없으면 새로 계산)

        Returns:
            List[RowSubsetResult]: 행별 서브셋 분할 결과
        """
        if boundary_result is None:
            boundary_result = self.check_boundary_table()

        results = []
        first_cols = self._sort_first_cols_by_position(boundary_result.first_cols)
        first_cols_set = set(first_cols)
        last_first_col = first_cols[-1] if first_cols else None

        self._log(f"first_cols (정렬): {first_cols}")
        self._log(f"마지막 first_col: {last_first_col}")

        for row_idx, start_cell in enumerate(first_cols):
            row_result = RowSubsetResult(
                row_index=row_idx,
                start_cell=start_cell
            )

            # 이미 방문한 셀 추적 (순환 방지)
            visited = set()

            # 시작 셀 정보 수집
            width, height = self._get_cell_size(start_cell)
            current_cell = SubCellInfo(
                list_id=start_cell,
                width=width,
                height=height,
                is_new_subcell=True  # 행의 첫 셀은 항상 새 서브셀
            )
            row_result.cells.append(current_cell)
            row_result.subcell_boundaries.append(start_cell)
            visited.add(start_cell)

            prev_height = height
            current_id = start_cell

            self._log(f"행 {row_idx} 시작: list_id={start_cell}, size=({width}, {height})")

            # 우측으로 이동하면서 셀 수집 (다음 first_col 만날 때까지)
            while True:
                # 우측으로 이동
                self.hwp.SetPos(current_id, 0, 0)
                self.hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
                next_id, _, _ = self.hwp.GetPos()

                # list_id가 동일하면 종료 (더 이상 우측 이동 불가)
                if next_id == current_id:
                    self._log(f"  행 {row_idx} 종료: 더 이상 이동 불가 (list_id={current_id})")
                    break

                # 이미 방문한 셀이면 계속 우측으로 이동해서 새 셀 찾기
                loop_count = 0
                while next_id in visited and loop_count < 100:
                    loop_count += 1

                    # first_cols 원소를 만나면 경계 지점으로 추가
                    if next_id in first_cols_set and next_id != start_cell:
                        if next_id not in row_result.subcell_boundaries:
                            row_result.subcell_boundaries.append(next_id)
                            self._log(f"  경계 추가 (first_col 재방문): list_id={next_id}")

                    prev_next = next_id
                    self.hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
                    next_id, _, _ = self.hwp.GetPos()

                    # 같은 셀에 머무르면 종료
                    if next_id == prev_next:
                        self._log(f"  행 {row_idx} 종료: 우측 끝 도달 (list_id={next_id})")
                        break

                    # first_col 만나면 종료 (다음 행의 시작)
                    if next_id in first_cols_set and next_id != start_cell:
                        break

                # first_col 만나면 행 종료
                if next_id in first_cols_set and next_id != start_cell:
                    self._log(f"  행 {row_idx} 종료: 다음 first_col 도달 (list_id={next_id})")
                    break

                # 여전히 방문한 셀이면 종료
                if next_id in visited:
                    self._log(f"  행 {row_idx} 종료: 순환 감지 (list_id={next_id})")
                    break

                visited.add(next_id)

                # 셀 크기 수집
                width, height = self._get_cell_size(next_id)

                # 높이가 달라지면 새로운 subcell로 정의
                is_new_subcell = (height != prev_height)

                cell_info = SubCellInfo(
                    list_id=next_id,
                    width=width,
                    height=height,
                    is_new_subcell=is_new_subcell
                )
                row_result.cells.append(cell_info)

                # 높이가 달라지면 새로운 subcell로 정의
                if is_new_subcell:
                    if next_id not in row_result.subcell_boundaries:
                        row_result.subcell_boundaries.append(next_id)
                        self._log(f"  새 서브셀: list_id={next_id}, height={height}")

                prev_height = height
                current_id = next_id

            results.append(row_result)

            # 조건 5: 마지막 first_cols 처리 완료 후 종료
            if start_cell == last_first_col:
                self._log(f"마지막 first_col 처리 완료, 전체 종료")
                break

        return results

    def print_sub_table_by_size(self, results: List[RowSubsetResult] = None):
        """행 단위 서브셋 분할 결과 출력"""
        if results is None:
            results = self.sub_table_by_size()

        print("\n=== 행 단위 서브셋 분할 결과 ===")
        for row in results:
            print(f"\n[행 {row.row_index}] 시작 셀: {row.start_cell}")
            print(f"  서브셀 경계: {row.subcell_boundaries}")
            print(f"  셀 목록:")
            for cell in row.cells:
                marker = " *" if cell.is_new_subcell else ""
                print(f"    list_id={cell.list_id}, size=({cell.width}, {cell.height}){marker}")

    def _find_matching_last_col(self, start_cell: int, last_cols: List[int]) -> Optional[int]:
        """
        start_cell에서 오른쪽으로 이동하여 해당 행의 last_col 찾기

        Args:
            start_cell: 행 시작 셀
            last_cols: 정렬된 last_cols 리스트

        Returns:
            해당 행의 last_col, 없으면 None
        """
        last_cols_set = set(last_cols)
        visited = set()
        current_id = start_cell
        visited.add(current_id)

        max_steps = 1000
        for _ in range(max_steps):
            if current_id in last_cols_set:
                return current_id

            self.hwp.SetPos(current_id, 0, 0)
            self.hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
            next_id, _, _ = self.hwp.GetPos()

            if next_id == current_id:
                return current_id  # 더 이상 이동 불가 = 마지막 열

            if next_id in visited:
                current_id = next_id
                continue

            visited.add(next_id)
            current_id = next_id

        return None

    def calculate_row_widths(self, boundary_result: TableBoundaryResult = None) -> TableWidthResult:
        """
        first_cols에서 last_cols까지 오른쪽으로 이동하면서 셀 너비 누적 계산

        알고리즘:
        1. first_cols와 last_cols를 y좌표 기준으로 정렬하여 행 매칭
        2. 각 first_col에서 시작하여 해당 행의 last_col까지만 이동
        3. 이동할 때마다 셀 너비를 수집하고 누적 (이미 방문한 셀은 건너뜀)
        4. 해당 행의 last_col에 도달하면 행 종료
        5. 다음 first_col을 만나면 종료 (다른 행으로 넘어감)

        Args:
            boundary_result: 경계 분석 결과 (없으면 새로 계산)

        Returns:
            TableWidthResult: 테이블 전체 셀 너비 결과
        """
        if boundary_result is None:
            boundary_result = self.check_boundary_table()

        result = TableWidthResult()

        # first_cols와 last_cols 정렬
        first_cols = self._sort_first_cols_by_position(boundary_result.first_cols)
        last_cols = self._sort_first_cols_by_position(boundary_result.last_cols)
        first_cols_set = set(first_cols)
        last_cols_set = set(last_cols)

        self._log(f"[너비 계산] first_cols: {first_cols}")
        self._log(f"[너비 계산] last_cols: {last_cols}")

        cumulative_y = 0  # y 좌표 누적

        for row_idx, start_cell in enumerate(first_cols):
            row_result = RowWidthResult(
                row_index=row_idx,
                start_cell=start_cell,
                start_y=cumulative_y
            )

            # 이번 행에서 이미 방문한 셀 추적
            visited_in_row = set()

            self._log(f"\n[너비 계산] 행 {row_idx} 시작: list_id={start_cell}")

            # 시작 셀 너비/높이 수집
            width, height = self._get_cell_size(start_cell)
            cumulative_x = width
            visited_in_row.add(start_cell)

            # 행 높이는 첫 번째 셀 기준
            row_result.row_height = height

            cell_info = CellWidthInfo(
                list_id=start_cell,
                width=width,
                height=height,
                start_x=0,
                end_x=cumulative_x,
                col_index=0
            )
            row_result.cells.append(cell_info)

            self._log(f"  col=0: list_id={start_cell}, width={width}, height={height}, x=[0~{cumulative_x}]")

            current_id = start_cell
            col_index = 0

            max_steps = 1000
            step = 0

            while step < max_steps:
                step += 1

                # last_col에 도달했으면 종료
                if current_id in last_cols_set:
                    row_result.end_cell = current_id
                    self._log(f"  행 {row_idx} 종료: last_col 도달 (list_id={current_id})")
                    break

                # 오른쪽으로 이동
                self.hwp.SetPos(current_id, 0, 0)
                self.hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
                next_id, _, _ = self.hwp.GetPos()

                # 같은 셀에 머무르면 종료
                if next_id == current_id:
                    row_result.end_cell = current_id
                    self._log(f"  행 {row_idx} 종료: 이동 불가 (list_id={current_id})")
                    break

                # 다른 행의 first_col을 만나면 종료 (시작 셀 제외)
                if next_id in first_cols_set and next_id != start_cell:
                    row_result.end_cell = current_id
                    self._log(f"  행 {row_idx} 종료: 다른 행 first_col 도달 (list_id={next_id})")
                    break

                # 이미 방문한 셀이면 종료 (순환 감지)
                if next_id in visited_in_row:
                    row_result.end_cell = current_id
                    self._log(f"  행 {row_idx} 종료: 순환 감지 (list_id={next_id})")
                    break

                # 새 셀 너비 수집
                visited_in_row.add(next_id)
                col_index += 1
                start_x = cumulative_x
                width, height = self._get_cell_size(next_id)
                cumulative_x += width

                cell_info = CellWidthInfo(
                    list_id=next_id,
                    width=width,
                    height=height,
                    start_x=start_x,
                    end_x=cumulative_x,
                    col_index=col_index
                )
                row_result.cells.append(cell_info)

                self._log(f"  col={col_index}: list_id={next_id}, width={width}, x=[{start_x}~{cumulative_x}]")

                current_id = next_id

            row_result.total_width = cumulative_x
            row_result.end_y = cumulative_y + row_result.row_height
            cumulative_y = row_result.end_y  # 다음 행의 시작 y
            result.rows.append(row_result)

            if cumulative_x > result.max_width:
                result.max_width = cumulative_x

        return result

    def print_row_widths(self, width_result: TableWidthResult = None):
        """행별 셀 너비 결과 출력"""
        if width_result is None:
            width_result = self.calculate_row_widths()

        print("\n=== 행별 셀 너비 결과 ===")
        print(f"최대 너비: {width_result.max_width} HWPUNIT")

        for row in width_result.rows:
            print(f"\n[행 {row.row_index}] 시작={row.start_cell} → 끝={row.end_cell}, 전체 너비={row.total_width}")
            for cell in row.cells:
                print(f"  col={cell.col_index}: list_id={cell.list_id}, "
                      f"width={cell.width}, cumulative={cell.cumulative_width}")



if __name__ == "__main__":
    from cursor import get_hwp_instance

    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        exit(1)

    boundary = TableBoundary(hwp, debug=True)

    # 테이블 내부인지 확인
    if not boundary._is_in_table():
        print("[오류] 커서가 테이블 내부에 있지 않습니다.")
        exit(1)

    # 경계 분석
    result = boundary.check_boundary_table()
    boundary.print_boundary_info(result)

    # 서브 테이블 분석
    subsets = boundary.sub_table(result)
    boundary.print_sub_table(subsets)

    # 크기 기반 서브셋 분할
    size_results = boundary.sub_table_by_size(result)
    boundary.print_sub_table_by_size(size_results)

    # 행별 셀 너비 계산
    width_result = boundary.calculate_row_widths(result)
    boundary.print_row_widths(width_result)
