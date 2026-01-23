# -*- coding: utf-8 -*-
"""
테이블 경계 판별 모듈

테이블의 경계(첫 번째 행, 마지막 행, 첫 번째 열, 마지막 열)를 판별하는 함수 제공
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from dataclasses import dataclass, field
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
    table_end: int = 0          # right_border_cells의 마지막 셀의 list_id
    table_cell_counts: int = 0  # 총 셀 개수

    # 4방향 경계 셀 리스트
    top_border_cells: List[int] = field(default_factory=list)   # 첫 번째 행 셀들
    bottom_border_cells: List[int] = field(default_factory=list)  # 마지막 행 셀들
    left_border_cells: List[int] = field(default_factory=list)   # 첫 번째 열 셀들
    right_border_cells: List[int] = field(default_factory=list)    # 마지막 열 셀들

    # 좌표 경계 (HWPUNIT)
    start_x: int = 0            # 테이블 시작 x (항상 0)
    start_y: int = 0            # 테이블 시작 y (항상 0)
    end_x: int = 0              # 테이블 끝 x (xend)
    end_y: int = 0              # 테이블 끝 y (yend)


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

    def _calc_xend_from_top_border_cells(self, top_border_cells: List[int]) -> int:
        """
        top_border_cells 셀들의 너비 합 = xend (테이블 전체 가로 크기)

        Args:
            top_border_cells: 첫 번째 행의 셀 list_id 리스트

        Returns:
            int: xend (HWPUNIT)
        """
        if not top_border_cells:
            return 0

        top_border_cells_set = set(top_border_cells)
        start_cell = top_border_cells[0]
        self.hwp.SetPos(start_cell, 0, 0)

        cumulative_x = 0
        current_id = start_cell

        while current_id in top_border_cells_set:
            w, _ = self._table_info.get_cell_dimensions()
            cumulative_x += w

            self.hwp.SetPos(current_id, 0, 0)
            self.hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
            next_id = self.hwp.GetPos()[0]

            if next_id == current_id:
                break

            current_id = next_id

        self._log(f"[xend] top_border_cells: {top_border_cells}, xend={cumulative_x}")
        return cumulative_x

    def _calc_yend_from_left_border_cells(self, left_border_cells: List[int]) -> int:
        """
        left_border_cells 셀들의 높이 합 = yend (테이블 전체 세로 크기)

        Args:
            left_border_cells: 첫 번째 열의 셀 list_id 리스트

        Returns:
            int: yend (HWPUNIT)
        """
        if not left_border_cells:
            return 0

        # 중복 제거 (순서 유지)
        seen = set()
        unique_left_border_cells = []
        for cell_id in left_border_cells:
            if cell_id not in seen:
                seen.add(cell_id)
                unique_left_border_cells.append(cell_id)

        cumulative_y = 0

        for cell_id in unique_left_border_cells:
            self.hwp.SetPos(cell_id, 0, 0)
            _, h = self._table_info.get_cell_dimensions()
            cumulative_y += h

        self._log(f"[yend] left_border_cells (unique): {unique_left_border_cells}, yend={cumulative_y}")
        return cumulative_y

    def _find_lastcols_by_xend(self, start_cell: int, xend: int, tolerance: int = 50) -> dict:
        """
        start_cell부터 우측으로 순회하면서 xend 기준으로 right_border_cells 찾기

        - xend 초과 시: 이전 셀 = last_col, 현재 셀 = first_col
        - xend 도달 시: 현재 셀 = last_col

        Args:
            start_cell: 시작 셀 list_id
            xend: 테이블 전체 가로 크기 (HWPUNIT)
            tolerance: 허용 오차 (HWPUNIT)

        Returns:
            dict: {'left_border_cells': [...], 'right_border_cells': [...]}
        """
        self.hwp.SetPos(start_cell, 0, 0)
        current_id = start_cell
        cumulative_x = 0
        right_border_cells = []
        left_border_cells = [start_cell]
        prev_id = None

        max_iterations = 1000

        for _ in range(max_iterations):
            self.hwp.SetPos(current_id, 0, 0)
            w, _ = self._table_info.get_cell_dimensions()
            cumulative_x += w

            # xend 초과 → 이전 셀이 last_col, 현재 셀은 새 행의 first_col
            if cumulative_x > xend + tolerance:
                if prev_id is not None:
                    right_border_cells.append(prev_id)

                left_border_cells.append(current_id)
                cumulative_x = w
                prev_id = None

            # xend 도달 (오차 범위 내)
            elif abs(cumulative_x - xend) <= tolerance:
                right_border_cells.append(current_id)

                # 우측 이동
                self.hwp.SetPos(current_id, 0, 0)
                self.hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
                next_id = self.hwp.GetPos()[0]

                if next_id == current_id:
                    break

                left_border_cells.append(next_id)
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
                right_border_cells.append(current_id)
                break

            current_id = next_id

        self._log(f"[xend 순회] left_border_cells: {left_border_cells}")
        self._log(f"[xend 순회] right_border_cells: {right_border_cells}")

        return {
            'left_border_cells': left_border_cells,
            'right_border_cells': right_border_cells,
        }

    def _sort_left_border_cells_by_position(self, left_border_cells: List[int]) -> List[int]:
        """left_border_cells를 y좌표(위치) 기준으로 정렬"""
        # 각 셀의 위치 정보 수집
        positions = []
        for list_id in left_border_cells:
            self.hwp.SetPos(list_id, 0, 0)
            # KeyIndicator로 페이지/줄 정보 획득
            key = self.hwp.KeyIndicator()
            page = key[3] if len(key) > 3 else 0
            line = key[4] if len(key) > 4 else 0
            positions.append((list_id, page, line))

        # 페이지, 줄 순서로 정렬
        positions.sort(key=lambda x: (x[1], x[2]))
        return [p[0] for p in positions]

    def check_boundary_table(self) -> TableBoundaryResult:
        """
        테이블의 모든 셀을 순회하면서 경계 정보 계산

        1. table_origin 계산 - 테이블 첫 번째 셀의 list_id
        2. top_border_cells, bottom_border_cells 계산 - 첫/마지막 행에 속한 셀들
        3. left_border_cells, right_border_cells 계산 - 첫/마지막 열에 속한 셀들
        4. table_end 계산 - bottom_border_cells의 마지막 list_id

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

        # 2. top_border_cells, bottom_border_cells 계산
        for list_id in all_list_ids:
            if self.check_first_row_cell(list_id):
                result.top_border_cells.append(list_id)
            if self.check_bottom_row_cell(list_id):
                result.bottom_border_cells.append(list_id)

        result.top_border_cells.sort()
        result.bottom_border_cells.sort()
        self._log(f"top_border_cells: {result.top_border_cells}")
        self._log(f"bottom_border_cells: {result.bottom_border_cells}")

        # 3. xend 계산 및 right_border_cells/left_border_cells 계산 (xend 기반 우측 순회)
        xend = self._calc_xend_from_top_border_cells(result.top_border_cells)

        if xend > 0 and result.top_border_cells:
            xend_result = self._find_lastcols_by_xend(result.top_border_cells[0], xend)
            result.left_border_cells = xend_result['left_border_cells']
            result.right_border_cells = xend_result['right_border_cells']
        else:
            # fallback: table_origin에서 아래로 내려가면서 수집
            self.hwp.SetPos(result.table_origin, 0, 0)
            while True:
                current_id, _, _ = self.hwp.GetPos()
                result.left_border_cells.append(current_id)
                before_id = current_id
                self.hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
                after_id, _, _ = self.hwp.GetPos()
                if after_id == before_id:
                    break

        # table_end = right_border_cells의 마지막 list_id
        if result.right_border_cells:
            result.table_end = result.right_border_cells[-1]

        # 4. 좌표 경계 계산
        result.start_x = 0
        result.start_y = 0
        result.end_x = xend
        result.end_y = self._calc_yend_from_left_border_cells(result.left_border_cells)

        self._log(f"left_border_cells: {result.left_border_cells}")
        self._log(f"right_border_cells: {result.right_border_cells}")
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
        print(f"top_border_cells: {result.top_border_cells}")
        print(f"bottom_border_cells: {result.bottom_border_cells}")
        print(f"left_border_cells: {result.left_border_cells}")
        print(f"right_border_cells: {result.right_border_cells}")
        print(f"좌표 경계: ({result.start_x}, {result.start_y}) ~ ({result.end_x}, {result.end_y})")

    # ============================================================
    # 삭제예정_의존성: 아래 메서드들은 외부 문서(docs/)에서 참조됨
    # ============================================================

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

            top_border_cells = []
            for list_id in all_list_ids:
                if self.check_first_row_cell(list_id):
                    top_border_cells.append(list_id)
            top_border_cells.sort()

            if not top_border_cells:
                return {'grid': {}, 'rowspan_positions': {}, 'col_positions': [],
                        'max_row': 0, 'max_col': 0, 'xend': 0}

            if start_cell is None:
                start_cell = top_border_cells[0]
            if xend is None:
                xend = self._calc_xend_from_top_border_cells(top_border_cells)

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
