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
    table_origin: int = 0       # 테이블 첫 번째 셀의 list_id
    table_end: int = 0          # last_cols의 마지막 셀의 list_id
    table_cell_counts: int = 0  # 총 셀 개수
    first_rows: List[int] = field(default_factory=list)   # 첫 번째 행 셀들
    bottom_rows: List[int] = field(default_factory=list)  # 마지막 행 셀들
    first_cols: List[int] = field(default_factory=list)   # 첫 번째 열 셀들
    last_cols: List[int] = field(default_factory=list)    # 마지막 열 셀들


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
class CellCoordinate:
    """셀 좌표 정보"""
    list_id: int
    row: int = 0        # 행 번호 (0부터 시작)
    col: int = 0        # 열 번호 (0부터 시작)
    visit_count: int = 0  # 해당 셀 방문 횟수


@dataclass
class CellCoordinateResult:
    """셀 좌표 매핑 결과"""
    cells: Dict[int, CellCoordinate] = field(default_factory=dict)  # list_id -> CellCoordinate
    max_row: int = 0    # 최대 행 번호
    max_col: int = 0    # 최대 열 번호
    traversal_order: List[int] = field(default_factory=list)  # 순회 순서


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

        # 3. first_cols 계산: table_origin에서 아래로 내려가면서 수집
        self.hwp.SetPos(result.table_origin, 0, 0)
        while True:
            current_id, _, _ = self.hwp.GetPos()
            result.first_cols.append(current_id)
            before_id = current_id
            self.hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
            after_id, _, _ = self.hwp.GetPos()
            if after_id == before_id:
                break

        # 4. last_cols 계산: first_rows 마지막에서 move_up_right_down 활용
        first_row_last = result.first_rows[-1] if result.first_rows else result.table_origin
        result.last_cols.append(first_row_last)  # 첫 행의 마지막 열은 last_col

        current_id = first_row_last
        while True:
            down_id, is_last_col = self.move_up_right_down(current_id)
            # 더 이상 아래로 못 가거나 테이블 밖으로 나감 → 마지막 행
            if down_id == current_id or down_id == 0:
                break
            if is_last_col:
                result.last_cols.append(down_id)
            else:
                # 다른 열로 갔지만 아래로 더 못 가면 마지막 행의 마지막 열
                self.hwp.SetPos(down_id, 0, 0)
                self.hwp.HAction.Run("MoveDown")
                next_id, _, _ = self.hwp.GetPos()
                if next_id == down_id:  # 마지막 행
                    result.last_cols.append(down_id)
            current_id = down_id

        result.first_cols.sort()
        result.last_cols.sort()

        # table_end = last_cols의 마지막 list_id
        if result.last_cols:
            result.table_end = result.last_cols[-1]

        self._log(f"first_cols: {result.first_cols}")
        self._log(f"last_cols: {result.last_cols}")
        self._log(f"table_end: {result.table_end}")

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

    def map_cell_coordinates(self, boundary_result: TableBoundaryResult = None) -> CellCoordinateResult:
        """
        first_col 기반으로 셀 좌표 매핑 (방문 횟수로 행 결정)

        알고리즘:
        1. first_cols의 첫 번째 원소부터 시작
        2. 오른쪽 키보드 이동으로 순회하면서 동일한 셀 방문 횟수 저장
        3. 다음 first_col 원소를 만나면:
           - 이전 서브셀을 처음 순회한 경우 → 첫 번째 행 (row=0)
           - 이전에 동일한 서브셀을 1번 순회한 경우 → 두 번째 행 (row=1)
           - 이전에 2번 순회한 경우 → 세 번째 행 (row=2)
        4. 열 번호는 가장 가까운 first_col로부터 지나온 수
        5. 이전에 방문한 셀은 최초 1회만 좌표 설정, 이후는 건너뜀

        Args:
            boundary_result: 경계 분석 결과 (없으면 새로 계산)

        Returns:
            CellCoordinateResult: 셀 좌표 매핑 결과
        """
        if boundary_result is None:
            boundary_result = self.check_boundary_table()

        result = CellCoordinateResult()
        first_cols = self._sort_first_cols_by_position(boundary_result.first_cols)
        first_cols_set = set(first_cols)

        self._log(f"first_cols (정렬): {first_cols}")

        # 방문 횟수 추적: list_id -> 방문 횟수
        visit_count: Dict[int, int] = {}

        # 좌표가 설정된 셀 추적: list_id -> True
        coord_assigned: Dict[int, bool] = {}

        current_row = 0  # 현재 행 번호
        current_col = 0  # 현재 열 번호

        # first_cols의 첫 번째 원소에서 시작
        if not first_cols:
            self._log("first_cols가 비어있습니다")
            return result

        start_cell = first_cols[0]
        self.hwp.SetPos(start_cell, 0, 0)
        current_id = start_cell

        self._log(f"시작 셀: {start_cell}")

        # 전체 순회 (first_cols의 마지막까지)
        max_iterations = boundary_result.table_cell_counts * 10  # 안전 장치
        iteration = 0

        while iteration < max_iterations:
            iteration += 1

            # 현재 셀 방문 횟수 증가
            visit_count[current_id] = visit_count.get(current_id, 0) + 1
            result.traversal_order.append(current_id)

            self._log(f"방문: list_id={current_id}, visit_count={visit_count[current_id]}, row={current_row}, col={current_col}")

            # 좌표가 아직 설정되지 않은 경우에만 설정
            if current_id not in coord_assigned:
                coord = CellCoordinate(
                    list_id=current_id,
                    row=current_row,
                    col=current_col,
                    visit_count=visit_count[current_id]
                )
                result.cells[current_id] = coord
                coord_assigned[current_id] = True

                # 최대 행/열 업데이트
                result.max_row = max(result.max_row, current_row)
                result.max_col = max(result.max_col, current_col)

                self._log(f"  좌표 설정: ({current_row}, {current_col})")

            # 오른쪽으로 이동
            self.hwp.SetPos(current_id, 0, 0)
            self.hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
            next_id, _, _ = self.hwp.GetPos()

            # 같은 셀에 머무르면 종료
            if next_id == current_id:
                self._log(f"종료: 더 이상 이동 불가 (list_id={current_id})")
                break

            # 다음 셀이 first_col 원소인지 확인
            if next_id in first_cols_set:
                # 이전 셀(current_id)의 방문 횟수로 행 번호 결정
                prev_visit = visit_count.get(current_id, 1)

                # 이전 서브셀을 처음 순회 → 첫 번째 행 (row=0)
                # 이전에 1번 순회 → 두 번째 행 (row=1)
                # 이전에 2번 순회 → 세 번째 행 (row=2)
                current_row = prev_visit - 1 + 1  # 다음 행으로
                # 실제로는: 이전 방문이 1번이면 다음은 row=1

                # first_col을 만나면 열 번호 리셋
                current_col = 0

                self._log(f"first_col 도달: {next_id}, 이전 방문횟수={prev_visit}, 새 행={current_row}")
            else:
                # 일반 이동: 열 번호 증가
                current_col += 1

            current_id = next_id

        return result

    def map_cell_coordinates_v2(self, boundary_result: TableBoundaryResult = None) -> CellCoordinateResult:
        """
        first_col 기반으로 셀 좌표 매핑 (개선 버전)

        알고리즘:
        1. first_cols 순회하면서 각 행의 시작점 결정
        2. 각 first_col에서 오른쪽으로 순회
        3. 방문 횟수가 n인 셀을 만나면 → 그 셀이 속한 행은 n번째 행
        4. first_col 간 순회에서 이미 좌표가 설정된 셀은 건너뜀
        5. 열 번호는 현재 first_col로부터의 거리

        Args:
            boundary_result: 경계 분석 결과 (없으면 새로 계산)

        Returns:
            CellCoordinateResult: 셀 좌표 매핑 결과
        """
        if boundary_result is None:
            boundary_result = self.check_boundary_table()

        result = CellCoordinateResult()
        first_cols = self._sort_first_cols_by_position(boundary_result.first_cols)
        first_cols_set = set(first_cols)

        self._log(f"first_cols (정렬): {first_cols}")

        if not first_cols:
            self._log("first_cols가 비어있습니다")
            return result

        # 방문 횟수 추적: list_id -> 방문 횟수
        visit_count: Dict[int, int] = {}

        # 좌표가 설정된 셀 추적
        coord_assigned: Dict[int, bool] = {}

        # 전체 순회: first_cols[0]에서 시작해서 다음 first_col까지 반복
        current_first_col_idx = 0
        current_row = 0

        while current_first_col_idx < len(first_cols):
            start_cell = first_cols[current_first_col_idx]
            self.hwp.SetPos(start_cell, 0, 0)

            current_id = start_cell
            current_col = 0
            row_for_this_pass = current_row

            self._log(f"\n=== first_col[{current_first_col_idx}] = {start_cell}, row={row_for_this_pass} ===")

            # 다음 first_col 또는 테이블 끝까지 순회
            max_steps = boundary_result.table_cell_counts * 2
            steps = 0

            while steps < max_steps:
                steps += 1

                # 현재 셀 방문 횟수 증가
                visit_count[current_id] = visit_count.get(current_id, 0) + 1
                curr_visit = visit_count[current_id]

                result.traversal_order.append(current_id)

                self._log(f"  방문: list_id={current_id}, visit={curr_visit}, col={current_col}")

                # 좌표가 아직 설정되지 않은 경우에만 설정
                if current_id not in coord_assigned:
                    # 행 번호 결정: 이 셀을 처음 방문(visit=1)이면 row_for_this_pass
                    # 두 번째 방문(visit=2)이면 이미 이전 행에서 설정됨
                    coord = CellCoordinate(
                        list_id=current_id,
                        row=row_for_this_pass,
                        col=current_col,
                        visit_count=curr_visit
                    )
                    result.cells[current_id] = coord
                    coord_assigned[current_id] = True

                    result.max_row = max(result.max_row, row_for_this_pass)
                    result.max_col = max(result.max_col, current_col)

                    self._log(f"    → 좌표 설정: ({row_for_this_pass}, {current_col})")

                # 오른쪽으로 이동
                self.hwp.SetPos(current_id, 0, 0)
                self.hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
                next_id, _, _ = self.hwp.GetPos()

                # 같은 셀에 머무르면 이 패스 종료
                if next_id == current_id:
                    self._log(f"  패스 종료: 더 이상 이동 불가")
                    break

                # 다음 first_col을 만나면 이 패스 종료
                if next_id in first_cols_set and next_id != start_cell:
                    # 다음 first_col의 인덱스 찾기
                    try:
                        next_fc_idx = first_cols.index(next_id)
                        # 다음 패스에서 처리할 first_col 업데이트
                        if next_fc_idx > current_first_col_idx:
                            self._log(f"  다음 first_col 도달: {next_id} (idx={next_fc_idx})")
                            break
                    except ValueError:
                        pass

                # 열 번호 증가
                current_col += 1
                current_id = next_id

            # 다음 first_col로 이동
            current_first_col_idx += 1
            current_row += 1

        return result

    def map_cell_coordinates_v3(self, boundary_result: TableBoundaryResult = None) -> CellCoordinateResult:
        """
        서브셀 기반 셀 좌표 매핑

        알고리즘:
        1. first_col[0]에서 시작하여 오른쪽으로 순회
        2. 새 셀(처음 방문)을 만나면 → 현재 행, 열+1
        3. 이미 방문한 셀(2회 이상)을 만나면 → 행+1 (한 번만), 열은 왼쪽 열+1 유지
        4. first_col을 만나면 → 열=0
        5. 좌표는 최초 방문 시에만 설정

        예시 (셀 2~10):
        - 2,3,4,5,6,7,8: 처음 방문 → (0,0)~(0,6)
        - 9: 처음 방문 → (0,7)
        - 10: 9 재방문 후 새 셀 → 행+1, 열=7+1=8 → (1,8)

        Args:
            boundary_result: 경계 분석 결과

        Returns:
            CellCoordinateResult: 셀 좌표 매핑 결과
        """
        if boundary_result is None:
            boundary_result = self.check_boundary_table()

        result = CellCoordinateResult()
        first_cols = self._sort_first_cols_by_position(boundary_result.first_cols)
        first_cols_set = set(first_cols)

        self._log(f"[v3] first_cols (정렬): {first_cols}")

        if not first_cols:
            self._log("[v3] first_cols가 비어있습니다")
            return result

        # 방문 횟수 추적: list_id -> 방문 횟수
        visit_count: Dict[int, int] = {}

        # 좌표가 설정된 셀 추적
        coord_assigned: Dict[int, bool] = {}

        # 시작
        start_cell = first_cols[0]
        self.hwp.SetPos(start_cell, 0, 0)
        current_id = start_cell

        current_row = 0
        current_col = 0
        row_changed = False  # 이번 재방문 구간에서 행 전환 여부

        max_iterations = boundary_result.table_cell_counts * 20
        iteration = 0

        self._log(f"[v3] 시작 셀: {start_cell}")

        # 재방문 구간에서의 열 카운터 (재방문 구간 끝나면 이 값+1이 새 셀의 열)
        revisit_col_count = 0

        while iteration < max_iterations:
            iteration += 1

            # 현재 셀 방문 횟수 증가
            prev_visit = visit_count.get(current_id, 0)
            visit_count[current_id] = prev_visit + 1
            curr_visit = visit_count[current_id]

            result.traversal_order.append(current_id)

            self._log(f"[v3] 방문: list_id={current_id}, visit={curr_visit}, row={current_row}, col={current_col}")

            # 2회 이상 방문한 셀이면 (재방문)
            if curr_visit >= 2:
                # 행 전환은 재방문 구간 시작 시 한 번만
                if not row_changed:
                    current_row += 1
                    row_changed = True
                    revisit_col_count = 0  # 재방문 구간 열 카운터 리셋
                    self._log(f"[v3]   → 행 전환: row={current_row} (재방문 시작)")
                else:
                    revisit_col_count += 1  # 재방문 구간 내 열 증가
            else:
                # 처음 방문한 셀
                if row_changed:
                    # 재방문 구간이 끝났다 → 재방문한 셀 수만큼 열 증가
                    current_col += revisit_col_count + 1
                    row_changed = False
                    self._log(f"[v3]   → 재방문 구간 종료, col={current_col}")

                if current_id not in coord_assigned:
                    coord = CellCoordinate(
                        list_id=current_id,
                        row=current_row,
                        col=current_col,
                        visit_count=curr_visit
                    )
                    result.cells[current_id] = coord
                    coord_assigned[current_id] = True

                    result.max_row = max(result.max_row, current_row)
                    result.max_col = max(result.max_col, current_col)

                    self._log(f"[v3]   → 좌표 설정: ({current_row}, {current_col})")

            # 오른쪽으로 이동
            self.hwp.SetPos(current_id, 0, 0)
            self.hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
            next_id, _, _ = self.hwp.GetPos()

            # 같은 셀에 머무르면 종료
            if next_id == current_id:
                self._log(f"[v3] 종료: 더 이상 이동 불가")
                break

            # 다음 셀이 first_col이면
            if next_id in first_cols_set:
                next_visit = visit_count.get(next_id, 0)
                if next_visit == 0:
                    # 처음 방문하는 새 first_col → 새 행 시작, 열=0
                    current_row += 1
                    current_col = 0
                    row_changed = False
                    self._log(f"[v3]   새 first_col 도달: {next_id}, row={current_row}, col=0")
                else:
                    # 이미 방문한 first_col 재방문
                    # 재방문 구간 중이면 그냥 열=0으로만 리셋 (행 증가 없음)
                    # 재방문 구간이 아니면 재방문 구간 시작
                    if not row_changed:
                        # 새로운 재방문 구간 시작 → 행 증가
                        current_row += 1
                        row_changed = True
                        revisit_col_count = 0
                        self._log(f"[v3]   기존 first_col 재방문 (새 구간): {next_id}, row={current_row}, col=0")
                    else:
                        # 이미 재방문 구간 중 → 행 증가 없이 열만 리셋
                        self._log(f"[v3]   기존 first_col 재방문 (구간 내): {next_id}, col=0 (행 유지)")
                    current_col = 0
            else:
                # 일반 셀로 이동
                next_visit = visit_count.get(next_id, 0)
                if next_visit == 0 and not row_changed:
                    # 처음 방문할 셀이고 재방문 구간이 아니면 → 열 증가
                    current_col += 1
                # 재방문 구간 중이거나 재방문할 셀이면 열 유지

            current_id = next_id

        return result

    def print_cell_coordinates(self, result: CellCoordinateResult = None):
        """셀 좌표 매핑 결과 출력"""
        if result is None:
            result = self.map_cell_coordinates_v3()

        print("\n=== 셀 좌표 매핑 결과 ===")
        print(f"최대 행: {result.max_row}, 최대 열: {result.max_col}")
        print(f"총 셀 수: {len(result.cells)}")
        print("\n좌표 매핑:")

        # 행 기준 정렬하여 출력
        sorted_cells = sorted(result.cells.values(), key=lambda c: (c.row, c.col))
        current_row = -1
        for cell in sorted_cells:
            if cell.row != current_row:
                current_row = cell.row
                print(f"\n[행 {current_row}]")
            print(f"  ({cell.row}, {cell.col}): list_id={cell.list_id}, visit={cell.visit_count}")


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

    # 셀 좌표 매핑 (v3: 서브셀 기반)
    coord_result = boundary.map_cell_coordinates_v3(result)
    boundary.print_cell_coordinates(coord_result)
