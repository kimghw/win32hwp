# -*- coding: utf-8 -*-
"""셀 위치 및 범위 계산 모듈"""

from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple


@dataclass
class CellRange:
    """셀의 위치 및 범위 정보"""
    list_id: int
    start_row: int
    start_col: int
    end_row: int
    end_col: int
    rowspan: int
    colspan: int
    # 물리적 좌표 (HWPUNIT)
    start_x: int = 0
    start_y: int = 0
    end_x: int = 0
    end_y: int = 0

    def is_merged(self) -> bool:
        """병합 셀인지 확인"""
        return self.rowspan > 1 or self.colspan > 1

    def contains(self, row: int, col: int) -> bool:
        """해당 좌표가 이 셀 범위에 포함되는지 확인"""
        return (self.start_row <= row <= self.end_row and
                self.start_col <= col <= self.end_col)

    def __str__(self):
        if self.is_merged():
            return f"({self.start_row},{self.start_col})~({self.end_row},{self.end_col}) [rowspan={self.rowspan}, colspan={self.colspan}]"
        return f"({self.start_row},{self.start_col})"


@dataclass
class CellPositionResult:
    """셀 위치 계산 결과"""
    cells: Dict[int, CellRange]  # list_id -> CellRange
    x_levels: List[int]  # X 레벨 목록
    y_levels: List[int]  # Y 레벨 목록
    max_row: int
    max_col: int


class CellPositionCalculator:
    """셀 위치 및 범위 계산기"""

    TOLERANCE = 3  # 레벨 병합 허용 오차

    def __init__(self, hwp, debug: bool = False):
        self.hwp = hwp
        self.debug = debug
        self._table_info = None
        self._boundary = None

    def _get_table_info(self):
        if self._table_info is None:
            from table.table_info import TableInfo
            self._table_info = TableInfo(self.hwp, debug=self.debug)
        return self._table_info

    def _get_boundary(self):
        if self._boundary is None:
            from table.table_boundary import TableBoundary
            self._boundary = TableBoundary(self.hwp, debug=self.debug)
        return self._boundary

    def _merge_close_levels(self, levels: List[int]) -> List[int]:
        """근접한 레벨들을 병합"""
        if not levels:
            return []
        sorted_levels = sorted(levels)
        merged = [sorted_levels[0]]
        for level in sorted_levels[1:]:
            if level - merged[-1] > self.TOLERANCE:
                merged.append(level)
        return merged

    def _find_level_index(self, value: int, levels: List[int]) -> int:
        """값에 해당하는 레벨 인덱스 반환"""
        for idx, level in enumerate(levels):
            if abs(value - level) <= self.TOLERANCE:
                return idx
        return -1

    def _find_end_level_index(self, value: int, levels: List[int]) -> int:
        """끝 좌표에 해당하는 레벨 인덱스 반환

        셀의 끝 좌표가 다음 셀의 시작 좌표와 같은 경우가 많으므로,
        끝 좌표가 레벨과 일치하면 이전 레벨 인덱스를 반환
        """
        # 정확히 일치하는 레벨 찾기 (끝 좌표는 경계에 있음)
        for idx, level in enumerate(levels):
            if abs(value - level) <= self.TOLERANCE:
                # 끝 좌표가 레벨 경계와 일치하면, 이전 레벨에 속함
                return max(0, idx - 1)

        # 일치하는 레벨이 없으면, value보다 작은 가장 큰 레벨 찾기
        for idx in range(len(levels) - 1, -1, -1):
            if levels[idx] < value - self.TOLERANCE:
                return idx
        return 0

    def calculate_grid(self, max_cells: int = 1000) -> CellPositionResult:
        """
        xline/yline 기반 그리드 레벨 생성

        알고리즘:
        1. 모든 셀을 순회하며 물리적 좌표(start_x, start_y, end_x, end_y)만 수집
        2. 수집된 모든 x 좌표에서 unique한 값들을 xline으로 추출
        3. 수집된 모든 y 좌표에서 unique한 값들을 yline으로 추출
        4. xline, yline으로 그리드를 만들고 각 셀을 매핑

        이 방식은 BFS 순회 순서에 의존하지 않고,
        물리적 좌표만으로 정확한 그리드를 생성합니다.
        """
        from table.table_info import MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL
        from collections import deque

        table_info = self._get_table_info()
        boundary = self._get_boundary()

        if not table_info.is_in_table():
            raise ValueError("커서가 테이블 내부에 있지 않습니다.")

        # 경계 분석으로 left_border_cells, right_border_cells 획득
        boundary_result = boundary.check_boundary_table()
        left_border_cells = boundary._sort_left_border_cells_by_position(boundary_result.left_border_cells)
        right_border_cells_set = set(boundary_result.right_border_cells)
        left_border_cells_set = set(left_border_cells)

        if self.debug:
            print(f"[calculate_grid] left_border_cells: {left_border_cells}")
            print(f"[calculate_grid] right_border_cells: {list(right_border_cells_set)}")

        # 1단계: 모든 셀의 물리적 좌표 수집
        cell_positions = {}  # list_id -> {start_x, end_x, start_y, end_y}
        all_x_coords = set()  # 모든 x 좌표 (start_x, end_x)
        all_y_coords = set()  # 모든 y 좌표 (start_y, end_y)

        cumulative_y = 0
        total_cells = 0

        for row_idx, row_start in enumerate(left_border_cells):
            if total_cells >= max_cells:
                if self.debug:
                    print(f"[경고] 최대 셀 수({max_cells}) 도달, 중단")
                break

            # 행 시작 셀의 높이 = 행 높이
            self.hwp.SetPos(row_start, 0, 0)
            _, row_height = table_info.get_cell_dimensions()

            row_start_y = cumulative_y
            row_end_y = cumulative_y + row_height

            if self.debug:
                print(f"[calculate_grid] 행 {row_idx}: y=[{row_start_y}~{row_end_y}], height={row_height}")

            # 행 내 셀들을 BFS로 수집 (물리 좌표 계산)
            visited_in_row = set()
            queue = deque()

            # 행 시작 셀
            self.hwp.SetPos(row_start, 0, 0)
            first_width, first_height = table_info.get_cell_dimensions()

            queue.append((row_start, 0, row_start_y))  # (list_id, start_x, start_y)
            visited_in_row.add(row_start)

            while queue and total_cells < max_cells:
                current_id, start_x, start_y = queue.popleft()
                total_cells += 1

                # 현재 셀 크기
                self.hwp.SetPos(current_id, 0, 0)
                width, height = table_info.get_cell_dimensions()

                end_x = start_x + width
                end_y = start_y + height

                # 물리 좌표 저장
                cell_positions[current_id] = {
                    'start_x': start_x, 'end_x': end_x,
                    'start_y': start_y, 'end_y': end_y,
                }

                # 모든 좌표 수집
                all_x_coords.add(start_x)
                all_x_coords.add(end_x)
                all_y_coords.add(start_y)
                all_y_coords.add(end_y)

                if self.debug:
                    print(f"  셀 {current_id}: x=[{start_x}~{end_x}], y=[{start_y}~{end_y}]")

                # 오른쪽 이동
                self.hwp.SetPos(current_id, 0, 0)
                self.hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
                right_id = self.hwp.GetPos()[0]

                if (right_id != current_id and
                    right_id not in visited_in_row and
                    right_id not in left_border_cells_set):
                    visited_in_row.add(right_id)
                    queue.append((right_id, end_x, start_y))

                # 아래 이동 (행 내 분할 셀 처리)
                if end_y < row_end_y:
                    self.hwp.SetPos(current_id, 0, 0)
                    self.hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
                    down_id = self.hwp.GetPos()[0]

                    if (down_id != current_id and
                        down_id not in visited_in_row and
                        down_id not in left_border_cells_set):
                        visited_in_row.add(down_id)
                        queue.append((down_id, start_x, end_y))

            cumulative_y = row_end_y

        # 1.5단계: 테이블 경계 좌표 계산 및 필터링
        # X 경계: left_border_cells 셀의 좌측(0) ~ right_border_cells 셀의 우측
        # Y 경계: top_border_cells의 상단(0) ~ bottom_border_cells의 하단(cumulative_y)
        table_min_x = 0
        table_max_x = 0
        table_min_y = 0
        table_max_y = cumulative_y

        # right_border_cells 셀들의 우측 좌표 중 최대값 계산
        for last_col_id in right_border_cells_set:
            if last_col_id in cell_positions:
                table_max_x = max(table_max_x, cell_positions[last_col_id]['end_x'])

        if self.debug:
            print(f"[calculate_grid] 테이블 경계: x=[{table_min_x}~{table_max_x}], y=[{table_min_y}~{table_max_y}]")
            print(f"[calculate_grid] 필터링 전: x좌표 {len(all_x_coords)}개, y좌표 {len(all_y_coords)}개")

        # X 좌표 필터링: 테이블 경계 내 좌표만 유지
        all_x_coords = {
            x for x in all_x_coords
            if table_min_x - self.TOLERANCE <= x <= table_max_x + self.TOLERANCE
        }

        # Y 좌표 필터링: 테이블 경계 내 좌표만 유지
        all_y_coords = {
            y for y in all_y_coords
            if table_min_y - self.TOLERANCE <= y <= table_max_y + self.TOLERANCE
        }

        if self.debug:
            print(f"[calculate_grid] 필터링 후: x좌표 {len(all_x_coords)}개, y좌표 {len(all_y_coords)}개")

        # 2단계: xline, yline 생성 (중복 제거 및 정렬)
        x_levels_list = self._merge_close_levels(list(all_x_coords))
        y_levels_list = self._merge_close_levels(list(all_y_coords))

        if self.debug:
            print(f"[calculate_grid] x_levels ({len(x_levels_list)}개): {x_levels_list}")
            print(f"[calculate_grid] y_levels ({len(y_levels_list)}개): {y_levels_list}")

        # 3단계: 각 셀을 그리드에 매핑
        cells = {}
        for list_id, pos in cell_positions.items():
            start_row = self._find_level_index(pos['start_y'], y_levels_list)
            start_col = self._find_level_index(pos['start_x'], x_levels_list)
            end_row = self._find_end_level_index(pos['end_y'], y_levels_list)
            end_col = self._find_end_level_index(pos['end_x'], x_levels_list)

            # 유효성 검사
            if start_row < 0 or start_col < 0:
                if self.debug:
                    print(f"[경고] 셀 {list_id} 레벨 매핑 실패: "
                          f"start=({start_row},{start_col}), pos={pos}")
                continue

            cells[list_id] = CellRange(
                list_id=list_id,
                start_row=start_row,
                start_col=start_col,
                end_row=end_row,
                end_col=end_col,
                rowspan=end_row - start_row + 1,
                colspan=end_col - start_col + 1,
                start_x=pos['start_x'],
                start_y=pos['start_y'],
                end_x=pos['end_x'],
                end_y=pos['end_y'],
            )

            if self.debug:
                print(f"  셀 {list_id} → ({start_row},{start_col})~({end_row},{end_col})")

        max_row = max(c.end_row for c in cells.values()) if cells else 0
        max_col = max(c.end_col for c in cells.values()) if cells else 0

        if self.debug:
            print(f"[calculate_grid] 결과: {len(cells)}개 셀, {max_row+1}행 x {max_col+1}열")

        return CellPositionResult(
            cells=cells,
            x_levels=x_levels_list,
            y_levels=y_levels_list,
            max_row=max_row,
            max_col=max_col,
        )

    def calculate(self, max_cells: int = 1000) -> CellPositionResult:
        """
        테이블의 모든 셀 위치 및 범위 계산 (메인 메서드)

        내부적으로 calculate_grid()를 호출합니다.
        xline/yline 기반으로 정확한 그리드를 생성합니다.

        Args:
            max_cells: 최대 처리할 셀 수 (기본 1000)

        Returns:
            CellPositionResult: 셀 위치 계산 결과
        """
        return self.calculate_grid(max_cells)

    def get_cell_at(self, result: CellPositionResult, row: int, col: int) -> Optional[CellRange]:
        """특정 좌표에 있는 셀 반환"""
        for cell in result.cells.values():
            if cell.contains(row, col):
                return cell
        return None

    def get_merged_cells(self, result: CellPositionResult) -> List[CellRange]:
        """병합된 셀 목록 반환"""
        return [c for c in result.cells.values() if c.is_merged()]

    def get_merge_info(self, result: CellPositionResult, list_id: int) -> Optional[dict]:
        """특정 셀의 병합 정보 반환"""
        cell = result.cells.get(list_id)
        if not cell:
            return None
        return {
            'list_id': list_id,
            'start': (cell.start_row, cell.start_col),
            'end': (cell.end_row, cell.end_col),
            'rowspan': cell.rowspan,
            'colspan': cell.colspan,
            'is_merged': cell.is_merged(),
            'covered_coords': self.get_covered_coords(cell),
        }

    def get_covered_coords(self, cell: CellRange) -> List[Tuple[int, int]]:
        """셀이 차지하는 모든 좌표 목록 반환"""
        coords = []
        for r in range(cell.start_row, cell.end_row + 1):
            for c in range(cell.start_col, cell.end_col + 1):
                coords.append((r, c))
        return coords

    def get_cell_by_coord(self, result: CellPositionResult, row: int, col: int) -> Optional[int]:
        """좌표로 list_id 찾기 (병합 셀 포함)"""
        cell = self.get_cell_at(result, row, col)
        return cell.list_id if cell else None

    def get_representative_coord(self, result: CellPositionResult, row: int, col: int) -> Optional[Tuple[int, int]]:
        """해당 좌표의 대표 좌표 반환 (병합 셀이면 좌상단)"""
        cell = self.get_cell_at(result, row, col)
        if cell:
            return (cell.start_row, cell.start_col)
        return None

    def build_coord_to_listid_map(self, result: CellPositionResult) -> Dict[Tuple[int, int], int]:
        """(row, col) → list_id 매핑 테이블 생성 (병합 셀 포함)"""
        coord_map = {}
        for cell in result.cells.values():
            for coord in self.get_covered_coords(cell):
                coord_map[coord] = cell.list_id
        return coord_map

    def get_all_merge_info(self, result: CellPositionResult) -> List[dict]:
        """모든 병합 셀 정보 목록"""
        merged_cells = self.get_merged_cells(result)
        return [self.get_merge_info(result, c.list_id) for c in merged_cells]

    def insert_coordinate_text(self, list_id: int, row: int, col: int):
        """셀에 좌표 텍스트를 파란색으로 삽입"""
        self.hwp.SetPos(list_id, 0, 0)
        self.hwp.MovePos(5, 0, 0)  # moveBottomOfList

        coord_text = f"\r({row}, {col})"

        act = self.hwp.CreateAction("InsertText")
        pset = act.CreateSet()
        act.GetDefault(pset)
        pset.SetItem("Text", coord_text)
        act.Execute(pset)

        # 삽입한 텍스트 선택
        coord_len = len(f"({row}, {col})")
        for _ in range(coord_len):
            self.hwp.HAction.Run("MoveSelPrevChar")

        # 파란색 적용
        act = self.hwp.CreateAction("CharShape")
        pset = act.CreateSet()
        act.GetDefault(pset)
        pset.SetItem("TextColor", 0xFF0000)  # BGR: 파란색
        act.Execute(pset)

        self.hwp.HAction.Run("Cancel")

    def insert_all_coordinates(self, result: CellPositionResult, verbose: bool = True):
        """모든 셀에 좌표 텍스트 삽입"""
        if verbose:
            print(f"총 {len(result.cells)}개 셀에 좌표 삽입 중...")

        for list_id, cell in result.cells.items():
            self.insert_coordinate_text(list_id, cell.start_row, cell.start_col)
            if verbose:
                print(f"  list_id={list_id} → ({cell.start_row}, {cell.start_col})")

        if verbose:
            print("완료!")

    def print_summary(self, result: CellPositionResult):
        """결과 요약 출력"""
        print(f"\n=== 셀 위치 계산 결과 ===")
        print(f"X 레벨: {len(result.x_levels)}개")
        print(f"Y 레벨: {len(result.y_levels)}개")
        print(f"총 셀: {len(result.cells)}개")
        print(f"테이블 크기: {result.max_row + 1}행 x {result.max_col + 1}열")

        merged = self.get_merged_cells(result)
        if merged:
            print(f"\n=== 병합 셀 ({len(merged)}개) ===")
            for cell in sorted(merged, key=lambda c: (c.start_row, c.start_col)):
                print(f"  list_id={cell.list_id}: {cell}")
