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
        """끝 좌표에 해당하는 레벨 인덱스 반환"""
        for idx, level in enumerate(levels):
            if abs(value - level) <= self.TOLERANCE:
                return idx - 1 if idx > 0 else 0
        # value보다 작은 가장 큰 레벨
        for idx in range(len(levels) - 1, -1, -1):
            if levels[idx] < value:
                return idx
        return len(levels) - 1

    def calculate(self, max_cells: int = 1000) -> CellPositionResult:
        """테이블의 모든 셀 위치 및 범위 계산"""
        from table.table_info import MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL

        boundary = self._get_boundary()
        table_info = self._get_table_info()

        if not boundary._is_in_table():
            raise ValueError("커서가 테이블 내부에 있지 않습니다.")

        # 경계 분석
        result = boundary.check_boundary_table()
        first_cols = boundary._sort_first_cols_by_position(result.first_cols)
        last_cols_set = set(result.last_cols)
        first_cols_set = set(first_cols)

        # X, Y 레벨 수집
        x_levels = {0}
        y_levels = {0}
        cell_positions = {}

        def collect_split_cells(cell_id, start_x, start_y, row_end_y, visited):
            """분할된 셀 재귀 순회"""
            self.hwp.SetPos(cell_id, 0, 0)
            width, height = table_info.get_cell_dimensions()
            end_x = start_x + width
            end_y = start_y + height

            x_levels.add(start_x)
            y_levels.add(start_y)

            if cell_id not in cell_positions:
                cell_positions[cell_id] = {
                    'start_x': start_x, 'end_x': end_x,
                    'start_y': start_y, 'end_y': end_y,
                }

            # 오른쪽 순회
            self.hwp.SetPos(cell_id, 0, 0)
            self.hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
            right_id, _, _ = self.hwp.GetPos()
            if right_id != cell_id and right_id not in visited and right_id not in first_cols_set:
                visited.add(right_id)
                collect_split_cells(right_id, end_x, start_y, row_end_y, visited)

            # 아래 순회
            if end_y < row_end_y:
                self.hwp.SetPos(cell_id, 0, 0)
                self.hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
                down_id, _, _ = self.hwp.GetPos()
                if down_id != cell_id and down_id not in visited and down_id not in first_cols_set:
                    visited.add(down_id)
                    collect_split_cells(down_id, start_x, end_y, row_end_y, visited)

        # 모든 셀 순회
        cumulative_y = 0
        total_cells = 0
        for row_start in first_cols:
            if total_cells >= max_cells:
                if self.debug:
                    print(f"[경고] 최대 셀 수({max_cells}) 도달, 중단")
                break

            self.hwp.SetPos(row_start, 0, 0)
            _, row_height = table_info.get_cell_dimensions()

            row_start_y = cumulative_y
            row_end_y = cumulative_y + row_height
            y_levels.add(row_end_y)

            cumulative_x = 0
            current_id = row_start
            visited = set()

            while True:
                if total_cells >= max_cells:
                    break
                if current_id in visited:
                    break
                visited.add(current_id)
                total_cells += 1

                self.hwp.SetPos(current_id, 0, 0)
                width, height = table_info.get_cell_dimensions()

                start_x = cumulative_x
                end_x = cumulative_x + width
                start_y = row_start_y
                end_y = row_start_y + height

                x_levels.add(start_x)
                y_levels.add(start_y)

                cell_positions[current_id] = {
                    'start_x': start_x, 'end_x': end_x,
                    'start_y': start_y, 'end_y': end_y,
                }

                # 분할 셀 처리
                if height < row_height:
                    self.hwp.SetPos(current_id, 0, 0)
                    self.hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
                    next_sub, _, _ = self.hwp.GetPos()
                    if next_sub != current_id and next_sub not in visited and next_sub not in first_cols_set:
                        visited.add(next_sub)
                        collect_split_cells(next_sub, start_x, end_y, row_end_y, visited)

                cumulative_x = end_x

                if current_id in last_cols_set:
                    break

                self.hwp.SetPos(current_id, 0, 0)
                self.hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
                next_id, _, _ = self.hwp.GetPos()

                if next_id == current_id:
                    break
                if next_id in first_cols_set and next_id != row_start:
                    break

                current_id = next_id

            cumulative_y = row_end_y

        # 레벨 병합
        x_levels_list = self._merge_close_levels(list(x_levels))
        y_levels_list = self._merge_close_levels(list(y_levels))

        # 셀 범위 계산
        cells = {}
        for list_id, pos in cell_positions.items():
            start_row = self._find_level_index(pos['start_y'], y_levels_list)
            start_col = self._find_level_index(pos['start_x'], x_levels_list)
            end_row = self._find_end_level_index(pos['end_y'], y_levels_list)
            end_col = self._find_end_level_index(pos['end_x'], x_levels_list)

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

        max_row = max(c.end_row for c in cells.values()) if cells else 0
        max_col = max(c.end_col for c in cells.values()) if cells else 0

        return CellPositionResult(
            cells=cells,
            x_levels=x_levels_list,
            y_levels=y_levels_list,
            max_row=max_row,
            max_col=max_col,
        )

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


# 직접 실행 시
if __name__ == "__main__":
    import sys
    import os
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

    from cursor import get_hwp_instance

    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        exit(1)

    calc = CellPositionCalculator(hwp)

    try:
        result = calc.calculate()
        calc.print_summary(result)

        print(f"\n=== 모든 셀 ===")
        for list_id, cell in sorted(result.cells.items()):
            print(f"  list_id={list_id}: {cell}")

    except ValueError as e:
        print(f"[오류] {e}")
