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

        # 경계 분석으로 first_cols, last_cols 획득
        boundary_result = boundary.check_boundary_table()
        first_cols = boundary._sort_first_cols_by_position(boundary_result.first_cols)
        last_cols_set = set(boundary_result.last_cols)
        first_cols_set = set(first_cols)

        if self.debug:
            print(f"[calculate_grid] first_cols: {first_cols}")
            print(f"[calculate_grid] last_cols: {list(last_cols_set)}")

        # 1단계: 모든 셀의 물리적 좌표 수집
        cell_positions = {}  # list_id -> {start_x, end_x, start_y, end_y}
        all_x_coords = set()  # 모든 x 좌표 (start_x, end_x)
        all_y_coords = set()  # 모든 y 좌표 (start_y, end_y)

        cumulative_y = 0
        total_cells = 0

        for row_idx, row_start in enumerate(first_cols):
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
                    right_id not in first_cols_set):
                    visited_in_row.add(right_id)
                    queue.append((right_id, end_x, start_y))

                # 아래 이동 (행 내 분할 셀 처리)
                if end_y < row_end_y:
                    self.hwp.SetPos(current_id, 0, 0)
                    self.hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
                    down_id = self.hwp.GetPos()[0]

                    if (down_id != current_id and
                        down_id not in visited_in_row and
                        down_id not in first_cols_set):
                        visited_in_row.add(down_id)
                        queue.append((down_id, start_x, end_y))

            cumulative_y = row_end_y

        # 1.5단계: 테이블 경계 좌표 계산 및 필터링
        # X 경계: first_cols 셀의 좌측(0) ~ last_cols 셀의 우측
        # Y 경계: first_row의 상단(0) ~ last_row의 하단(cumulative_y)
        table_min_x = 0
        table_max_x = 0
        table_min_y = 0
        table_max_y = cumulative_y

        # last_cols 셀들의 우측 좌표 중 최대값 계산
        for last_col_id in last_cols_set:
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

        # 4단계, 5단계 비활성화 (순수 좌표 기반 매핑이 더 정확함)
        # cells, max_row, max_col = self._fix_empty_positions(
        #     cells, x_levels_list, y_levels_list, max_row, max_col
        # )
        # cells = self._fix_overlaps(cells, max_row, max_col)

        if self.debug:
            print(f"[calculate_grid] 결과: {len(cells)}개 셀, {max_row+1}행 x {max_col+1}열")

        return CellPositionResult(
            cells=cells,
            x_levels=x_levels_list,
            y_levels=y_levels_list,
            max_row=max_row,
            max_col=max_col,
        )

    def _fix_empty_positions(
        self,
        cells: Dict[int, CellRange],
        x_levels: List[int],
        y_levels: List[int],
        max_row: int,
        max_col: int
    ) -> tuple:
        """
        빈 위치를 인접 셀 커서 이동으로 보완

        알고리즘:
        1. 그리드 점유 맵 생성
        2. 빈 위치 찾기
        3. 빈 위치에서 인접 셀 기반 커서 이동으로 실제 셀 찾기
        4. 찾은 셀의 start_row를 빈 위치로 확장
        """
        from table.table_info import MOVE_RIGHT_OF_CELL, MOVE_LEFT_OF_CELL, MOVE_UP_OF_CELL, MOVE_DOWN_OF_CELL

        # 그리드 점유 맵 생성
        grid = {}  # (row, col) -> list_id
        for list_id, cell in cells.items():
            for r in range(cell.start_row, cell.end_row + 1):
                for c in range(cell.start_col, cell.end_col + 1):
                    grid[(r, c)] = list_id

        # 빈 위치 찾기
        empty_positions = []
        for r in range(max_row + 1):
            for c in range(max_col + 1):
                if (r, c) not in grid:
                    empty_positions.append((r, c))

        if not empty_positions:
            return cells, max_row, max_col

        if self.debug:
            print(f"[_fix_empty_positions] 빈 위치 {len(empty_positions)}개: {empty_positions}")

        # 빈 위치별로 실제 셀 찾기
        fixed_cells = {}  # list_id -> 확장할 row 범위

        for empty_row, empty_col in empty_positions:
            # 인접 셀 찾기
            adjacent_cells = []

            # 왼쪽 인접 셀
            for c in range(empty_col - 1, -1, -1):
                if (empty_row, c) in grid:
                    adjacent_cells.append(('left', grid[(empty_row, c)]))
                    break

            # 오른쪽 인접 셀
            for c in range(empty_col + 1, max_col + 1):
                if (empty_row, c) in grid:
                    adjacent_cells.append(('right', grid[(empty_row, c)]))
                    break

            # 위쪽 인접 셀
            for r in range(empty_row - 1, -1, -1):
                if (r, empty_col) in grid:
                    adjacent_cells.append(('up', grid[(r, empty_col)]))
                    break

            # 아래쪽 인접 셀
            for r in range(empty_row + 1, max_row + 1):
                if (r, empty_col) in grid:
                    adjacent_cells.append(('down', grid[(r, empty_col)]))
                    break

            if not adjacent_cells:
                continue

            # 인접 셀에서 커서 이동으로 빈 위치의 실제 셀 찾기
            found_cell_id = None
            move_map = {
                'left': MOVE_RIGHT_OF_CELL,
                'right': MOVE_LEFT_OF_CELL,
                'up': MOVE_DOWN_OF_CELL,
                'down': MOVE_UP_OF_CELL,
            }

            for direction, adj_id in adjacent_cells:
                self.hwp.SetPos(adj_id, 0, 0)
                self.hwp.MovePos(move_map[direction], 0, 0)
                target_id = self.hwp.GetPos()[0]

                if target_id != adj_id and target_id in cells:
                    found_cell_id = target_id
                    break

            if found_cell_id:
                # 찾은 셀의 row 범위 확장
                if found_cell_id not in fixed_cells:
                    fixed_cells[found_cell_id] = {
                        'min_row': empty_row,
                        'max_row': empty_row,
                        'cols': {empty_col}
                    }
                else:
                    fixed_cells[found_cell_id]['min_row'] = min(
                        fixed_cells[found_cell_id]['min_row'], empty_row
                    )
                    fixed_cells[found_cell_id]['max_row'] = max(
                        fixed_cells[found_cell_id]['max_row'], empty_row
                    )
                    fixed_cells[found_cell_id]['cols'].add(empty_col)

        # 셀 범위 확장 적용
        for list_id, fix_info in fixed_cells.items():
            cell = cells[list_id]
            new_start_row = min(cell.start_row, fix_info['min_row'])
            new_end_row = max(cell.end_row, fix_info['max_row'])

            if new_start_row != cell.start_row or new_end_row != cell.end_row:
                if self.debug:
                    print(f"[_fix_empty_positions] 셀 {list_id} 확장: "
                          f"row {cell.start_row}~{cell.end_row} → {new_start_row}~{new_end_row}")

                cells[list_id] = CellRange(
                    list_id=list_id,
                    start_row=new_start_row,
                    start_col=cell.start_col,
                    end_row=new_end_row,
                    end_col=cell.end_col,
                    rowspan=new_end_row - new_start_row + 1,
                    colspan=cell.colspan,
                    start_x=cell.start_x,
                    start_y=cell.start_y,
                    end_x=cell.end_x,
                    end_y=cell.end_y,
                )

        # max_row/max_col 재계산
        if cells:
            max_row = max(c.end_row for c in cells.values())
            max_col = max(c.end_col for c in cells.values())

        return cells, max_row, max_col

    def _fix_overlaps(
        self,
        cells: Dict[int, CellRange],
        max_row: int,
        max_col: int
    ) -> Dict[int, CellRange]:
        """
        중복 위치 해결 (커서 이동 기반 인접 관계 확인)

        알고리즘:
        1. 중복 위치 찾기
        2. 중복된 셀 쌍에 대해 커서 이동으로 인접 관계 확인
        3. 위아래 인접이면, 위 셀의 end_row를 아래 셀의 start_row - 1로 조정
        """
        from table.table_info import MOVE_DOWN_OF_CELL, MOVE_UP_OF_CELL

        # 그리드에서 중복 찾기
        grid = {}  # (row, col) -> [list_ids]
        for list_id, cell in cells.items():
            for r in range(cell.start_row, cell.end_row + 1):
                for c in range(cell.start_col, cell.end_col + 1):
                    key = (r, c)
                    if key not in grid:
                        grid[key] = []
                    grid[key].append(list_id)

        # 중복 위치 찾기
        overlaps = {k: v for k, v in grid.items() if len(v) > 1}

        if not overlaps:
            return cells

        if self.debug:
            print(f"[_fix_overlaps] 중복 위치 {len(overlaps)}개 발견")

        # 중복 셀 쌍 수집
        overlap_pairs = set()
        for pos, ids in overlaps.items():
            if len(ids) == 2:
                overlap_pairs.add(tuple(sorted(ids)))

        # 각 쌍에 대해 인접 관계 확인 및 조정
        for id1, id2 in overlap_pairs:
            cell1 = cells[id1]
            cell2 = cells[id2]

            # 커서 이동으로 인접 관계 확인
            # id1 → 아래 → id2 이면 id1이 위에 있음
            self.hwp.SetPos(id1, 0, 0)
            self.hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
            down_of_1 = self.hwp.GetPos()[0]

            self.hwp.SetPos(id2, 0, 0)
            self.hwp.MovePos(MOVE_UP_OF_CELL, 0, 0)
            up_of_2 = self.hwp.GetPos()[0]

            if down_of_1 == id2 and up_of_2 == id1:
                # id1이 위, id2가 아래로 인접
                # id1의 end_row를 id2의 start_row - 1로 조정
                new_end_row = cell2.start_row - 1

                if new_end_row >= cell1.start_row:
                    if self.debug:
                        print(f"[_fix_overlaps] 셀 {id1} end_row 조정: "
                              f"{cell1.end_row} → {new_end_row}")

                    cells[id1] = CellRange(
                        list_id=id1,
                        start_row=cell1.start_row,
                        start_col=cell1.start_col,
                        end_row=new_end_row,
                        end_col=cell1.end_col,
                        rowspan=new_end_row - cell1.start_row + 1,
                        colspan=cell1.colspan,
                        start_x=cell1.start_x,
                        start_y=cell1.start_y,
                        end_x=cell1.end_x,
                        end_y=cell1.end_y,
                    )
            elif down_of_1 == id1 and up_of_2 == id2:
                # 반대로 id2가 위, id1이 아래
                new_end_row = cell1.start_row - 1

                if new_end_row >= cell2.start_row:
                    if self.debug:
                        print(f"[_fix_overlaps] 셀 {id2} end_row 조정: "
                              f"{cell2.end_row} → {new_end_row}")

                    cells[id2] = CellRange(
                        list_id=id2,
                        start_row=cell2.start_row,
                        start_col=cell2.start_col,
                        end_row=new_end_row,
                        end_col=cell2.end_col,
                        rowspan=new_end_row - cell2.start_row + 1,
                        colspan=cell2.colspan,
                        start_x=cell2.start_x,
                        start_y=cell2.start_y,
                        end_x=cell2.end_x,
                        end_y=cell2.end_y,
                    )

        return cells

    def _calculate_bfs(self, max_cells: int = 1000) -> CellPositionResult:
        """
        [DEPRECATED] BFS 방식으로 테이블의 모든 셀 위치 및 범위 계산

        주의: 이 메서드는 더 이상 사용되지 않습니다.
        calculate_grid() 또는 calculate()를 사용하세요.

        문제점:
        - 순회 방향에 따라 좌표 계산이 부정확할 수 있음
        - 복잡한 병합 셀 구조에서 좌표 오류 발생
        """
        from table.table_info import (
            MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL,
            MOVE_LEFT_OF_CELL, MOVE_UP_OF_CELL,
            MOVE_START_OF_CELL, MOVE_TOP_OF_CELL
        )
        from collections import deque

        table_info = self._get_table_info()

        if not table_info.is_in_table():
            raise ValueError("커서가 테이블 내부에 있지 않습니다.")

        # 첫 번째 셀로 이동
        table_info.move_to_first_cell()
        first_id = self.hwp.GetPos()[0]

        # 데이터 구조
        cell_positions = {}  # list_id -> {start_x, end_x, start_y, end_y}
        x_levels = {0}
        y_levels = {0}
        visited = set()

        # BFS 큐: (list_id, start_x, start_y, from_direction)
        # from_direction: 어느 방향에서 왔는지 ('right', 'down', 'left', 'up', 'start')
        queue = deque([(first_id, 0, 0, 'start')])
        visited.add(first_id)

        # 첫 셀 크기 구하기
        self.hwp.SetPos(first_id, 0, 0)
        first_width, first_height = table_info.get_cell_dimensions()

        cell_positions[first_id] = {
            'start_x': 0, 'end_x': first_width,
            'start_y': 0, 'end_y': first_height,
        }
        x_levels.add(0)
        y_levels.add(0)

        total_cells = 1

        while queue and total_cells < max_cells:
            current_id, cur_start_x, cur_start_y, from_dir = queue.popleft()

            # 현재 셀 정보
            cur_pos = cell_positions.get(current_id)
            if not cur_pos:
                continue

            cur_end_x = cur_pos['end_x']
            cur_end_y = cur_pos['end_y']
            cur_width = cur_end_x - cur_start_x
            cur_height = cur_end_y - cur_start_y

            # 3방향 탐색 (좌측 제외 - 우측 이동으로만 X 매핑)
            directions = [
                (MOVE_RIGHT_OF_CELL, 'right'),
                (MOVE_DOWN_OF_CELL, 'down'),
                (MOVE_UP_OF_CELL, 'up'),
            ]

            for move_cmd, direction in directions:
                self.hwp.SetPos(current_id, 0, 0)
                self.hwp.MovePos(move_cmd, 0, 0)
                next_id = self.hwp.GetPos()[0]

                if next_id == current_id or next_id in visited:
                    continue

                visited.add(next_id)
                total_cells += 1

                # 다음 셀 크기
                self.hwp.SetPos(next_id, 0, 0)
                next_width, next_height = table_info.get_cell_dimensions()

                # 좌표 계산 - 우측 이동 시에만 X 누적
                if direction == 'right':
                    # 우측 이동: X 누적
                    next_start_x = cur_end_x
                    next_start_y = cur_start_y
                elif direction == 'down':
                    # 아래 이동: Y 누적, X는 현재 셀과 동일
                    next_start_x = cur_start_x
                    next_start_y = cur_end_y
                elif direction == 'up':
                    # 위 이동: Y는 현재 - 이전 셀 높이
                    next_start_x = cur_start_x
                    next_start_y = cur_start_y - next_height

                next_end_x = next_start_x + next_width
                next_end_y = next_start_y + next_height

                cell_positions[next_id] = {
                    'start_x': next_start_x, 'end_x': next_end_x,
                    'start_y': next_start_y, 'end_y': next_end_y,
                }

                x_levels.add(next_start_x)
                y_levels.add(next_start_y)

                if self.debug:
                    print(f"[BFS] {current_id} --{direction}--> {next_id}: "
                          f"({next_start_x}, {next_start_y})")

                queue.append((next_id, next_start_x, next_start_y, direction))

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

    def _calculate_legacy(self, max_cells: int = 1000) -> CellPositionResult:
        """
        [DEPRECATED] 레거시 셀 위치 계산 (first_cols 기반)

        calculate_grid() 사용을 권장합니다.
        """
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
