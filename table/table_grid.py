# -*- coding: utf-8 -*-
"""
테이블 그리드 좌표 생성 모듈

# table 관련 정보는 우선하여 table_boundary.py에서 가져옵니다.
# 이 스크립트에서 계산하는 것은 아래한글 테이블의 grid 정보입니다.
# - 우리가 받을 수 있는 정보는 table_boundary.py에서의 경계선 정보입니다.
# - 그리고 테이블 셀의 list_id, 높이, 너비 정보입니다.
# 마우스 커서를 이용해서 오른쪽으로 이동하면서 셀의 corners를 계산합니다.
# - 기본적으로 마우스 우측으로 계속 이동하면 모든 셀을 다 방문하게 됩니다.
# - 우측으로 이동하면서 셀의 너비를 조회해서 누적해서 셀 좌표를 계산합니다.
# - 일부 병합되거나 분할되는 셀이 있는데 반복은 하지만, 너비는 계속 더합니다.  left_border_cell, right_border_cell을 이용해서 행을 분리합니다.
# 진행합니다
# - left_border_cell에 만나면 만나는 셀의 좌표는 이미 왔던거니깐 너비만 계산하고, 셀의 우측좌측 모서리 좌표는 상단 셀의 좌측하단 이랑 동일해야지요.
# - 테이블 원점에서 우측으로 이동하면서 셀의 좌표를 계산합니다. 셀의 list_id가 right_border_cell에 포함되면 행을 바꿉니다.
# - 바뀐 행은 이전에 방문한 셀일 수 있습니다. 오버랩되는데 신경쓰지 않아도 됩니다.
#   우측으로 가면서 셀 너비를 누적하고 right_border_cell을 만나면 행을 바꿉니다.
# - 이런 방법으로 cell_corners/2d q배열를 계산합니다.
"""
# 셀의 높이도 추정해서 y축도 관리 합니다.

# cell_line을 계산합니다.
# cell_corners를 받아서 이 좌표들을 지나는 xline, yline 들을 계산합니다.
# 그렇면  엑셀 스타일의  그리드 셀이 만들어 집니다. 
# 추가하자면 border 라인 밖으로 생성된 포인트나 라인은 제거 하고 셀정보를 반환합니다.


import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from dataclasses import dataclass, field
from typing import List, Tuple, Optional

try:
    from .table_boundary import TableBoundary, TableBoundaryResult
    from .table_info import TableInfo, MOVE_RIGHT_OF_CELL
except ImportError:
    from table_boundary import TableBoundary, TableBoundaryResult
    from table_info import TableInfo, MOVE_RIGHT_OF_CELL


@dataclass
class CellCorner:
    """셀의 4개 모서리 좌표 (HWPUNIT)"""
    top_left: Tuple[int, int]       # (x, y)
    top_right: Tuple[int, int]
    bottom_left: Tuple[int, int]
    bottom_right: Tuple[int, int]


@dataclass
class CellLine:
    """셀의 4개 테두리선 좌표 (HWPUNIT)
    각 선은 (x1, y1, x2, y2) 형식
    """
    top: Tuple[int, int, int, int]
    bottom: Tuple[int, int, int, int]
    left: Tuple[int, int, int, int]
    right: Tuple[int, int, int, int]


@dataclass
class GridCell:
    """그리드 셀 정보"""
    list_id: int
    row: int
    col: int
    corners: CellCorner
    lines: CellLine

    @property
    def width(self) -> int:
        return self.corners.top_right[0] - self.corners.top_left[0]

    @property
    def height(self) -> int:
        return self.corners.bottom_left[1] - self.corners.top_left[1]


@dataclass
class TableGridResult:
    """테이블 그리드 결과"""
    cells: List[GridCell] = field(default_factory=list)
    row_count: int = 0
    col_count: int = 0

    def get_by_position(self, row: int, col: int) -> Optional[GridCell]:
        """(row, col)로 셀 조회"""
        for cell in self.cells:
            if cell.row == row and cell.col == col:
                return cell
        return None

    def get_by_list_id(self, list_id: int) -> Optional[GridCell]:
        """list_id로 셀 조회"""
        for cell in self.cells:
            if cell.list_id == list_id:
                return cell
        return None


@dataclass
class GridLines:
    """엑셀 스타일 그리드 라인 정보"""
    x_lines: List[int] = field(default_factory=list)  # 세로선 x좌표들
    y_lines: List[int] = field(default_factory=list)  # 가로선 y좌표들
    table_width: int = 0
    table_height: int = 0

    @property
    def col_count(self) -> int:
        return max(0, len(self.x_lines) - 1)

    @property
    def row_count(self) -> int:
        return max(0, len(self.y_lines) - 1)


@dataclass
class ExcelStyleCell:
    """엑셀 스타일 그리드 셀"""
    row: int
    col: int
    x1: int
    y1: int
    x2: int
    y2: int

    @property
    def width(self) -> int:
        return self.x2 - self.x1

    @property
    def height(self) -> int:
        return self.y2 - self.y1


@dataclass
class ExcelStyleGrid:
    """엑셀 스타일 그리드 결과"""
    lines: GridLines = None
    cells: List[List[ExcelStyleCell]] = field(default_factory=list)

    def get(self, row: int, col: int) -> Optional[ExcelStyleCell]:
        if self.cells and 0 <= row < len(self.cells) and 0 <= col < len(self.cells[row]):
            return self.cells[row][col]
        return None


@dataclass
class CellGridMapping:
    """테이블 셀과 그리드 셀 매칭 결과"""
    list_id: int
    table_cell: GridCell
    grid_cells: List[Tuple[int, int]]  # (row, col) 리스트
    row_span: Tuple[int, int]  # (start_row, end_row)
    col_span: Tuple[int, int]  # (start_col, end_col)

    @property
    def row_count(self) -> int:
        return self.row_span[1] - self.row_span[0] + 1

    @property
    def col_count(self) -> int:
        return self.col_span[1] - self.col_span[0] + 1


class TableGrid:
    """테이블 그리드 생성 클래스

    테이블 원점에서 우측으로 이동하면서 셀 좌표를 계산합니다.
    - 셀의 list_id가 right_border_cell에 포함되면 행을 바꿉니다.
    - 바뀐 행은 이전에 방문한 셀일 수 있지만 신경쓰지 않습니다.
    - 우측으로 가면서 셀 너비를 누적하고 right_border_cell을 만나면 행을 바꿉니다.
    """

    def __init__(self, hwp=None, debug: bool = False):
        from cursor import get_hwp_instance
        self.hwp = hwp or get_hwp_instance()
        self.debug = debug
        self._table_info = TableInfo(self.hwp, debug)
        self._boundary = TableBoundary(self.hwp, debug)

    def _log(self, msg: str):
        if self.debug:
            print(f"[TableGrid] {msg}")

    def build_grid(self, boundary_result: TableBoundaryResult = None) -> TableGridResult:
        """
        테이블 원점에서 우측으로 이동하면서 셀의 corners를 계산합니다.

        - 셀 방문 여부를 추적하여 중첩 셀(세로 병합 등)을 처리합니다.
        - 중첩 셀: x는 기존 corners에서, y는 현재 행의 상단 y좌표 기준으로 계산

        Returns:
            TableGridResult: 셀별 list_id, row, col, corners, lines 정보
        """
        if boundary_result is None:
            boundary_result = self._boundary.check_boundary_table()

        right_border_set = set(boundary_result.right_border_cells)

        result = TableGridResult()

        # 테이블 원점으로 이동
        self.hwp.SetPos(boundary_result.table_origin, 0, 0)

        current_row = 0
        current_col = 0
        cumulative_x = 0
        cumulative_y = 0       # 현재 행의 상단 y좌표
        row_heights = []       # 현재 행의 비중첩 셀 높이 목록
        max_col = 0

        # 방문한 셀 정보: {list_id: CellCorner}
        visited = {}

        while True:
            list_id = self.hwp.GetPos()[0]

            # 이미 방문한 셀인지 확인 (중첩 셀 처리)
            if list_id in visited:
                prev_corners = visited[list_id]

                # x는 이전 방문 시의 x좌표 사용 (너비 유지)
                x1 = prev_corners.top_left[0]
                x2 = prev_corners.top_right[0]

                # cumulative_x 갱신 (다음 셀 위치)
                cumulative_x = x2

                # 중첩 셀은 row_heights 계산에 포함하지 않음 (세로 병합 셀이므로)
                # y 좌표는 비중첩 셀들에 의해 결정됨

                self._log(f"중첩 ({current_row},{current_col}) list_id={list_id} x={x1}~{x2} (skip y)")

                current_col += 1
                max_col = max(max_col, current_col)

                # right_border_cell이면 행 변경
                if list_id in right_border_set:
                    # 비중첩 셀의 최소 높이를 행 높이로 사용 (세로 병합 셀 영향 제거)
                    row_height = min(row_heights) if row_heights else 0
                    current_row += 1
                    current_col = 0
                    cumulative_x = 0
                    cumulative_y += row_height
                    row_heights = []
                    self._log(f"행 변경 → {current_row}행 (y={cumulative_y})")

                # 다음 셀로 이동
                self.hwp.SetPos(list_id, 0, 0)
                self.hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
                next_id = self.hwp.GetPos()[0]

                if next_id == list_id:
                    break
                continue

            # 셀 크기 조회
            cell_width, cell_height = self._table_info.get_cell_dimensions()

            # corners 계산
            x1 = cumulative_x
            y1 = cumulative_y
            x2 = cumulative_x + cell_width
            y2 = cumulative_y + cell_height

            corners = CellCorner(
                top_left=(x1, y1),
                top_right=(x2, y1),
                bottom_left=(x1, y2),
                bottom_right=(x2, y2)
            )

            lines = CellLine(
                top=(x1, y1, x2, y1),
                bottom=(x1, y2, x2, y2),
                left=(x1, y1, x1, y2),
                right=(x2, y1, x2, y2)
            )

            cell = GridCell(
                list_id=list_id,
                row=current_row,
                col=current_col,
                corners=corners,
                lines=lines
            )
            result.cells.append(cell)

            # 방문 기록 (corners 저장)
            visited[list_id] = corners

            self._log(f"셀 ({current_row},{current_col}) list_id={list_id} x={x1}~{x2} y={y1}~{y2}")

            # x 좌표 누적
            cumulative_x += cell_width
            current_col += 1
            max_col = max(max_col, current_col)

            # 행 높이 목록에 추가
            row_heights.append(cell_height)

            # right_border_cell이면 행 변경
            if list_id in right_border_set:
                # 비중첩 셀의 최소 높이를 행 높이로 사용 (세로 병합 셀 영향 제거)
                row_height = min(row_heights) if row_heights else 0
                current_row += 1
                current_col = 0
                cumulative_x = 0
                cumulative_y += row_height
                row_heights = []
                self._log(f"행 변경 → {current_row}행 (y={cumulative_y})")

            # 다음 셀로 이동
            self.hwp.SetPos(list_id, 0, 0)
            self.hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
            next_id = self.hwp.GetPos()[0]

            if next_id == list_id:
                # 더 이상 이동 불가
                break

        result.row_count = current_row + 1 if current_col > 0 else current_row
        result.col_count = max_col

        self._log(f"완료: {result.row_count}행 x {result.col_count}열, 총 {len(result.cells)}셀")

        return result

    def build_grid_lines(self, grid_result: TableGridResult = None,
                         boundary_result: TableBoundaryResult = None) -> ExcelStyleGrid:
        """
        cell_corners를 받아서 x_lines, y_lines를 계산합니다.
        엑셀 스타일의 그리드 셀을 생성합니다.
        border 라인 밖의 포인트/라인은 제거합니다.
        """
        if boundary_result is None:
            boundary_result = self._boundary.check_boundary_table()

        if grid_result is None:
            grid_result = self.build_grid(boundary_result)

        # 테이블 경계
        table_width = boundary_result.end_x - boundary_result.start_x
        table_height = boundary_result.end_y - boundary_result.start_y

        self._log(f"테이블 경계: width={table_width}, height={table_height}")

        # 모든 셀의 corners에서 x, y 좌표 수집
        all_x = set()
        all_y = set()
        for cell in grid_result.cells:
            c = cell.corners
            all_x.add(c.top_left[0])
            all_x.add(c.top_right[0])
            all_y.add(c.top_left[1])
            all_y.add(c.bottom_left[1])

        # border 밖 좌표 제거
        x_lines = sorted([x for x in all_x if 0 <= x <= table_width])
        y_lines = sorted([y for y in all_y if 0 <= y <= table_height])

        self._log(f"x_lines: {len(x_lines)}개, y_lines: {len(y_lines)}개")

        # GridLines 생성
        lines = GridLines(
            x_lines=x_lines,
            y_lines=y_lines,
            table_width=table_width,
            table_height=table_height
        )

        # 엑셀 스타일 그리드 셀 생성
        cells = []
        for row in range(lines.row_count):
            row_cells = []
            for col in range(lines.col_count):
                cell = ExcelStyleCell(
                    row=row,
                    col=col,
                    x1=x_lines[col],
                    y1=y_lines[row],
                    x2=x_lines[col + 1],
                    y2=y_lines[row + 1]
                )
                row_cells.append(cell)
            cells.append(row_cells)

        result = ExcelStyleGrid(lines=lines, cells=cells)
        self._log(f"엑셀 스타일 그리드: {lines.row_count}행 x {lines.col_count}열")

        return result

    def map_cells_to_grid(self, grid_result: TableGridResult,
                          excel_grid: ExcelStyleGrid,
                          tolerance: int = 30) -> List[CellGridMapping]:
        """
        테이블 셀과 그리드 셀을 매칭합니다.

        각 테이블 셀의 corners를 기준으로 해당 범위에 포함되는 그리드 셀을 찾습니다.
        오차 허용을 위해 테이블 셀 범위를 전방향으로 tolerance만큼 확장합니다.

        Args:
            grid_result: build_grid() 결과 (테이블 셀 정보)
            excel_grid: build_grid_lines() 결과 (그리드 셀 정보)
            tolerance: 오차 허용 범위 (기본 30 HWPUNIT)

        Returns:
            List[CellGridMapping]: 각 테이블 셀에 매칭된 그리드 셀 정보
        """
        mappings = []

        for table_cell in grid_result.cells:
            # 테이블 셀 범위 (tolerance 적용)
            t_x1 = table_cell.corners.top_left[0] - tolerance
            t_y1 = table_cell.corners.top_left[1] - tolerance
            t_x2 = table_cell.corners.bottom_right[0] + tolerance
            t_y2 = table_cell.corners.bottom_right[1] + tolerance

            matched_cells = []
            min_row, max_row = float('inf'), -1
            min_col, max_col = float('inf'), -1

            # 모든 그리드 셀을 순회하며 매칭
            for row_idx, row_cells in enumerate(excel_grid.cells):
                for col_idx, grid_cell in enumerate(row_cells):
                    # 그리드 셀이 테이블 셀 범위 안에 포함되는지 확인
                    if (grid_cell.x1 >= t_x1 and grid_cell.x2 <= t_x2 and
                        grid_cell.y1 >= t_y1 and grid_cell.y2 <= t_y2):
                        matched_cells.append((row_idx, col_idx))
                        min_row = min(min_row, row_idx)
                        max_row = max(max_row, row_idx)
                        min_col = min(min_col, col_idx)
                        max_col = max(max_col, col_idx)

            # 매칭된 셀이 있으면 결과 추가
            if matched_cells:
                mapping = CellGridMapping(
                    list_id=table_cell.list_id,
                    table_cell=table_cell,
                    grid_cells=matched_cells,
                    row_span=(min_row, max_row),
                    col_span=(min_col, max_col)
                )
                mappings.append(mapping)
                self._log(f"list_id={table_cell.list_id} → 그리드 ({min_row},{min_col})~({max_row},{max_col}) [{len(matched_cells)}셀]")

        return mappings

    def build_grid_with_log(self, boundary_result: TableBoundaryResult = None) -> TableGridResult:
        """build_grid() 래핑 - 요약 출력 후 반환"""
        result = self.build_grid(boundary_result)
        self._print_summary(result)
        return result

    def _print_summary(self, result: TableGridResult):
        """그리드 요약 출력"""
        print(f"\n=== 테이블 그리드 ===")
        print(f"크기: {result.row_count}행 x {result.col_count}열")
        print(f"총 셀 수: {len(result.cells)}")

    def print_all_cells(self, result: TableGridResult = None):
        """모든 셀 출력"""
        if result is None:
            result = self.build_grid()

        self._print_summary(result)
        print(f"\n=== 셀 목록 ===")
        for cell in result.cells:
            print(f"({cell.row},{cell.col}) list_id={cell.list_id} | {cell.corners.top_left} ~ {cell.corners.bottom_right} | {cell.width}x{cell.height}")


if __name__ == "__main__":
    from cursor import get_hwp_instance

    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        exit(1)

    grid = TableGrid(hwp, debug=True)

    if not grid._boundary._is_in_table():
        print("[오류] 커서가 테이블 내부에 있지 않습니다.")
        exit(1)

    # boundary 정보 먼저 출력
    boundary = grid._boundary.check_boundary_table()
    print(f"\n=== 테이블 경계 정보 ===")
    print(f"원점: list_id={boundary.table_origin}")
    print(f"크기: {boundary.end_x - boundary.start_x} x {boundary.end_y - boundary.start_y}")
    print(f"우측 경계 셀: {boundary.right_border_cells}")

    # 그리드 빌드
    result = grid.build_grid_with_log(boundary)

    # 셀별 상세 출력
    print(f"\n=== 셀 상세 정보 ===")
    for cell in result.cells:
        print(f"({cell.row},{cell.col}) list_id={cell.list_id:3d} | "
              f"x: {cell.corners.top_left[0]:6d} ~ {cell.corners.top_right[0]:6d} | "
              f"y: {cell.corners.top_left[1]:6d} ~ {cell.corners.bottom_left[1]:6d} | "
              f"{cell.width:6d} x {cell.height:6d}")

    # 그리드 라인 테스트
    print(f"\n=== 그리드 라인 ===")
    excel_grid = grid.build_grid_lines(result, boundary)
    print(f"x_lines ({len(excel_grid.lines.x_lines)}개): {excel_grid.lines.x_lines}")
    print(f"y_lines ({len(excel_grid.lines.y_lines)}개): {excel_grid.lines.y_lines}")

    # 2D 배열 출력
    print(f"\n=== 2D 그리드 배열 ({len(excel_grid.cells)}행 x {len(excel_grid.cells[0]) if excel_grid.cells else 0}열) ===")
    for row_idx, row_cells in enumerate(excel_grid.cells):
        row_str = f"Row {row_idx:2d}: "
        cell_strs = []
        for cell in row_cells:
            cell_strs.append(f"({cell.x1},{cell.y1})-({cell.x2},{cell.y2})")
        print(row_str + " | ".join(cell_strs[:6]) + (" ..." if len(cell_strs) > 6 else ""))

    # 테이블 셀 → 그리드 셀 매칭
    print(f"\n=== 테이블 셀 → 그리드 셀 매칭 (tolerance=30) ===")
    mappings = grid.map_cells_to_grid(result, excel_grid, tolerance=30)
    for m in mappings:
        print(f"list_id={m.list_id:3d} → 그리드 row:{m.row_span[0]}~{m.row_span[1]} col:{m.col_span[0]}~{m.col_span[1]} "
              f"({m.row_count}x{m.col_count}={len(m.grid_cells)}셀)")

    # 매칭되지 않은 그리드 셀 확인
    print(f"\n=== 매칭되지 않은 그리드 셀 확인 ===")
    all_matched = set()
    for m in mappings:
        for cell in m.grid_cells:
            all_matched.add(cell)

    total_grid_cells = len(excel_grid.cells) * len(excel_grid.cells[0]) if excel_grid.cells else 0
    unmatched = []
    for row_idx in range(len(excel_grid.cells)):
        for col_idx in range(len(excel_grid.cells[0])):
            if (row_idx, col_idx) not in all_matched:
                unmatched.append((row_idx, col_idx))

    print(f"전체 그리드 셀: {total_grid_cells}개")
    print(f"매칭된 셀: {len(all_matched)}개")
    print(f"매칭 안된 셀: {len(unmatched)}개")
    if unmatched:
        print(f"매칭 안된 셀 목록: {unmatched[:20]}{'...' if len(unmatched) > 20 else ''}")
