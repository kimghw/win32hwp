# -*- coding: utf-8 -*-
"""
셀 좌표/범위 조회 모듈

테이블 셀의 좌표, 범위, 병합 정보 등을 조회합니다.
"""

from typing import Dict, List, Optional, Tuple

# table/cell_position.py의 클래스를 재사용 (중복 정의 방지)
from table.cell_position import CellRange, CellPositionResult


def calculate_cell_positions(hwp, max_cells: int = 1000) -> CellPositionResult:
    """
    테이블의 모든 셀 위치 및 범위 계산

    xline/yline 기반으로 정확한 그리드를 생성합니다.

    Args:
        hwp: HWP COM 객체
        max_cells: 최대 처리할 셀 수 (기본 1000)

    Returns:
        CellPositionResult: 셀 위치 계산 결과
    """
    from table.cell_position import CellPositionCalculator
    calc = CellPositionCalculator(hwp, debug=False)
    return calc.calculate(max_cells)


def get_cell_at(result: CellPositionResult, row: int, col: int) -> Optional[CellRange]:
    """
    특정 좌표에 있는 셀 반환

    Args:
        result: 셀 위치 계산 결과
        row: 행 좌표
        col: 열 좌표

    Returns:
        CellRange 또는 None
    """
    for cell in result.cells.values():
        if cell.contains(row, col):
            return cell
    return None


def get_cell_by_coord(result: CellPositionResult, row: int, col: int) -> Optional[int]:
    """
    좌표로 list_id 찾기 (병합 셀 포함)

    Args:
        result: 셀 위치 계산 결과
        row: 행 좌표
        col: 열 좌표

    Returns:
        list_id 또는 None
    """
    cell = get_cell_at(result, row, col)
    return cell.list_id if cell else None


def get_representative_coord(result: CellPositionResult, row: int, col: int) -> Optional[Tuple[int, int]]:
    """
    해당 좌표의 대표 좌표 반환 (병합 셀이면 좌상단)

    Args:
        result: 셀 위치 계산 결과
        row: 행 좌표
        col: 열 좌표

    Returns:
        (row, col) 대표 좌표 또는 None
    """
    cell = get_cell_at(result, row, col)
    if cell:
        return (cell.start_row, cell.start_col)
    return None


def get_covered_coords(cell: CellRange) -> List[Tuple[int, int]]:
    """
    셀이 차지하는 모든 좌표 목록 반환

    Args:
        cell: CellRange 객체

    Returns:
        [(row, col), ...] 좌표 리스트
    """
    coords = []
    for r in range(cell.start_row, cell.end_row + 1):
        for c in range(cell.start_col, cell.end_col + 1):
            coords.append((r, c))
    return coords


def build_coord_to_listid_map(result: CellPositionResult) -> Dict[Tuple[int, int], int]:
    """
    (row, col) → list_id 매핑 테이블 생성 (병합 셀 포함)

    Args:
        result: 셀 위치 계산 결과

    Returns:
        {(row, col): list_id} 매핑
    """
    coord_map = {}
    for cell in result.cells.values():
        for coord in get_covered_coords(cell):
            coord_map[coord] = cell.list_id
    return coord_map


def get_merged_cells(result: CellPositionResult) -> List[CellRange]:
    """
    병합된 셀 목록 반환

    Args:
        result: 셀 위치 계산 결과

    Returns:
        CellRange 리스트
    """
    return [c for c in result.cells.values() if c.is_merged()]


def get_merge_info(result: CellPositionResult, list_id: int) -> Optional[dict]:
    """
    특정 셀의 병합 정보 반환

    Args:
        result: 셀 위치 계산 결과
        list_id: 셀의 list_id

    Returns:
        {
            'list_id': list_id,
            'start': (row, col),
            'end': (row, col),
            'rowspan': int,
            'colspan': int,
            'is_merged': bool,
            'covered_coords': [(row, col), ...]
        }
    """
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
        'covered_coords': get_covered_coords(cell),
    }


def get_all_merge_info(result: CellPositionResult) -> List[dict]:
    """
    모든 병합 셀 정보 목록

    Args:
        result: 셀 위치 계산 결과

    Returns:
        [병합 정보 dict, ...]
    """
    merged_cells = get_merged_cells(result)
    return [get_merge_info(result, c.list_id) for c in merged_cells]


def print_cell_summary(result: CellPositionResult):
    """
    셀 위치 계산 결과 요약 출력

    Args:
        result: 셀 위치 계산 결과
    """
    print(f"\n=== 셀 위치 계산 결과 ===")
    print(f"X 레벨: {len(result.x_levels)}개")
    print(f"Y 레벨: {len(result.y_levels)}개")
    print(f"총 셀: {len(result.cells)}개")
    print(f"테이블 크기: {result.max_row + 1}행 x {result.max_col + 1}열")

    merged = get_merged_cells(result)
    if merged:
        print(f"\n=== 병합 셀 ({len(merged)}개) ===")
        for cell in sorted(merged, key=lambda c: (c.start_row, c.start_col)):
            print(f"  list_id={cell.list_id}: {cell}")
