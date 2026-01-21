# -*- coding: utf-8 -*-
"""HWP 데이터 추출 통합 모듈

한글 문서에서 데이터를 추출하는 모든 함수를 통합합니다.
- 페이지 정보 추출
- 셀 스타일 추출
- 필드 이름 생성

이 모듈은 Excel/openpyxl 관련 코드를 포함하지 않습니다.
"""

from dataclasses import dataclass, field
from typing import Optional, List, Dict, Any

# 기존 모듈에서 데이터 클래스 및 함수 임포트
from ..page_meta import HwpPageMeta, Unit
from .page import PageMatchResult, extract_page_info
from .cell import (
    CellStyleData,
    CellMatchResult,
    extract_cell_style,
    get_cell_text,
)
from .field import (
    FieldInfo,
    generate_field_names,
    get_cell_bookmark,
    find_left_header_cell,
    find_top_header_cell,
    clean_text_for_field_name,
    generate_random_field_name,
)


@dataclass
class HwpExtractedData:
    """HWP에서 추출된 모든 데이터를 담는 컨테이너

    Attributes:
        page_result: 페이지 정보 추출 결과 (PageMatchResult)
        cells_data: 셀 스타일 데이터 리스트 (List[CellStyleData])
        fields: 필드 정보 리스트 (List[FieldInfo])
        row_heights: 행 높이 리스트 (HWPUNIT)
        col_widths: 열 너비 리스트 (HWPUNIT)
        success: 추출 성공 여부
        error: 오류 메시지 (실패 시)
    """
    page_result: Optional[PageMatchResult] = None
    cells_data: List[CellStyleData] = field(default_factory=list)
    fields: List[FieldInfo] = field(default_factory=list)
    row_heights: List[int] = field(default_factory=list)
    col_widths: List[int] = field(default_factory=list)
    success: bool = False
    error: Optional[str] = None

    def get_cells_by_position(self) -> Dict[tuple, CellStyleData]:
        """(row, col) -> CellStyleData 매핑 반환"""
        return {(c.row, c.col): c for c in self.cells_data}

    def get_fields_by_position(self) -> Dict[tuple, FieldInfo]:
        """(row, col) -> FieldInfo 매핑 반환"""
        return {(f.row, f.col): f for f in self.fields}

    def get_field_by_name(self, field_name: str) -> Optional[FieldInfo]:
        """필드 이름으로 FieldInfo 검색"""
        for f in self.fields:
            if f.field_name == field_name:
                return f
        return None

    def to_dict(self) -> Dict[str, Any]:
        """딕셔너리로 변환"""
        return {
            'page_result': {
                'page_meta': self.page_result.page_meta.to_dict() if self.page_result and self.page_result.page_meta else None,
                'success': self.page_result.success if self.page_result else False,
                'error': self.page_result.error if self.page_result else None,
            } if self.page_result else None,
            'cells_count': len(self.cells_data),
            'fields_count': len(self.fields),
            'row_count': len(self.row_heights),
            'col_count': len(self.col_widths),
            'success': self.success,
            'error': self.error,
        }


def extract_hwp_data(
    hwp,
    calc_result=None,
    config=None,
    include_page_info: bool = True,
    include_cell_styles: bool = True,
    include_fields: bool = True,
    tolerance: int = 50,
) -> HwpExtractedData:
    """한글 문서에서 모든 데이터 추출

    Args:
        hwp: HWP 객체
        calc_result: CellPositionResult 객체 (None이면 내부에서 계산)
        config: ExportConfig 객체 (선택)
        include_page_info: 페이지 정보 추출 여부
        include_cell_styles: 셀 스타일 추출 여부
        include_fields: 필드 이름 생성 여부
        tolerance: 필드 이름 생성 시 크기 비교 허용 오차 (HWPUNIT)

    Returns:
        HwpExtractedData 객체
    """
    result = HwpExtractedData()

    try:
        # 1. 페이지 정보 추출
        if include_page_info:
            result.page_result = extract_page_info(hwp)
            if not result.page_result.success:
                # 페이지 정보 추출 실패해도 계속 진행
                pass

        # 2. 셀 위치 계산 (calc_result가 없으면 직접 계산)
        if calc_result is None and (include_cell_styles or include_fields):
            from table.table_info import TableInfo
            from table.cell_position import CellPositionCalculator

            table_info = TableInfo(hwp, debug=False)
            if not table_info.is_in_table():
                result.error = "커서가 표 안에 있지 않습니다."
                return result

            calc = CellPositionCalculator(hwp, debug=False)
            calc_result = calc.calculate()

        # 3. 셀 스타일 추출
        if include_cell_styles and calc_result:
            cells_data = []
            row_heights_set = set()
            col_widths_set = set()

            for list_id, cell_range in calc_result.cells.items():
                # 셀 스타일 추출
                style = extract_cell_style(hwp, list_id)

                # 위치 정보 설정
                style.list_id = list_id
                style.row = cell_range.start_row
                style.col = cell_range.start_col
                style.end_row = cell_range.end_row
                style.end_col = cell_range.end_col
                style.rowspan = cell_range.rowspan
                style.colspan = cell_range.colspan

                # 물리 좌표 설정
                style.x = cell_range.start_x
                style.y = cell_range.start_y
                style.width = cell_range.end_x - cell_range.start_x
                style.height = cell_range.end_y - cell_range.start_y

                # 텍스트 추출
                style.text = get_cell_text(hwp, list_id)

                cells_data.append(style)

            result.cells_data = cells_data

            # 행 높이 / 열 너비 계산
            # y_levels에서 행 높이 계산
            y_levels = sorted(calc_result.y_levels)
            for i in range(len(y_levels) - 1):
                row_heights_set.add(y_levels[i + 1] - y_levels[i])
            result.row_heights = _calculate_row_heights(calc_result)

            # x_levels에서 열 너비 계산
            result.col_widths = _calculate_col_widths(calc_result)

        # 4. 필드 이름 생성
        if include_fields and result.cells_data:
            result.fields = generate_field_names(hwp, result.cells_data, tolerance)

            # 셀 데이터에 필드 정보 추가
            fields_dict = {(f.list_id): f for f in result.fields}
            for cell in result.cells_data:
                if cell.list_id in fields_dict:
                    field_info = fields_dict[cell.list_id]
                    cell.field_name = field_info.field_name
                    cell.field_source = field_info.source

        result.success = True

    except Exception as e:
        result.error = str(e)
        result.success = False

    return result


def _calculate_row_heights(calc_result) -> List[int]:
    """CellPositionResult에서 행 높이 계산

    Args:
        calc_result: CellPositionResult 객체

    Returns:
        행 높이 리스트 (HWPUNIT)
    """
    y_levels = sorted(calc_result.y_levels)
    row_heights = []

    for i in range(len(y_levels) - 1):
        height = y_levels[i + 1] - y_levels[i]
        row_heights.append(height)

    return row_heights


def _calculate_col_widths(calc_result) -> List[int]:
    """CellPositionResult에서 열 너비 계산

    Args:
        calc_result: CellPositionResult 객체

    Returns:
        열 너비 리스트 (HWPUNIT)
    """
    x_levels = sorted(calc_result.x_levels)
    col_widths = []

    for i in range(len(x_levels) - 1):
        width = x_levels[i + 1] - x_levels[i]
        col_widths.append(width)

    return col_widths


def extract_cells_only(hwp, calc_result=None) -> List[CellStyleData]:
    """셀 스타일만 추출 (간편 함수)

    Args:
        hwp: HWP 객체
        calc_result: CellPositionResult 객체 (선택)

    Returns:
        CellStyleData 리스트
    """
    result = extract_hwp_data(
        hwp,
        calc_result=calc_result,
        include_page_info=False,
        include_cell_styles=True,
        include_fields=False,
    )
    return result.cells_data if result.success else []


def extract_fields_only(hwp, cells_data: List[CellStyleData] = None, tolerance: int = 50) -> List[FieldInfo]:
    """필드 이름만 추출 (간편 함수)

    Args:
        hwp: HWP 객체
        cells_data: CellStyleData 리스트 (None이면 내부에서 추출)
        tolerance: 크기 비교 허용 오차 (HWPUNIT)

    Returns:
        FieldInfo 리스트
    """
    if cells_data is None:
        cells_data = extract_cells_only(hwp)

    if not cells_data:
        return []

    return generate_field_names(hwp, cells_data, tolerance)


# ============================================================
# 하위 호환성을 위한 re-export
# ============================================================

__all__ = [
    # 데이터 클래스
    'HwpExtractedData',
    'PageMatchResult',
    'CellStyleData',
    'CellMatchResult',
    'FieldInfo',
    'HwpPageMeta',
    'Unit',

    # 메인 함수
    'extract_hwp_data',
    'extract_cells_only',
    'extract_fields_only',

    # 페이지 관련
    'extract_page_info',

    # 셀 관련
    'extract_cell_style',
    'get_cell_text',

    # 필드 관련
    'generate_field_names',
    'get_cell_bookmark',
    'find_left_header_cell',
    'find_top_header_cell',
    'clean_text_for_field_name',
    'generate_random_field_name',
]


# ============================================================
# 테스트
# ============================================================

if __name__ == "__main__":
    import sys
    import os
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

    from cursor import get_hwp_instance

    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        exit(1)

    print("=== HWP 데이터 추출 테스트 ===\n")

    result = extract_hwp_data(hwp)

    if result.success:
        print("추출 성공!")
        print(f"\n페이지 정보:")
        if result.page_result and result.page_result.success:
            meta = result.page_result.page_meta
            print(f"  용지 크기: {Unit.hwpunit_to_cm(meta.page_size.width):.1f} x {Unit.hwpunit_to_cm(meta.page_size.height):.1f} cm")
            print(f"  용지 방향: {meta.page_size.orientation}")

        print(f"\n셀 데이터: {len(result.cells_data)}개")
        for cell in result.cells_data[:5]:  # 처음 5개만 출력
            bg = f"bg={cell.bg_color_rgb}" if cell.bg_color_rgb else "no_bg"
            print(f"  ({cell.row},{cell.col}): {bg}, text='{cell.text[:20]}...' " if len(cell.text) > 20 else f"  ({cell.row},{cell.col}): {bg}, text='{cell.text}'")

        print(f"\n필드 정보: {len(result.fields)}개")
        for f in result.fields[:5]:  # 처음 5개만 출력
            print(f"  ({f.row},{f.col}): {f.field_name} [{f.source}]")

        print(f"\n행 높이: {len(result.row_heights)}개")
        print(f"  {result.row_heights[:5]}..." if len(result.row_heights) > 5 else f"  {result.row_heights}")

        print(f"\n열 너비: {len(result.col_widths)}개")
        print(f"  {result.col_widths[:5]}..." if len(result.col_widths) > 5 else f"  {result.col_widths}")
    else:
        print(f"[오류] {result.error}")
