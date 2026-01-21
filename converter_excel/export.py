# -*- coding: utf-8 -*-
"""HWP 표 → Excel 변환 메인 실행 파일

사용법:
    cmd.exe /c "cd /d C:\\win32hwp && python converter_excel\\export.py"
    cmd.exe /c "cd /d C:\\win32hwp && python converter_excel\\export.py --config my_config.yaml"
    cmd.exe /c "cd /d C:\\win32hwp && python converter_excel\\export.py --output result.xlsx"
"""

import sys
import os
import argparse

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor import get_hwp_instance
from table.table_info import TableInfo
from table.cell_position import CellPositionCalculator

try:
    from openpyxl import Workbook
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

from converter_excel.config import load_config, get_default_config, ExportConfig
from converter_excel.extract_data_hwp import (
    HwpExtractedData,
    extract_hwp_data,
    extract_page_info,
    set_cell_field_names,
)
from converter_excel.apply_excel import (
    apply_to_excel,
    create_main_sheet,
    write_page_info_to_sheet,
    write_cell_styles_to_sheet,
    write_row_col_sizes_to_sheet,
)


class HwpToExcelExporter:
    """HWP 표 → Excel 변환기

    추출과 엑셀 적용이 분리된 구조:
    - extract_data(): 한글에서 데이터만 추출
    - apply_to_excel(): 추출된 데이터를 엑셀로 저장
    - export(): 기존 호환성 유지 (추출 + 저장)
    """

    def __init__(self, config: ExportConfig = None):
        self.config = config or get_default_config()
        self.hwp = None
        self.table_info = None
        self._extracted_data = None  # 추출된 데이터 캐시

    def connect(self) -> bool:
        """한글 인스턴스 연결"""
        self.hwp = get_hwp_instance()
        if not self.hwp:
            print("[오류] 한글이 실행 중이지 않습니다.")
            return False

        self.table_info = TableInfo(self.hwp, debug=False)
        if not self.table_info.is_in_table():
            print("[오류] 커서가 표 안에 있지 않습니다.")
            return False

        return True

    def extract_data(self) -> HwpExtractedData:
        """한글에서 데이터만 추출 (엑셀 생성 없음)

        Returns:
            HwpExtractedData 객체
        """
        if not self.hwp:
            if not self.connect():
                return HwpExtractedData(success=False, error="한글 연결 실패")

        cfg = self.config
        print("=" * 60)
        print("HWP 데이터 추출")
        print("=" * 60)

        # CellPositionCalculator로 셀 위치 계산
        print("\n[1] 셀 위치 계산...")
        calc = CellPositionCalculator(self.hwp, debug=False)
        calc_result = calc.calculate()

        print(f"    행 수: {len(calc_result.y_levels) - 1}")
        print(f"    열 수: {len(calc_result.x_levels) - 1}")
        print(f"    셀 수: {len(calc_result.cells)}")

        # extract_hwp_data로 통합 추출
        print("\n[2] 데이터 추출...")
        result = extract_hwp_data(
            self.hwp,
            calc_result=calc_result,
            config=cfg,
            include_page_info=True,
            include_cell_styles=True,
            include_fields=cfg.field.enabled,
            tolerance=cfg.field.naming.pattern.tolerance if cfg.field.enabled else 50,
        )

        if result.success:
            colored_count = sum(1 for c in result.cells_data if c.bg_color_rgb)
            print(f"    셀 수: {len(result.cells_data)}개 (배경색 있는 셀: {colored_count}개)")
            print(f"    필드 수: {len(result.fields)}개")

            if result.page_result and result.page_result.success:
                meta = result.page_result.page_meta
                print(f"    용지: {meta.page_size.width/2834.6:.1f} x {meta.page_size.height/2834.6:.1f} cm ({meta.page_size.orientation})")

            # 한글 셀에 필드 설정 (설정에 따라)
            if cfg.field.enabled and cfg.field.set_hwp_field and result.fields:
                print("\n[3] 한글 셀 필드 설정...")
                success_count = set_cell_field_names(self.hwp, result.fields)
                print(f"    설정 완료: {success_count}/{len(result.fields)}개")
        else:
            print(f"    [오류] {result.error}")

        self._extracted_data = result
        return result

    def save_to_excel(self, output_path: str, sheet_name: str = "표",
                      extracted_data: HwpExtractedData = None) -> bool:
        """추출된 데이터를 엑셀로 저장

        Args:
            output_path: 출력 파일 경로
            sheet_name: 메인 시트 이름
            extracted_data: HwpExtractedData 객체 (None이면 캐시된 데이터 사용)

        Returns:
            성공 여부
        """
        if not HAS_OPENPYXL:
            print("[오류] openpyxl이 필요합니다: pip install openpyxl")
            return False

        data = extracted_data or self._extracted_data
        if not data:
            print("[오류] 추출된 데이터가 없습니다. extract_data()를 먼저 호출하세요.")
            return False

        if not data.success:
            print(f"[오류] 추출 실패: {data.error}")
            return False

        print("\n" + "=" * 60)
        print("Excel 파일 생성")
        print("=" * 60)

        # apply_to_excel 호출
        return apply_to_excel(
            extracted_data={
                'cells': data.cells_data,
                'row_heights': data.row_heights,
                'col_widths': data.col_widths,
                'page_result': data.page_result,
                'fields': data.fields,
            },
            output_path=output_path,
            sheet_name=sheet_name,
            config=self.config,
        )

    def export(self, output_path: str, sheet_name: str = "표") -> bool:
        """변환 실행 (기존 호환성 유지)

        Args:
            output_path: 출력 파일 경로
            sheet_name: 메인 시트 이름

        Returns:
            성공 여부
        """
        # 1. 데이터 추출
        result = self.extract_data()
        if not result.success:
            return False

        # 2. 엑셀 저장
        return self.save_to_excel(output_path, sheet_name, result)


# =============================================================================
# 메인 실행
# =============================================================================

def main():
    parser = argparse.ArgumentParser(description='HWP 표 → Excel 변환')
    parser.add_argument('--config', '-c', type=str, default=None,
                       help='설정 파일 경로 (YAML)')
    parser.add_argument('--output', '-o', type=str, default='C:\\win32hwp\\export_test.xlsx',
                       help='출력 파일 경로')
    parser.add_argument('--sheet', '-s', type=str, default='표',
                       help='메인 시트 이름')

    args = parser.parse_args()

    # 설정 로드
    if args.config:
        print(f"설정 파일 로드: {args.config}")
        config = load_config(args.config)
    else:
        config = get_default_config()

    # 변환 실행
    exporter = HwpToExcelExporter(config)
    exporter.export(args.output, args.sheet)


if __name__ == "__main__":
    main()
