# -*- coding: utf-8 -*-
"""HWP 테이블 ↔ 엑셀 변환 모듈"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from dataclasses import dataclass
from typing import Dict, List, Optional, Any
from table.cell_position import CellPositionCalculator, CellPositionResult, CellRange

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.utils import get_column_letter
    from openpyxl.styles import Alignment, Border, Side, Font
    from openpyxl.cell.cell import MergedCell
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


@dataclass
class CellData:
    """셀 데이터"""
    list_id: int
    row: int
    col: int
    end_row: int
    end_col: int
    text: str
    rowspan: int = 1
    colspan: int = 1


class TableExcelConverter:
    """HWP 테이블 ↔ 엑셀 변환기"""

    def __init__(self, hwp, debug: bool = False):
        self.hwp = hwp
        self.debug = debug
        self._calc = CellPositionCalculator(hwp, debug=debug)

    def _get_cell_text(self, list_id: int) -> str:
        """셀의 텍스트 내용 가져오기"""
        try:
            self.hwp.SetPos(list_id, 0, 0)

            # 셀 시작으로 이동
            self.hwp.MovePos(4, 0, 0)  # moveTopOfList

            # 셀 끝까지 선택
            self.hwp.HAction.Run("MoveSelBottomOfList")

            # 선택된 텍스트 가져오기
            text = self.hwp.GetTextFile("TEXT", "")

            # 선택 해제
            self.hwp.HAction.Run("Cancel")

            # 텍스트 정리
            if text:
                text = text.strip().replace('\r\n', ' ').replace('\r', ' ').replace('\n', ' ')
            return text or ""
        except Exception as e:
            if self.debug:
                print(f"[경고] list_id={list_id} 텍스트 추출 실패: {e}")
            return ""

    def extract_table_data(self, max_cells: int = 1000) -> List[CellData]:
        """현재 테이블의 모든 셀 데이터 추출"""
        result = self._calc.calculate()

        cells_data = []
        count = 0
        for list_id, cell in result.cells.items():
            if count >= max_cells:
                print(f"[경고] 최대 셀 수({max_cells}) 도달, 중단")
                break

            text = self._get_cell_text(list_id)
            cells_data.append(CellData(
                list_id=list_id,
                row=cell.start_row,
                col=cell.start_col,
                end_row=cell.end_row,
                end_col=cell.end_col,
                text=text,
                rowspan=cell.rowspan,
                colspan=cell.colspan,
            ))
            count += 1

            if self.debug and count % 50 == 0:
                print(f"  {count}개 셀 처리됨...")

        return cells_data

    def to_excel(self, filepath: str, sheet_name: str = "Sheet1", with_text: bool = True):
        """HWP 테이블을 엑셀 파일로 저장

        Args:
            filepath: 저장할 엑셀 파일 경로
            sheet_name: 시트 이름
            with_text: True면 셀 텍스트 포함, False면 셀 범위만 저장
        """
        if not HAS_OPENPYXL:
            raise ImportError("openpyxl이 설치되어 있지 않습니다. pip install openpyxl")

        # 셀 위치 계산
        result = self._calc.calculate()

        # 텍스트 포함 여부에 따라 데이터 준비
        if with_text:
            cells_data = self.extract_table_data()
        else:
            # 텍스트 없이 셀 범위만
            cells_data = [
                CellData(
                    list_id=cell.list_id,
                    row=cell.start_row,
                    col=cell.start_col,
                    end_row=cell.end_row,
                    end_col=cell.end_col,
                    text="",
                    rowspan=cell.rowspan,
                    colspan=cell.colspan,
                )
                for cell in result.cells.values()
            ]

        # 워크북 생성
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name

        # 테두리 스타일
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        # 먼저 모든 병합 처리
        for cell_data in cells_data:
            if cell_data.rowspan > 1 or cell_data.colspan > 1:
                excel_row = cell_data.row + 1
                excel_col = cell_data.col + 1
                end_row = cell_data.end_row + 1
                end_col = cell_data.end_col + 1
                ws.merge_cells(
                    start_row=excel_row,
                    start_column=excel_col,
                    end_row=end_row,
                    end_column=end_col
                )

        # 셀 데이터 쓰기 (병합 후 좌상단 셀에만 값 입력)
        for cell_data in cells_data:
            excel_row = cell_data.row + 1
            excel_col = cell_data.col + 1

            cell = ws.cell(row=excel_row, column=excel_col)

            # MergedCell인 경우 건너뛰기 (이미 다른 셀에 병합됨)
            if isinstance(cell, MergedCell):
                continue

            cell.value = cell_data.text
            cell.border = thin_border
            cell.alignment = Alignment(wrap_text=True, vertical='center')

            # 병합 영역 테두리 적용
            if cell_data.rowspan > 1 or cell_data.colspan > 1:
                end_row = cell_data.end_row + 1
                end_col = cell_data.end_col + 1
                for r in range(excel_row, end_row + 1):
                    for c in range(excel_col, end_col + 1):
                        try:
                            ws.cell(row=r, column=c).border = thin_border
                        except AttributeError:
                            pass  # MergedCell은 건너뜀

        # 열 너비 자동 조정
        for col_idx in range(1, result.max_col + 2):
            max_length = 0
            column_letter = get_column_letter(col_idx)
            for row_idx in range(1, result.max_row + 2):
                cell = ws.cell(row=row_idx, column=col_idx)
                if cell.value:
                    # 한글은 2배 너비로 계산
                    cell_length = sum(2 if ord(c) > 127 else 1 for c in str(cell.value))
                    max_length = max(max_length, cell_length)
            ws.column_dimensions[column_letter].width = min(max_length + 2, 50)

        wb.save(filepath)
        print(f"엑셀 파일 저장: {filepath}")
        return filepath

    def to_dict(self) -> Dict[tuple, str]:
        """테이블을 딕셔너리로 변환 {(row, col): text}"""
        cells_data = self.extract_table_data()
        result = {}

        for cell_data in cells_data:
            # 병합 셀은 모든 좌표에 같은 텍스트
            for r in range(cell_data.row, cell_data.end_row + 1):
                for c in range(cell_data.col, cell_data.end_col + 1):
                    result[(r, c)] = cell_data.text

        return result

    def to_2d_array(self) -> List[List[str]]:
        """테이블을 2차원 배열로 변환"""
        result = self._calc.calculate()
        cells_data = self.extract_table_data()

        # 빈 2D 배열 생성
        rows = result.max_row + 1
        cols = result.max_col + 1
        array = [["" for _ in range(cols)] for _ in range(rows)]

        # 데이터 채우기
        for cell_data in cells_data:
            for r in range(cell_data.row, cell_data.end_row + 1):
                for c in range(cell_data.col, cell_data.end_col + 1):
                    array[r][c] = cell_data.text

        return array

    def print_table(self):
        """테이블 내용을 콘솔에 출력"""
        array = self.to_2d_array()
        result = self._calc.calculate()

        print(f"\n=== 테이블 내용 ({result.max_row + 1}행 x {result.max_col + 1}열) ===")
        for row_idx, row in enumerate(array):
            row_str = " | ".join(cell[:10].ljust(10) if cell else "".ljust(10) for cell in row)
            print(f"[{row_idx:2d}] {row_str}")


# 직접 실행 시
if __name__ == "__main__":
    from cursor import get_hwp_instance

    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        exit(1)

    converter = TableExcelConverter(hwp)

    try:
        # 셀 위치만 계산 (텍스트 추출 없음)
        result = converter._calc.calculate()
        converter._calc.print_summary(result)

        # 엑셀로 저장 (텍스트 없이 셀 범위만)
        if HAS_OPENPYXL:
            converter.to_excel("C:\\win32hwp\\output_table.xlsx", with_text=False)
        else:
            print("[경고] openpyxl이 없어서 엑셀 저장을 건너뜁니다.")
            print("설치: pip install openpyxl")

    except ValueError as e:
        print(f"[오류] {e}")
