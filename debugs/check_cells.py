# -*- coding: utf-8 -*-
"""셀 데이터 확인"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor import get_hwp_instance
from table.cell_position import CellPositionCalculator

hwp = get_hwp_instance()
if not hwp:
    print("한글이 실행 중이지 않습니다.")
    exit(1)

calc = CellPositionCalculator(hwp, debug=False)
result = calc.calculate()

print(f"그리드: {result.max_row + 1}행 x {result.max_col + 1}열")
print(f"x_levels: {result.x_levels[:10]}...")
print(f"y_levels: {result.y_levels[:10]}...")
print()

print("=== 처음 10개 셀 (list_id 순) ===")
for list_id, cell in sorted(result.cells.items())[:10]:
    print(f"list_id={list_id}: ({cell.start_row},{cell.start_col})~({cell.end_row},{cell.end_col})")
    print(f"  물리: x=[{cell.start_x}~{cell.end_x}], y=[{cell.start_y}~{cell.end_y}]")
