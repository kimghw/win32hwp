# -*- coding: utf-8 -*-
"""셀 매핑 디버그"""
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

print(f"y_levels: {result.y_levels}")
print()

# 문제 셀들 확인
for list_id in [2, 7, 12, 16, 21]:
    if list_id in result.cells:
        cell = result.cells[list_id]
        print(f"list_id={list_id}:")
        print(f"  grid: ({cell.start_row},{cell.start_col})~({cell.end_row},{cell.end_col})")
        print(f"  물리 y: [{cell.start_y}~{cell.end_y}]")

        # y_levels에서 인덱스 확인
        for i, y in enumerate(result.y_levels):
            if abs(y - cell.start_y) <= 3:
                print(f"  start_y={cell.start_y} → y_levels[{i}]={y}")
            if abs(y - cell.end_y) <= 3:
                print(f"  end_y={cell.end_y} → y_levels[{i}]={y}")
        print()
