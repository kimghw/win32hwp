# -*- coding: utf-8 -*-
"""병합 충돌 분석 스크립트"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from table.table_excel_converter import TableExcelConverter
from cursor import get_hwp_instance

hwp = get_hwp_instance()
if not hwp:
    print("[오류] 한글이 실행 중이지 않습니다.")
    exit(1)

converter = TableExcelConverter(hwp, debug=False)
result = converter._calc.calculate()

print("=== 셀 위치 검증 ===")
validation = converter.validate_cell_positions(result)

print("\n=== 충돌 상세 분석 ===")

# 좌표별로 어떤 셀들이 겹치는지 확인
coord_cells = {}
for lid, cell in result.cells.items():
    for r in range(cell.start_row, cell.end_row + 1):
        for c in range(cell.start_col, cell.end_col + 1):
            coord = (r, c)
            if coord not in coord_cells:
                coord_cells[coord] = []
            coord_cells[coord].append({
                'list_id': lid,
                'range': f"({cell.start_row},{cell.start_col})~({cell.end_row},{cell.end_col})",
                'phys': f"x={cell.start_x}~{cell.end_x}, y={cell.start_y}~{cell.end_y}"
            })

# 충돌 좌표 출력
conflicts = [(coord, cells) for coord, cells in coord_cells.items() if len(cells) > 1]
print(f"총 {len(conflicts)}개 좌표에서 충돌 발생")

for coord, cells in conflicts[:10]:
    print(f"\n좌표 {coord}:")
    for c in cells:
        print(f"  list_id={c['list_id']}: {c['range']}")
        print(f"    {c['phys']}")
