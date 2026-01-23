# -*- coding: utf-8 -*-
"""테이블 그리드 시각화 (Pillow)"""

import sys
import os

_this_dir = os.path.dirname(os.path.abspath(__file__))
_parent_dir = os.path.dirname(_this_dir)
if _this_dir not in sys.path:
    sys.path.insert(0, _this_dir)
if _parent_dir not in sys.path:
    sys.path.insert(0, _parent_dir)

from PIL import Image, ImageDraw

from table_grid import TableGrid
from cursor import get_hwp_instance


# 색상 팔레트 (20개)
COLORS = [
    (255, 99, 71), (60, 179, 113), (65, 105, 225), (255, 215, 0),
    (238, 130, 238), (0, 206, 209), (255, 165, 0), (147, 112, 219),
    (144, 238, 144), (255, 182, 193), (135, 206, 235), (240, 128, 128),
    (152, 251, 152), (221, 160, 221), (250, 250, 210), (176, 224, 230),
    (255, 228, 181), (216, 191, 216), (175, 238, 238), (245, 222, 179)
]


def visualize_table_grid(mappings, excel_grid, output_path="table_grid.jpg"):
    """테이블 그리드를 시각화"""

    rows = len(excel_grid.cells)
    cols = len(excel_grid.cells[0]) if excel_grid.cells else 0

    cell_size = 25
    margin = 30
    width = cols * cell_size + margin * 2
    height = rows * cell_size + margin * 2

    img = Image.new('RGB', (width, height), 'white')
    draw = ImageDraw.Draw(img)

    # 그리드 셀 → 테이블 셀 매핑
    grid_to_table = {}
    for idx, m in enumerate(mappings):
        for (r, c) in m.grid_cells:
            grid_to_table[(r, c)] = (m.list_id, idx)

    # 각 그리드 셀 그리기
    for row_idx in range(rows):
        for col_idx in range(cols):
            x1 = margin + col_idx * cell_size
            y1 = margin + row_idx * cell_size
            x2 = x1 + cell_size
            y2 = y1 + cell_size

            if (row_idx, col_idx) in grid_to_table:
                list_id, color_idx = grid_to_table[(row_idx, col_idx)]
                color = COLORS[color_idx % len(COLORS)]
            else:
                color = (200, 200, 200)

            draw.rectangle([x1, y1, x2, y2], fill=color, outline='gray')

    # 테이블 셀 경계 (두꺼운 선) + list_id 표시
    for m in mappings:
        min_row, max_row = m.row_span
        min_col, max_col = m.col_span

        x1 = margin + min_col * cell_size
        y1 = margin + min_row * cell_size
        x2 = margin + (max_col + 1) * cell_size
        y2 = margin + (max_row + 1) * cell_size

        # 두꺼운 경계선
        for i in range(2):
            draw.rectangle([x1-i, y1-i, x2+i, y2+i], outline='black')

        # list_id 텍스트
        cx = (x1 + x2) // 2
        cy = (y1 + y2) // 2
        draw.text((cx, cy), str(m.list_id), fill='black', anchor='mm')

    # 축 레이블
    for c in range(0, cols, 4):
        draw.text((margin + c * cell_size + cell_size//2, 10), str(c), fill='black', anchor='mm')
    for r in range(0, rows, 4):
        draw.text((10, margin + r * cell_size + cell_size//2), str(r), fill='black', anchor='mm')

    img.save(output_path, quality=95)
    print(f"저장됨: {output_path}")


if __name__ == "__main__":
    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        exit(1)

    grid = TableGrid(hwp, debug=False)

    if not grid._boundary._is_in_table():
        print("[오류] 커서가 테이블 내부에 있지 않습니다.")
        exit(1)

    boundary = grid._boundary.check_boundary_table()
    result = grid.build_grid(boundary)
    excel_grid = grid.build_grid_lines(result, boundary)
    mappings = grid.map_cells_to_grid(result, excel_grid, tolerance=5)

    print(f"테이블 셀: {len(mappings)}개")
    print(f"그리드: {len(excel_grid.cells)}행 x {len(excel_grid.cells[0])}열")

    output_path = os.path.join(os.path.dirname(__file__), "table_grid.jpg")
    visualize_table_grid(mappings, excel_grid, output_path)
