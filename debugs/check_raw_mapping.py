# -*- coding: utf-8 -*-
"""원본 그리드 매핑 확인 (_fix 전)"""
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

# _find_level_index 테스트
y_levels = [0, 2112, 4224, 6270]  # 예시

def find_level_index(value, levels, tolerance=3):
    for idx, level in enumerate(levels):
        if abs(value - level) <= tolerance:
            return idx
    return -1

def find_end_level_index(value, levels, tolerance=3):
    for idx, level in enumerate(levels):
        if abs(value - level) <= tolerance:
            return max(0, idx - 1)
    for idx in range(len(levels) - 1, -1, -1):
        if levels[idx] < value - tolerance:
            return idx
    return 0

print("=== _find_level_index 테스트 ===")
test_cases = [
    (0, "list_id=2 start_y"),
    (2112, "list_id=2 end_y / list_id=7 start_y"),
    (4224, "list_id=7 end_y"),
]

for y, desc in test_cases:
    start_idx = find_level_index(y, y_levels)
    end_idx = find_end_level_index(y, y_levels)
    print(f"y={y} ({desc})")
    print(f"  find_level_index → {start_idx}")
    print(f"  find_end_level_index → {end_idx}")
    print()

print("=== 예상 그리드 매핑 ===")
print("list_id=2: start_y=0→row0, end_y=2112→row0")
print("list_id=7: start_y=2112→row1, end_y=4224→row1")
