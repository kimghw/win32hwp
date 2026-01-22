"""xend 기반 check_boundary_table 테스트"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from table.table_boundary import TableBoundary
from cursor import get_hwp_instance

hwp = get_hwp_instance()
if not hwp:
    print("한글이 실행 중이 아닙니다")
    exit(1)

b = TableBoundary(hwp, debug=True)
r = b.check_boundary_table()

print()
print("=" * 50)
print("=== 결과 ===")
print(f"first_cols ({len(r.first_cols)}): {r.first_cols}")
print(f"last_cols  ({len(r.last_cols)}): {r.last_cols}")
