# 현재 열린 한글 파일의 테이블에 좌표 필드 설정
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor import get_hwp_instance
from table import TableInfo, TableField

def p(msg):
    print(msg, flush=True)

# ROT에서 한글 연결
p("[1] 한글 연결...")
hwp = get_hwp_instance()
if not hwp:
    p("한글이 실행 중이 아닙니다.")
    sys.exit(1)

# 테이블 정보 수집
p("[2] 테이블 분석...")
ti = TableInfo(hwp, debug=False)

if not ti.enter_table(0):
    p("테이블이 없습니다.")
    sys.exit(1)

ti.collect_cells_bfs()
coord_map = ti.build_coordinate_map()

size = ti.get_table_size()
p(f"  테이블 크기: {size['rows']}행 x {size['cols']}열")
p(f"  셀 개수: {len(ti.cells)}개")

# 각 셀에 필드 설정 및 좌표 값 입력
p("\n[3] 셀 필드 설정 및 좌표 입력...")

for list_id in sorted(ti.cells.keys()):
    # 대표 좌표 가져오기
    rep_coord = ti.get_representative_coord(list_id)
    if rep_coord is None:
        continue

    row, col = rep_coord
    all_coords = ti.get_cell_coords(list_id)

    # 필드 이름 생성
    field_name = f"cell_{row}_{col}"

    # 좌표 문자열 (병합 셀은 범위 표시)
    if len(all_coords) > 1:
        min_r = min(r for r, c in all_coords)
        max_r = max(r for r, c in all_coords)
        min_c = min(c for r, c in all_coords)
        max_c = max(c for r, c in all_coords)
        coord_str = f"({min_r},{min_c})~({max_r},{max_c})"
    else:
        coord_str = f"({row},{col})"

    # 셀로 이동하여 필드 설정
    hwp.SetPos(list_id, 0, 0)
    hwp.HAction.Run('TableCellBlock')
    result = hwp.SetCurFieldName(field_name, 1, '', '')
    hwp.HAction.Run('Cancel')

    # 필드 값으로 좌표 입력
    hwp.PutFieldText(field_name, coord_str)

    p(f"  {field_name} = '{coord_str}'")

# 결과 확인
p("\n[4] 필드 목록 확인...")
field_list = hwp.GetFieldList(1, 1)  # 셀 필드만
if field_list:
    fields = field_list.split(chr(2))
    p(f"  총 {len([f for f in fields if f])}개 필드 생성됨")

p("\n완료! 한글 문서를 확인하세요.")
