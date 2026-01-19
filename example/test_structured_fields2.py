# 구조화된 필드 이름 테스트 (디버그)
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor import get_hwp_instance
from table import TableField

def p(msg):
    print(msg, flush=True)

# ROT에서 한글 연결
p("[1] 한글 연결...")
hwp = get_hwp_instance()
if not hwp:
    p("한글이 실행 중이 아닙니다.")
    sys.exit(1)

# TableField 사용 (디버그 모드)
p("[2] 테이블 진입...")
tf = TableField(hwp, debug=True)

if not tf.enter_table(0):
    p("테이블이 없습니다.")
    sys.exit(1)

p(f"\n[3] 테이블 정보:")
size = tf.get_table_size()
p(f"  크기: {size['rows']}행 x {size['cols']}열")
p(f"  cells 개수: {len(tf.table_info.cells)}")
p(f"  coord_map 개수: {len(tf._coord_map)}")
p(f"  representative_coords 개수: {len(tf.table_info._representative_coords)}")

# 캡션 확인
caption = tf.table_info.get_table_caption()
p(f"  캡션: '{caption}'")

# 셀 목록
p("\n[4] 셀 목록:")
for list_id in sorted(tf.table_info.cells.keys())[:5]:  # 처음 5개만
    rep = tf.table_info.get_representative_coord(list_id)
    p(f"  list_id={list_id}, rep_coord={rep}")

# 구조화된 필드 설정
p("\n[5] 구조화된 필드 설정...")
result = tf.set_structured_field_names(
    caption="TEST",    # 명시적 캡션
    header_rows=1,
    footer_rows=1
)

p(f"\n[6] 결과:")
p(f"  head: {result['head']}")
p(f"  body: {result['body']}")
p(f"  foot: {result['foot']}")
