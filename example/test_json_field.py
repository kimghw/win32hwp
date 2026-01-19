# JSON 필드 이름 테스트
# 필드 이름: {"table":"표이름","coord":[row,col]}
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor import get_hwp_instance
from table import TableField

def p(msg):
    print(msg, flush=True)

# 한글 연결
p("[1] 한글 연결...")
hwp = get_hwp_instance()
if not hwp:
    p("한글이 실행 중이 아닙니다.")
    sys.exit(1)

# 테이블 진입
p("[2] 테이블 진입...")
tf = TableField(hwp, debug=True)

if not tf.enter_table(0):
    p("테이블이 없습니다.")
    sys.exit(1)

size = tf.get_table_size()
p(f"  크기: {size['rows']}행 x {size['cols']}열")

# JSON 필드 설정
p("\n[3] JSON 필드 이름 설정...")
result = tf.set_table_fields_json()

p(f"\n  표 이름: {result['table']}")
p(f"  필드 개수: {len(result['fields'])}개")
p("  예시:")
for f in result['fields'][:3]:
    p(f"    {f}")

# JSON 필드 조회
p("\n[4] JSON 필드 조회...")
fields = tf.get_json_fields()
p(f"  조회된 필드: {len(fields)}개")
for f in fields[:3]:
    p(f"    table={f['table']}, coord={f['coord']}")

p("\n완료!")
