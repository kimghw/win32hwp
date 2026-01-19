# 구조화된 필드 이름 테스트: {캡션}_{영역}_{row}_{col}
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

# TableField 사용
p("[2] 테이블 진입...")
tf = TableField(hwp, debug=False)

if not tf.enter_table(0):
    p("테이블이 없습니다.")
    sys.exit(1)

size = tf.get_table_size()
p(f"  테이블 크기: {size['rows']}행 x {size['cols']}열")

# 구조화된 필드 설정 (헤더 1행, 푸터 1행)
p("\n[3] 구조화된 필드 설정...")
p("  - 캡션: 자동 추출 (없으면 TBL)")
p("  - 헤더: 1행")
p("  - 푸터: 1행")

result = tf.set_structured_field_values(
    caption=None,      # 자동 추출
    header_rows=1,     # 첫 1행 = head
    footer_rows=1,     # 마지막 1행 = foot
    show_coords=True   # 좌표 값 입력
)

p("\n[4] 결과:")
p(f"  head 필드: {len(result['head'])}개")
for f in result['head']:
    p(f"    - {f}")

p(f"  body 필드: {len(result['body'])}개")
for f in result['body']:
    p(f"    - {f}")

p(f"  foot 필드: {len(result['foot'])}개")
for f in result['foot']:
    p(f"    - {f}")

p("\n완료! 한글 문서를 확인하세요.")
