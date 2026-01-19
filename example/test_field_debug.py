# GetFieldText 디버그
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor import get_hwp_instance

hwp = get_hwp_instance()
if not hwp:
    print("한글이 실행 중이 아닙니다.")
    sys.exit(1)

# 필드 목록
field_list = hwp.GetFieldList(1, 1)  # 셀 필드
print(f"필드 목록: {repr(field_list)}")

if field_list:
    fields = field_list.split('\x02')
    print(f"필드 개수: {len(fields)}")
    for f in fields[:5]:
        print(f"  - {f}")

# GetFieldText 결과
text = hwp.GetFieldText("TBL")
print(f"\nGetFieldText('TBL'): {repr(text)}")

# 필드 목록으로 조회
text2 = hwp.GetFieldText(field_list)
print(f"\nGetFieldText(field_list): {repr(text2[:100])}...")
