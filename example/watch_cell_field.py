# 셀 클릭 시 셀 필드명 출력
import time
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor import get_hwp_instance

hwp = get_hwp_instance()
if not hwp:
    print("한글이 실행 중이 아닙니다.")
    sys.exit(1)

print("=" * 50)
print("셀 필드 모니터링")
print("테이블 셀을 클릭하면 필드명이 출력됩니다.")
print("종료: Ctrl+C")
print("=" * 50)

last_list_id = None

try:
    while True:
        pos = hwp.GetPos()
        list_id = pos[0]

        if list_id != last_list_id:
            last_list_id = list_id

            # 셀 필드명 확인 (option=1: 셀 필드)
            field_name = hwp.GetCurFieldName(1)

            if field_name:
                text = hwp.GetFieldText(field_name)
                text = text.strip(chr(2)) if text else ""
                print(f"셀 필드: {field_name} = '{text}'  (list_id={list_id})")
            elif list_id > 1:  # 테이블 내부 (list_id > 1)
                print(f"(필드 없는 셀) list_id={list_id}")

        time.sleep(0.15)

except KeyboardInterrupt:
    print("\n종료")
