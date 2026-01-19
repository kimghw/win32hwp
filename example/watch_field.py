# 셀 클릭 시 필드명 출력 (폴링 방식)
import time
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor import get_hwp_instance

def watch_field():
    """커서 위치 변경 시 필드명 출력"""
    hwp = get_hwp_instance()
    if not hwp:
        print("한글이 실행 중이 아닙니다.")
        return

    print("=" * 50)
    print("셀 필드 모니터링 시작")
    print("셀을 클릭하면 필드명이 출력됩니다.")
    print("종료: Ctrl+C")
    print("=" * 50)

    last_pos = None
    last_field = None

    try:
        while True:
            try:
                # 현재 위치
                pos = hwp.GetPos()
                list_id, para_id, char_pos = pos[0], pos[1], pos[2]

                # 위치가 변경되었을 때만 처리
                if pos != last_pos:
                    last_pos = pos

                    # 현재 필드명 확인
                    field_name = hwp.GetCurFieldName(1)  # option=1 (셀 필드)

                    # 필드 상태 확인
                    state = hwp.GetCurFieldState
                    field_type_map = {0: "없음", 1: "셀", 2: "누름틀"}
                    field_type = field_type_map.get(state & 0x0F, "알수없음")
                    has_name = bool(state & 0x10)

                    if field_name and field_name != last_field:
                        last_field = field_name
                        # 필드 텍스트도 가져오기
                        field_text = hwp.GetFieldText(field_name)
                        field_text = field_text.strip(chr(2)) if field_text else ""

                        print(f"\n[필드 감지]")
                        print(f"  필드명: {field_name}")
                        print(f"  타입: {field_type}")
                        print(f"  값: '{field_text}'")
                        print(f"  위치: list_id={list_id}, para={para_id}, pos={char_pos}")
                    elif not field_name and last_field:
                        last_field = None
                        print(f"\n[필드 영역 벗어남] list_id={list_id}")

            except Exception as e:
                pass  # 에러 무시 (문서 전환 등)

            time.sleep(0.2)  # 200ms 간격으로 체크

    except KeyboardInterrupt:
        print("\n\n모니터링 종료")


if __name__ == "__main__":
    watch_field()
