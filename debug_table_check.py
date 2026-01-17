"""테이블 인식 방법 디버그"""
from cursor_utils import get_hwp_instance

hwp = get_hwp_instance()
if not hwp:
    print("한글이 실행 중이 아닙니다")
    exit()

print("=== 테이블 인식 디버그 ===\n")

# 1. GetPos로 list_id 확인
pos = hwp.GetPos()
print(f"1. GetPos(): list_id={pos[0]}, para_id={pos[1]}, char_pos={pos[2]}")

# 2. FindCtrl 확인
try:
    ctrl_id = hwp.FindCtrl()
    print(f"2. FindCtrl(): '{ctrl_id}'")
except Exception as e:
    print(f"2. FindCtrl() 오류: {e}")

# 3. KeyIndicator 확인
try:
    key = hwp.KeyIndicator()
    print(f"3. KeyIndicator(): {key}")
    if len(key) > 7:
        print(f"   ctrl_name (key[7]): '{key[7]}'")
except Exception as e:
    print(f"3. KeyIndicator() 오류: {e}")

# 4. MovePos 테스트 (셀 이동이 가능한지)
try:
    result = hwp.MovePos(100)  # MOVE_LEFT_OF_CELL
    print(f"4. MovePos(100) 좌측 셀 이동: {result}")
except Exception as e:
    print(f"4. MovePos(100) 오류: {e}")

# 5. 테이블 관련 액션 실행 가능 여부
try:
    # TableCellBlock 액션 시도
    result = hwp.HAction.Run("TableCellBlock")
    print(f"5. TableCellBlock 액션: {result}")
    hwp.HAction.Run("Cancel")  # 선택 해제
except Exception as e:
    print(f"5. TableCellBlock 오류: {e}")

print("\n=== 디버그 완료 ===")
