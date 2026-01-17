"""행 끝에서 MovePos(MOVE_RIGHT_OF_CELL) 계속 테스트 - 끝까지"""

from cursor_utils import get_hwp_instance

MOVE_RIGHT_OF_CELL = 101

hwp = get_hwp_instance()
if not hwp:
    print("HWP 연결 실패")
    exit()

print("=== 오른쪽 이동 테스트 (False 나올때까지) ===\n")

for i in range(100):
    prev_list_id = hwp.GetPos()[0]
    result = hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
    new_list_id = hwp.GetPos()[0]
    print(f"{i+1:3d}. result={result}, list_id: {prev_list_id} -> {new_list_id}")

    if not result:
        print(f"\n>>> {i+1}번째에서 False 반환")
        break
else:
    print("\n>>> 100회 모두 True")
