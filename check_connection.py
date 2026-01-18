# 현재 연결된 HWP 문서 확인
from cursor_utils import get_hwp_instance

hwp = get_hwp_instance()
if hwp:
    pos = hwp.GetPos()
    print(f"연결 성공 - para_id={pos[1]}, char_pos={pos[2]}")

    # 현재 문단 텍스트 확인
    hwp.HAction.Run("MoveParaBegin")
    hwp.HAction.Run("MoveSelParaEnd")
    text = hwp.GetTextFile("TEXT", "saveblock")
    hwp.HAction.Run("Cancel")

    if text:
        print(f"현재 문단: {text[:80]}...")
    else:
        print("현재 문단: (텍스트 없음)")
else:
    print("HWP 연결 실패")
