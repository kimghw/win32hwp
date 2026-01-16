"""26번 문단만 text_align 처리 테스트"""

from cursor_position_monitor import get_hwp_instance
from text_align import TextAlign

hwp = get_hwp_instance()
if not hwp:
    print("[ERROR] 한글을 찾을 수 없습니다.")
    exit()

print("[OK] 한글 연결됨")

# 25번 문단으로 이동 (페이지 걸친 문단)
list_id = hwp.GetPos()[0]
hwp.SetPos(list_id, 25, 0)
hwp.HAction.Run("MoveParaBegin")

pos = hwp.GetPos()
print(f"현재 위치: para_id={pos[1]}, pos={pos[2]}")

# text_align 실행
align = TextAlign(hwp, debug=True)
result = align.align_paragraph(spacing_step=-1.0, min_spacing=-100)

print(f"\n결과: {result['message']}")
print(f"조정: {result['adjusted_lines']}, 건너뜀: {result['skipped_lines']}, 실패: {result['failed_lines']}")

# 로그 저장
json_path = align.save_debug_log(result)
txt_path = align.save_text_log(result)
print(f"\n로그 저장: {json_path}")
print(f"로그 저장: {txt_path}")
