# 글자 너비 측정 - MovePos 2번으로 화면 좌표 비교
import sys
sys.stdout.reconfigure(encoding='utf-8')

from cursor_utils import get_hwp_instance

hwp = get_hwp_instance()
if not hwp:
    print("han-gul not running")
    exit()

# MovePos 상수
MOVE_DOC_BEGIN = 2
MOVE_DOC_END = 3
MOVE_LINE_BEGIN = 4
MOVE_LINE_END = 5
MOVE_RIGHT = 8
MOVE_LEFT = 9

def get_font_size_pt():
    """현재 글꼴 크기(pt) 반환"""
    hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
    height = hwp.HParameterSet.HCharShape.Height
    if height == 0:
        return 10.0
    return height / 100

def measure_width_by_cursor_move(text):
    """텍스트 삽입 후 커서 이동으로 너비 측정

    방법: 텍스트 삽입 -> 줄 시작으로 -> 오른쪽 이동하며 char_pos 비교
    """
    # 현재 위치 저장
    saved_pos = hwp.GetPos()

    # 텍스트 삽입
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

    # 줄 시작으로 이동
    hwp.MovePos(MOVE_LINE_BEGIN, 0, 0)

    # 각 문자별 오른쪽 이동하며 위치 변화 확인
    positions = []
    for i in range(len(text) + 1):
        pos = hwp.GetPos()
        positions.append(pos[2])  # char_pos
        if i < len(text):
            hwp.MovePos(MOVE_RIGHT, 0, 0)

    # 삭제하고 원위치
    hwp.SetPos(*saved_pos)
    hwp.HAction.Run("SelectAll")
    hwp.HAction.Run("Delete")
    hwp.SetPos(*saved_pos)

    return positions

# 테스트
print("=== Font size ===")
font_size = get_font_size_pt()
print(f"  {font_size} pt")
print()

# 테스트 문자열
print("=== Test: cursor positions after each char ===")
test_text = "□ A가"
positions = measure_width_by_cursor_move(test_text)
print(f"  Text: '{test_text}'")
print(f"  Positions: {positions}")
print(f"  Diffs: {[positions[i+1]-positions[i] for i in range(len(positions)-1)]}")

print()
print("=== Prefix width measurement ===")

# 새 문서에서 테스트를 위해 기존 내용 지우고 시작
prefixes = ["□ ", "○ ", "- "]

for prefix in prefixes:
    # 삽입
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = prefix + "TEST"
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

    # 줄 시작으로
    hwp.MovePos(MOVE_LINE_BEGIN, 0, 0)

    # prefix 길이만큼 이동하며 측정
    char_pos_list = []
    for i in range(len(prefix) + 1):
        pos = hwp.GetPos()
        char_pos_list.append(pos[2])
        hwp.MovePos(MOVE_RIGHT, 0, 0)

    diffs = [char_pos_list[i+1] - char_pos_list[i] for i in range(len(char_pos_list)-1)]

    print(f"  '{prefix}': char_pos={char_pos_list[:len(prefix)+1]}, diffs={diffs[:len(prefix)]}")

    # 줄 삭제
    hwp.MovePos(MOVE_LINE_BEGIN, 0, 0)
    hwp.HAction.Run("SelectAll")
    hwp.HAction.Run("Delete")

print()
print("=== Result ===")
print("char_pos diff shows character count, not pixel width.")
print("To get actual pt width, use font_size * (fullwidth=1, halfwidth=0.5)")
print()
print("For 10pt font, recommended hanging:")
print("  □  (box + space): 10 + 5 = 15 pt")
print("  ○  (circle + space): 10 + 5 = 15 pt")
print("  -  (dash + space): 5 + 5 = 10 pt")

