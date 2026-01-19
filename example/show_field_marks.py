# 필드 표시 옵션 설정
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor import get_hwp_instance

hwp = get_hwp_instance()
if not hwp:
    print("한글이 실행 중이 아닙니다.")
    sys.exit(1)

# 필드 표시 옵션
# 1: 『』 숨김
# 2: 『』 빨간색 표시 (기본, 누름틀) / 흰색 (다른 필드)
# 3: 『』 흰색 표시

print("필드 표시 옵션:")
print("  1 = 숨김")
print("  2 = 빨간색 (기본)")
print("  3 = 흰색")
print()

# 현재 옵션 확인 (0 입력으로)
current = hwp.SetFieldViewOption(0)
print(f"현재 옵션: {current}")

# 빨간색으로 설정
result = hwp.SetFieldViewOption(2)
print(f"빨간색 표시로 변경: {result}")

# 필드 목록 출력
print("\n[문서 내 필드 목록]")
field_list = hwp.GetFieldList(1, 0)  # 모든 필드
if field_list:
    fields = field_list.split(chr(2))
    for f in fields:
        if f:
            text = hwp.GetFieldText(f)
            text = text.strip(chr(2)) if text else ""
            print(f"  {f}: '{text}'")
else:
    print("  필드 없음")

print("\n한글에서 『』 기호가 빨간색으로 보이면 필드입니다.")
print("메뉴: [보기] → [조판 부호] 체크하면 더 잘 보입니다.")
