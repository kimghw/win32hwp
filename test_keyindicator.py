"""
KeyIndicator 반환값 테스트
실제 값을 확인하여 인덱스 매핑 검증
"""
from cursor_utils import get_hwp_instance

hwp = get_hwp_instance()
if not hwp:
    print("한글이 실행 중이 아닙니다")
    exit()

hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')

key = hwp.KeyIndicator()

print("KeyIndicator() 원본 값:")
print(f"  {key}")
print()
print("인덱스별 값:")
for i, val in enumerate(key):
    print(f"  key[{i}] = {val}")
print()
print("문서 상 매핑:")
print("  key[0] = 총 구역 (seccnt)")
print("  key[1] = 현재 구역 (secno)")
print("  key[2] = 쪽 (prnpageno)")
print("  key[3] = 단 (colno)")
print("  key[4] = 줄 (line)")
print("  key[5] = 칸 (pos)")
print("  key[6] = 수정모드 (over)")
print()
print("한글 상태바에 표시된 값과 비교해보세요!")
