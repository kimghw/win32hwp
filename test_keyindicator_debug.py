"""
KeyIndicator() 반환값 상세 디버깅 테스트
실제로 어떤 값이 어느 인덱스에 들어오는지 확인
"""

from cursor_utils import get_hwp_instance
import time

def test_keyindicator():
    hwp = get_hwp_instance()
    if not hwp:
        print("한글이 실행되지 않았습니다.")
        return

    hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')

    print("=" * 60)
    print("KeyIndicator() 반환값 디버깅")
    print("=" * 60)
    print()
    print("한글 화면 하단의 상태바를 확인하세요.")
    print("예: '쪽 1/1  줄 3  칸 15' 같은 형식")
    print()
    print("커서를 여러 위치로 이동하면서 비교해보세요.")
    print("Ctrl+C로 종료")
    print("=" * 60)
    print()

    try:
        while True:
            # KeyIndicator 호출
            key = hwp.KeyIndicator()

            # GetPos도 함께 호출
            pos = hwp.GetPos()

            print("\r" + " " * 100, end="")  # 이전 출력 지우기

            # KeyIndicator 전체 출력
            print(f"\rKeyIndicator(): {key}", end="")
            print(f" | GetPos(): list={pos[0]}, para={pos[1]}, pos={pos[2]}", end="")

            # 각 인덱스별 상세 출력
            print("\n")
            print(f"  key[0] = {key[0]:3d}  (총 구역?)")
            print(f"  key[1] = {key[1]:3d}  (현재 구역?)")
            print(f"  key[2] = {key[2]:3d}  (페이지?)")
            print(f"  key[3] = {key[3]:3d}  (단?)")
            print(f"  key[4] = {key[4]:3d}  (줄?)")
            print(f"  key[5] = {key[5]:3d}  (칸?)")
            if len(key) > 6:
                print(f"  key[6] = {key[6]:3d}  (수정/삽입 모드?)")
            if len(key) > 7:
                print(f"  key[7] = {key[7]}  (컨트롤 이름?)")

            print("\n상태바와 비교: 쪽=?, 줄=?, 칸=?")
            print("(Ctrl+C로 종료)", end="")

            time.sleep(0.5)

            # 화면 지우기 (10줄 위로)
            for _ in range(12):
                print("\033[F\033[K", end="")

    except KeyboardInterrupt:
        print("\n\n종료되었습니다.")

if __name__ == "__main__":
    test_keyindicator()
