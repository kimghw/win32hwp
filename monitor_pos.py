"""
실시간 커서 위치 모니터링
터미널에서 커서 위치 변화를 실시간으로 표시
"""
import time
import os
from cursor_utils import get_hwp_instance


def clear_screen():
    """터미널 화면 지우기"""
    os.system('cls' if os.name == 'nt' else 'clear')


def monitor_position(interval=0.1, show_raw=False):
    """
    커서 위치를 실시간으로 모니터링

    Args:
        interval: 갱신 주기 (초, 기본 0.1초)
        show_raw: True면 원본 튜플도 표시
    """
    hwp = get_hwp_instance()
    if not hwp:
        print("한글이 실행 중이 아닙니다")
        return

    # 보안 모듈 등록
    hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')

    print("실시간 커서 위치 모니터링 시작")
    print("종료: Ctrl+C\n")
    time.sleep(1)

    prev_pos = None
    prev_key = None
    first_run = True

    try:
        while True:
            # 현재 위치 가져오기
            pos = hwp.GetPos()
            key = hwp.KeyIndicator()

            # 위치가 변경되었을 때만 화면 갱신 (첫 실행은 무조건 표시)
            if first_run or pos != prev_pos or key != prev_key:
                if not first_run:
                    clear_screen()
                first_run = False

                print("=" * 70)
                print("실시간 커서 위치 모니터링 (Ctrl+C로 종료)")
                print("=" * 70)

                # GetPos
                print("\n[GetPos]")
                if show_raw:
                    print(f"  원본: {pos}")
                print(f"  list_id:  {pos[0]:3d}  ", end="")
                if pos[0] == 0:
                    print("(본문)")
                elif pos[0] >= 10:
                    print("(테이블/특수영역)")
                else:
                    print("(머리말/꼬리말 등)")

                print(f"  para_id:  {pos[1]:3d}")
                print(f"  char_pos: {pos[2]:3d}")

                # KeyIndicator
                print("\n[KeyIndicator]")
                print(f"  튜플 길이: {len(key)}")
                if show_raw:
                    print(f"  원본: {key}")
                print(f"  [0] total_section:   {key[0]:3d}")
                print(f"  [1] current_section: {key[1]:3d}")
                print(f"  [2] page:            {key[2]:3d}")
                print(f"  [3] column_num (단): {key[3]:3d}")
                print(f"  [4] line:            {key[4]:3d}")
                print(f"  [5] column (칸):     {key[5]:3d}")
                print(f"  [6] insert_mode:     {key[6]:3d} ({'수정' if key[6] else '삽입'})")
                if len(key) > 7:
                    print(f"  [7] ctrlname:        {key[7]}")

                # 테이블 확인
                try:
                    act = hwp.CreateAction("TableCellBorder")
                    pset = act.GetDefault("TableCellBorder", hwp.HParameterSet.HTableCellBorder.HSet)
                    in_table = act.Execute(pset)
                except:
                    in_table = False

                print("\n[위치 유형]")
                if in_table:
                    print("  테이블 내부")
                elif pos[0] == 0:
                    print("  본문")
                else:
                    print(f"  특수 영역 (list_id={pos[0]})")

                print("\n" + "=" * 70)
                print(f"갱신 주기: {interval}초")

                prev_pos = pos
                prev_key = key

            time.sleep(interval)

    except KeyboardInterrupt:
        print("\n\n모니터링을 종료합니다.")


if __name__ == "__main__":
    import sys

    # 명령줄 인수 처리
    show_raw = "--raw" in sys.argv or "-r" in sys.argv
    interval = 0.1

    for arg in sys.argv[1:]:
        if arg.startswith("--interval="):
            try:
                interval = float(arg.split("=")[1])
            except:
                pass
        elif arg.startswith("-i"):
            try:
                interval = float(arg[2:])
            except:
                pass

    monitor_position(interval=interval, show_raw=show_raw)
