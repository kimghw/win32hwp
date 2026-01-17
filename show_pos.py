"""
현재 GetPos()와 KeyIndicator() 값을 간단히 출력
"""
from cursor_utils import get_hwp_instance


def show_pos(hwp=None):
    """
    현재 GetPos()와 KeyIndicator() 원본 값을 출력

    Args:
        hwp: HWP 인스턴스 (None이면 자동으로 가져옴)

    Returns:
        tuple: (pos, key) - GetPos()와 KeyIndicator() 반환값
    """
    if hwp is None:
        hwp = get_hwp_instance()
        if not hwp:
            print("한글이 실행 중이 아닙니다")
            return None, None

    # 보안 모듈 등록
    hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')

    # GetPos
    pos = hwp.GetPos()

    # KeyIndicator
    key = hwp.KeyIndicator()

    # 출력
    print("GetPos():")
    print(f"  {pos}")
    print(f"  [0] list_id:  {pos[0]}")
    print(f"  [1] para_id:  {pos[1]}")
    print(f"  [2] char_pos: {pos[2]}")
    print()

    print("KeyIndicator():")
    print(f"  {key}")
    print(f"  [0] total_section:   {key[0]}")
    print(f"  [1] current_section: {key[1]}")
    print(f"  [2] page:            {key[2]}")
    print(f"  [3] column_num (단): {key[3]}")
    print(f"  [4] line:            {key[4]}")
    print(f"  [5] column (칸):     {key[5]}")
    print(f"  [6] insert_mode:     {key[6]} ({'수정' if key[6] else '삽입'})")
    if len(key) > 7:
        print(f"  [7] ctrlname:        {key[7]}")

    return pos, key


if __name__ == "__main__":
    show_pos()
