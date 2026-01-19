"""
빈 문단 글자 크기 변경 디버그
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor_utils import get_hwp_instance
from separated_para import SeparatedPara


def debug_empty_para():
    hwp = get_hwp_instance()
    if not hwp:
        print("[ERROR] 한글이 실행 중이지 않습니다.")
        return

    print("=" * 60)
    print("빈 문단 글자 크기 변경 디버그")
    print("=" * 60)

    helper = SeparatedPara(hwp)

    # 1. para_page_map 조회
    helper.ParaAlignWords()
    print(f"\n[1] 전체 문단 정보:")
    for para_id, info in sorted(SeparatedPara.para_page_map.items()):
        empty_mark = " (빈)" if info.get('is_empty') else ""
        spanning = " [걸침]" if info['start_page'] != info['end_page'] else ""
        print(f"  para_id={para_id}: 페이지 {info['start_page']}-{info['end_page']}{empty_mark}{spanning}")

    # 2. 페이지 1로 고정 (걸친 문단이 있는 페이지)
    current_page = 1
    print(f"\n[2] 대상 페이지: {current_page}")

    # 3. 걸친 문단 찾기
    spanning_para = None
    for para_id, info in SeparatedPara.para_page_map.items():
        if info['start_page'] == current_page and info['start_page'] != info['end_page']:
            if not info.get('is_empty'):
                spanning_para = para_id
                break

    if not spanning_para:
        print("\n[3] 현재 페이지에 걸친 문단 없음")
        return

    print(f"\n[3] 걸친 문단: para_id={spanning_para}")

    # 4. 같은 페이지의 빈 문단 찾기
    empty_paras = []
    for pid, info in SeparatedPara.para_page_map.items():
        if info.get('is_empty') and info['start_page'] == current_page:
            if pid < spanning_para:
                empty_paras.append(pid)
    empty_paras.sort()

    print(f"\n[4] 같은 페이지의 빈 문단 (걸친 문단 앞): {empty_paras}")

    if not empty_paras:
        print("    빈 문단 없음!")
        return

    # 5. 각 빈 문단의 글자 크기 확인
    saved_pos = hwp.GetPos()
    list_id = saved_pos[0]

    print(f"\n[5] 빈 문단 글자 크기 확인:")
    for empty_para_id in empty_paras:
        hwp.SetPos(list_id, empty_para_id, 0)
        hwp.HAction.Run("MoveParaBegin")

        # 방법 1: MoveSelParaEnd
        hwp.HAction.Run("MoveSelParaEnd")
        pset = hwp.HParameterSet.HCharShape
        hwp.HAction.GetDefault("CharShape", pset.HSet)
        height1 = pset.Height
        hwp.HAction.Run("Cancel")

        # 방법 2: MoveSelNextParaBegin
        hwp.HAction.Run("MoveParaBegin")
        hwp.HAction.Run("MoveSelNextParaBegin")
        hwp.HAction.GetDefault("CharShape", pset.HSet)
        height2 = pset.Height
        hwp.HAction.Run("Cancel")

        # 방법 3: SelectAll (문단 내)
        hwp.HAction.Run("MoveParaBegin")
        hwp.HAction.Run("MoveRight")  # 한 글자(줄바꿈) 선택
        hwp.HAction.Run("MoveSelLeft")
        hwp.HAction.GetDefault("CharShape", pset.HSet)
        height3 = pset.Height
        hwp.HAction.Run("Cancel")

        print(f"  para_id={empty_para_id}:")
        print(f"    MoveSelParaEnd: {height1} ({height1/100 if height1 else 0}pt)")
        print(f"    MoveSelNextParaBegin: {height2} ({height2/100 if height2 else 0}pt)")
        print(f"    MoveRight+MoveSelLeft: {height3} ({height3/100 if height3 else 0}pt)")

    # 6. 테스트: 첫 번째 빈 문단 글자 크기 줄이기
    if empty_paras:
        test_para = empty_paras[0]
        print(f"\n[6] 테스트: para_id={test_para} 글자 크기 줄이기")

        hwp.SetPos(list_id, test_para, 0)
        hwp.HAction.Run("MoveParaBegin")

        # 문단 전체 선택 (다음 문단 시작 전까지)
        hwp.HAction.Run("MoveSelNextParaBegin")

        # 현재 글자 크기
        pset = hwp.HParameterSet.HCharShape
        hwp.HAction.GetDefault("CharShape", pset.HSet)
        before_height = pset.Height
        print(f"    변경 전: {before_height} ({before_height/100 if before_height else 0}pt)")

        # 글자 크기 줄이기 (5pt로)
        pset.Height = 500
        hwp.HAction.Execute("CharShape", pset.HSet)
        hwp.HAction.Run("Cancel")

        # 확인
        hwp.SetPos(list_id, test_para, 0)
        hwp.HAction.Run("MoveParaBegin")
        hwp.HAction.Run("MoveSelNextParaBegin")
        hwp.HAction.GetDefault("CharShape", pset.HSet)
        after_height = pset.Height
        print(f"    변경 후: {after_height} ({after_height/100 if after_height else 0}pt)")

        if before_height != after_height:
            print("    -> 성공!")
        else:
            print("    -> 실패! 글자 크기가 변경되지 않음")

    # 원래 위치 복원
    hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])

    # 7. 걸친 문단 상태 다시 확인
    print(f"\n[7] 걸친 문단 상태 재확인:")
    helper.ParaAlignWords()
    info = SeparatedPara.para_page_map.get(spanning_para)
    if info:
        print(f"  para_id={spanning_para}: 페이지 {info['start_page']}-{info['end_page']}")
        if info['start_page'] == info['end_page']:
            print("    -> 걸침 해소!")
        else:
            print("    -> 여전히 걸침")


if __name__ == "__main__":
    debug_empty_para()
