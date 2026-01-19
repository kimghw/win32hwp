"""
SeparatedPara 디버그 모듈 - 통합 CLI 인터페이스

기능:
1. empty_para: 빈 문단 글자 크기 변경 디버그
2. fix_all: 모든 걸친 문단 처리 테스트
3. fix_page: 특정 페이지 fix_page 함수 디버그
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor import get_hwp_instance
from separated_para import SeparatedPara


def get_hwp_and_helper():
    """공통: hwp 인스턴스와 SeparatedPara 헬퍼 반환"""
    hwp = get_hwp_instance()
    if not hwp:
        print("[ERROR] 한글이 실행 중이지 않습니다.")
        return None, None
    helper = SeparatedPara(hwp)
    return hwp, helper


def print_spanning_paragraphs(para_page_map, title="걸친 문단"):
    """공통: 걸친 문단 출력"""
    spanning_count = 0
    for para_id, info in sorted(para_page_map.items()):
        if info['start_page'] != info['end_page']:
            spanning_count += 1
            print(f"  para_id={para_id}: 페이지 {info['start_page']}-{info['end_page']} [걸침]")
    if spanning_count == 0:
        print(f"  ({title} 없음)")
    return spanning_count


def debug_empty_para():
    """빈 문단 글자 크기 변경 디버그"""
    hwp, helper = get_hwp_and_helper()
    if not hwp:
        return

    print("=" * 60)
    print("빈 문단 글자 크기 변경 디버그")
    print("=" * 60)

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


def debug_fix_all():
    """fix_all_paragraphs 함수 테스트 - 모든 걸친 문단 처리"""
    hwp, helper = get_hwp_and_helper()
    if not hwp:
        return

    print("=" * 60)
    print("fix_all_paragraphs 테스트 - 전체 걸침 문단 처리")
    print("=" * 60)

    # 처리 전 상태
    print("\n[처리 전 상태]")
    helper.ParaAlignWords()
    print_spanning_paragraphs(SeparatedPara.para_page_map)

    # fix_all_paragraphs 실행
    print("\n" + "-" * 50)
    print("[실행] fix_all_paragraphs()")
    print("-" * 50)

    result = helper.fix_all_paragraphs(page=None, min_font_size=4, max_rounds=50)

    print("-" * 50)
    print(f"\n[결과]")
    print(f"  반복 횟수: {result['rounds']}")
    print(f"  처리 문단: {result['processed']}")
    print(f"  성공: {result['success']}")
    print(f"  실패: {result['failed']}")
    print(f"  남은 걸침: {result['remaining_spanning']}")

    # 처리 후 상태
    print("\n[처리 후 상태]")
    helper.ParaAlignWords()
    print_spanning_paragraphs(SeparatedPara.para_page_map)


def debug_fix_page(page=1):
    """fix_page 함수 디버그"""
    hwp, helper = get_hwp_and_helper()
    if not hwp:
        return

    print("=" * 60)
    print("fix_page 함수 디버그")
    print("=" * 60)

    print(f"\n[실행] 페이지 {page} fix_page 호출")
    print("-" * 50)

    def log_callback(msg):
        print(msg)

    result = helper.fix_page(page, max_iterations=50, strategy='empty_font',
                             log_callback=log_callback)

    print("-" * 50)
    print(f"\n[결과]")
    print(f"  성공: {result.get('success')}")
    print(f"  반복: {result.get('iterations')}")
    print(f"  전략: {result.get('strategy_used')}")
    print(f"  메시지: {result.get('message', '')}")

    if 'original_lines_info' in result:
        print(f"  원본 줄분포: {result['original_lines_info'].get('lines_per_page')}")
    if 'final_lines_info' in result:
        print(f"  최종 줄분포: {result['final_lines_info'].get('lines_per_page')}")

    # 최종 상태 확인
    print("\n[최종 상태]")
    helper.ParaAlignWords()
    print_spanning_paragraphs(SeparatedPara.para_page_map)


def show_menu():
    """메뉴 출력"""
    print("\n" + "=" * 60)
    print("SeparatedPara 디버그 메뉴")
    print("=" * 60)
    print("1. empty_para  - 빈 문단 글자 크기 변경 디버그")
    print("2. fix_all     - 모든 걸친 문단 처리 테스트")
    print("3. fix_page    - 특정 페이지 fix_page 함수 디버그")
    print("q. 종료")
    print("-" * 60)


def main():
    """CLI 메인 함수"""
    # 커맨드라인 인자 처리
    if len(sys.argv) > 1:
        cmd = sys.argv[1].lower()
        if cmd in ('1', 'empty_para', 'empty'):
            debug_empty_para()
        elif cmd in ('2', 'fix_all', 'all'):
            debug_fix_all()
        elif cmd in ('3', 'fix_page', 'page'):
            page = int(sys.argv[2]) if len(sys.argv) > 2 else 1
            debug_fix_page(page)
        else:
            print(f"알 수 없는 명령: {cmd}")
            print("사용법: python debug_separated.py [empty_para|fix_all|fix_page [page]]")
        return

    # 대화형 모드
    while True:
        show_menu()
        choice = input("선택 (1-3, q): ").strip().lower()

        if choice in ('q', 'quit', 'exit'):
            print("종료합니다.")
            break
        elif choice in ('1', 'empty_para', 'empty'):
            debug_empty_para()
        elif choice in ('2', 'fix_all', 'all'):
            debug_fix_all()
        elif choice in ('3', 'fix_page', 'page'):
            page_input = input("페이지 번호 (기본값=1): ").strip()
            page = int(page_input) if page_input else 1
            debug_fix_page(page)
        else:
            print(f"알 수 없는 선택: {choice}")


if __name__ == "__main__":
    main()
