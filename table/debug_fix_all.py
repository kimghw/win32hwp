"""
fix_all_paragraphs 함수 테스트 - 모든 걸친 문단 처리
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor_utils import get_hwp_instance
from separated_para import SeparatedPara


def debug_fix_all():
    hwp = get_hwp_instance()
    if not hwp:
        print("[ERROR] 한글이 실행 중이지 않습니다.")
        return

    print("=" * 60)
    print("fix_all_paragraphs 테스트 - 전체 걸침 문단 처리")
    print("=" * 60)

    helper = SeparatedPara(hwp)

    # 처리 전 상태
    print("\n[처리 전 상태]")
    helper.ParaAlignWords()
    for para_id, info in sorted(SeparatedPara.para_page_map.items()):
        if info['start_page'] != info['end_page']:
            print(f"  para_id={para_id}: 페이지 {info['start_page']}-{info['end_page']} [걸침]")

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
    spanning_count = 0
    for para_id, info in sorted(SeparatedPara.para_page_map.items()):
        if info['start_page'] != info['end_page']:
            spanning_count += 1
            print(f"  para_id={para_id}: 페이지 {info['start_page']}-{info['end_page']} [걸침]")

    if spanning_count == 0:
        print("  (걸친 문단 없음)")


if __name__ == "__main__":
    debug_fix_all()
