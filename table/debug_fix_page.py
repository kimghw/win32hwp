"""
fix_page 함수 디버그
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor_utils import get_hwp_instance
from separated_para import SeparatedPara


def debug_fix_page():
    hwp = get_hwp_instance()
    if not hwp:
        print("[ERROR] 한글이 실행 중이지 않습니다.")
        return

    print("=" * 60)
    print("fix_page 함수 디버그")
    print("=" * 60)

    helper = SeparatedPara(hwp)

    # 페이지 1 처리
    page = 1
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
    for para_id, info in sorted(SeparatedPara.para_page_map.items()):
        if info['start_page'] != info['end_page']:
            print(f"  para_id={para_id}: 페이지 {info['start_page']}-{info['end_page']} [걸침]")


if __name__ == "__main__":
    debug_fix_page()
