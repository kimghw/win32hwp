"""align_page 1, 2페이지 실행 테스트"""

from cursor_position_monitor import get_hwp_instance
from text_align_page import TextAlignPage

hwp = get_hwp_instance()
if not hwp:
    print("[ERROR] 한글을 찾을 수 없습니다.")
    exit()

print("[OK] 한글 연결됨")

helper = TextAlignPage(hwp)

for page in [1, 2]:
    print(f"\n{'='*50}")
    print(f"페이지 {page} 처리 시작")
    print('='*50)

    result = helper.align_page(page)

    print(f"처리 문단: {result['total_paragraphs']}개")
    print(f"조정: {result['total_adjusted']}, 건너뜀: {result['total_skipped']}, 실패: {result['total_failed']}")

    print("\n문단별 결과:")
    for r in result['results']:
        print(f"  para_id={r['para_id']}: 조정={r['adjusted']}, 건너뜀={r['skipped']}, 실패={r['failed']}")
