"""HWP 문서 정렬 도구

실행: python run.py
"""

from cursor_utils import get_hwp_instance


def main():
    print("=" * 50)
    print("HWP 문서 정렬 도구")
    print("=" * 50)
    print()
    print("  1. 분리된 단어 처리")
    print("     (한 단어가 두 줄에 걸친 경우)")
    print()
    print("  2. 분리된 문단 처리")
    print("     (문단이 두 페이지에 걸친 경우)")
    print()
    print("  0. 종료")
    print()
    print("=" * 50)

    choice = input("선택: ").strip()

    if choice == "0":
        print("종료")
        return

    if choice not in ["1", "2"]:
        print("잘못된 선택")
        return

    # HWP 연결
    hwp = get_hwp_instance()
    if not hwp:
        print()
        print("[오류] 한글을 찾을 수 없습니다.")
        print("       한글을 먼저 실행하고 문서를 열어주세요.")
        return

    print()
    print("[연결] 한글에 연결됨")
    print()

    if choice == "1":
        from separated_word import SeparatedWord

        print("분리된 단어 처리 중...")
        print("-" * 50)

        fixer = SeparatedWord(hwp, debug=False)
        result = fixer.fix_paragraph()

        print()
        print("[결과]")
        print(f"  조정: {result['adjusted_lines']}줄")
        print(f"  건너뜀: {result['skipped_lines']}줄")
        print(f"  실패: {result['failed_lines']}줄")

    elif choice == "2":
        from separated_para import SeparatedPara

        page = hwp.KeyIndicator()[3]
        print(f"현재 페이지: {page}")
        print()
        print("분리된 문단 처리 중...")
        print("-" * 50)

        fixer = SeparatedPara(hwp)
        result = fixer.fix_page(page, strategy='empty_font')

        print()
        print("[결과]")
        print(f"  성공: {result['success']}")
        print(f"  반복: {result.get('iterations', 0)}회")

    print()
    print("완료")


if __name__ == "__main__":
    main()
