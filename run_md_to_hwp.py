# 마크다운 → 한글 문서 변환 + 분리단어/문단 처리
import os
import win32com.client as win32
from md_to_hwp import markdown_to_hwp, set_hwp
from separated_word import SeparatedWord
from separated_para import SeparatedPara

# 마크다운 파일 경로
MD_FILE = os.path.join(os.path.dirname(__file__), "table", "md_content.md")


def fix_all_paragraphs_word(hwp, debug=False):
    """문서의 모든 문단에 분리단어 처리 적용"""
    print("\n[분리단어 처리] 시작...")

    # 문서 처음으로 이동
    hwp.MovePos(2)  # moveDocBegin

    processed = 0
    adjusted_total = 0

    while True:
        pos = hwp.GetPos()
        para_id = pos[1]

        # 문단 시작으로 이동
        hwp.HAction.Run("MoveParaBegin")

        # 분리단어 처리
        sw = SeparatedWord(hwp, debug=debug)
        result = sw.fix_paragraph()

        processed += 1
        adjusted_total += result.get('adjusted_lines', 0)

        if debug:
            print(f"  문단 {para_id}: 조정 {result.get('adjusted_lines', 0)}줄")

        # 다음 문단으로 이동
        before_para = para_id
        hwp.HAction.Run("MoveNextParaBegin")
        after_pos = hwp.GetPos()

        if after_pos[1] == before_para:
            break

    print(f"[분리단어 처리] 완료: {processed}개 문단, 총 {adjusted_total}줄 조정")
    return {'processed': processed, 'adjusted': adjusted_total}


def fix_all_pages_para(hwp, debug=False):
    """문서의 모든 페이지에서 걸친 문단 처리"""
    print("\n[걸친 문단 처리] 시작...")

    helper = SeparatedPara(hwp)

    # 전체 페이지 수 확인
    total_pages = hwp.PageCount
    print(f"  전체 페이지 수: {total_pages}")

    fixed_count = 0

    for page in range(1, total_pages + 1):
        def log_callback(msg):
            if debug:
                print(f"    {msg}")

        result = helper.fix_page(page, log_callback=log_callback if debug else None)

        if result.get('success') and result.get('iterations', 0) > 0:
            fixed_count += 1
            print(f"  페이지 {page}: 걸침 해소 (반복 {result['iterations']}회)")

    print(f"[걸친 문단 처리] 완료: {fixed_count}개 페이지 처리됨")
    return {'pages': total_pages, 'fixed': fixed_count}


if __name__ == "__main__":
    print("=" * 60)
    print("마크다운 → 한글 문서 변환 + 분리단어/문단 처리")
    print("=" * 60)

    # 마크다운 파일 읽기
    print(f"\n[1] 마크다운 파일 읽기: {MD_FILE}")
    with open(MD_FILE, "r", encoding="utf-8") as f:
        test_markdown = f.read()

    # 한글 새 문서 열기
    print("\n[2] 한글 새 문서 열기...")
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
    hwp.XHwpWindows.Item(0).Visible = True
    set_hwp(hwp)

    print("\n[3] 마크다운 삽입 중...")
    result = markdown_to_hwp(test_markdown)

    if not result:
        print("\n삽입 실패!")
        exit(1)

    print("\n[4] 후처리 시작...")

    # 분리단어 처리 (줄 끝에서 단어가 잘린 경우 자간 조정)
    word_result = fix_all_paragraphs_word(hwp, debug=False)

    # 걸친 문단 처리 (페이지를 넘어가는 문단 처리)
    para_result = fix_all_pages_para(hwp, debug=False)

    print("\n" + "=" * 60)
    print("완료!")
    print("=" * 60)
    print(f"  마크다운 삽입: 성공")
    print(f"  분리단어 처리: {word_result['processed']}개 문단, {word_result['adjusted']}줄 조정")
    print(f"  걸친 문단 처리: {para_result['fixed']}개 페이지 처리")
