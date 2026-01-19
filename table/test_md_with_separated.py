"""
마크다운 -> 한글 변환 + 분리된 단어 처리 테스트

table_md_2_hwp와 separated_word를 조합하여 사용하는 테스트 파일
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor import get_hwp_instance
from md_to_hwp import (
    markdown_to_hwp,
    markdown_to_hwp_in_table,
    set_hwp,
    get_hwp,
    parse_markdown_line,
    convert_line,
    set_para_shape,
    insert_picture,
    is_picture_line,
    parse_picture_line,
)
from separated_word import SeparatedWord
from separated_para import SeparatedPara


def markdown_to_hwp_with_fix(markdown_text, fix_separated=True, spacing_step=-1.0,
                              min_spacing=-100, max_iterations=100, debug=False):
    """
    마크다운 텍스트를 한글 문서에 삽입하고, 분리된 단어를 자동으로 처리

    Args:
        markdown_text: 마크다운 텍스트
        fix_separated: 분리된 단어 처리 여부 (기본 True)
        spacing_step: 자간 감소 단위 (기본 -1.0)
        min_spacing: 최소 자간 값 (기본 -100)
        max_iterations: 최대 반복 횟수 (기본 100)
        debug: 디버그 모드 (기본 False)

    Returns:
        {
            'success': bool,
            'lines_inserted': int,
            'pictures_inserted': int,
            'paragraphs_fixed': int,
            'fix_results': list  # 각 문단별 fix_paragraph 결과
        }
    """
    hwp = get_hwp()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return {'success': False, 'error': '한글이 실행 중이지 않습니다.'}

    lines = markdown_text.strip().split('\n')
    picture_count = 0
    paragraphs_inserted = []  # 삽입된 문단 ID 추적

    # 시작 위치 저장
    start_pos = hwp.GetPos()
    start_para = start_pos[1]

    for i, line in enumerate(lines):
        # [picture] 태그 처리
        if is_picture_line(line):
            pic_path = parse_picture_line(line)
            if pic_path:
                insert_picture(hwp, pic_path)
                picture_count += 1
            # 마지막 줄이 아니면 줄바꿈
            if i < len(lines) - 1:
                hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
                hwp.HParameterSet.HInsertText.Text = "\r\n"
                hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
            continue

        level, text = parse_markdown_line(line)
        converted = convert_line(level, text)

        # 문단 모양 먼저 설정 (내어쓰기)
        set_para_shape(hwp, level)

        # 텍스트 삽입
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = converted
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

        # 현재 문단 ID 기록
        current_pos = hwp.GetPos()
        current_para = current_pos[1]
        if current_para not in paragraphs_inserted:
            paragraphs_inserted.append(current_para)

        # 마지막 줄이 아니면 줄바꿈
        if i < len(lines) - 1:
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
            hwp.HParameterSet.HInsertText.Text = "\r\n"
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

    print(f"총 {len(lines)}줄 삽입 완료 (이미지 {picture_count}개)")

    # 분리된 단어 처리
    fix_results = []
    paragraphs_fixed = 0

    if fix_separated and paragraphs_inserted:
        print(f"\n[분리된 단어 처리] {len(paragraphs_inserted)}개 문단 처리 시작...")

        sw = SeparatedWord(hwp, debug=debug)
        list_id = start_pos[0]

        for para_id in paragraphs_inserted:
            # 해당 문단으로 이동
            hwp.SetPos(list_id, para_id, 0)
            hwp.HAction.Run("MoveParaBegin")

            # 분리된 단어 처리
            result = sw.fix_paragraph(
                spacing_step=spacing_step,
                min_spacing=min_spacing,
                max_iterations=max_iterations
            )

            fix_results.append({
                'para_id': para_id,
                'result': result
            })

            if result['adjusted_lines'] > 0:
                paragraphs_fixed += 1
                print(f"  문단 {para_id}: {result['adjusted_lines']}줄 조정, "
                      f"{result['skipped_lines']}줄 건너뜀, {result['failed_lines']}줄 실패")

        print(f"[분리된 단어 처리] 완료: {paragraphs_fixed}개 문단 수정됨")

    return {
        'success': True,
        'lines_inserted': len(lines),
        'pictures_inserted': picture_count,
        'paragraphs_fixed': paragraphs_fixed,
        'fix_results': fix_results
    }


def fix_current_paragraph(spacing_step=-1.0, min_spacing=-100,
                          max_iterations=100, debug=False):
    """
    현재 커서 위치의 문단에 대해 분리된 단어 처리 실행

    Args:
        spacing_step: 자간 감소 단위 (기본 -1.0)
        min_spacing: 최소 자간 값 (기본 -100)
        max_iterations: 최대 반복 횟수 (기본 100)
        debug: 디버그 모드 (기본 False)

    Returns:
        SeparatedWord.fix_paragraph() 결과
    """
    hwp = get_hwp()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return {'success': False, 'error': '한글이 실행 중이지 않습니다.'}

    sw = SeparatedWord(hwp, debug=debug)
    result = sw.fix_paragraph(
        spacing_step=spacing_step,
        min_spacing=min_spacing,
        max_iterations=max_iterations
    )

    print(f"[결과] 조정: {result['adjusted_lines']}줄, "
          f"건너뜀: {result['skipped_lines']}줄, "
          f"실패: {result['failed_lines']}줄")

    return result


def fix_all_paragraphs_in_document(spacing_step=-1.0, min_spacing=-100,
                                    max_iterations=100, debug=False):
    """
    문서 전체의 모든 문단에 대해 분리된 단어 처리 실행

    Args:
        spacing_step: 자간 감소 단위 (기본 -1.0)
        min_spacing: 최소 자간 값 (기본 -100)
        max_iterations: 최대 반복 횟수 (기본 100)
        debug: 디버그 모드 (기본 False)

    Returns:
        {
            'total_paragraphs': int,
            'paragraphs_fixed': int,
            'total_adjusted': int,
            'total_skipped': int,
            'total_failed': int,
            'results': list
        }
    """
    hwp = get_hwp()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return {'success': False, 'error': '한글이 실행 중이지 않습니다.'}

    # 시작 위치 저장
    saved_pos = hwp.GetPos()

    # 문서 처음으로 이동
    hwp.MovePos(2)  # moveDocBegin

    sw = SeparatedWord(hwp, debug=debug)

    total_paragraphs = 0
    paragraphs_fixed = 0
    total_adjusted = 0
    total_skipped = 0
    total_failed = 0
    results = []

    while True:
        pos = hwp.GetPos()
        para_id = pos[1]
        total_paragraphs += 1

        print(f"\n[문단 {total_paragraphs}] para_id={para_id} 처리 중...")

        # 분리된 단어 처리
        result = sw.fix_paragraph(
            spacing_step=spacing_step,
            min_spacing=min_spacing,
            max_iterations=max_iterations
        )

        results.append({
            'para_id': para_id,
            'result': result
        })

        total_adjusted += result['adjusted_lines']
        total_skipped += result['skipped_lines']
        total_failed += result['failed_lines']

        if result['adjusted_lines'] > 0:
            paragraphs_fixed += 1
            print(f"  -> 조정: {result['adjusted_lines']}줄")

        # 다음 문단으로 이동
        before_para = para_id
        hwp.HAction.Run("MoveNextParaBegin")
        after_pos = hwp.GetPos()
        after_para = after_pos[1]

        # 문단 변경 없으면 종료 (문서 끝)
        if before_para == after_para:
            break

    # 원래 위치 복원
    hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])

    print(f"\n{'='*50}")
    print(f"[전체 결과]")
    print(f"  처리 문단: {total_paragraphs}개")
    print(f"  수정된 문단: {paragraphs_fixed}개")
    print(f"  조정된 줄: {total_adjusted}줄")
    print(f"  건너뛴 줄: {total_skipped}줄")
    print(f"  실패한 줄: {total_failed}줄")
    print(f"{'='*50}")

    return {
        'total_paragraphs': total_paragraphs,
        'paragraphs_fixed': paragraphs_fixed,
        'total_adjusted': total_adjusted,
        'total_skipped': total_skipped,
        'total_failed': total_failed,
        'results': results
    }


def fix_page_spanning_paragraphs(page=None, strategy='empty_font', max_iterations=50, debug=False):
    """
    페이지를 걸쳐있는 문단 처리 (빈 문단 글자 크기 줄이기 등)

    Args:
        page: 대상 페이지 (None이면 현재 페이지)
        strategy: 전략 ('empty_font', 'char_spacing')
        max_iterations: 최대 반복 횟수
        debug: 디버그 모드

    Returns:
        SeparatedPara.fix_page() 결과
    """
    hwp = get_hwp()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return {'success': False, 'error': '한글이 실행 중이지 않습니다.'}

    # 현재 페이지 확인
    if page is None:
        key_info = hwp.KeyIndicator()
        page = key_info[3]

    print(f"[페이지 걸침 문단 처리] 페이지 {page}")
    print("-" * 50)

    helper = SeparatedPara(hwp)

    def log_callback(msg):
        if debug:
            print(msg)

    result = helper.fix_page(page, max_iterations=max_iterations,
                             strategy=strategy, log_callback=log_callback)

    print("-" * 50)
    print(f"[결과] 성공: {result.get('success', False)}, "
          f"반복: {result.get('iterations', 0)}회")

    if 'strategy_used' in result:
        print(f"  사용 전략: {result['strategy_used']}")
    if 'empty_line_removed' in result:
        print(f"  빈줄 제거: {result['empty_line_removed']}")

    return result


def fix_all_spanning_paragraphs(page=None, min_font_size=4, max_rounds=50):
    """
    모든 걸친 문단 처리 (걸침이 없을 때까지 반복)

    Args:
        page: 대상 페이지 (None이면 전체)
        min_font_size: 빈 문단 최소 글자 크기
        max_rounds: 최대 반복 횟수

    Returns:
        SeparatedPara.fix_all_paragraphs() 결과
    """
    hwp = get_hwp()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return {'success': False, 'error': '한글이 실행 중이지 않습니다.'}

    helper = SeparatedPara(hwp)

    print(f"[전체 걸침 문단 처리] 페이지: {page if page else '전체'}")
    print("-" * 50)

    result = helper.fix_all_paragraphs(page=page, min_font_size=min_font_size,
                                        max_rounds=max_rounds)

    print("-" * 50)
    print(f"[결과]")
    print(f"  반복 횟수: {result['rounds']}")
    print(f"  처리 문단: {result['processed']}")
    print(f"  성공: {result['success']}")
    print(f"  실패: {result['failed']}")
    print(f"  남은 걸침: {result['remaining_spanning']}")

    return result


def fix_words_in_page(page=None, spacing_step=-1.0, min_spacing=-100, max_iterations=100):
    """
    페이지 내 모든 문단의 분리된 단어 처리

    Args:
        page: 대상 페이지 (None이면 현재 페이지)
        spacing_step: 자간 감소 단위
        min_spacing: 최소 자간 값
        max_iterations: 문단당 최대 반복 횟수

    Returns:
        SeparatedPara.fix_all_words_in_page() 결과
    """
    hwp = get_hwp()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return {'success': False, 'error': '한글이 실행 중이지 않습니다.'}

    # 현재 페이지 확인
    if page is None:
        key_info = hwp.KeyIndicator()
        page = key_info[3]

    helper = SeparatedPara(hwp)

    print(f"[페이지 분리단어 처리] 페이지 {page}")
    print("-" * 50)

    result = helper.fix_all_words_in_page(page, spacing_step=spacing_step,
                                          min_spacing=min_spacing,
                                          max_iterations=max_iterations)

    print("-" * 50)
    print(f"[결과]")
    print(f"  처리 문단: {result['total_paragraphs']}")
    print(f"  조정된 줄: {result['total_adjusted']}")
    print(f"  건너뛴 줄: {result['total_skipped']}")
    print(f"  실패한 줄: {result['total_failed']}")

    return result


def get_page_info(page=None):
    """
    페이지 정보 조회 (문단 수, 걸침 여부 등)

    Args:
        page: 대상 페이지 (None이면 현재 페이지)

    Returns:
        페이지 정보 딕셔너리
    """
    hwp = get_hwp()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return None

    # 현재 페이지 확인
    if page is None:
        key_info = hwp.KeyIndicator()
        page = key_info[3]

    helper = SeparatedPara(hwp)

    # para_page_map 갱신
    helper.ParaAlignWords()

    # 페이지 문단 정보
    page_info = helper.get_page_paragraph_count(page)

    # 걸침 문단 확인
    spanning_info = helper.get_page_last_spanning_para(page)

    print(f"[페이지 {page} 정보]")
    print(f"  문단 수: {page_info['paragraph_count']}")
    print(f"  마지막 문단 걸침: {spanning_info['has_spanning']}")

    if spanning_info['has_spanning'] and spanning_info['lines_info']:
        lines_info = spanning_info['lines_info']
        print(f"  걸친 문단 ID: {lines_info['para_id']}")
        print(f"  줄 분포: {lines_info['lines_per_page']}")

    return {
        'page': page,
        'page_info': page_info,
        'spanning_info': spanning_info
    }


# 테스트용 마크다운 예시
TEST_MARKDOWN = """
# (배경) 자율운항선박이 실질적으로 운용되기 위해서는 육상과 연계되는 기술이 필수적으로 요구됨
## 우리나라에서 개발한 자율운항선박이 육상과 연계하여 실질적으로 상용화될 수 있도록 하기 위해 본 사업 추진
### 자율운항선박을 육상에서 관제하기 위한 기술
### 자율운항선박 육상제어사에 대한 면허체계
## 본 사업의 목표
### 단기 목표
### 장기 목표
"""


def create_hwp_from_md(md_file, output_file=None, fix_separated=True, debug=False):
    """
    마크다운 파일을 읽어서 한글 문서로 생성

    Args:
        md_file: 마크다운 파일 경로
        output_file: 출력 파일 경로 (None이면 저장 안 함)
        fix_separated: 분리된 단어 처리 여부
        debug: 디버그 모드

    Returns:
        hwp 인스턴스 (후속 작업용)
    """
    import win32com.client as win32

    print("=" * 60)
    print("마크다운 → 한글 문서 생성")
    print("=" * 60)

    # 1. 마크다운 파일 읽기
    print(f"\n[1] 마크다운 파일 읽기: {md_file}")
    with open(md_file, "r", encoding="utf-8") as f:
        markdown_text = f.read()
    print(f"    {len(markdown_text.splitlines())}줄 읽음")

    # 2. 한글 새 문서 열기
    print("\n[2] 한글 새 문서 열기...")
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
    hwp.XHwpWindows.Item(0).Visible = True
    set_hwp(hwp)
    print("    한글 연결 완료")

    # 3. 마크다운 삽입 + 분리단어 처리
    print("\n[3] 마크다운 내용 삽입 중...")
    result = markdown_to_hwp_with_fix(markdown_text, fix_separated=fix_separated, debug=debug)

    if not result.get('success'):
        print("    삽입 실패!")
        return hwp

    # 4. 파일 저장 (지정된 경우)
    if output_file:
        print(f"\n[4] 파일 저장: {output_file}")
        hwp.SaveAs(output_file, "HWP", "")
        print("    저장 완료!")

    print("\n" + "=" * 60)
    print(f"결과: {result['lines_inserted']}줄, 이미지 {result['pictures_inserted']}개")
    if fix_separated:
        print(f"분리단어 처리: {result['paragraphs_fixed']}개 문단 수정됨")
    print("=" * 60)

    return hwp


if __name__ == "__main__":
    import win32com.client as win32

    # 파일 경로 설정
    MD_FILE = os.path.join(os.path.dirname(__file__), "md_content.md")
    OUTPUT_FILE = r"C:\win32hwp\table\output_with_fix.hwp"

    # 마크다운 파일로 한글 문서 생성
    hwp = create_hwp_from_md(MD_FILE, OUTPUT_FILE, fix_separated=True, debug=False)

    # 페이지 걸침 문단 자동 처리
    print("\n[페이지 걸침 문단 처리]")
    result = fix_all_spanning_paragraphs()
    if result.get('remaining_spanning', 0) == 0:
        print("  -> 모든 걸침 해소됨")
    else:
        print(f"  -> {result['remaining_spanning']}개 걸침 남음")

    # 파일 다시 저장
    hwp.Save()

    print("\n[OK] 한글 인스턴스가 준비되었습니다.")
