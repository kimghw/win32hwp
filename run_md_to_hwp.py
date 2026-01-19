# 마크다운 → 한글 문서 변환 + 분리단어/문단 처리 (섹션별)
import os
import win32com.client as win32
from md_to_hwp import (
    set_hwp,
    split_by_section,
    get_sections,
    parse_markdown_line,
    convert_line,
    set_para_shape,
    is_picture_line,
    parse_picture_line,
    insert_picture,
)
from separated_word import SeparatedWord
from separated_para import SeparatedPara

# 마크다운 파일 경로
MD_FILE = os.path.join(os.path.dirname(__file__), "table", "md_content.md")


def fix_section_paragraphs(hwp, start_para, end_para, debug=False):
    """특정 범위의 문단에 분리단어 처리 적용"""
    adjusted_total = 0
    processed = 0

    list_id = hwp.GetPos()[0]

    for para_id in range(start_para, end_para + 1):
        hwp.SetPos(list_id, para_id, 0)
        hwp.HAction.Run("MoveParaBegin")

        sw = SeparatedWord(hwp, debug=debug)
        result = sw.fix_paragraph()

        processed += 1
        adjusted_total += result.get('adjusted_lines', 0)

        if debug and result.get('adjusted_lines', 0) > 0:
            print(f"    문단 {para_id}: 조정 {result.get('adjusted_lines', 0)}줄")

    return {'processed': processed, 'adjusted': adjusted_total}


def fix_all_pages_para(hwp, debug=False):
    """문서의 모든 페이지에서 걸친 문단 처리"""
    print("\n[걸친 문단 처리] 시작...")

    helper = SeparatedPara(hwp)

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


def render_section_content(hwp, section_content):
    """섹션 내용을 한글에 렌더링하고 삽입된 문단 범위 반환"""
    lines = section_content.strip().split('\n')
    picture_count = 0

    # 시작 문단 ID 저장
    start_para = hwp.GetPos()[1]

    for i, line in enumerate(lines):
        # [picture] 태그 처리
        if is_picture_line(line):
            pic_path = parse_picture_line(line)
            if pic_path:
                insert_picture(hwp, pic_path)
                picture_count += 1
            if i < len(lines) - 1:
                hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
                hwp.HParameterSet.HInsertText.Text = "\r\n"
                hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
            continue

        level, text = parse_markdown_line(line)
        converted = convert_line(level, text)

        # 문단 모양 설정
        set_para_shape(hwp, level)

        # 텍스트 삽입
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = converted
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

        # 줄바꿈
        if i < len(lines) - 1:
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
            hwp.HParameterSet.HInsertText.Text = "\r\n"
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

    # 끝 문단 ID
    end_para = hwp.GetPos()[1]

    return {
        'lines': len(lines),
        'pictures': picture_count,
        'start_para': start_para,
        'end_para': end_para
    }


def process_markdown_by_section(hwp, markdown_text, debug=False):
    """마크다운을 섹션별로 렌더링하고 각 섹션마다 분리단어 처리"""
    sections = split_by_section(markdown_text)
    total_sections = len(sections)

    print(f"\n[섹션 분리] 총 {total_sections}개 섹션")
    for i, s in enumerate(sections):
        line_count = len(s['content'].split('\n'))
        print(f"  [{i}] {s['title']} ({line_count}줄)")

    total_lines = 0
    total_pictures = 0
    total_adjusted = 0
    section_results = []

    for idx, section in enumerate(sections):
        print(f"\n{'='*50}")
        print(f"[섹션 {idx}/{total_sections-1}] {section['title']}")
        print(f"{'='*50}")

        # 첫 번째 섹션이 아니면 상단에 빈줄 추가
        if idx > 0:
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
            hwp.HParameterSet.HInsertText.Text = "\r\n"
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

        # 1. 섹션 렌더링
        print(f"  [렌더링]...")
        render_result = render_section_content(hwp, section['content'])
        total_lines += render_result['lines']
        total_pictures += render_result['pictures']
        print(f"    {render_result['lines']}줄, 이미지 {render_result['pictures']}개")
        print(f"    문단 범위: {render_result['start_para']} ~ {render_result['end_para']}")

        # 2. 분리단어 처리
        print(f"  [분리단어 처리]...")
        fix_result = fix_section_paragraphs(
            hwp,
            render_result['start_para'],
            render_result['end_para'],
            debug=debug
        )
        total_adjusted += fix_result['adjusted']
        print(f"    {fix_result['processed']}개 문단, {fix_result['adjusted']}줄 조정")

        section_results.append({
            'index': idx,
            'title': section['title'],
            'lines': render_result['lines'],
            'pictures': render_result['pictures'],
            'adjusted': fix_result['adjusted']
        })

    return {
        'total_sections': total_sections,
        'total_lines': total_lines,
        'total_pictures': total_pictures,
        'total_adjusted': total_adjusted,
        'sections': section_results
    }


if __name__ == "__main__":
    print("=" * 60)
    print("마크다운 → 한글 문서 변환 (섹션별 처리)")
    print("=" * 60)

    # 마크다운 파일 읽기
    print(f"\n[1] 마크다운 파일 읽기: {MD_FILE}")
    with open(MD_FILE, "r", encoding="utf-8") as f:
        markdown_text = f.read()

    # 한글 새 문서 열기
    print("\n[2] 한글 새 문서 열기...")
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
    hwp.XHwpWindows.Item(0).Visible = True
    set_hwp(hwp)

    # 섹션별 처리 (렌더링 + 분리단어)
    print("\n[3] 섹션별 렌더링 + 분리단어 처리...")
    result = process_markdown_by_section(hwp, markdown_text, debug=False)

    # 걸친 문단 처리 (페이지를 넘어가는 문단)
    print("\n[4] 걸친 문단 처리...")
    para_result = fix_all_pages_para(hwp, debug=False)

    # 결과 요약
    print("\n" + "=" * 60)
    print("완료!")
    print("=" * 60)
    print(f"  총 섹션: {result['total_sections']}개")
    print(f"  총 줄 수: {result['total_lines']}줄")
    print(f"  총 이미지: {result['total_pictures']}개")
    print(f"  분리단어 조정: {result['total_adjusted']}줄")
    print(f"  걸친 문단 처리: {para_result['fixed']}개 페이지")
    print()
    print("[섹션별 결과]")
    for s in result['sections']:
        print(f"  [{s['index']}] {s['title']}: {s['lines']}줄, 조정 {s['adjusted']}줄")
