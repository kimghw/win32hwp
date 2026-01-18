# 마크다운 → 한글 문서 형식 변환
# # → □ (배경)
# ## → ○
# ### → -

import sys
from cursor_utils import get_hwp_instance

# 전역 hwp 인스턴스
_hwp_instance = None

def get_hwp():
    """전역 hwp 인스턴스 반환"""
    global _hwp_instance
    if _hwp_instance is None:
        _hwp_instance = get_hwp_instance()
    return _hwp_instance

def set_hwp(hwp):
    """전역 hwp 인스턴스 설정"""
    global _hwp_instance
    _hwp_instance = hwp


# 마크다운 레벨별 변환 설정
HEADING_STYLES = {
    1: {"prefix": "□ ", "indent": 0},      # # → □
    2: {"prefix": "○ ", "indent": 3},      # ## → ○ (3칸 들여쓰기)
    3: {"prefix": "- ", "indent": 5},       # ### → - (5칸 들여쓰기)
    4: {"prefix": "· ", "indent": 7},       # #### → · (7칸 들여쓰기)
    5: {"prefix": "▸ ", "indent": 9},       # ##### → ▸ (9칸 들여쓰기)
}


def parse_markdown_line(line):
    """마크다운 라인 파싱하여 레벨과 텍스트 반환"""
    stripped = line.lstrip()

    if not stripped.startswith('#'):
        return 0, line  # 일반 텍스트

    # # 개수 세기
    level = 0
    for ch in stripped:
        if ch == '#':
            level += 1
        else:
            break

    # # 뒤의 텍스트 추출
    text = stripped[level:].lstrip()

    return level, text


def convert_line(level, text):
    """레벨에 따라 한글 문서 형식으로 변환"""
    if level == 0:
        return text  # 일반 텍스트는 그대로

    style = HEADING_STYLES.get(level, HEADING_STYLES[5])
    indent = " " * style["indent"]
    prefix = style["prefix"]

    return f"{indent}{prefix}{text}"


def markdown_to_hwp_text(markdown_text):
    """마크다운 텍스트를 한글 문서 형식으로 변환"""
    lines = markdown_text.strip().split('\n')
    result = []

    for line in lines:
        level, text = parse_markdown_line(line)
        converted = convert_line(level, text)
        result.append(converted)

    return '\n'.join(result)


def insert_text_to_hwp(text):
    """한글 문서에 텍스트 삽입"""
    hwp = get_hwp()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return False

    # 텍스트 삽입
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

    return True


def markdown_to_hwp(markdown_text):
    """마크다운 텍스트를 한글 문서에 삽입"""
    hwp = get_hwp()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return False

    lines = markdown_text.strip().split('\n')

    for i, line in enumerate(lines):
        level, text = parse_markdown_line(line)
        converted = convert_line(level, text)

        # 텍스트 삽입
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = converted
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

        # 마지막 줄이 아니면 줄바꿈
        if i < len(lines) - 1:
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
            hwp.HParameterSet.HInsertText.Text = "\r\n"
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

    print(f"총 {len(lines)}줄 삽입 완료")
    return True


def markdown_to_hwp_in_table(markdown_text, table_index=0, row=0, col=0):
    """마크다운 텍스트를 테이블 셀에 삽입"""
    from table_info import TableInfo

    hwp = get_hwp()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return False

    table = TableInfo(hwp, debug=False)

    # 테이블 찾기
    tables = table.find_all_tables()
    if not tables:
        print("[오류] 문서에 테이블이 없습니다.")
        return False

    if table_index >= len(tables):
        print(f"[오류] 테이블 #{table_index}이 없습니다. (총 {len(tables)}개)")
        return False

    # 테이블 진입
    table.enter_table(table_index)

    # 해당 셀로 이동
    table.move_to_first_cell()

    # row, col 위치로 이동
    from table_info import MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL

    for _ in range(row):
        hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
    for _ in range(col):
        hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)

    # 마크다운 텍스트 변환 및 삽입
    lines = markdown_text.strip().split('\n')

    for i, line in enumerate(lines):
        level, text = parse_markdown_line(line)
        converted = convert_line(level, text)

        # 텍스트 삽입
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = converted
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

        # 마지막 줄이 아니면 줄바꿈
        if i < len(lines) - 1:
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
            hwp.HParameterSet.HInsertText.Text = "\r\n"
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

    # 테이블 밖으로 나가기
    hwp.HAction.Run("MoveParentList")
    hwp.HAction.Run("Cancel")

    print(f"테이블 #{table_index} 셀({row},{col})에 {len(lines)}줄 삽입 완료")
    return True


# 테스트
if __name__ == "__main__":
    # 테스트 마크다운
    test_markdown = """
# (배경) 자율운항선박이 실질적으로 운용되기 위해서는 육상과 연계되는 기술이 필수적으로 요구됨
## 우리나라에서 개발한 자율운항선박이 육상과 연계하여 실질적으로 상용화될 수 있도록 하기 위해 본 사업 추진
### 자율운항선박을 육상에서 관제하기 위한 기술
### 자율운항선박 육상제어사에 대한 면허체계
## 본 사업의 목표
### 단기 목표
### 장기 목표
"""

    print("=== 마크다운 → 한글 문서 형식 변환 테스트 ===\n")
    print("[입력 마크다운]")
    print(test_markdown)
    print("\n[변환 결과]")
    result = markdown_to_hwp_text(test_markdown)
    print(result)

    print("\n" + "="*50)
    print("\n한글 문서에 삽입하려면:")
    print("  from table_cell_generate import markdown_to_hwp")
    print("  markdown_to_hwp(마크다운텍스트)")
    print("\n테이블 셀에 삽입하려면:")
    print("  from table_cell_generate import markdown_to_hwp_in_table")
    print("  markdown_to_hwp_in_table(마크다운텍스트, table_index=0, row=0, col=0)")
