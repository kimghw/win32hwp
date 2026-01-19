# 마크다운 → 한글 문서 형식 변환
# # → □ (배경)
# ## → ○
# ### → -
# [picture] 경로 → 이미지 삽입

import sys
import os
import re
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
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
# left_margin: 왼쪽 여백 (pt)
# hanging: 내어쓰기 (pt) - prefix 너비만큼
# 10pt 글꼴 기준: 전각=10pt, 반각=5pt
# □, ○, · = 전각(10pt), - = 반각(5pt), 공백 = 반각(5pt)
HEADING_STYLES = {
    1: {"prefix": "□ ", "left_margin": 0, "hanging": 15},           # # → □  (공백0개 + □ + 공백 = 10+5=15pt)
    2: {"prefix": "  ○ ", "left_margin": 0, "hanging": 25},         # ## → ○ (공백2개 + ○ + 공백 = 5+5+10+5=25pt)
    3: {"prefix": "   - ", "left_margin": 0, "hanging": 25},        # ### → - (공백3개 + - + 공백 = 5+5+5+5+5=25pt)
    4: {"prefix": "    · ", "left_margin": 0, "hanging": 35},       # #### → · (공백4개 + · + 공백 = 20+10+5=35pt)
    5: {"prefix": "     ▸ ", "left_margin": 0, "hanging": 40},      # ##### → ▸ (공백5개 + ▸ + 공백 = 25+10+5=40pt)
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
    """레벨에 따라 한글 문서 형식으로 변환 (텍스트만)"""
    if level == 0:
        return text  # 일반 텍스트는 그대로

    style = HEADING_STYLES.get(level, HEADING_STYLES[5])
    prefix = style["prefix"]

    return f"{prefix}{text}"


def set_para_shape(hwp, level):
    """문단 모양 설정 (왼쪽 여백 + 내어쓰기)"""
    if level == 0:
        return

    style = HEADING_STYLES.get(level, HEADING_STYLES[5])
    left_margin = style["left_margin"]
    hanging = style["hanging"]

    # pt → HWPUNIT 변환 (1pt = 200 HWPUNIT)
    left_margin_hwp = int(left_margin * 200)
    hanging_hwp = int(hanging * 200)

    # 문단 모양 설정
    hwp.HAction.GetDefault("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)
    hwp.HParameterSet.HParaShape.LeftMargin = left_margin_hwp
    hwp.HParameterSet.HParaShape.Indentation = -hanging_hwp  # 음수 = 내어쓰기
    hwp.HAction.Execute("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)


# 이미지 삽입 관련 상수
# HwpSizeOption
SIZE_REAL = 0           # 원본 크기
SIZE_SPECIFIC = 1       # 지정 크기 (width, height 사용)
SIZE_CELL = 2           # 셀 크기에 맞춤
SIZE_CELL_RATIO = 3     # 셀 크기에 맞춤 (비율 유지)

# HwpPictureEffect
EFFECT_REAL = 0         # 실제 이미지
EFFECT_GRAYSCALE = 1    # 그레이스케일
EFFECT_BLACKWHITE = 2   # 흑백


def is_picture_line(line):
    """[picture] 태그가 있는 줄인지 확인"""
    stripped = line.strip()
    return stripped.lower().startswith("[picture]")


def parse_picture_line(line):
    """[picture] 줄에서 경로 추출

    형식: [picture] 경로
    예: [picture] C:/images/test.jpg
        [picture] /mnt/c/win32hwp/test.jpg

    Returns:
        경로 문자열 또는 None
    """
    stripped = line.strip()
    if not stripped.lower().startswith("[picture]"):
        return None

    # [picture] 뒤의 경로 추출
    path = stripped[9:].strip()  # "[picture]" = 9글자

    if not path:
        return None

    return path


def insert_picture(hwp, path, width=None, height=None, embed=True, in_cell=False, caption=True):
    """한글 문서에 이미지 삽입

    Args:
        hwp: HWP 객체
        path: 이미지 파일 경로
        width: 가로 크기 (mm), None이면 자동
        height: 세로 크기 (mm), None이면 자동
        embed: 문서에 포함 여부 (기본 True)
        in_cell: 테이블 셀 안에 삽입 여부
        caption: 캡션에 파일명 삽입 여부 (기본 True)

    Returns:
        성공 시 Ctrl 객체, 실패 시 None
    """
    # 경로 변환 (WSL 경로 → Windows 경로)
    win_path = path
    if path.startswith("/mnt/"):
        # /mnt/c/... → C:/...
        parts = path.split("/")
        if len(parts) >= 3:
            drive = parts[2].upper()
            rest = "/".join(parts[3:])
            win_path = f"{drive}:/{rest}"

    # 파일 존재 확인 (원본 경로로)
    check_path = path
    if not os.path.exists(check_path) and os.path.exists(win_path.replace("/", "\\")):
        check_path = win_path.replace("/", "\\")

    # 크기 옵션 결정
    if in_cell:
        size_option = SIZE_CELL_RATIO  # 셀 크기에 맞춤 (비율 유지)
    elif width is not None and height is not None:
        size_option = SIZE_SPECIFIC
    else:
        size_option = SIZE_REAL

    # 기본값 설정
    if width is None:
        width = 0
    if height is None:
        height = 0

    try:
        # 이미지 삽입 전 문단 가운데 정렬
        hwp.HAction.GetDefault("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)
        hwp.HParameterSet.HParaShape.AlignType = 3  # 3 = 가운데 정렬
        hwp.HAction.Execute("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)

        ctrl = hwp.InsertPicture(
            win_path,       # path
            embed,          # embedded
            size_option,    # sizeoption
            False,          # reverse
            False,          # watermark
            EFFECT_REAL,    # effect
            width,          # width (mm)
            height          # height (mm)
        )

        if ctrl:
            # 캡션 추가 (파일명)
            if caption:
                try:
                    # 파일명 추출
                    filename = os.path.basename(win_path)

                    # 이미지 선택 (방금 삽입한 개체)
                    hwp.HAction.Run("SelectCtrlFront")

                    # 캡션 넣기
                    hwp.HAction.Run("ShapeObjAttachCaption")

                    # 캡션에 파일명 입력
                    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
                    hwp.HParameterSet.HInsertText.Text = filename
                    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

                    # 선택 해제
                    hwp.HAction.Run("Cancel")

                    print(f"[이미지 삽입] {win_path} (가운데 정렬, 캡션: {filename})")
                except Exception as cap_e:
                    print(f"[경고] 캡션 삽입 실패: {cap_e}")
                    print(f"[이미지 삽입] {win_path} (가운데 정렬)")
            else:
                print(f"[이미지 삽입] {win_path} (가운데 정렬)")
            return ctrl
        else:
            print(f"[오류] 이미지 삽입 실패: {win_path}")
            return None

    except Exception as e:
        print(f"[오류] 이미지 삽입 중 예외 발생: {e}")
        print(f"       경로: {win_path}")
        return None


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
    """마크다운 텍스트를 한글 문서에 삽입 (내어쓰기 적용, 이미지 지원)"""
    hwp = get_hwp()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return False

    lines = markdown_text.strip().split('\n')
    picture_count = 0

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

        # 마지막 줄이 아니면 줄바꿈
        if i < len(lines) - 1:
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
            hwp.HParameterSet.HInsertText.Text = "\r\n"
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

    print(f"총 {len(lines)}줄 삽입 완료 (이미지 {picture_count}개)")
    return True


def markdown_to_hwp_in_table(markdown_text, table_index=0, row=0, col=0):
    """마크다운 텍스트를 테이블 셀에 삽입"""
    try:
        from .table_info import TableInfo
    except ImportError:
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
    try:
        from .table_info import MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL
    except ImportError:
        from table_info import MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL

    for _ in range(row):
        hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
    for _ in range(col):
        hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)

    # 마크다운 텍스트 변환 및 삽입
    lines = markdown_text.strip().split('\n')
    picture_count = 0

    for i, line in enumerate(lines):
        # [picture] 태그 처리
        if is_picture_line(line):
            pic_path = parse_picture_line(line)
            if pic_path:
                insert_picture(hwp, pic_path, in_cell=True)
                picture_count += 1
            # 마지막 줄이 아니면 줄바꿈
            if i < len(lines) - 1:
                hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
                hwp.HParameterSet.HInsertText.Text = "\r\n"
                hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
            continue

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

    print(f"테이블 #{table_index} 셀({row},{col})에 {len(lines)}줄 삽입 완료 (이미지 {picture_count}개)")
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
    print("  from table.table_md_2_hwp import markdown_to_hwp")
    print("  markdown_to_hwp(마크다운텍스트)")
    print("\n테이블 셀에 삽입하려면:")
    print("  from table.table_md_2_hwp import markdown_to_hwp_in_table")
    print("  markdown_to_hwp_in_table(마크다운텍스트, table_index=0, row=0, col=0)")
