# -*- coding: utf-8 -*-
"""
한글(HWP) 서식/스타일 관리 모듈

주요 기능:
1. 글자 모양 (CharShape) - 굵게, 기울임, 색상, 크기, 밑줄 등
2. 문단 모양 (ParaShape) - 정렬, 줄간격, 여백 등
3. 서식 복사/붙여넣기 - 글자/문단 모양 복사
4. 스타일 적용/제거
"""

import pythoncom
import win32com.client as win32


def get_hwp_instance():
    """실행 중인 한글 인스턴스에 연결"""
    try:
        context = pythoncom.CreateBindCtx(0)
        rot = pythoncom.GetRunningObjectTable()

        for moniker in rot:
            name = moniker.GetDisplayName(context, None)
            if 'HwpObject' in name:
                obj = rot.GetObject(moniker)
                return win32.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
    except:
        pass
    return None


# ========== 유틸리티 함수 ==========

def rgb_to_bgr(r, g, b):
    """RGB를 HWP BGR 색상값으로 변환"""
    return (b << 16) | (g << 8) | r


def bgr_to_rgb(bgr):
    """HWP BGR 색상값을 RGB 튜플로 변환"""
    r = bgr & 0xFF
    g = (bgr >> 8) & 0xFF
    b = (bgr >> 16) & 0xFF
    return (r, g, b)


def pt_to_hwpunit(pt):
    """포인트를 HWPUNIT으로 변환 (1pt = 100)"""
    return int(pt * 100)


def hwpunit_to_pt(hwpunit):
    """HWPUNIT을 포인트로 변환"""
    return hwpunit / 100


# ========== 스타일 조회 함수 (독립 함수) ==========

def get_char_shape(hwp):
    """
    현재 커서 위치의 글자 모양 정보 반환

    Returns:
        dict: {
            'font_hangul': 한글 글꼴,
            'font_latin': 영문 글꼴,
            'font_hanja': 한자 글꼴,
            'size_pt': 글자 크기 (pt),
            'bold': 굵게 여부,
            'italic': 기울임 여부,
            'underline': 밑줄 타입,
            'strikeout': 취소선 타입,
            'text_color': 글자색 (BGR),
            'text_color_rgb': 글자색 (R, G, B),
            'shade_color': 음영색,
            'spacing': 자간 (한글),
            'ratio': 장평 (한글),
            'superscript': 위첨자,
            'subscript': 아래첨자
        }
    """
    hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
    pset = hwp.HParameterSet.HCharShape

    text_color = pset.TextColor
    r = text_color & 0xFF
    g = (text_color >> 8) & 0xFF
    b = (text_color >> 16) & 0xFF

    return {
        'font_hangul': pset.FaceNameHangul,
        'font_latin': pset.FaceNameLatin,
        'font_hanja': pset.FaceNameHanja,
        'size_pt': pset.Height / 100,
        'bold': bool(pset.Bold),
        'italic': bool(pset.Italic),
        'underline': pset.UnderlineType,
        'strikeout': pset.StrikeOutType,
        'text_color': text_color,
        'text_color_rgb': (r, g, b),
        'shade_color': pset.ShadeColor,
        'spacing': pset.SpacingHangul,
        'ratio': pset.RatioHangul,
        'superscript': bool(pset.SuperScript),
        'subscript': bool(pset.SubScript)
    }


def get_para_shape(hwp):
    """
    현재 커서 위치의 문단 모양 정보 반환

    Returns:
        dict: {
            'align': 정렬 타입 (0=양쪽, 1=왼쪽, 2=오른쪽, 3=가운데, 4=배분, 5=나눔),
            'align_name': 정렬 이름,
            'line_spacing': 줄간격 값,
            'line_spacing_type': 줄간격 종류 (0=글자에따라, 1=고정값, 2=여백만, 3=최소),
            'line_spacing_type_name': 줄간격 종류 이름,
            'prev_spacing_pt': 문단 위 여백 (pt),
            'next_spacing_pt': 문단 아래 여백 (pt),
            'left_margin_pt': 왼쪽 여백 (pt),
            'right_margin_pt': 오른쪽 여백 (pt),
            'indent_pt': 들여쓰기 (pt)
        }
    """
    hwp.HAction.GetDefault("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)
    pset = hwp.HParameterSet.HParaShape

    align_names = {0: '양쪽', 1: '왼쪽', 2: '오른쪽', 3: '가운데', 4: '배분', 5: '나눔'}
    spacing_type_names = {0: '글자에따라(%)', 1: '고정값', 2: '여백만지정', 3: '최소'}

    return {
        'align': pset.AlignType,
        'align_name': align_names.get(pset.AlignType, '알수없음'),
        'line_spacing': pset.LineSpacing,
        'line_spacing_type': pset.LineSpacingType,
        'line_spacing_type_name': spacing_type_names.get(pset.LineSpacingType, '알수없음'),
        'prev_spacing_pt': pset.PrevSpacing / 100,
        'next_spacing_pt': pset.NextSpacing / 100,
        'left_margin_pt': pset.LeftMargin / 100,
        'right_margin_pt': pset.RightMargin / 100,
        'indent_pt': pset.Indentation / 100
    }


def get_cursor_style(hwp):
    """
    현재 커서 위치의 글자 모양 + 문단 모양 통합 조회

    Returns:
        dict: {
            'char': 글자 모양 정보,
            'para': 문단 모양 정보,
            'position': 커서 위치 정보 (para_id, char_pos, page, line, column)
        }
    """
    pos = hwp.GetPos()
    key = hwp.KeyIndicator()

    position = {
        'list_id': pos[0],
        'para_id': pos[1],
        'char_pos': pos[2],
        'page': key[2],
        'line': key[4],
        'column': key[5]
    }

    return {
        'char': get_char_shape(hwp),
        'para': get_para_shape(hwp),
        'position': position
    }


def print_cursor_style(hwp):
    """현재 커서 위치의 스타일 정보를 보기 좋게 출력"""
    style = get_cursor_style(hwp)
    char = style['char']
    para = style['para']
    pos = style['position']

    print("=" * 50)
    print("[ 커서 위치 ]")
    print(f"  문단: {pos['para_id']}, 위치: {pos['char_pos']}, 페이지: {pos['page']}, 줄: {pos['line']}")
    print()
    print("[ 글자 모양 ]")
    print(f"  글꼴(한글): {char['font_hangul']}")
    print(f"  글꼴(영문): {char['font_latin']}")
    print(f"  크기: {char['size_pt']}pt")
    print(f"  굵게: {'예' if char['bold'] else '아니오'}, 기울임: {'예' if char['italic'] else '아니오'}")
    print(f"  밑줄: {char['underline']}, 취소선: {char['strikeout']}")
    print(f"  글자색(RGB): {char['text_color_rgb']}")
    print(f"  자간: {char['spacing']}, 장평: {char['ratio']}%")
    print()
    print("[ 문단 모양 ]")
    print(f"  정렬: {para['align_name']}")
    print(f"  줄간격: {para['line_spacing']}% ({para['line_spacing_type_name']})")
    print(f"  문단 위 여백: {para['prev_spacing_pt']}pt")
    print(f"  문단 아래 여백: {para['next_spacing_pt']}pt")
    print(f"  왼쪽 여백: {para['left_margin_pt']}pt, 오른쪽 여백: {para['right_margin_pt']}pt")
    print(f"  들여쓰기: {para['indent_pt']}pt")
    print("=" * 50)


# ========== 색상 상수 ==========

COLORS = {
    'black': 0x000000,
    'red': 0x0000FF,
    'green': 0x00FF00,
    'blue': 0xFF0000,
    'yellow': 0x00FFFF,
    'cyan': 0xFFFF00,
    'magenta': 0xFF00FF,
    'white': 0xFFFFFF,
    'gray': 0x808080,
    'dark_red': 0x000080,
    'dark_green': 0x008000,
    'dark_blue': 0x800000,
    'orange': 0x00A5FF,
    'purple': 0x800080,
}


class StyleFormat:
    """한글 서식/스타일 관리 클래스"""

    def __init__(self, hwp):
        self.hwp = hwp
        self._saved_styles = {}  # 스타일 저장소 {이름: {char: {...}, para: {...}}}

    # ========== 스타일 저장/불러오기 ==========

    def save_style(self, name, include_char=True, include_para=True):
        """
        현재 커서 위치의 스타일을 이름으로 저장

        Args:
            name: 스타일 이름
            include_char: 글자 모양 포함 여부
            include_para: 문단 모양 포함 여부

        Returns:
            dict: 저장된 스타일 정보
        """
        style = {}
        if include_char:
            style['char'] = get_char_shape(self.hwp)
        if include_para:
            style['para'] = get_para_shape(self.hwp)

        self._saved_styles[name] = style
        return style

    def save_char_style(self, name):
        """글자 모양만 저장"""
        return self.save_style(name, include_char=True, include_para=False)

    def save_para_style(self, name):
        """문단 모양만 저장"""
        return self.save_style(name, include_char=False, include_para=True)

    def load_style(self, name, apply_char=True, apply_para=True):
        """
        저장된 스타일을 불러와서 적용 (선택 영역에 적용)

        Args:
            name: 스타일 이름
            apply_char: 글자 모양 적용 여부
            apply_para: 문단 모양 적용 여부

        Returns:
            bool: 성공 여부
        """
        if name not in self._saved_styles:
            return False

        style = self._saved_styles[name]

        if apply_char and 'char' in style:
            self._apply_char_style(style['char'])
        if apply_para and 'para' in style:
            self._apply_para_style(style['para'])

        return True

    def load_char_style(self, name):
        """저장된 글자 모양만 적용"""
        return self.load_style(name, apply_char=True, apply_para=False)

    def load_para_style(self, name):
        """저장된 문단 모양만 적용"""
        return self.load_style(name, apply_char=False, apply_para=True)

    def _apply_char_style(self, char_info):
        """글자 모양 정보를 적용"""
        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
        pset = self.hwp.HParameterSet.HCharShape

        pset.FaceNameHangul = char_info.get('font_hangul', '맑은 고딕')
        pset.FaceNameLatin = char_info.get('font_latin', '맑은 고딕')
        pset.FaceNameHanja = char_info.get('font_hanja', '맑은 고딕')
        pset.Height = int(char_info.get('size_pt', 10) * 100)
        pset.Bold = int(char_info.get('bold', False))
        pset.Italic = int(char_info.get('italic', False))
        pset.UnderlineType = char_info.get('underline', 0)
        pset.StrikeOutType = char_info.get('strikeout', 0)
        pset.TextColor = char_info.get('text_color', 0)
        pset.ShadeColor = char_info.get('shade_color', -1)
        pset.SpacingHangul = char_info.get('spacing', 0)
        pset.SpacingLatin = char_info.get('spacing', 0)
        pset.RatioHangul = char_info.get('ratio', 100)
        pset.RatioLatin = char_info.get('ratio', 100)
        pset.SuperScript = int(char_info.get('superscript', False))
        pset.SubScript = int(char_info.get('subscript', False))

        self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)

    def _apply_para_style(self, para_info):
        """문단 모양 정보를 적용"""
        self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        pset = self.hwp.HParameterSet.HParaShape

        pset.AlignType = para_info.get('align', 0)
        pset.LineSpacing = para_info.get('line_spacing', 160)
        pset.LineSpacingType = para_info.get('line_spacing_type', 0)
        pset.PrevSpacing = int(para_info.get('prev_spacing_pt', 0) * 100)
        pset.NextSpacing = int(para_info.get('next_spacing_pt', 0) * 100)
        pset.LeftMargin = int(para_info.get('left_margin_pt', 0) * 100)
        pset.RightMargin = int(para_info.get('right_margin_pt', 0) * 100)
        pset.Indentation = int(para_info.get('indent_pt', 0) * 100)

        self.hwp.HAction.Execute("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)

    def get_saved_style(self, name):
        """저장된 스타일 정보 조회"""
        return self._saved_styles.get(name)

    def get_all_saved_styles(self):
        """모든 저장된 스타일 이름 목록"""
        return list(self._saved_styles.keys())

    def delete_saved_style(self, name):
        """저장된 스타일 삭제"""
        if name in self._saved_styles:
            del self._saved_styles[name]
            return True
        return False

    def copy_style_to(self, name):
        """현재 커서 위치의 스타일을 복사하여 저장 (copy_style 별칭)"""
        return self.save_style(name)

    # ========== HWP 문서 스타일 관리 ==========

    def create_document_style(self, style_name, char_style=None, para_style=None):
        """
        HWP 문서에 새 스타일 정의

        Args:
            style_name: 스타일 이름
            char_style: 글자 모양 딕셔너리 (없으면 현재 커서 위치 사용)
            para_style: 문단 모양 딕셔너리 (없으면 현재 커서 위치 사용)

        Note:
            HWP API의 CreateStyle 사용
        """
        self.hwp.HAction.GetDefault("Style", self.hwp.HParameterSet.HStyle.HSet)
        pset = self.hwp.HParameterSet.HStyle

        # 스타일 이름 설정
        pset.Name = style_name

        self.hwp.HAction.Execute("Style", self.hwp.HParameterSet.HStyle.HSet)

    def get_document_style_count(self):
        """문서에 정의된 스타일 개수"""
        return self.hwp.StyleCount

    def get_document_style_name(self, index):
        """인덱스로 스타일 이름 조회"""
        return self.hwp.StyleName(index)

    def get_all_document_styles(self):
        """문서에 정의된 모든 스타일 이름 목록"""
        count = self.hwp.StyleCount
        return [self.hwp.StyleName(i) for i in range(count)]

    def apply_document_style_by_name(self, style_name):
        """
        문서 스타일을 이름으로 적용

        Args:
            style_name: 적용할 스타일 이름
        """
        styles = self.get_all_document_styles()
        if style_name in styles:
            index = styles.index(style_name)
            self.apply_style(index)
            return True
        return False

    # ========== 글자 모양 (CharShape) ==========

    def set_bold(self, enable=True):
        """굵게 설정 (선택 영역에 적용)"""
        if enable:
            self.hwp.HAction.Run("CharShapeBold")
        else:
            # 굵게 해제: CharShape에서 Bold=0 설정
            self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
            self.hwp.HParameterSet.HCharShape.Bold = 0
            self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)

    def set_italic(self, enable=True):
        """기울임 설정"""
        if enable:
            self.hwp.HAction.Run("CharShapeItalic")
        else:
            self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
            self.hwp.HParameterSet.HCharShape.Italic = 0
            self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)

    def set_underline(self, underline_type=1, shape=0, color=None):
        """
        밑줄 설정

        Args:
            underline_type: 0=없음, 1=글자아래, 2=어절아래, 3=위쪽
            shape: 0=실선, 1=점선, 2=굵은실선 등
            color: BGR 색상값 또는 None(글자색)
        """
        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
        self.hwp.HParameterSet.HCharShape.UnderlineType = underline_type
        self.hwp.HParameterSet.HCharShape.UnderlineShape = shape
        if color is not None:
            self.hwp.HParameterSet.HCharShape.UnderlineColor = color
        self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)

    def set_strikeout(self, strikeout_type=1):
        """
        취소선 설정

        Args:
            strikeout_type: 0=없음, 1~7=다양한 스타일
        """
        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
        self.hwp.HParameterSet.HCharShape.StrikeOutType = strikeout_type
        self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)

    def set_text_color(self, color):
        """
        글자색 설정

        Args:
            color: BGR 색상값 또는 COLORS 딕셔너리 키
        """
        if isinstance(color, str):
            color = COLORS.get(color, 0x000000)

        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
        self.hwp.HParameterSet.HCharShape.TextColor = color
        self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)

    def set_font_size(self, pt):
        """글자 크기 설정 (포인트)"""
        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
        self.hwp.HParameterSet.HCharShape.Height = pt_to_hwpunit(pt)
        self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)

    def set_font(self, hangul=None, latin=None, hanja=None):
        """
        글꼴 설정

        Args:
            hangul: 한글 글꼴명
            latin: 영문 글꼴명
            hanja: 한자 글꼴명
        """
        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
        if hangul:
            self.hwp.HParameterSet.HCharShape.FaceNameHangul = hangul
        if latin:
            self.hwp.HParameterSet.HCharShape.FaceNameLatin = latin
        if hanja:
            self.hwp.HParameterSet.HCharShape.FaceNameHanja = hanja
        self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)

    def set_char_shape(self, **kwargs):
        """
        글자 모양 한번에 설정

        Args:
            bold: 굵게 (0/1)
            italic: 기울임 (0/1)
            underline: 밑줄 타입 (0~3)
            strikeout: 취소선 타입 (0~7)
            color: 글자색 (BGR)
            size: 글자 크기 (pt)
            font_hangul: 한글 글꼴
            font_latin: 영문 글꼴
            superscript: 위첨자 (0/1)
            subscript: 아래첨자 (0/1)
            spacing: 자간 (-50~50)
            ratio: 장평 (50~200)
        """
        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
        pset = self.hwp.HParameterSet.HCharShape

        if 'bold' in kwargs:
            pset.Bold = kwargs['bold']
        if 'italic' in kwargs:
            pset.Italic = kwargs['italic']
        if 'underline' in kwargs:
            pset.UnderlineType = kwargs['underline']
        if 'strikeout' in kwargs:
            pset.StrikeOutType = kwargs['strikeout']
        if 'color' in kwargs:
            color = kwargs['color']
            if isinstance(color, str):
                color = COLORS.get(color, 0x000000)
            pset.TextColor = color
        if 'size' in kwargs:
            pset.Height = pt_to_hwpunit(kwargs['size'])
        if 'font_hangul' in kwargs:
            pset.FaceNameHangul = kwargs['font_hangul']
        if 'font_latin' in kwargs:
            pset.FaceNameLatin = kwargs['font_latin']
        if 'superscript' in kwargs:
            pset.SuperScript = kwargs['superscript']
        if 'subscript' in kwargs:
            pset.SubScript = kwargs['subscript']
        if 'spacing' in kwargs:
            pset.SpacingHangul = kwargs['spacing']
            pset.SpacingLatin = kwargs['spacing']
        if 'ratio' in kwargs:
            pset.RatioHangul = kwargs['ratio']
            pset.RatioLatin = kwargs['ratio']

        self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)

    def reset_char_shape(self):
        """글자 모양 초기화 (보통 모양)"""
        self.hwp.HAction.Run("CharShapeNormal")

    # ========== 문단 모양 (ParaShape) ==========

    def set_align(self, align_type):
        """
        문단 정렬 설정

        Args:
            align_type: 'left', 'center', 'right', 'justify', 'distribute'
                       또는 숫자 (0=양쪽, 1=왼쪽, 2=오른쪽, 3=가운데, 4=배분, 5=나눔)
        """
        align_map = {
            'justify': 0, 'left': 1, 'right': 2, 'center': 3, 'distribute': 4, 'divide': 5
        }

        if isinstance(align_type, str):
            align_type = align_map.get(align_type, 0)

        action_map = {
            0: "ParagraphShapeAlignJustify",
            1: "ParagraphShapeAlignLeft",
            2: "ParagraphShapeAlignRight",
            3: "ParagraphShapeAlignCenter",
        }

        if align_type in action_map:
            self.hwp.HAction.Run(action_map[align_type])
        else:
            self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
            self.hwp.HParameterSet.HParaShape.AlignType = align_type
            self.hwp.HAction.Execute("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)

    def set_line_spacing(self, value, spacing_type=0):
        """
        줄간격 설정

        Args:
            value: 줄간격 값 (%, pt 등)
            spacing_type: 0=글자에따라(%), 1=고정값, 2=여백만지정, 3=최소
        """
        self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        self.hwp.HParameterSet.HParaShape.LineSpacingType = spacing_type
        self.hwp.HParameterSet.HParaShape.LineSpacing = value
        self.hwp.HAction.Execute("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)

    def set_para_margin(self, left=None, right=None, indent=None, before=None, after=None):
        """
        문단 여백/간격 설정 (포인트 단위)

        Args:
            left: 왼쪽 여백
            right: 오른쪽 여백
            indent: 첫줄 들여쓰기 (음수면 내어쓰기)
            before: 문단 위 간격
            after: 문단 아래 간격
        """
        self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        pset = self.hwp.HParameterSet.HParaShape

        if left is not None:
            pset.LeftMargin = pt_to_hwpunit(left)
        if right is not None:
            pset.RightMargin = pt_to_hwpunit(right)
        if indent is not None:
            pset.Indentation = pt_to_hwpunit(indent)
        if before is not None:
            pset.PrevSpacing = pt_to_hwpunit(before)
        if after is not None:
            pset.NextSpacing = pt_to_hwpunit(after)

        self.hwp.HAction.Execute("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)

    def set_para_shape(self, **kwargs):
        """
        문단 모양 한번에 설정

        Args:
            align: 정렬 (0~5 또는 문자열)
            line_spacing: 줄간격 값
            line_spacing_type: 줄간격 종류 (0~3)
            left_margin: 왼쪽 여백 (pt)
            right_margin: 오른쪽 여백 (pt)
            indent: 들여쓰기 (pt)
            before: 문단 위 간격 (pt)
            after: 문단 아래 간격 (pt)
        """
        self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        pset = self.hwp.HParameterSet.HParaShape

        if 'align' in kwargs:
            align = kwargs['align']
            align_map = {'justify': 0, 'left': 1, 'right': 2, 'center': 3, 'distribute': 4}
            if isinstance(align, str):
                align = align_map.get(align, 0)
            pset.AlignType = align

        if 'line_spacing' in kwargs:
            pset.LineSpacing = kwargs['line_spacing']
        if 'line_spacing_type' in kwargs:
            pset.LineSpacingType = kwargs['line_spacing_type']
        if 'left_margin' in kwargs:
            pset.LeftMargin = pt_to_hwpunit(kwargs['left_margin'])
        if 'right_margin' in kwargs:
            pset.RightMargin = pt_to_hwpunit(kwargs['right_margin'])
        if 'indent' in kwargs:
            pset.Indentation = pt_to_hwpunit(kwargs['indent'])
        if 'before' in kwargs:
            pset.PrevSpacing = pt_to_hwpunit(kwargs['before'])
        if 'after' in kwargs:
            pset.NextSpacing = pt_to_hwpunit(kwargs['after'])

        self.hwp.HAction.Execute("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)

    # ========== 서식 복사/붙여넣기 ==========

    def copy_char_shape(self):
        """글자 모양 복사 (현재 커서 위치)"""
        self.hwp.HAction.GetDefault("ShapeCopyPaste", self.hwp.HParameterSet.HShapeCopyPaste.HSet)
        self.hwp.HParameterSet.HShapeCopyPaste.Type = 0  # 글자 모양만
        self.hwp.HAction.Execute("ShapeCopyPaste", self.hwp.HParameterSet.HShapeCopyPaste.HSet)

    def copy_para_shape(self):
        """문단 모양 복사 (현재 커서 위치)"""
        self.hwp.HAction.GetDefault("ShapeCopyPaste", self.hwp.HParameterSet.HShapeCopyPaste.HSet)
        self.hwp.HParameterSet.HShapeCopyPaste.Type = 1  # 문단 모양만
        self.hwp.HAction.Execute("ShapeCopyPaste", self.hwp.HParameterSet.HShapeCopyPaste.HSet)

    def copy_all_shape(self):
        """글자 + 문단 모양 복사"""
        self.hwp.HAction.GetDefault("ShapeCopyPaste", self.hwp.HParameterSet.HShapeCopyPaste.HSet)
        self.hwp.HParameterSet.HShapeCopyPaste.Type = 2  # 글자 + 문단
        self.hwp.HAction.Execute("ShapeCopyPaste", self.hwp.HParameterSet.HShapeCopyPaste.HSet)

    def paste_shape(self, shape_type=0):
        """
        서식 붙여넣기 (선택 영역에 적용)

        Args:
            shape_type: 0=글자, 1=문단, 2=둘다
        """
        self.hwp.HAction.GetDefault("ShapeCopyPaste", self.hwp.HParameterSet.HShapeCopyPaste.HSet)
        self.hwp.HParameterSet.HShapeCopyPaste.Type = shape_type
        self.hwp.HAction.Execute("ShapeCopyPaste", self.hwp.HParameterSet.HShapeCopyPaste.HSet)

    # ========== 스타일 관리 ==========

    def apply_style(self, style_index):
        """
        스타일 적용 (인덱스로)

        Args:
            style_index: 스타일 인덱스 (0부터)
        """
        self.hwp.HAction.GetDefault("Style", self.hwp.HParameterSet.HStyle.HSet)
        self.hwp.HParameterSet.HStyle.Apply = style_index
        self.hwp.HAction.Execute("Style", self.hwp.HParameterSet.HStyle.HSet)

    def apply_style_shortcut(self, num):
        """
        스타일 단축키 적용 (Ctrl+1~0)

        Args:
            num: 1~10 (10은 Ctrl+0)
        """
        if 1 <= num <= 10:
            self.hwp.HAction.Run(f"StyleShortcut{num}")

    def clear_char_style(self):
        """글자 스타일 해제"""
        self.hwp.HAction.Run("StyleClearCharStyle")

    def clear_all_formatting(self):
        """
        모든 서식 제거 (글자 모양 + 문단 모양 초기화)
        """
        # 글자 모양 초기화
        self.hwp.HAction.Run("CharShapeNormal")

        # 문단 모양 초기화
        self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        pset = self.hwp.HParameterSet.HParaShape
        pset.AlignType = 0  # 양쪽 정렬
        pset.LeftMargin = 0
        pset.RightMargin = 0
        pset.Indentation = 0
        pset.PrevSpacing = 0
        pset.NextSpacing = 0
        pset.LineSpacingType = 0
        pset.LineSpacing = 160  # 기본 160%
        self.hwp.HAction.Execute("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)

    def clear_char_formatting(self):
        """글자 서식만 제거 (보통 모양으로)"""
        self.hwp.HAction.Run("CharShapeNormal")

    def clear_para_formatting(self):
        """문단 서식만 제거 (기본값으로)"""
        self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        pset = self.hwp.HParameterSet.HParaShape
        pset.AlignType = 0
        pset.LeftMargin = 0
        pset.RightMargin = 0
        pset.Indentation = 0
        pset.PrevSpacing = 0
        pset.NextSpacing = 0
        pset.LineSpacingType = 0
        pset.LineSpacing = 160
        self.hwp.HAction.Execute("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)

    # ========== 줄간격 / 문단 여백 증감 ==========

    def _get_current_para_shape(self):
        """현재 커서 위치의 문단 모양 정보 가져오기"""
        self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        pset = self.hwp.HParameterSet.HParaShape
        return {
            'line_spacing': pset.LineSpacing,
            'line_spacing_type': pset.LineSpacingType,
            'prev_spacing': pset.PrevSpacing,  # 문단 위 여백
            'next_spacing': pset.NextSpacing,  # 문단 아래 여백
        }

    def adjust_line_spacing(self, delta_percent):
        """
        줄간격 증감 (현재 값 기준 delta% 만큼 증감)

        Args:
            delta_percent: 증감량 (예: 5 = 5% 증가, -5 = 5% 감소)
        """
        self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        pset = self.hwp.HParameterSet.HParaShape
        current = pset.LineSpacing
        new_value = max(50, current + delta_percent)  # 최소 50%
        pset.LineSpacing = new_value
        self.hwp.HAction.Execute("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        return new_value

    def increase_line_spacing(self, percent=5):
        """줄간격 증가 (기본 5%)"""
        return self.adjust_line_spacing(percent)

    def decrease_line_spacing(self, percent=5):
        """줄간격 감소 (기본 5%)"""
        return self.adjust_line_spacing(-percent)

    def adjust_para_prev_spacing(self, delta_pt):
        """
        문단 위 여백(간격) 증감 (포인트 단위)

        Args:
            delta_pt: 증감량 (pt)
        """
        self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        pset = self.hwp.HParameterSet.HParaShape
        current = pset.PrevSpacing  # HWPUNIT
        new_value = max(0, current + pt_to_hwpunit(delta_pt))
        pset.PrevSpacing = new_value
        self.hwp.HAction.Execute("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        return hwpunit_to_pt(new_value)

    def adjust_para_next_spacing(self, delta_pt):
        """
        문단 아래 여백(간격) 증감 (포인트 단위)

        Args:
            delta_pt: 증감량 (pt)
        """
        self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        pset = self.hwp.HParameterSet.HParaShape
        current = pset.NextSpacing  # HWPUNIT
        new_value = max(0, current + pt_to_hwpunit(delta_pt))
        pset.NextSpacing = new_value
        self.hwp.HAction.Execute("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        return hwpunit_to_pt(new_value)

    def adjust_para_both_spacing(self, delta_pt):
        """
        문단 위/아래 여백 동시 증감 (포인트 단위)

        Args:
            delta_pt: 증감량 (pt)
        """
        self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        pset = self.hwp.HParameterSet.HParaShape
        prev_current = pset.PrevSpacing
        next_current = pset.NextSpacing
        new_prev = max(0, prev_current + pt_to_hwpunit(delta_pt))
        new_next = max(0, next_current + pt_to_hwpunit(delta_pt))
        pset.PrevSpacing = new_prev
        pset.NextSpacing = new_next
        self.hwp.HAction.Execute("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        return (hwpunit_to_pt(new_prev), hwpunit_to_pt(new_next))

    def increase_para_prev_spacing(self, pt=0.5):
        """문단 위 여백 증가 (기본 0.5pt)"""
        return self.adjust_para_prev_spacing(pt)

    def decrease_para_prev_spacing(self, pt=0.5):
        """문단 위 여백 감소 (기본 0.5pt)"""
        return self.adjust_para_prev_spacing(-pt)

    def increase_para_next_spacing(self, pt=0.5):
        """문단 아래 여백 증가 (기본 0.5pt)"""
        return self.adjust_para_next_spacing(pt)

    def decrease_para_next_spacing(self, pt=0.5):
        """문단 아래 여백 감소 (기본 0.5pt)"""
        return self.adjust_para_next_spacing(-pt)

    def increase_para_both_spacing(self, pt=0.5):
        """문단 위/아래 여백 동시 증가 (기본 0.5pt)"""
        return self.adjust_para_both_spacing(pt)

    def decrease_para_both_spacing(self, pt=0.5):
        """문단 위/아래 여백 동시 감소 (기본 0.5pt)"""
        return self.adjust_para_both_spacing(-pt)

    # ========== 5% 비율 기반 증감 함수 ==========

    def adjust_line_spacing_by_percent(self, ratio_percent):
        """
        줄간격을 현재 값의 비율로 증감

        Args:
            ratio_percent: 비율 (예: 5 = 현재값의 5% 증가, -5 = 5% 감소)
        """
        self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        pset = self.hwp.HParameterSet.HParaShape
        current = pset.LineSpacing
        delta = current * ratio_percent / 100
        new_value = max(50, int(current + delta))
        pset.LineSpacing = new_value
        self.hwp.HAction.Execute("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        return new_value

    def increase_line_spacing_5percent(self):
        """줄간격 5% 증가 (현재값 기준)"""
        return self.adjust_line_spacing_by_percent(5)

    def decrease_line_spacing_5percent(self):
        """줄간격 5% 감소 (현재값 기준)"""
        return self.adjust_line_spacing_by_percent(-5)

    def adjust_para_prev_spacing_by_percent(self, ratio_percent, min_value_pt=0.5):
        """
        문단 위 여백을 현재 값의 비율로 증감

        Args:
            ratio_percent: 비율 (예: 5 = 현재값의 5% 증가)
            min_value_pt: 최소 증감폭 (pt) - 현재값이 0에 가까울 때 사용
        """
        self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        pset = self.hwp.HParameterSet.HParaShape
        current = pset.PrevSpacing
        if current == 0:
            delta = pt_to_hwpunit(min_value_pt) if ratio_percent > 0 else 0
        else:
            delta = current * ratio_percent / 100
            if abs(delta) < pt_to_hwpunit(min_value_pt):
                delta = pt_to_hwpunit(min_value_pt) * (1 if ratio_percent > 0 else -1)
        new_value = max(0, int(current + delta))
        pset.PrevSpacing = new_value
        self.hwp.HAction.Execute("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        return hwpunit_to_pt(new_value)

    def adjust_para_next_spacing_by_percent(self, ratio_percent, min_value_pt=0.5):
        """
        문단 아래 여백을 현재 값의 비율로 증감

        Args:
            ratio_percent: 비율 (예: 5 = 현재값의 5% 증가)
            min_value_pt: 최소 증감폭 (pt)
        """
        self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        pset = self.hwp.HParameterSet.HParaShape
        current = pset.NextSpacing
        if current == 0:
            delta = pt_to_hwpunit(min_value_pt) if ratio_percent > 0 else 0
        else:
            delta = current * ratio_percent / 100
            if abs(delta) < pt_to_hwpunit(min_value_pt):
                delta = pt_to_hwpunit(min_value_pt) * (1 if ratio_percent > 0 else -1)
        new_value = max(0, int(current + delta))
        pset.NextSpacing = new_value
        self.hwp.HAction.Execute("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        return hwpunit_to_pt(new_value)

    def adjust_para_both_spacing_by_percent(self, ratio_percent, min_value_pt=0.5):
        """
        문단 위/아래 여백을 현재 값의 비율로 동시 증감

        Args:
            ratio_percent: 비율 (예: 5 = 현재값의 5% 증가)
            min_value_pt: 최소 증감폭 (pt)
        """
        self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        pset = self.hwp.HParameterSet.HParaShape

        # 위 여백
        prev_current = pset.PrevSpacing
        if prev_current == 0:
            prev_delta = pt_to_hwpunit(min_value_pt) if ratio_percent > 0 else 0
        else:
            prev_delta = prev_current * ratio_percent / 100
            if abs(prev_delta) < pt_to_hwpunit(min_value_pt):
                prev_delta = pt_to_hwpunit(min_value_pt) * (1 if ratio_percent > 0 else -1)

        # 아래 여백
        next_current = pset.NextSpacing
        if next_current == 0:
            next_delta = pt_to_hwpunit(min_value_pt) if ratio_percent > 0 else 0
        else:
            next_delta = next_current * ratio_percent / 100
            if abs(next_delta) < pt_to_hwpunit(min_value_pt):
                next_delta = pt_to_hwpunit(min_value_pt) * (1 if ratio_percent > 0 else -1)

        new_prev = max(0, int(prev_current + prev_delta))
        new_next = max(0, int(next_current + next_delta))
        pset.PrevSpacing = new_prev
        pset.NextSpacing = new_next
        self.hwp.HAction.Execute("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        return (hwpunit_to_pt(new_prev), hwpunit_to_pt(new_next))

    def increase_para_prev_spacing_5percent(self):
        """문단 위 여백 5% 증가"""
        return self.adjust_para_prev_spacing_by_percent(5)

    def decrease_para_prev_spacing_5percent(self):
        """문단 위 여백 5% 감소"""
        return self.adjust_para_prev_spacing_by_percent(-5)

    def increase_para_next_spacing_5percent(self):
        """문단 아래 여백 5% 증가"""
        return self.adjust_para_next_spacing_by_percent(5)

    def decrease_para_next_spacing_5percent(self):
        """문단 아래 여백 5% 감소"""
        return self.adjust_para_next_spacing_by_percent(-5)

    def increase_para_both_spacing_5percent(self):
        """문단 위/아래 여백 5% 동시 증가"""
        return self.adjust_para_both_spacing_by_percent(5)

    def decrease_para_both_spacing_5percent(self):
        """문단 위/아래 여백 5% 동시 감소"""
        return self.adjust_para_both_spacing_by_percent(-5)

    # ========== 편의 메서드 ==========

    def select_all(self):
        """전체 선택"""
        self.hwp.HAction.Run("SelectAll")

    def cancel_selection(self):
        """선택 해제"""
        self.hwp.HAction.Run("Cancel")

    def select_paragraph(self):
        """현재 문단 선택"""
        self.hwp.HAction.Run("MoveParaBegin")
        self.hwp.HAction.Run("MoveSelParaEnd")

    def select_line(self):
        """현재 줄 선택"""
        self.hwp.HAction.Run("MoveLineBegin")
        self.hwp.HAction.Run("MoveSelLineEnd")

    # ========== 커서 위치 줄/문단 여백 조회 ==========

    def get_current_line_spacing_info(self):
        """
        현재 커서가 있는 줄의 여백 정보 조회

        Returns:
            dict: {
                'line_spacing': 줄간격 값,
                'line_spacing_type': 줄간격 종류 (0=글자에따라, 1=고정값, 2=여백만, 3=최소),
                'line_spacing_type_name': 줄간격 종류 이름,
                'prev_spacing_pt': 문단 위 여백 (pt),
                'next_spacing_pt': 문단 아래 여백 (pt),
                'position': 현재 위치 정보
            }
        """
        pos = self.hwp.GetPos()
        key = self.hwp.KeyIndicator()

        self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        pset = self.hwp.HParameterSet.HParaShape

        spacing_type_names = {0: '글자에따라(%)', 1: '고정값', 2: '여백만지정', 3: '최소'}

        return {
            'line_spacing': pset.LineSpacing,
            'line_spacing_type': pset.LineSpacingType,
            'line_spacing_type_name': spacing_type_names.get(pset.LineSpacingType, '알수없음'),
            'prev_spacing_pt': pset.PrevSpacing / 100,
            'next_spacing_pt': pset.NextSpacing / 100,
            'position': {
                'para_id': pos[1],
                'char_pos': pos[2],
                'page': key[2],
                'line': key[4],
                'column': key[5]
            }
        }

    def get_current_para_spacing_info(self):
        """
        현재 커서가 있는 문단의 여백 정보 조회

        Returns:
            dict: {
                'prev_spacing_pt': 문단 위 여백 (pt),
                'next_spacing_pt': 문단 아래 여백 (pt),
                'left_margin_pt': 왼쪽 여백 (pt),
                'right_margin_pt': 오른쪽 여백 (pt),
                'indent_pt': 들여쓰기 (pt),
                'line_spacing': 줄간격,
                'align': 정렬,
                'align_name': 정렬 이름
            }
        """
        self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        pset = self.hwp.HParameterSet.HParaShape

        align_names = {0: '양쪽', 1: '왼쪽', 2: '오른쪽', 3: '가운데', 4: '배분', 5: '나눔'}

        return {
            'prev_spacing_pt': pset.PrevSpacing / 100,
            'next_spacing_pt': pset.NextSpacing / 100,
            'left_margin_pt': pset.LeftMargin / 100,
            'right_margin_pt': pset.RightMargin / 100,
            'indent_pt': pset.Indentation / 100,
            'line_spacing': pset.LineSpacing,
            'line_spacing_type': pset.LineSpacingType,
            'align': pset.AlignType,
            'align_name': align_names.get(pset.AlignType, '알수없음')
        }

    def print_current_line_info(self):
        """현재 줄의 여백 정보 출력"""
        info = self.get_current_line_spacing_info()
        print("=" * 40)
        print("[ 현재 줄 여백 정보 ]")
        print(f"  위치: 문단 {info['position']['para_id']}, "
              f"페이지 {info['position']['page']}, 줄 {info['position']['line']}")
        print(f"  줄간격: {info['line_spacing']}% ({info['line_spacing_type_name']})")
        print(f"  문단 위 여백: {info['prev_spacing_pt']}pt")
        print(f"  문단 아래 여백: {info['next_spacing_pt']}pt")
        print("=" * 40)

    def print_current_para_info(self):
        """현재 문단의 여백 정보 출력"""
        info = self.get_current_para_spacing_info()
        print("=" * 40)
        print("[ 현재 문단 여백 정보 ]")
        print(f"  정렬: {info['align_name']}")
        print(f"  줄간격: {info['line_spacing']}%")
        print(f"  문단 위 여백: {info['prev_spacing_pt']}pt")
        print(f"  문단 아래 여백: {info['next_spacing_pt']}pt")
        print(f"  왼쪽 여백: {info['left_margin_pt']}pt")
        print(f"  오른쪽 여백: {info['right_margin_pt']}pt")
        print(f"  들여쓰기: {info['indent_pt']}pt")
        print("=" * 40)

    # ========== 커서 위치 스타일 정보 보기 ==========

    def get_cursor_style_info(self):
        """
        현재 커서 위치의 전체 스타일 정보 조회

        Returns:
            dict: {
                'char': 글자 모양 정보,
                'para': 문단 모양 정보,
                'position': 위치 정보,
                'document_style': 문서 스타일 정보 (있을 경우)
            }
        """
        pos = self.hwp.GetPos()
        key = self.hwp.KeyIndicator()

        return {
            'char': get_char_shape(self.hwp),
            'para': get_para_shape(self.hwp),
            'position': {
                'list_id': pos[0],
                'para_id': pos[1],
                'char_pos': pos[2],
                'page': key[2],
                'line': key[4],
                'column': key[5]
            }
        }

    def print_cursor_style_info(self):
        """현재 커서 위치의 스타일 정보를 상세 출력"""
        info = self.get_cursor_style_info()
        char = info['char']
        para = info['para']
        pos = info['position']

        print("=" * 60)
        print("[ 커서 위치 스타일 정보 ]")
        print("=" * 60)

        print("\n▶ 위치 정보")
        print(f"   문단: {pos['para_id']}, 위치: {pos['char_pos']}")
        print(f"   페이지: {pos['page']}, 줄: {pos['line']}, 칸: {pos['column']}")

        print("\n▶ 글자 모양")
        print(f"   글꼴(한글): {char['font_hangul']}")
        print(f"   글꼴(영문): {char['font_latin']}")
        print(f"   크기: {char['size_pt']}pt")
        print(f"   굵게: {'예' if char['bold'] else '아니오'}")
        print(f"   기울임: {'예' if char['italic'] else '아니오'}")
        print(f"   밑줄: {char['underline']} (0=없음)")
        print(f"   취소선: {char['strikeout']} (0=없음)")
        print(f"   글자색(RGB): {char['text_color_rgb']}")
        print(f"   자간: {char['spacing']}")
        print(f"   장평: {char['ratio']}%")
        print(f"   위첨자: {'예' if char['superscript'] else '아니오'}")
        print(f"   아래첨자: {'예' if char['subscript'] else '아니오'}")

        print("\n▶ 문단 모양")
        print(f"   정렬: {para['align_name']}")
        print(f"   줄간격: {para['line_spacing']}% ({para['line_spacing_type_name']})")
        print(f"   문단 위 여백: {para['prev_spacing_pt']}pt")
        print(f"   문단 아래 여백: {para['next_spacing_pt']}pt")
        print(f"   왼쪽 여백: {para['left_margin_pt']}pt")
        print(f"   오른쪽 여백: {para['right_margin_pt']}pt")
        print(f"   들여쓰기: {para['indent_pt']}pt")

        print("=" * 60)

    def print_saved_style_info(self, name):
        """저장된 스타일 정보 출력"""
        style = self.get_saved_style(name)
        if not style:
            print(f"'{name}' 스타일이 저장되어 있지 않습니다.")
            return

        print("=" * 60)
        print(f"[ 저장된 스타일: {name} ]")
        print("=" * 60)

        if 'char' in style:
            char = style['char']
            print("\n▶ 글자 모양")
            print(f"   글꼴(한글): {char.get('font_hangul', '-')}")
            print(f"   글꼴(영문): {char.get('font_latin', '-')}")
            print(f"   크기: {char.get('size_pt', '-')}pt")
            print(f"   굵게: {'예' if char.get('bold') else '아니오'}")
            print(f"   기울임: {'예' if char.get('italic') else '아니오'}")
            print(f"   글자색(RGB): {char.get('text_color_rgb', '-')}")

        if 'para' in style:
            para = style['para']
            print("\n▶ 문단 모양")
            print(f"   정렬: {para.get('align_name', '-')}")
            print(f"   줄간격: {para.get('line_spacing', '-')}%")
            print(f"   문단 위 여백: {para.get('prev_spacing_pt', '-')}pt")
            print(f"   문단 아래 여백: {para.get('next_spacing_pt', '-')}pt")

        print("=" * 60)


if __name__ == "__main__":
    hwp = get_hwp_instance()
    if hwp is None:
        print("한글이 실행 중이지 않습니다.")
        exit(1)

    sf = StyleFormat(hwp)

    print("=== StyleFormat 테스트 ===\n")

    # 새 문서
    hwp.HAction.Run("FileNew")

    # 테스트 텍스트 삽입
    for text in ["첫 번째 문단입니다.", "두 번째 문단입니다.", "세 번째 문단입니다."]:
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = text
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HAction.Run("BreakPara")

    # 테스트 1: 첫 문단 굵게 + 빨강
    print("[1] 첫 문단: 굵게 + 빨강")
    hwp.HAction.Run("MoveDocBegin")
    sf.select_paragraph()
    sf.set_char_shape(bold=1, color='red')
    sf.cancel_selection()

    # 테스트 2: 두 번째 문단 가운데 정렬 + 파랑
    print("[2] 두 번째 문단: 가운데 정렬 + 파랑")
    hwp.HAction.Run("MoveDown")
    sf.select_paragraph()
    sf.set_align('center')
    sf.set_text_color('blue')
    sf.cancel_selection()

    # 테스트 3: 서식 복사 후 세 번째 문단에 붙여넣기
    print("[3] 서식 복사 -> 세 번째 문단에 붙여넣기")
    hwp.HAction.Run("MoveDocBegin")
    sf.copy_all_shape()
    hwp.HAction.Run("MoveDown")
    hwp.HAction.Run("MoveDown")
    sf.select_paragraph()
    sf.paste_shape(2)
    sf.cancel_selection()

    print("\n테스트 완료! 문서를 확인하세요.")
