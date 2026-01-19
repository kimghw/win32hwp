# -*- coding: utf-8 -*-
"""
한글(HWP) 스타일 관리 클래스

글자 모양(CharShape), 문단 모양(ParaShape) 설정 및 조회 기능 제공.

주요 기능:
1. 글자 모양: 굵게, 기울임, 밑줄, 글꼴, 크기, 색상
2. 문단 모양: 정렬, 줄간격, 여백
3. 모양 복사/붙여넣기
4. 유틸리티: RGB-BGR 변환, pt-HWPUNIT 변환

사용 예:
    from style_para import StylePara
    from cursor_utils import get_hwp_instance

    hwp = get_hwp_instance()
    style = StylePara(hwp)

    # 글자 모양 설정
    style.set_bold(True)
    style.set_font_size(12)
    style.set_text_color(255, 0, 0)  # 빨강

    # 문단 모양 설정
    style.set_align('center')
    style.set_line_spacing(160)
"""

import os
import yaml

from cursor_utils import get_hwp_instance

# 기본 스타일 파일 경로
DEFAULT_STYLES_PATH = os.path.join(os.path.dirname(__file__), 'styles.yaml')


class StylePara:
    """HWP 스타일(글자 모양, 문단 모양) 관리 클래스"""

    # 정렬 타입 매핑
    ALIGN_TYPES = {
        'justify': 0,   # 양쪽 정렬
        'left': 1,      # 왼쪽 정렬
        'right': 2,     # 오른쪽 정렬
        'center': 3,    # 가운데 정렬
        'distribute': 4, # 배분 정렬
        'divide': 5     # 나눔 정렬
    }

    # 줄간격 타입 매핑
    LINE_SPACING_TYPES = {
        'percent': 0,   # 글자에 따라 (%)
        'fixed': 1,     # 고정값
        'spacing': 2,   # 여백만 지정
        'minimum': 3    # 최소
    }

    def __init__(self, hwp=None):
        """
        StylePara 초기화

        Args:
            hwp: 한글 인스턴스 (None이면 get_hwp_instance() 호출)
        """
        self.hwp = hwp if hwp else get_hwp_instance()
        if not self.hwp:
            raise RuntimeError("한글 인스턴스에 연결할 수 없습니다.")

    # ========== 유틸리티 함수 ==========

    @staticmethod
    def rgb_to_bgr(r, g, b):
        """
        RGB 값을 HWP COLORREF(BGR) 형식으로 변환

        Args:
            r: 빨강 (0-255)
            g: 초록 (0-255)
            b: 파랑 (0-255)

        Returns:
            int: BGR 형식 색상값 (0x00BBGGRR)
        """
        return (b << 16) | (g << 8) | r

    @staticmethod
    def pt_to_hwpunit(pt):
        """
        포인트를 HWPUNIT으로 변환

        Args:
            pt: 포인트 값

        Returns:
            int: HWPUNIT 값 (1pt = 100 HWPUNIT)
        """
        return int(pt * 100)

    # ========== 글자 모양 (CharShape) ==========

    def set_bold(self, enabled=True):
        """
        굵게 설정

        Args:
            enabled: True=굵게, False=보통
        """
        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
        self.hwp.HParameterSet.HCharShape.Bold = 1 if enabled else 0
        self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)

    def set_italic(self, enabled=True):
        """
        기울임 설정

        Args:
            enabled: True=기울임, False=보통
        """
        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
        self.hwp.HParameterSet.HCharShape.Italic = 1 if enabled else 0
        self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)

    def set_underline(self, enabled=True, line_type=1, shape=0, color=None):
        """
        밑줄 설정

        Args:
            enabled: True=밑줄, False=없음
            line_type: 밑줄 위치 (1=글자아래, 2=어절아래, 3=위쪽)
            shape: 밑줄 모양 (0=실선, 1=점선, 2=굵은실선 등)
            color: (r, g, b) 튜플 또는 None(글자색)
        """
        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
        self.hwp.HParameterSet.HCharShape.UnderlineType = line_type if enabled else 0
        if enabled:
            self.hwp.HParameterSet.HCharShape.UnderlineShape = shape
            if color:
                self.hwp.HParameterSet.HCharShape.UnderlineColor = self.rgb_to_bgr(*color)
        self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)

    def set_font_size(self, pt):
        """
        글자 크기 설정

        Args:
            pt: 포인트 크기 (예: 10, 12, 14)
        """
        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
        self.hwp.HParameterSet.HCharShape.Height = self.pt_to_hwpunit(pt)
        self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)

    def set_font(self, hangul=None, latin=None):
        """
        글꼴 설정

        Args:
            hangul: 한글 글꼴명 (예: "맑은 고딕")
            latin: 영문 글꼴명 (예: "Arial")
        """
        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
        if hangul:
            self.hwp.HParameterSet.HCharShape.FaceNameHangul = hangul
        if latin:
            self.hwp.HParameterSet.HCharShape.FaceNameLatin = latin
        self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)

    def set_text_color(self, r, g, b):
        """
        글자 색상 설정

        Args:
            r: 빨강 (0-255)
            g: 초록 (0-255)
            b: 파랑 (0-255)
        """
        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
        self.hwp.HParameterSet.HCharShape.TextColor = self.rgb_to_bgr(r, g, b)
        self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)

    def set_strikeout(self, enabled=True, strikeout_type=1, shape=0, color=None):
        """
        취소선 설정

        Args:
            enabled: True=취소선, False=없음
            strikeout_type: 1=단선, 2=이중선
            shape: 선 모양
            color: (r, g, b) 튜플 또는 None
        """
        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
        self.hwp.HParameterSet.HCharShape.StrikeOutType = strikeout_type if enabled else 0
        if enabled:
            self.hwp.HParameterSet.HCharShape.StrikeOutShape = shape
            if color:
                self.hwp.HParameterSet.HCharShape.StrikeOutColor = self.rgb_to_bgr(*color)
        self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)

    def set_outline(self, outline_type=0):
        """
        외곽선 설정

        Args:
            outline_type: 0=없음, 1=실선, 2=점선, 3=굵은실선
        """
        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
        self.hwp.HParameterSet.HCharShape.OutlineType = outline_type
        self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)

    def set_shadow(self, shadow_type=0, offset_x=10, offset_y=10, color=None):
        """
        그림자 설정

        Args:
            shadow_type: 0=없음, 1=비연속, 2=연속
            offset_x: X 오프셋 (%)
            offset_y: Y 오프셋 (%)
            color: (r, g, b) 튜플 또는 None
        """
        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
        self.hwp.HParameterSet.HCharShape.ShadowType = shadow_type
        if shadow_type > 0:
            self.hwp.HParameterSet.HCharShape.ShadowOffsetX = offset_x
            self.hwp.HParameterSet.HCharShape.ShadowOffsetY = offset_y
            if color:
                self.hwp.HParameterSet.HCharShape.ShadowColor = self.rgb_to_bgr(*color)
        self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)

    def set_emboss(self, enabled=True):
        """양각 설정"""
        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
        self.hwp.HParameterSet.HCharShape.Emboss = 1 if enabled else 0
        self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)

    def set_engrave(self, enabled=True):
        """음각 설정"""
        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
        self.hwp.HParameterSet.HCharShape.Engrave = 1 if enabled else 0
        self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)

    def set_superscript(self, enabled=True):
        """위첨자 설정"""
        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
        self.hwp.HParameterSet.HCharShape.SuperScript = 1 if enabled else 0
        self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)

    def set_subscript(self, enabled=True):
        """아래첨자 설정"""
        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
        self.hwp.HParameterSet.HCharShape.SubScript = 1 if enabled else 0
        self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)

    def set_char_spacing(self, spacing):
        """
        자간 설정

        Args:
            spacing: 자간 (%, 음수=좁게, 양수=넓게)
        """
        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
        self.hwp.HParameterSet.HCharShape.CharSpacing = spacing
        self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)

    def set_char_ratio(self, ratio):
        """
        장평 설정

        Args:
            ratio: 장평 (%, 100=기본)
        """
        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
        self.hwp.HParameterSet.HCharShape.CharRatio = ratio
        self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)

    def set_char_offset(self, offset):
        """
        글자 위치 (상하) 설정

        Args:
            offset: 오프셋 (%, 양수=위, 음수=아래)
        """
        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
        self.hwp.HParameterSet.HCharShape.CharOffset = offset
        self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)

    def set_all_fonts(self, hangul=None, latin=None, hanja=None, japanese=None, other=None, symbol=None):
        """
        모든 언어 글꼴 설정

        Args:
            hangul: 한글 글꼴
            latin: 영문 글꼴
            hanja: 한자 글꼴
            japanese: 일본어 글꼴
            other: 기타 글꼴
            symbol: 기호 글꼴
        """
        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
        cs = self.hwp.HParameterSet.HCharShape
        if hangul:
            cs.FaceNameHangul = hangul
        if latin:
            cs.FaceNameLatin = latin
        if hanja:
            cs.FaceNameHanja = hanja
        if japanese:
            cs.FaceNameJapanese = japanese
        if other:
            cs.FaceNameOther = other
        if symbol:
            cs.FaceNameSymbol = symbol
        self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)

    def set_shade_color(self, r, g, b):
        """음영색 설정"""
        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
        self.hwp.HParameterSet.HCharShape.ShadeColor = self.rgb_to_bgr(r, g, b)
        self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)

    def get_char_shape(self):
        """
        현재 글자 모양 조회

        Returns:
            dict: {
                'font_hangul': 한글 글꼴,
                'font_latin': 영문 글꼴,
                'height': 글자 크기 (HWPUNIT),
                'height_pt': 글자 크기 (pt),
                'bold': 굵게 여부,
                'italic': 기울임 여부,
                'underline_type': 밑줄 종류,
                'text_color': 글자색 (BGR)
            }
        """
        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
        cs = self.hwp.HParameterSet.HCharShape

        return {
            'font_hangul': cs.FaceNameHangul,
            'font_latin': cs.FaceNameLatin,
            'height': cs.Height,
            'height_pt': cs.Height / 100,
            'bold': bool(cs.Bold),
            'italic': bool(cs.Italic),
            'underline_type': cs.UnderlineType,
            'text_color': cs.TextColor
        }

    # ========== 문단 모양 (ParaShape) ==========

    def set_align(self, align_type):
        """
        문단 정렬 설정

        Args:
            align_type: 정렬 방식
                - 'left': 왼쪽 정렬
                - 'center': 가운데 정렬
                - 'right': 오른쪽 정렬
                - 'justify': 양쪽 정렬
                - 'distribute': 배분 정렬
                - 'divide': 나눔 정렬
        """
        if align_type not in self.ALIGN_TYPES:
            raise ValueError(f"잘못된 정렬 타입: {align_type}. "
                           f"가능한 값: {list(self.ALIGN_TYPES.keys())}")

        self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        self.hwp.HParameterSet.HParaShape.AlignType = self.ALIGN_TYPES[align_type]
        self.hwp.HAction.Execute("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)

    def set_line_spacing(self, value, spacing_type='percent'):
        """
        줄간격 설정

        Args:
            value: 줄간격 값 (percent 타입이면 160=160%)
            spacing_type: 줄간격 종류
                - 'percent': 글자에 따라 (%)
                - 'fixed': 고정값 (pt)
                - 'spacing': 여백만 지정
                - 'minimum': 최소
        """
        if spacing_type not in self.LINE_SPACING_TYPES:
            raise ValueError(f"잘못된 줄간격 타입: {spacing_type}. "
                           f"가능한 값: {list(self.LINE_SPACING_TYPES.keys())}")

        self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        self.hwp.HParameterSet.HParaShape.LineSpacingType = self.LINE_SPACING_TYPES[spacing_type]

        # fixed 타입이면 pt를 HWPUNIT으로 변환
        if spacing_type == 'fixed':
            self.hwp.HParameterSet.HParaShape.LineSpacing = self.pt_to_hwpunit(value)
        else:
            self.hwp.HParameterSet.HParaShape.LineSpacing = value

        self.hwp.HAction.Execute("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)

    def set_para_margin(self, left=None, right=None, indent=None):
        """
        문단 여백 설정

        Args:
            left: 왼쪽 여백 (pt)
            right: 오른쪽 여백 (pt)
            indent: 첫줄 들여쓰기 (pt, 음수면 내어쓰기)
        """
        self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)

        if left is not None:
            self.hwp.HParameterSet.HParaShape.LeftMargin = self.pt_to_hwpunit(left)
        if right is not None:
            self.hwp.HParameterSet.HParaShape.RightMargin = self.pt_to_hwpunit(right)
        if indent is not None:
            self.hwp.HParameterSet.HParaShape.Indentation = self.pt_to_hwpunit(indent)

        self.hwp.HAction.Execute("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)

    def set_para_spacing(self, before=None, after=None):
        """
        문단 간격 설정

        Args:
            before: 문단 앞 간격 (pt)
            after: 문단 뒤 간격 (pt)
        """
        self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)

        if before is not None:
            self.hwp.HParameterSet.HParaShape.SpaceBeforePara = self.pt_to_hwpunit(before)
        if after is not None:
            self.hwp.HParameterSet.HParaShape.SpaceAfterPara = self.pt_to_hwpunit(after)

        self.hwp.HAction.Execute("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)

    def set_widow_orphan(self, enabled=True):
        """외톨이줄 보호 설정"""
        self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        self.hwp.HParameterSet.HParaShape.WidowOrphan = 1 if enabled else 0
        self.hwp.HAction.Execute("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)

    def set_keep_with_next(self, enabled=True):
        """다음 문단과 함께 설정"""
        self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        self.hwp.HParameterSet.HParaShape.KeepWithNext = 1 if enabled else 0
        self.hwp.HAction.Execute("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)

    def set_keep_lines(self, enabled=True):
        """문단 분리 금지 설정"""
        self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        self.hwp.HParameterSet.HParaShape.KeepLines = 1 if enabled else 0
        self.hwp.HAction.Execute("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)

    def set_page_break_before(self, enabled=True):
        """문단 앞에서 페이지 나눔 설정"""
        self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        self.hwp.HParameterSet.HParaShape.PageBreakBefore = 1 if enabled else 0
        self.hwp.HAction.Execute("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)

    def set_snap_to_grid(self, enabled=True):
        """줄 격자에 맞춤 설정"""
        self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        self.hwp.HParameterSet.HParaShape.SnapToGrid = 1 if enabled else 0
        self.hwp.HAction.Execute("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)

    def set_break_latin_word(self, enabled=True):
        """영어 단어 나눔 설정"""
        self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        self.hwp.HParameterSet.HParaShape.BreakLatinWord = 1 if enabled else 0
        self.hwp.HAction.Execute("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)

    def set_break_non_latin_word(self, enabled=True):
        """한글 단어 나눔 설정"""
        self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        self.hwp.HParameterSet.HParaShape.BreakNonLatinWord = 1 if enabled else 0
        self.hwp.HAction.Execute("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)

    def get_para_shape(self):
        """
        현재 문단 모양 조회

        Returns:
            dict: {
                'align_type': 정렬 타입 (0-5),
                'align_name': 정렬 이름,
                'line_spacing': 줄간격 값,
                'line_spacing_type': 줄간격 타입 (0-3),
                'left_margin': 왼쪽 여백 (HWPUNIT),
                'right_margin': 오른쪽 여백 (HWPUNIT),
                'indentation': 들여쓰기 (HWPUNIT)
            }
        """
        self.hwp.HAction.GetDefault("ParagraphShape", self.hwp.HParameterSet.HParaShape.HSet)
        ps = self.hwp.HParameterSet.HParaShape

        # 정렬 타입 이름 찾기
        align_name = None
        for name, value in self.ALIGN_TYPES.items():
            if value == ps.AlignType:
                align_name = name
                break

        return {
            'align_type': ps.AlignType,
            'align_name': align_name,
            'line_spacing': ps.LineSpacing,
            'line_spacing_type': ps.LineSpacingType,
            'left_margin': ps.LeftMargin,
            'right_margin': ps.RightMargin,
            'indentation': ps.Indentation
        }

    # ========== 모양 복사/붙여넣기 ==========

    def copy_char_shape(self):
        """글자 모양 복사 (클립보드에 저장)"""
        self.hwp.HAction.GetDefault("ShapeCopyPaste", self.hwp.HParameterSet.HShapeCopyPaste.HSet)
        self.hwp.HParameterSet.HShapeCopyPaste.Type = 0  # 글자 모양
        self.hwp.HAction.Execute("ShapeCopyPaste", self.hwp.HParameterSet.HShapeCopyPaste.HSet)

    def copy_para_shape(self):
        """문단 모양 복사 (클립보드에 저장)"""
        self.hwp.HAction.GetDefault("ShapeCopyPaste", self.hwp.HParameterSet.HShapeCopyPaste.HSet)
        self.hwp.HParameterSet.HShapeCopyPaste.Type = 1  # 문단 모양
        self.hwp.HAction.Execute("ShapeCopyPaste", self.hwp.HParameterSet.HShapeCopyPaste.HSet)

    def paste_shape(self):
        """복사된 모양 붙여넣기 (선택 영역에 적용)"""
        self.hwp.HAction.GetDefault("ShapeCopyPaste", self.hwp.HParameterSet.HShapeCopyPaste.HSet)
        self.hwp.HAction.Execute("ShapeCopyPaste", self.hwp.HParameterSet.HShapeCopyPaste.HSet)

    # ========== YAML 스타일 로드/적용 ==========

    def load_styles(self, yaml_path=None):
        """
        YAML 파일에서 스타일 프리셋 로드

        Args:
            yaml_path: YAML 파일 경로 (None이면 기본 styles.yaml)

        Returns:
            dict: 로드된 스타일 데이터
        """
        path = yaml_path or DEFAULT_STYLES_PATH
        with open(path, 'r', encoding='utf-8') as f:
            self._styles = yaml.safe_load(f)
        return self._styles

    def get_style_names(self, style_type='char'):
        """
        사용 가능한 스타일 이름 목록 반환

        Args:
            style_type: 'char', 'para', 'combined' 중 하나

        Returns:
            list: 스타일 이름 목록
        """
        if not hasattr(self, '_styles'):
            self.load_styles()

        key_map = {
            'char': 'char_styles',
            'para': 'para_styles',
            'combined': 'combined_styles'
        }
        key = key_map.get(style_type, 'char_styles')
        return list(self._styles.get(key, {}).keys())

    def apply_char_style(self, style_name):
        """
        YAML에 정의된 글자 스타일 적용

        Args:
            style_name: 스타일 이름 (예: 'body', 'heading1', 'emphasis')
        """
        if not hasattr(self, '_styles'):
            self.load_styles()

        style = self._styles.get('char_styles', {}).get(style_name)
        if not style:
            raise ValueError(f"글자 스타일 '{style_name}'을 찾을 수 없습니다.")

        # 글꼴 설정
        self.set_all_fonts(
            hangul=style.get('font_hangul'),
            latin=style.get('font_latin'),
            hanja=style.get('font_hanja'),
            japanese=style.get('font_japanese'),
            other=style.get('font_other'),
            symbol=style.get('font_symbol')
        )

        # 크기 설정
        if 'size_pt' in style:
            self.set_font_size(style['size_pt'])

        # 장평/자간/위치
        if 'ratio' in style:
            self.set_char_ratio(style['ratio'])
        if 'spacing' in style:
            self.set_char_spacing(style['spacing'])
        if 'offset' in style:
            self.set_char_offset(style['offset'])

        # 굵게/기울임
        if 'bold' in style:
            self.set_bold(style['bold'])
        if 'italic' in style:
            self.set_italic(style['italic'])

        # 밑줄
        if 'underline' in style:
            self.set_underline(
                enabled=style['underline'],
                line_type=style.get('underline_type', 1),
                shape=style.get('underline_shape', 0),
                color=style.get('underline_color_rgb')
            )

        # 취소선
        if 'strikeout' in style:
            self.set_strikeout(
                enabled=style['strikeout'],
                strikeout_type=style.get('strikeout_type', 1),
                shape=style.get('strikeout_shape', 0),
                color=style.get('strikeout_color_rgb')
            )

        # 외곽선
        if 'outline_type' in style:
            self.set_outline(style['outline_type'])

        # 그림자
        if 'shadow_type' in style:
            self.set_shadow(
                shadow_type=style['shadow_type'],
                offset_x=style.get('shadow_offset_x', 10),
                offset_y=style.get('shadow_offset_y', 10),
                color=style.get('shadow_color_rgb')
            )

        # 양각/음각
        if 'emboss' in style:
            self.set_emboss(style['emboss'])
        if 'engrave' in style:
            self.set_engrave(style['engrave'])

        # 위첨자/아래첨자
        if 'superscript' in style:
            self.set_superscript(style['superscript'])
        if 'subscript' in style:
            self.set_subscript(style['subscript'])

        # 색상
        if 'color_rgb' in style:
            r, g, b = style['color_rgb']
            self.set_text_color(r, g, b)

        # 음영색
        if 'shade_color_rgb' in style and style['shade_color_rgb']:
            r, g, b = style['shade_color_rgb']
            self.set_shade_color(r, g, b)

    def apply_para_style(self, style_name):
        """
        YAML에 정의된 문단 스타일 적용

        Args:
            style_name: 스타일 이름 (예: 'body', 'heading', 'center')
        """
        if not hasattr(self, '_styles'):
            self.load_styles()

        style = self._styles.get('para_styles', {}).get(style_name)
        if not style:
            raise ValueError(f"문단 스타일 '{style_name}'을 찾을 수 없습니다.")

        # 정렬
        if 'align' in style:
            self.set_align(style['align'])

        # 줄간격
        if 'line_spacing' in style:
            spacing_type = style.get('line_spacing_type', 'percent')
            self.set_line_spacing(style['line_spacing'], spacing_type)

        # 여백
        left = style.get('left_margin_pt')
        right = style.get('right_margin_pt')
        indent = style.get('indent_pt')
        if left is not None or right is not None or indent is not None:
            self.set_para_margin(left, right, indent)

        # 문단 간격
        before = style.get('space_before_pt')
        after = style.get('space_after_pt')
        if before is not None or after is not None:
            self.set_para_spacing(before, after)

        # 편집 옵션
        if 'widow_orphan' in style:
            self.set_widow_orphan(style['widow_orphan'])
        if 'keep_with_next' in style:
            self.set_keep_with_next(style['keep_with_next'])
        if 'keep_lines' in style:
            self.set_keep_lines(style['keep_lines'])
        if 'page_break_before' in style:
            self.set_page_break_before(style['page_break_before'])

        # 격자/단어 나눔
        if 'snap_to_grid' in style:
            self.set_snap_to_grid(style['snap_to_grid'])
        if 'break_latin_word' in style:
            self.set_break_latin_word(style['break_latin_word'])
        if 'break_non_latin_word' in style:
            self.set_break_non_latin_word(style['break_non_latin_word'])

    def apply_style(self, style_name):
        """
        복합 스타일 적용 (글자 + 문단)

        Args:
            style_name: combined_styles의 스타일 이름
        """
        if not hasattr(self, '_styles'):
            self.load_styles()

        style = self._styles.get('combined_styles', {}).get(style_name)
        if not style:
            raise ValueError(f"복합 스타일 '{style_name}'을 찾을 수 없습니다.")

        # 글자 스타일 적용
        if 'char' in style:
            self.apply_char_style(style['char'])

        # 문단 스타일 적용
        if 'para' in style:
            self.apply_para_style(style['para'])

    def get_color(self, color_name):
        """
        YAML에 정의된 색상값 반환

        Args:
            color_name: 색상 이름 (예: 'red', 'blue')

        Returns:
            tuple: (r, g, b)
        """
        if not hasattr(self, '_styles'):
            self.load_styles()

        colors = self._styles.get('colors', {})
        if color_name not in colors:
            raise ValueError(f"색상 '{color_name}'을 찾을 수 없습니다.")
        return tuple(colors[color_name])


def main():
    """독립 실행 시 사용법 출력"""
    print("=" * 60)
    print("StylePara - HWP 스타일 관리 클래스")
    print("=" * 60)
    print()
    print("사용법:")
    print("-" * 60)
    print("""
from style_para import StylePara
from cursor_utils import get_hwp_instance

hwp = get_hwp_instance()
style = StylePara(hwp)

# 1. 글자 모양 설정 (선택 영역에 적용)
hwp.HAction.Run("SelectAll")  # 또는 특정 영역 선택

style.set_bold(True)           # 굵게
style.set_italic(True)         # 기울임
style.set_underline(True)      # 밑줄
style.set_font_size(12)        # 12pt
style.set_font("맑은 고딕", "Arial")  # 글꼴
style.set_text_color(255, 0, 0)       # 빨강

hwp.HAction.Run("Cancel")      # 선택 해제

# 2. 문단 모양 설정
style.set_align('center')      # 가운데 정렬
style.set_align('left')        # 왼쪽 정렬
style.set_align('right')       # 오른쪽 정렬
style.set_align('justify')     # 양쪽 정렬

style.set_line_spacing(160)    # 160% 줄간격
style.set_line_spacing(20, 'fixed')  # 20pt 고정 줄간격

style.set_para_margin(left=10, right=10, indent=5)  # 여백

# 3. 현재 모양 조회
char_info = style.get_char_shape()
print(f"글꼴: {char_info['font_hangul']}, 크기: {char_info['height_pt']}pt")

para_info = style.get_para_shape()
print(f"정렬: {para_info['align_name']}, 줄간격: {para_info['line_spacing']}")

# 4. 모양 복사/붙여넣기
# 원본 선택 후 복사
style.copy_char_shape()  # 또는 copy_para_shape()

# 대상 선택 후 붙여넣기
style.paste_shape()

# 5. 유틸리티 함수
bgr = StylePara.rgb_to_bgr(255, 0, 0)  # RGB -> BGR
hwpunit = StylePara.pt_to_hwpunit(12)  # pt -> HWPUNIT

# 6. YAML 스타일 프리셋 사용
style.load_styles()  # styles.yaml 로드
# 또는 style.load_styles('/path/to/custom.yaml')

# 스타일 이름 확인
print(style.get_style_names('char'))     # 글자 스타일 목록
print(style.get_style_names('para'))     # 문단 스타일 목록
print(style.get_style_names('combined')) # 복합 스타일 목록

# 스타일 적용
style.apply_char_style('heading1')  # 글자 스타일만
style.apply_para_style('center')    # 문단 스타일만
style.apply_style('document_title') # 복합 스타일 (글자+문단)

# 색상 사용
r, g, b = style.get_color('red')
style.set_text_color(r, g, b)
""")
    print("-" * 60)
    print()
    print("정렬 타입:")
    for name, value in StylePara.ALIGN_TYPES.items():
        print(f"  '{name}': {value}")
    print()
    print("줄간격 타입:")
    for name, value in StylePara.LINE_SPACING_TYPES.items():
        print(f"  '{name}': {value}")
    print()

    # YAML 스타일 목록 출력
    try:
        with open(DEFAULT_STYLES_PATH, 'r', encoding='utf-8') as f:
            styles = yaml.safe_load(f)

        print("YAML 프리셋 스타일:")
        print()
        print("  글자 스타일 (char_styles):")
        for name in styles.get('char_styles', {}).keys():
            print(f"    - {name}")
        print()
        print("  문단 스타일 (para_styles):")
        for name in styles.get('para_styles', {}).keys():
            print(f"    - {name}")
        print()
        print("  복합 스타일 (combined_styles):")
        for name in styles.get('combined_styles', {}).keys():
            print(f"    - {name}")
        print()
    except FileNotFoundError:
        print("  (styles.yaml 파일 없음)")
        print()

    print("=" * 60)


if __name__ == "__main__":
    main()
