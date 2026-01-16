# -*- coding: utf-8 -*-
"""
한글(HWP) 개요 수준 관리 클래스

마크다운 헤딩(#, ##, ###)을 HWP 개요 수준으로 변환.

사용 예:
    from style_numb import StyleNumb
    from cursor_utils import get_hwp_instance

    hwp = get_hwp_instance()
    numb = StyleNumb(hwp)

    # 문서에서 마크다운 헤딩을 찾아 개요 수준 적용
    result = numb.개요수준정의()

    # 또는 새 문서에 마크다운 텍스트 입력 후 개요 적용
    numb.새문서()
    numb.텍스트입력('''
    # 1장 서론
    ## 1.1 배경
    ### 1.1.1 세부 내용
    ''')
    numb.개요수준정의()
"""

import os
try:
    import yaml
    HAS_YAML = True
except ImportError:
    HAS_YAML = False
from cursor_utils import get_hwp_instance

DEFAULT_STYLES_PATH = os.path.join(os.path.dirname(__file__), 'styles.yaml')


class StyleNumb:
    """HWP 개요 수준 관리 클래스"""

    def __init__(self, hwp=None, styles_path=None):
        """
        초기화

        Args:
            hwp: 한글 인스턴스 (None이면 get_hwp_instance() 호출)
            styles_path: styles.yaml 경로
        """
        self.hwp = hwp if hwp else get_hwp_instance()
        if not self.hwp:
            raise RuntimeError("한글 인스턴스에 연결할 수 없습니다.")

        self.styles_path = styles_path or DEFAULT_STYLES_PATH
        self.styles = self._load_styles()
        self.heading_config = self.styles.get('heading_levels', {})

    def _load_styles(self):
        """styles.yaml 로드"""
        if not HAS_YAML:
            return {}
        if os.path.exists(self.styles_path):
            with open(self.styles_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f) or {}
        return {}

    def get_heading_config(self, level):
        """헤딩 레벨 설정 조회"""
        key = f'h{level}'
        return self.heading_config.get(key, {
            'outline_level': level,
            'char_style': None,
            'para_style': None
        })

    # ========== 핵심 기능 ==========

    def 새문서(self):
        """새 문서 열기"""
        self.hwp.HAction.Run("FileNew")

    def 텍스트입력(self, text):
        """
        마크다운 텍스트 입력 (개요 적용 전)

        Args:
            text: 마크다운 형식 텍스트
        """
        lines = [line for line in text.strip().split('\n') if line.strip()]

        for line in lines:
            self.hwp.HAction.GetDefault("InsertText",
                                        self.hwp.HParameterSet.HInsertText.HSet)
            self.hwp.HParameterSet.HInsertText.Text = line.strip()
            self.hwp.HAction.Execute("InsertText",
                                     self.hwp.HParameterSet.HInsertText.HSet)
            self.hwp.HAction.Run("BreakPara")

    def set_outline_level(self, level, debug=False):
        """
        개요 수준 설정 (PutOutlineNumber + LevelDown 방식)

        Args:
            level: 개요 수준 (1~7)
            debug: True면 결과 출력

        Returns:
            bool: 성공 여부
        """
        if not 1 <= level <= 7:
            raise ValueError(f"개요 수준은 1~7 사이여야 합니다: {level}")

        # 1. 개요번호 달기 (1수준으로 시작)
        result = self.hwp.HAction.Run("PutOutlineNumber")

        # 2. 필요한 만큼 수준 내리기
        for _ in range(level - 1):
            self.hwp.HAction.Run("ParaNumberBulletLevelDown")

        if debug:
            print(f"  set_outline_level({level}) Run: {result}")
        return result

    def remove(self):
        """개요 해제"""
        self.hwp.HAction.GetDefault("ParagraphShape",
                                    self.hwp.HParameterSet.HParaShape.HSet)
        self.hwp.HParameterSet.HParaShape.HeadingType = 0
        self.hwp.HAction.Execute("ParagraphShape",
                                 self.hwp.HParameterSet.HParaShape.HSet)

    def get_outline_level(self):
        """
        현재 문단의 개요 수준 조회

        Returns:
            dict: {
                'heading_string': 개요 번호 문자열 (예: "1.", "1.1", ""),
                'has_outline': 개요가 있는지 여부
            }
        """
        # GetHeadingString()으로 개요 번호 문자열 조회
        heading_str = self.hwp.GetHeadingString()
        return {
            'heading_string': heading_str,
            'has_outline': bool(heading_str)
        }

    def get_outline_shape(self):
        """
        개요 번호 설정 조회 (SecDef.OutlineShape)

        Returns:
            dict: 개요 번호 설정 정보
        """
        try:
            # HParameterSet 방식으로 시도
            self.hwp.HAction.GetDefault("OutlineNumber",
                                        self.hwp.HParameterSet.HSecDef.HSet)
            sec_set = self.hwp.HParameterSet.HSecDef

            result = {}
            # OutlineShape 서브셋 접근 시도
            try:
                outline_set = sec_set.OutlineShape

                # 모든 속성 조회 시도
                attrs = ['Type', 'StartNumber']
                for attr in attrs:
                    try:
                        val = getattr(outline_set, attr, None)
                        if val is not None:
                            result[attr] = val
                    except:
                        pass

                # 각 수준별 설정
                for i in range(7):
                    level_info = {}
                    for attr in ['StrFormatLevel', 'NumFormatLevel', 'ParaShapeLevel',
                                 'CharShapeLevel', 'StartNumberLevel']:
                        try:
                            val = getattr(outline_set, f"{attr}{i}", None)
                            if val is not None:
                                level_info[attr] = val
                        except:
                            pass
                    if level_info:
                        result[f'level{i+1}'] = level_info

            except Exception as e:
                result['outline_error'] = str(e)

            return result
        except Exception as e:
            return {'error': str(e)}

    def apply_style(self, style_name):
        """
        스타일 적용

        Args:
            style_name: 스타일 이름 (예: "개요 1")
        """
        try:
            act = self.hwp.CreateAction("Style")
            pset = act.CreateSet()
            act.GetDefault(pset)
            pset.SetItem("Apply", 0)
            pset.SetItem("StyleName", style_name)
            return act.Execute(pset)
        except:
            return False

    @staticmethod
    def parse_heading_level(text):
        """
        마크다운 헤딩 파싱

        Args:
            text: 텍스트 (예: "## 소제목")

        Returns:
            tuple: (level, clean_text)
        """
        if not text or not text.startswith('#'):
            return 0, text

        level = 0
        for char in text:
            if char == '#':
                level += 1
            else:
                break

        level = min(level, 7)
        clean_text = text[level:].lstrip()

        return level, clean_text

    def scan_headings(self, debug=False):
        """
        문서 전체에서 마크다운 헤딩 스캔

        Returns:
            list: 헤딩 정보 리스트 (list_id, para_id 포함)
        """
        headings = []
        counters = [0, 0, 0, 0, 0]
        processed_paras = set()
        max_iterations = 10000

        self.hwp.HAction.Run("MoveDocBegin")

        for _ in range(max_iterations):
            self.hwp.HAction.Run("MoveParaBegin")
            pos = self.hwp.GetPos()
            list_id, para_id = pos[0], pos[1]

            if para_id in processed_paras:
                break
            processed_paras.add(para_id)

            self.hwp.HAction.Run("MoveSelParaEnd")
            text = self.hwp.GetTextFile("TEXT", "saveblock")
            self.hwp.HAction.Run("Cancel")

            if text:
                text = text.strip()
                level, clean_text = self.parse_heading_level(text)

                if level > 0 and level <= 5:
                    counters[level - 1] += 1
                    for i in range(level, 5):
                        counters[i] = 0

                    headings.append({
                        'list_id': list_id,
                        'para_id': para_id,
                        'level': level,
                        'text': text,
                        'clean_text': clean_text,
                        'numbering': counters.copy()
                    })

                    if debug:
                        print(f"  [{para_id}] lv={level} {counters}")

            self.hwp.HAction.Run("MoveParaEnd")
            self.hwp.HAction.Run("MoveRight")

            new_pos = self.hwp.GetPos()
            if new_pos[1] == para_id:
                break

        if debug:
            print(f"[스캔] {len(headings)}개 헤딩 발견")

        return headings

    def get_heading_number_string(self, numbering, separator='.'):
        """
        번호 배열을 문자열로 변환

        Args:
            numbering: [1, 2, 1, 0, 0] 형식

        Returns:
            str: "1.2.1" 형식
        """
        nums = []
        for n in numbering:
            if n == 0:
                break
            nums.append(str(n))
        return separator.join(nums)

    def 개요수준정의(self, remove_marker=True, debug=False):
        """
        문서의 마크다운 헤딩을 스캔하고 개요 수준 적용

        Args:
            remove_marker: True면 # 마커 제거
            debug: 디버그 출력

        Returns:
            dict: {
                'processed': 처리된 헤딩 수,
                'headings': 헤딩 정보 리스트
            }

        Example:
            # 제목      → 개요 1수준
            ## 소제목   → 개요 2수준
            ### 세부    → 개요 3수준
        """
        headings = self.scan_headings(debug=debug)

        if not headings:
            return {'processed': 0, 'headings': []}

        if debug:
            print(f"\n[적용] {len(headings)}개 헤딩 처리 시작")

        # 역순으로 처리 (para_id 유지)
        for h in reversed(headings):
            list_id = h['list_id']
            para_id = h['para_id']
            level = h['level']
            clean_text = h['clean_text']

            # 새 텍스트
            new_text = clean_text if remove_marker else h['text']

            # SetPos로 직접 이동
            self.hwp.SetPos(list_id, para_id, 0)

            # 문단 선택 및 텍스트 교체
            self.hwp.HAction.Run("MoveSelParaEnd")

            self.hwp.HAction.GetDefault("InsertText",
                                        self.hwp.HParameterSet.HInsertText.HSet)
            self.hwp.HParameterSet.HInsertText.Text = new_text
            self.hwp.HAction.Execute("InsertText",
                                     self.hwp.HParameterSet.HInsertText.HSet)

            # 개요 수준 적용 (문단 전체 선택 후 적용)
            self.hwp.HAction.Run("MoveParaBegin")
            self.hwp.HAction.Run("MoveSelParaEnd")
            config = self.get_heading_config(level)
            outline_level = config.get('outline_level', level)
            self.set_outline_level(outline_level, debug=debug)
            self.hwp.HAction.Run("Cancel")

            if debug:
                num_str = self.get_heading_number_string(h['numbering'])
                print(f"  {num_str:10} → {new_text[:20]}")

        if debug:
            print(f"[완료] {len(headings)}개 처리됨")

        return {
            'processed': len(headings),
            'headings': headings
        }


def main():
    """테스트: 개요 번호 설정 조회"""
    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글을 먼저 실행하세요")
        return

    print("[연결] 한글에 연결됨")

    numb = StyleNumb(hwp)

    # 개요 번호 설정 조회
    print("\n[조회] 개요 번호 설정:")
    shape = numb.get_outline_shape()
    for key, value in shape.items():
        print(f"  {key}: {value}")


if __name__ == "__main__":
    main()
