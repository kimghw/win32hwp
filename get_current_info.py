"""
한글(HWP) 현재 위치 정보 조회 모듈

KeyIndicator, GetPos, CtrlID 정보를 각각 클래스로 제공
"""

import win32com.client as win32
from typing import Optional, Dict, Tuple, Any
from cursor_utils import get_hwp_instance


class KeyIndicatorInfo:
    """
    상태바 위치 정보 클래스

    사용자 친화적 위치 정보 (1부터 시작)
    """

    def __init__(self, hwp):
        """
        Args:
            hwp: HWP 인스턴스
        """
        self.hwp = hwp

    def get_info(self) -> Dict[str, Any]:
        """
        현재 커서의 상태바 정보 조회

        주의: 공식 문서와 실제 반환값이 다름!
        - key[3] = 실제 페이지 번호 (문서에는 '단 번호'라고 되어 있음)
        - key[5] = 실제 줄 정보 (문서에는 '칸 번호'라고 되어 있음)

        Returns:
            dict: {
                'total_sections': int,      # 총 구역 수 (bool일 수도 있음)
                'current_section': int,     # 현재 구역
                'page': int,                # 페이지 번호 (key[3])
                'unknown_2': Any,           # key[2] - 용도 불명
                'line': int,                # 줄 번호 (key[4])
                'line_info': int,           # 줄 정보 (key[5])
                'insert_mode': str,         # '삽입' or '수정'
                'ctrl_name': str            # 컨트롤 이름 (있을 경우)
            }
        """
        try:
            key = self.hwp.KeyIndicator()

            return {
                'total_sections': key[0],
                'current_section': key[1],
                'unknown_2': key[2],
                'page': key[3],              # 실제 페이지
                'line': key[4],
                'line_info': key[5],         # 실제 줄 정보
                'insert_mode': '삽입' if key[6] == 0 else '수정',
                'ctrl_name': key[7] if len(key) > 7 else ''
            }
        except Exception as e:
            print(f"KeyIndicator 조회 실패: {e}")
            return {}

    def print_info(self):
        """상태바 정보 출력"""
        info = self.get_info()
        if not info:
            return

        print("=== 상태바 위치 정보 ===")
        # total_sections이 bool로 반환되는 경우 처리 (True=1, False=0)
        total = info['total_sections']
        if isinstance(total, bool):
            total = 1 if total else 0
        print(f"구역: {info['current_section']}/{total}")
        print(f"페이지: {info['page']}")
        print(f"줄: {info['line']}")
        print(f"줄 정보: {info['line_info']}")
        print(f"모드: {info['insert_mode']}")
        if info.get('unknown_2') is not None:
            print(f"key[2] (용도불명): {info['unknown_2']}")
        if info['ctrl_name']:
            print(f"컨트롤: {info['ctrl_name']}")


class PosInfo:
    """
    내부 위치 정보 클래스

    API 위치 지정/이동용 내부 인덱스 (0부터 시작)
    """

    def __init__(self, hwp):
        """
        Args:
            hwp: HWP 인스턴스
        """
        self.hwp = hwp

    def get_pos(self) -> Tuple[int, int, int]:
        """
        현재 커서의 내부 위치 조회

        Returns:
            tuple: (list_id, para_id, char_pos)
                - list_id: 리스트 ID (0=본문, 10+=테이블)
                - para_id: 문단 ID
                - char_pos: 문단 내 문자 위치 (한글=2, 영문=1)
        """
        try:
            pos = self.hwp.GetPos()
            return (pos[0], pos[1], pos[2])
        except Exception as e:
            print(f"GetPos 조회 실패: {e}")
            return (0, 0, 0)

    def get_info(self) -> Dict[str, Any]:
        """
        내부 위치 정보를 딕셔너리로 반환

        Returns:
            dict: {
                'list_id': int,      # 리스트 ID
                'para_id': int,      # 문단 ID
                'char_pos': int,     # 문자 위치
                'list_type': str     # 리스트 타입 설명
            }
        """
        list_id, para_id, char_pos = self.get_pos()

        # 리스트 타입 설명
        if list_id == 0:
            list_type = '본문'
        elif 1 <= list_id <= 9:
            list_type = '머리말/꼬리말/각주/미주'
        else:
            list_type = f'테이블 셀 (ID: {list_id})'

        return {
            'list_id': list_id,
            'para_id': para_id,
            'char_pos': char_pos,
            'list_type': list_type
        }

    def set_pos(self, list_id: int, para_id: int, char_pos: int) -> bool:
        """
        커서를 특정 위치로 이동

        Args:
            list_id: 리스트 ID
            para_id: 문단 ID
            char_pos: 문자 위치

        Returns:
            bool: 성공 여부
        """
        try:
            self.hwp.SetPos(list_id, para_id, char_pos)
            return True
        except Exception as e:
            print(f"SetPos 실패: {e}")
            return False

    def print_info(self):
        """내부 위치 정보 출력"""
        info = self.get_info()

        print("=== 내부 위치 정보 ===")
        print(f"리스트 타입: {info['list_type']}")
        print(f"리스트 ID: {info['list_id']}")
        print(f"문단 ID: {info['para_id']}")
        print(f"문자 위치: {info['char_pos']}")


class CtrlInfo:
    """
    컨트롤 정보 클래스

    문서 내 표, 그림, 각주 등의 컨트롤 정보 조회
    """

    # 컨트롤 ID 설명 매핑
    CTRL_NAMES = {
        'cold': '단',
        'secd': '구역',
        'fn': '각주',
        'en': '미주',
        'tbl': '표',
        'eqed': '수식',
        'atno': '번호넣기',
        'head': '머리말',
        'foot': '꼬리말',
        '%dte': '현재 날짜/시간',
        '%pat': '문서 경로',
        '%mmg': '메일 머지',
        '%xrf': '상호 참조',
        '%clk': '누름틀',
        '%hlk': '하이퍼링크',
        'bokm': '책갈피',
        'idxm': '찾아보기',
        '$pic': '그림',
        'form': '양식 개체',
        '+pbt': '명령 단추',
        '+rbt': '라디오 단추',
        '+cbt': '선택 상자',
        '+cob': '콤보 상자',
        '+edt': '입력 상자'
    }

    def __init__(self, hwp):
        """
        Args:
            hwp: HWP 인스턴스
        """
        self.hwp = hwp

    def find_ctrl(self) -> str:
        """
        현재 커서 위치의 컨트롤 ID 조회

        Returns:
            str: 컨트롤 ID (없으면 빈 문자열)
        """
        try:
            ctrl_id = self.hwp.FindCtrl()
            return ctrl_id if ctrl_id else ''
        except Exception as e:
            print(f"FindCtrl 실패: {e}")
            return ''

    def get_info(self) -> Dict[str, str]:
        """
        현재 위치의 컨트롤 정보 조회

        Returns:
            dict: {
                'ctrl_id': str,        # 컨트롤 ID
                'ctrl_name': str,      # 컨트롤 이름
                'has_ctrl': bool       # 컨트롤 존재 여부
            }
        """
        ctrl_id = self.find_ctrl()

        if ctrl_id:
            ctrl_name = self.CTRL_NAMES.get(ctrl_id, '알 수 없는 컨트롤')
            return {
                'ctrl_id': ctrl_id,
                'ctrl_name': ctrl_name,
                'has_ctrl': True
            }
        else:
            return {
                'ctrl_id': '',
                'ctrl_name': '',
                'has_ctrl': False
            }

    def select_ctrl(self, ctrl_list: str, option: int = 1) -> bool:
        """
        특정 컨트롤 선택

        Args:
            ctrl_list: 선택할 컨트롤 ID (여러 개는 0x02로 구분)
            option: 0=추가 선택, 1=취소 후 재선택

        Returns:
            bool: 성공 여부
        """
        try:
            self.hwp.SelectCtrl(ctrl_list, option)
            return True
        except Exception as e:
            print(f"SelectCtrl 실패: {e}")
            return False

    def print_info(self):
        """컨트롤 정보 출력"""
        info = self.get_info()

        print("=== 컨트롤 정보 ===")
        if info['has_ctrl']:
            print(f"컨트롤 ID: {info['ctrl_id']}")
            print(f"컨트롤 이름: {info['ctrl_name']}")
        else:
            print("컨트롤 없음")


class CurrentInfo:
    """
    통합 위치 정보 클래스

    KeyIndicator, GetPos, CtrlID 정보를 모두 제공
    """

    def __init__(self, hwp=None):
        """
        Args:
            hwp: HWP 인스턴스 (None이면 자동 연결)
        """
        self.hwp = hwp if hwp else get_hwp_instance()
        if not self.hwp:
            raise Exception("HWP 인스턴스를 가져올 수 없습니다")

        self.key_indicator = KeyIndicatorInfo(self.hwp)
        self.pos = PosInfo(self.hwp)
        self.ctrl = CtrlInfo(self.hwp)

    def get_all_info(self) -> Dict[str, Any]:
        """
        모든 위치 정보 조회

        Returns:
            dict: {
                'key_indicator': dict,  # 상태바 정보
                'pos': dict,            # 내부 위치 정보
                'ctrl': dict            # 컨트롤 정보
            }
        """
        return {
            'key_indicator': self.key_indicator.get_info(),
            'pos': self.pos.get_info(),
            'ctrl': self.ctrl.get_info()
        }

    def print_all(self):
        """모든 위치 정보 출력"""
        print("\n" + "="*50)
        print("현재 커서 위치 정보")
        print("="*50 + "\n")

        self.key_indicator.print_info()
        print()
        self.pos.print_info()
        print()
        self.ctrl.print_info()
        print()


def main():
    """테스트 실행"""
    try:
        # HWP 인스턴스 가져오기
        hwp = get_hwp_instance()
        if not hwp:
            print("한글이 실행 중이 아닙니다")
            return

        # 통합 정보 출력
        current_info = CurrentInfo(hwp)
        current_info.print_all()

        # 개별 클래스 사용 예시
        print("="*50)
        print("개별 클래스 사용 예시")
        print("="*50 + "\n")

        # KeyIndicator
        key_info = KeyIndicatorInfo(hwp)
        key_data = key_info.get_info()
        print(f"현재 페이지: {key_data['page']}")

        # Pos
        pos_info = PosInfo(hwp)
        list_id, para_id, char_pos = pos_info.get_pos()
        print(f"문단 ID: {para_id}")

        # Ctrl
        ctrl_info = CtrlInfo(hwp)
        ctrl_data = ctrl_info.get_info()
        if ctrl_data['has_ctrl']:
            print(f"컨트롤: {ctrl_data['ctrl_name']}")

    except Exception as e:
        print(f"오류 발생: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
