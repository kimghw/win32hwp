"""
HWP API 래퍼

Win32 HWP API를 Python에서 사용하기 편하도록 래핑한 모듈.
API의 불일치나 버그를 보정하고 일관된 인터페이스 제공.
"""
from cursor_utils import get_hwp_instance
from typing import Dict, Tuple, Optional


class HwpPosition:
    """
    커서 위치 정보를 담는 클래스
    GetPos()와 KeyIndicator()를 통합
    """

    def __init__(self, hwp):
        """
        Args:
            hwp: HWP 인스턴스
        """
        self.hwp = hwp
        self._update()

    def _update(self):
        """위치 정보 갱신"""
        # GetPos - 내부 위치 정보
        pos = self.hwp.GetPos()
        self.list_id = pos[0]
        self.para_id = pos[1]
        self.char_pos = pos[2]

        # KeyIndicator - 상태바 표시 정보
        # 주의: Win32 문서와 실제 반환 순서가 다름!
        key = self.hwp.KeyIndicator()
        self.total_section = key[0]
        self.current_section = key[1]
        # key[2] = 미확인 (문서에 없음)
        self.page = key[3]          # 페이지 (실제 테스트 확인)
        self.column_num = key[4]    # 단 (다단 편집)
        self.line = key[5]          # 줄 (실제 테스트 확인)
        self.column = key[6]        # 칸
        # key[7] = ctrlname (Python에서는 반환 안 됨)

        # 테이블 확인
        self.in_table = self._check_in_table()

    def _check_in_table(self) -> bool:
        """테이블 내부 여부 확인"""
        try:
            act = self.hwp.CreateAction("TableCellBorder")
            pset = act.GetDefault("TableCellBorder", self.hwp.HParameterSet.HTableCellBorder.HSet)
            return act.Execute(pset)
        except:
            return False

    def to_dict(self) -> Dict:
        """딕셔너리로 변환"""
        return {
            'list_id': self.list_id,
            'para_id': self.para_id,
            'char_pos': self.char_pos,
            'total_section': self.total_section,
            'current_section': self.current_section,
            'page': self.page,
            'column_num': self.column_num,
            'line': self.line,
            'column': self.column,
            'in_table': self.in_table
        }

    def __repr__(self):
        return (f"HwpPosition(list={self.list_id}, para={self.para_id}, "
                f"char={self.char_pos}, page={self.page}, "
                f"line={self.line}, col={self.column})")


class HwpAPI:
    """
    HWP API 래퍼 클래스

    Win32 API의 불일치를 보정하고 사용하기 편한 인터페이스 제공
    """

    def __init__(self, hwp=None):
        """
        Args:
            hwp: HWP 인스턴스 (None이면 자동 연결)
        """
        if hwp is None:
            hwp = get_hwp_instance()
            if not hwp:
                raise RuntimeError("한글이 실행 중이 아닙니다")

        self.hwp = hwp
        self.hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')

    def get_position(self) -> HwpPosition:
        """
        현재 커서 위치 정보 반환

        Returns:
            HwpPosition: 위치 정보 객체
        """
        return HwpPosition(self.hwp)

    def get_pos_tuple(self) -> Tuple[int, int, int]:
        """
        GetPos() 원본 반환

        Returns:
            Tuple[list_id, para_id, char_pos]
        """
        return self.hwp.GetPos()

    def get_key_indicator(self) -> Tuple:
        """
        KeyIndicator() 원본 반환

        주의: Win32 문서와 실제 반환값이 다름!
        실제 순서: (total_sec, cur_sec, ?, page, col_num, line, column)

        Returns:
            Tuple: 7개 요소 (ctrlname 제외)
        """
        return self.hwp.KeyIndicator()

    def set_pos(self, list_id: int, para_id: int, char_pos: int) -> bool:
        """
        커서 위치 이동

        Args:
            list_id: 리스트 ID
            para_id: 문단 ID
            char_pos: 문자 위치

        Returns:
            bool: 성공 여부
        """
        return self.hwp.SetPos(list_id, para_id, char_pos)

    def is_in_table(self) -> bool:
        """
        테이블 내부 여부 확인

        Returns:
            bool: 테이블 내부이면 True
        """
        try:
            act = self.hwp.CreateAction("TableCellBorder")
            pset = act.GetDefault("TableCellBorder", self.hwp.HParameterSet.HTableCellBorder.HSet)
            return act.Execute(pset)
        except:
            return False

    def get_field_name(self) -> str:
        """
        현재 필드 이름 반환

        Returns:
            str: 필드 이름 (테이블="Cell", 본문="", 등)
        """
        return self.hwp.GetCurFieldName()


# 편의 함수
def get_api(hwp=None) -> HwpAPI:
    """
    HWP API 래퍼 인스턴스 생성

    Args:
        hwp: HWP 인스턴스 (None이면 자동 연결)

    Returns:
        HwpAPI: API 래퍼 객체
    """
    return HwpAPI(hwp)


if __name__ == "__main__":
    # 테스트
    api = get_api()

    print("=== HWP API 래퍼 테스트 ===\n")

    # 위치 정보
    pos = api.get_position()
    print("현재 위치:")
    print(f"  {pos}")
    print()

    # 딕셔너리 변환
    pos_dict = pos.to_dict()
    print("딕셔너리 형식:")
    for key, value in pos_dict.items():
        print(f"  {key}: {value}")
    print()

    # 테이블 확인
    print(f"테이블 내부: {api.is_in_table()}")
    print(f"필드 이름: '{api.get_field_name()}'")
    print()

    # 원본 API
    print("원본 API:")
    print(f"  GetPos(): {api.get_pos_tuple()}")
    print(f"  KeyIndicator(): {api.get_key_indicator()}")
