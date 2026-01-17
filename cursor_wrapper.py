"""
한글(HWP) 커서 조작 및 정보 조회 통합 래퍼

커서 이동, 위치 정보, 선택 등을 쉽게 사용할 수 있는 통합 클래스
"""

from typing import Optional, Dict, Tuple, Any
from cursor_utils import get_hwp_instance
from get_current_info import KeyIndicatorInfo, PosInfo, CtrlInfo


class CursorWrapper:
    """
    커서 조작 및 정보 조회 통합 클래스

    위치 정보 조회, 커서 이동, 텍스트 선택 등을 하나의 클래스로 제공
    """

    def __init__(self, hwp=None):
        """
        Args:
            hwp: HWP 인스턴스 (None이면 자동 연결)
        """
        self.hwp = hwp if hwp else get_hwp_instance()
        if not self.hwp:
            raise Exception("HWP 인스턴스를 가져올 수 없습니다")

        # 정보 조회 클래스
        self.key_indicator = KeyIndicatorInfo(self.hwp)
        self.pos_info = PosInfo(self.hwp)
        self.ctrl_info = CtrlInfo(self.hwp)

    # ==================== 위치 정보 조회 ====================

    def get_pos(self) -> Tuple[int, int, int]:
        """
        현재 커서의 내부 위치 조회

        Returns:
            tuple: (list_id, para_id, char_pos)
        """
        return self.pos_info.get_pos()

    def get_page(self) -> int:
        """
        현재 페이지 번호 조회

        Returns:
            int: 페이지 번호 (1부터 시작)
        """
        info = self.key_indicator.get_info()
        return info.get('page', 0)

    def get_line(self) -> int:
        """
        현재 줄 번호 조회

        Returns:
            int: 줄 번호 (1부터 시작)
        """
        info = self.key_indicator.get_info()
        return info.get('line', 0)

    def get_ctrl_id(self) -> str:
        """
        현재 위치의 컨트롤 ID 조회

        Returns:
            str: 컨트롤 ID ('tbl', '$pic' 등) 또는 빈 문자열
        """
        return self.ctrl_info.find_ctrl()

    def is_in_table(self) -> bool:
        """
        커서가 표 안에 있는지 확인

        Returns:
            bool: 표 안이면 True
        """
        return self.get_ctrl_id() == 'tbl'

    def is_in_ctrl(self, ctrl_id: str) -> bool:
        """
        커서가 특정 컨트롤 안에 있는지 확인

        Args:
            ctrl_id: 확인할 컨트롤 ID ('tbl', '$pic', 'fn' 등)

        Returns:
            bool: 해당 컨트롤 안이면 True
        """
        return self.get_ctrl_id() == ctrl_id

    def print_info(self):
        """현재 커서의 모든 정보 출력"""
        print("\n" + "="*50)
        print("현재 커서 위치 정보")
        print("="*50 + "\n")

        self.key_indicator.print_info()
        print()
        self.pos_info.print_info()
        print()
        self.ctrl_info.print_info()
        print()

    # ==================== 커서 이동 (문단) ====================

    def move_para_begin(self) -> bool:
        """
        문단 시작으로 이동

        Returns:
            bool: 성공 여부
        """
        try:
            self.hwp.HAction.Run("MoveParaBegin")
            return True
        except Exception as e:
            print(f"MoveParaBegin 실패: {e}")
            return False

    def move_para_end(self) -> bool:
        """
        문단 끝으로 이동

        Returns:
            bool: 성공 여부
        """
        try:
            self.hwp.HAction.Run("MoveParaEnd")
            return True
        except Exception as e:
            print(f"MoveParaEnd 실패: {e}")
            return False

    # ==================== 커서 이동 (줄) ====================

    def move_line_begin(self) -> bool:
        """
        줄 시작으로 이동 (Home 키)

        Returns:
            bool: 성공 여부
        """
        try:
            self.hwp.HAction.Run("MoveLineBegin")
            return True
        except Exception as e:
            print(f"MoveLineBegin 실패: {e}")
            return False

    def move_line_end(self) -> bool:
        """
        줄 끝으로 이동 (End 키)

        Returns:
            bool: 성공 여부
        """
        try:
            self.hwp.HAction.Run("MoveLineEnd")
            return True
        except Exception as e:
            print(f"MoveLineEnd 실패: {e}")
            return False

    # ==================== 커서 이동 (문서) ====================

    def move_doc_begin(self) -> bool:
        """
        문서 처음으로 이동 (Ctrl+Home)

        Returns:
            bool: 성공 여부
        """
        try:
            self.hwp.HAction.Run("MoveDocBegin")
            return True
        except Exception as e:
            print(f"MoveDocBegin 실패: {e}")
            return False

    def move_doc_end(self) -> bool:
        """
        문서 끝으로 이동 (Ctrl+End)

        Returns:
            bool: 성공 여부
        """
        try:
            self.hwp.HAction.Run("MoveDocEnd")
            return True
        except Exception as e:
            print(f"MoveDocEnd 실패: {e}")
            return False

    # ==================== 커서 이동 (글자/줄 단위) ====================

    def move_left(self, count: int = 1) -> bool:
        """
        왼쪽으로 이동

        Args:
            count: 이동할 글자 수 (기본값: 1)

        Returns:
            bool: 성공 여부
        """
        try:
            for _ in range(count):
                self.hwp.HAction.Run("MoveLeft")
            return True
        except Exception as e:
            print(f"MoveLeft 실패: {e}")
            return False

    def move_right(self, count: int = 1) -> bool:
        """
        오른쪽으로 이동

        Args:
            count: 이동할 글자 수 (기본값: 1)

        Returns:
            bool: 성공 여부
        """
        try:
            for _ in range(count):
                self.hwp.HAction.Run("MoveRight")
            return True
        except Exception as e:
            print(f"MoveRight 실패: {e}")
            return False

    def move_up(self, count: int = 1) -> bool:
        """
        위로 이동

        Args:
            count: 이동할 줄 수 (기본값: 1)

        Returns:
            bool: 성공 여부
        """
        try:
            for _ in range(count):
                self.hwp.HAction.Run("MoveUp")
            return True
        except Exception as e:
            print(f"MoveUp 실패: {e}")
            return False

    def move_down(self, count: int = 1) -> bool:
        """
        아래로 이동

        Args:
            count: 이동할 줄 수 (기본값: 1)

        Returns:
            bool: 성공 여부
        """
        try:
            for _ in range(count):
                self.hwp.HAction.Run("MoveDown")
            return True
        except Exception as e:
            print(f"MoveDown 실패: {e}")
            return False

    # ==================== 위치 설정 ====================

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
        return self.pos_info.set_pos(list_id, para_id, char_pos)

    def save_pos(self) -> Tuple[int, int, int]:
        """
        현재 위치 저장

        Returns:
            tuple: (list_id, para_id, char_pos)
        """
        return self.get_pos()

    def restore_pos(self, pos: Tuple[int, int, int]) -> bool:
        """
        저장된 위치로 복원

        Args:
            pos: save_pos()로 저장한 위치 튜플

        Returns:
            bool: 성공 여부
        """
        return self.set_pos(pos[0], pos[1], pos[2])

    # ==================== 텍스트 선택 ====================

    def select_all(self) -> bool:
        """
        전체 선택 (Ctrl+A)

        Returns:
            bool: 성공 여부
        """
        try:
            self.hwp.HAction.Run("SelectAll")
            return True
        except Exception as e:
            print(f"SelectAll 실패: {e}")
            return False

    def select_line(self) -> bool:
        """
        현재 줄 선택 (줄 시작부터 끝까지)

        Returns:
            bool: 성공 여부
        """
        try:
            self.hwp.HAction.Run("MoveLineBegin")
            self.hwp.HAction.Run("MoveSelLineEnd")
            return True
        except Exception as e:
            print(f"select_line 실패: {e}")
            return False

    def select_para(self) -> bool:
        """
        현재 문단 선택 (문단 시작부터 끝까지)

        Returns:
            bool: 성공 여부
        """
        try:
            self.hwp.HAction.Run("MoveParaBegin")
            self.hwp.HAction.Run("MoveSelParaEnd")
            return True
        except Exception as e:
            print(f"select_para 실패: {e}")
            return False

    def cancel_selection(self) -> bool:
        """
        선택 해제 (Esc)

        Returns:
            bool: 성공 여부
        """
        try:
            self.hwp.HAction.Run("Cancel")
            return True
        except Exception as e:
            print(f"Cancel 실패: {e}")
            return False

    # ==================== 텍스트 조작 ====================

    def get_selected_text(self) -> str:
        """
        선택된 텍스트 가져오기

        Returns:
            str: 선택된 텍스트 (선택 없으면 빈 문자열)
        """
        try:
            text = self.hwp.GetTextFile("TEXT", "saveblock")
            return text if text else ""
        except Exception as e:
            print(f"GetTextFile 실패: {e}")
            return ""

    def get_char_at_cursor(self) -> str:
        """
        커서 위치의 한 글자 가져오기

        Returns:
            str: 현재 위치 글자
        """
        try:
            # 현재 위치 저장
            saved_pos = self.save_pos()

            # 한 글자 선택
            self.hwp.HAction.Run("MoveSelRight")
            text = self.get_selected_text()

            # 위치 복원
            self.restore_pos(saved_pos)

            return text
        except Exception as e:
            print(f"get_char_at_cursor 실패: {e}")
            return ""

    def insert_text(self, text: str) -> bool:
        """
        텍스트 삽입

        Args:
            text: 삽입할 텍스트

        Returns:
            bool: 성공 여부
        """
        try:
            self.hwp.HAction.GetDefault("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
            self.hwp.HParameterSet.HInsertText.Text = text
            self.hwp.HAction.Execute("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
            return True
        except Exception as e:
            print(f"insert_text 실패: {e}")
            return False

    def delete_char(self, forward: bool = True) -> bool:
        """
        글자 삭제

        Args:
            forward: True면 오른쪽 삭제(Delete), False면 왼쪽 삭제(Backspace)

        Returns:
            bool: 성공 여부
        """
        try:
            action = "Delete" if forward else "DeleteBack"
            self.hwp.HAction.Run(action)
            return True
        except Exception as e:
            print(f"delete_char 실패: {e}")
            return False


def main():
    """테스트 실행"""
    try:
        # 커서 래퍼 생성
        cursor = CursorWrapper()

        print("=== 커서 래퍼 테스트 ===\n")

        # 위치 정보 출력
        cursor.print_info()

        # 개별 정보 조회
        print("=== 개별 정보 조회 ===")
        print(f"페이지: {cursor.get_page()}")
        print(f"줄: {cursor.get_line()}")
        print(f"내부 위치: {cursor.get_pos()}")
        print(f"컨트롤 ID: {cursor.get_ctrl_id()}")
        print(f"표 안에 있나요? {cursor.is_in_table()}")
        print()

        # 커서 이동 테스트
        print("=== 커서 이동 테스트 ===")

        # 위치 저장
        saved = cursor.save_pos()
        print(f"현재 위치 저장: {saved}")

        # 줄 끝으로 이동
        cursor.move_line_end()
        print(f"줄 끝으로 이동: {cursor.get_pos()}")

        # 줄 시작으로 이동
        cursor.move_line_begin()
        print(f"줄 시작으로 이동: {cursor.get_pos()}")

        # 위치 복원
        cursor.restore_pos(saved)
        print(f"위치 복원: {cursor.get_pos()}")
        print()

        # 텍스트 조작 테스트
        print("=== 텍스트 조작 테스트 ===")
        char = cursor.get_char_at_cursor()
        print(f"현재 위치 글자: '{char}'")

    except Exception as e:
        print(f"오류 발생: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
