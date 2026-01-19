# -*- coding: utf-8 -*-
"""
한글(HWP) 커서/위치/상태 통합 모듈

이 모듈은 다음 기능을 제공합니다:
1. HWP 인스턴스 연결 (ROT)
2. 위치 정보 조회 (GetPos, KeyIndicator)
3. 범위 정보 (문단, 줄, 문장)
4. 커서 이동 및 선택
5. 컨트롤 정보 조회
6. 실시간 모니터링

사용 예시:
    from cursor import get_hwp_instance, Cursor

    hwp = get_hwp_instance()
    cursor = Cursor(hwp)
    cursor.print_info()
"""

import os
import sys
import time
import pythoncom
import win32com.client as win32
from typing import Optional, Dict, Tuple, Any, List


# =============================================================================
# 1. 상수 및 설정
# =============================================================================

# 컨트롤 ID → 이름 매핑
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
    'gso': '그리기 개체',
    'form': '양식 개체',
    '+pbt': '명령 단추',
    '+rbt': '라디오 단추',
    '+cbt': '선택 상자',
    '+cob': '콤보 상자',
    '+edt': '입력 상자'
}


# =============================================================================
# 2. 유틸리티 함수
# =============================================================================

def get_hwp_instance():
    """
    실행 중인 한글 인스턴스에 연결 (ROT 사용)

    Returns:
        hwp: HWP COM 객체 또는 None
    """
    context = pythoncom.CreateBindCtx(0)
    rot = pythoncom.GetRunningObjectTable()

    for moniker in rot:
        name = moniker.GetDisplayName(context, None)
        if 'HwpObject' in name:
            obj = rot.GetObject(moniker)
            return win32.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))

    return None


def get_current_pos(hwp) -> Dict[str, Any]:
    """
    현재 커서 위치 정보 반환

    Returns:
        dict: {
            'list_id': 리스트 ID,
            'para_id': 문단 ID,
            'char_pos': 문단 내 글자 위치,
            'page': 페이지,
            'line': 줄,
            'column': 칸,
            'insert_mode': '삽입' 또는 '수정'
        }
    """
    pos = hwp.GetPos()
    key = hwp.KeyIndicator()

    return {
        'list_id': pos[0],
        'para_id': pos[1],
        'char_pos': pos[2],
        'page': key[2],
        'line': key[4],
        'column': key[5],
        'insert_mode': '수정' if key[6] else '삽입'
    }


def get_para_range(hwp) -> Dict[str, Any]:
    """
    문단 시작/끝 pos 반환

    Returns:
        dict: {
            'current': (list_id, para_id, char_pos),
            'start': 문단 시작 pos,
            'end': 문단 끝 pos
        }
    """
    current = hwp.GetPos()
    hwp.HAction.Run("MoveParaBegin")
    start = hwp.GetPos()[2]
    hwp.HAction.Run("MoveParaEnd")
    end = hwp.GetPos()[2]
    hwp.SetPos(current[0], current[1], current[2])

    return {
        'current': current,
        'start': start,
        'end': end
    }


def get_line_range(hwp) -> Dict[str, Any]:
    """
    현재 줄의 시작/끝 pos 반환

    Returns:
        dict: {
            'current': (list_id, para_id, char_pos),
            'start': 줄 시작 pos,
            'end': 줄 끝 pos,
            'line_starts': 문단 내 모든 줄 시작 pos 목록
        }
    """
    current = hwp.GetPos()
    init_para, init_pos = current[1], current[2]

    hwp.HAction.Run("MoveParaEnd")
    para_end = hwp.GetPos()[2]
    hwp.HAction.Run("MoveParaBegin")

    line_starts = [0]
    current_line_start = 0
    next_line_start = None

    while True:
        hwp.HAction.Run("MoveLineDown")
        pos = hwp.GetPos()
        if pos[1] != init_para:
            break
        if pos[2] > init_pos:
            next_line_start = pos[2]
            break
        line_starts.append(pos[2])
        current_line_start = pos[2]

    line_end = next_line_start - 1 if next_line_start else para_end
    hwp.SetPos(current[0], current[1], current[2])

    return {
        'current': current,
        'start': current_line_start,
        'end': line_end,
        'line_starts': line_starts
    }


def get_sentences(hwp, include_text: bool = False):
    """
    문단 내 문장 경계 반환 ('.' 기준)

    Args:
        hwp: 한글 인스턴스
        include_text: True면 문단 텍스트도 함께 반환

    Returns:
        include_text=False:
            list: [{'index': 1, 'start': 0, 'end': 15}, ...]
        include_text=True:
            tuple: (sentences_list, para_text)
    """
    current = hwp.GetPos()

    hwp.HAction.Run("MoveParaBegin")
    start_pos = hwp.GetPos()
    hwp.HAction.Run("MoveParaEnd")
    end_pos = hwp.GetPos()

    hwp.SelectText(start_pos[1], start_pos[2], end_pos[1], end_pos[2])
    para_text = hwp.GetTextFile("TEXT", "saveblock")
    hwp.HAction.Run("Cancel")
    hwp.SetPos(current[0], current[1], current[2])

    if para_text is None or not para_text.strip():
        return ([], "") if include_text else []

    para_text = para_text.replace('\r\n', '').replace('\r', '').replace('\n', '')

    sentences = []
    sentence_index = 1
    sentence_start = 0
    i = 0

    while i < len(para_text):
        if para_text[i] == '.':
            sentences.append({'index': sentence_index, 'start': sentence_start, 'end': i})

            next_start = i + 1
            skip_count = 0
            while next_start < len(para_text) and skip_count < 3:
                if para_text[next_start] == ' ':
                    next_start += 1
                    skip_count += 1
                else:
                    break

            if next_start >= len(para_text) or (skip_count >= 3 and para_text[next_start] == ' '):
                break

            sentence_index += 1
            sentence_start = next_start
            i = next_start
        else:
            i += 1

    if sentence_start < len(para_text) and (not sentences or sentences[-1]['end'] < len(para_text) - 1):
        sentences.append({'index': sentence_index, 'start': sentence_start, 'end': len(para_text) - 1})

    return (sentences, para_text) if include_text else sentences


def get_cursor_index(hwp, pos: Optional[int] = None) -> Optional[Dict[str, Any]]:
    """
    현재 커서가 몇 번째 문장의 몇 번째 단어인지 반환

    Args:
        hwp: 한글 인스턴스
        pos: 문단 내 위치 (None이면 현재 커서 위치)

    Returns:
        dict: {
            'sentence_index': 문장 번호 (1부터),
            'word_index': 단어 번호 (0부터),
            'sentence_start': 문장 시작 pos,
            'sentence_end': 문장 끝 pos
        }
        또는 None (위치를 찾을 수 없는 경우)
    """
    if pos is None:
        pos = hwp.GetPos()[2]

    sentences, para_text = get_sentences(hwp, include_text=True)
    if not sentences or not para_text:
        return None

    current_sentence = None
    for s in sentences:
        if s['start'] <= pos <= s['end']:
            current_sentence = s
            break

    if current_sentence is None:
        return None

    word_index = 0
    for i in range(current_sentence['start'], pos):
        if i < len(para_text) and para_text[i] == ' ':
            word_index += 1

    return {
        'sentence_index': current_sentence['index'],
        'word_index': word_index,
        'sentence_start': current_sentence['start'],
        'sentence_end': current_sentence['end']
    }


# =============================================================================
# 3. 정보 클래스
# =============================================================================

class KeyIndicatorInfo:
    """상태바 위치 정보 클래스"""

    def __init__(self, hwp):
        self.hwp = hwp

    def get_info(self) -> Dict[str, Any]:
        """
        현재 커서의 상태바 정보 조회

        Returns:
            dict: 상태바 정보
        """
        try:
            key = self.hwp.KeyIndicator()

            return {
                'total_sections': key[0],
                'current_section': key[1],
                'unknown_2': key[2],
                'page': key[3],
                'line': key[4],
                'line_info': key[5],
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
    """내부 위치 정보 클래스"""

    def __init__(self, hwp):
        self.hwp = hwp

    def get_pos(self) -> Tuple[int, int, int]:
        """
        현재 커서의 내부 위치 조회

        Returns:
            tuple: (list_id, para_id, char_pos)
        """
        try:
            pos = self.hwp.GetPos()
            return (pos[0], pos[1], pos[2])
        except Exception as e:
            print(f"GetPos 조회 실패: {e}")
            return (0, 0, 0)

    def get_info(self) -> Dict[str, Any]:
        """내부 위치 정보를 딕셔너리로 반환"""
        list_id, para_id, char_pos = self.get_pos()

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
        """커서를 특정 위치로 이동"""
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
    """컨트롤 정보 클래스"""

    def __init__(self, hwp):
        self.hwp = hwp

    def find_ctrl(self) -> str:
        """현재 커서 위치의 컨트롤 ID 조회"""
        try:
            ctrl_id = self.hwp.FindCtrl()
            return ctrl_id if ctrl_id else ''
        except Exception as e:
            print(f"FindCtrl 실패: {e}")
            return ''

    def get_info(self) -> Dict[str, Any]:
        """현재 위치의 컨트롤 정보 조회"""
        ctrl_id = self.find_ctrl()

        if ctrl_id:
            ctrl_name = CTRL_NAMES.get(ctrl_id, '알 수 없는 컨트롤')
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
        """특정 컨트롤 선택"""
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


# =============================================================================
# 4. 메인 클래스 (통합)
# =============================================================================

class Cursor:
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
        """현재 커서의 내부 위치 조회"""
        return self.pos_info.get_pos()

    def get_page(self) -> int:
        """현재 페이지 번호 조회"""
        info = self.key_indicator.get_info()
        return info.get('page', 0)

    def get_line(self) -> int:
        """현재 줄 번호 조회"""
        info = self.key_indicator.get_info()
        return info.get('line', 0)

    def get_ctrl_id(self) -> str:
        """현재 위치의 컨트롤 ID 조회"""
        return self.ctrl_info.find_ctrl()

    def is_in_table(self) -> bool:
        """커서가 표 안에 있는지 확인"""
        return self.get_ctrl_id() == 'tbl'

    def is_in_ctrl(self, ctrl_id: str) -> bool:
        """커서가 특정 컨트롤 안에 있는지 확인"""
        return self.get_ctrl_id() == ctrl_id

    def get_all_info(self) -> Dict[str, Any]:
        """모든 위치 정보 조회"""
        return {
            'key_indicator': self.key_indicator.get_info(),
            'pos': self.pos_info.get_info(),
            'ctrl': self.ctrl_info.get_info()
        }

    def print_info(self):
        """현재 커서의 모든 정보 출력"""
        print("\n" + "=" * 50)
        print("현재 커서 위치 정보")
        print("=" * 50 + "\n")

        self.key_indicator.print_info()
        print()
        self.pos_info.print_info()
        print()
        self.ctrl_info.print_info()
        print()

    # ==================== 커서 이동 (문단) ====================

    def move_para_begin(self) -> bool:
        """문단 시작으로 이동"""
        try:
            self.hwp.HAction.Run("MoveParaBegin")
            return True
        except Exception as e:
            print(f"MoveParaBegin 실패: {e}")
            return False

    def move_para_end(self) -> bool:
        """문단 끝으로 이동"""
        try:
            self.hwp.HAction.Run("MoveParaEnd")
            return True
        except Exception as e:
            print(f"MoveParaEnd 실패: {e}")
            return False

    # ==================== 커서 이동 (줄) ====================

    def move_line_begin(self) -> bool:
        """줄 시작으로 이동 (Home 키)"""
        try:
            self.hwp.HAction.Run("MoveLineBegin")
            return True
        except Exception as e:
            print(f"MoveLineBegin 실패: {e}")
            return False

    def move_line_end(self) -> bool:
        """줄 끝으로 이동 (End 키)"""
        try:
            self.hwp.HAction.Run("MoveLineEnd")
            return True
        except Exception as e:
            print(f"MoveLineEnd 실패: {e}")
            return False

    # ==================== 커서 이동 (문서) ====================

    def move_doc_begin(self) -> bool:
        """문서 처음으로 이동 (Ctrl+Home)"""
        try:
            self.hwp.HAction.Run("MoveDocBegin")
            return True
        except Exception as e:
            print(f"MoveDocBegin 실패: {e}")
            return False

    def move_doc_end(self) -> bool:
        """문서 끝으로 이동 (Ctrl+End)"""
        try:
            self.hwp.HAction.Run("MoveDocEnd")
            return True
        except Exception as e:
            print(f"MoveDocEnd 실패: {e}")
            return False

    # ==================== 커서 이동 (글자/줄 단위) ====================

    def move_left(self, count: int = 1) -> bool:
        """왼쪽으로 이동"""
        try:
            for _ in range(count):
                self.hwp.HAction.Run("MoveLeft")
            return True
        except Exception as e:
            print(f"MoveLeft 실패: {e}")
            return False

    def move_right(self, count: int = 1) -> bool:
        """오른쪽으로 이동"""
        try:
            for _ in range(count):
                self.hwp.HAction.Run("MoveRight")
            return True
        except Exception as e:
            print(f"MoveRight 실패: {e}")
            return False

    def move_up(self, count: int = 1) -> bool:
        """위로 이동"""
        try:
            for _ in range(count):
                self.hwp.HAction.Run("MoveUp")
            return True
        except Exception as e:
            print(f"MoveUp 실패: {e}")
            return False

    def move_down(self, count: int = 1) -> bool:
        """아래로 이동"""
        try:
            for _ in range(count):
                self.hwp.HAction.Run("MoveDown")
            return True
        except Exception as e:
            print(f"MoveDown 실패: {e}")
            return False

    # ==================== 위치 설정 ====================

    def set_pos(self, list_id: int, para_id: int, char_pos: int) -> bool:
        """커서를 특정 위치로 이동"""
        return self.pos_info.set_pos(list_id, para_id, char_pos)

    def save_pos(self) -> Tuple[int, int, int]:
        """현재 위치 저장"""
        return self.get_pos()

    def restore_pos(self, pos: Tuple[int, int, int]) -> bool:
        """저장된 위치로 복원"""
        return self.set_pos(pos[0], pos[1], pos[2])

    # ==================== 텍스트 선택 ====================

    def select_all(self) -> bool:
        """전체 선택 (Ctrl+A)"""
        try:
            self.hwp.HAction.Run("SelectAll")
            return True
        except Exception as e:
            print(f"SelectAll 실패: {e}")
            return False

    def select_line(self) -> bool:
        """현재 줄 선택"""
        try:
            self.hwp.HAction.Run("MoveLineBegin")
            self.hwp.HAction.Run("MoveSelLineEnd")
            return True
        except Exception as e:
            print(f"select_line 실패: {e}")
            return False

    def select_para(self) -> bool:
        """현재 문단 선택"""
        try:
            self.hwp.HAction.Run("MoveParaBegin")
            self.hwp.HAction.Run("MoveSelParaEnd")
            return True
        except Exception as e:
            print(f"select_para 실패: {e}")
            return False

    def cancel_selection(self) -> bool:
        """선택 해제 (Esc)"""
        try:
            self.hwp.HAction.Run("Cancel")
            return True
        except Exception as e:
            print(f"Cancel 실패: {e}")
            return False

    # ==================== 텍스트 조작 ====================

    def get_selected_text(self) -> str:
        """선택된 텍스트 가져오기"""
        try:
            text = self.hwp.GetTextFile("TEXT", "saveblock")
            return text if text else ""
        except Exception as e:
            print(f"GetTextFile 실패: {e}")
            return ""

    def get_char_at_cursor(self) -> str:
        """커서 위치의 한 글자 가져오기"""
        try:
            saved_pos = self.save_pos()
            self.hwp.HAction.Run("MoveSelRight")
            text = self.get_selected_text()
            self.restore_pos(saved_pos)
            return text
        except Exception as e:
            print(f"get_char_at_cursor 실패: {e}")
            return ""

    def insert_text(self, text: str) -> bool:
        """텍스트 삽입"""
        try:
            self.hwp.HAction.GetDefault("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
            self.hwp.HParameterSet.HInsertText.Text = text
            self.hwp.HAction.Execute("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
            return True
        except Exception as e:
            print(f"insert_text 실패: {e}")
            return False

    def delete_char(self, forward: bool = True) -> bool:
        """글자 삭제"""
        try:
            action = "Delete" if forward else "DeleteBack"
            self.hwp.HAction.Run(action)
            return True
        except Exception as e:
            print(f"delete_char 실패: {e}")
            return False


# =============================================================================
# 5. 모니터링 함수
# =============================================================================

def clear_screen():
    """터미널 화면 지우기"""
    os.system('cls' if os.name == 'nt' else 'clear')


def monitor_position(interval: float = 0.1, show_raw: bool = False):
    """
    커서 위치를 실시간으로 모니터링

    Args:
        interval: 갱신 주기 (초, 기본 0.1초)
        show_raw: True면 원본 튜플도 표시
    """
    hwp = get_hwp_instance()
    if not hwp:
        print("한글이 실행 중이 아닙니다")
        return

    hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')

    print("실시간 커서 위치 모니터링 시작")
    print("종료: Ctrl+C\n")
    time.sleep(1)

    prev_pos = None
    prev_key = None
    first_run = True

    try:
        while True:
            pos = hwp.GetPos()
            key = hwp.KeyIndicator()

            if first_run or pos != prev_pos or key != prev_key:
                if not first_run:
                    clear_screen()
                first_run = False

                print("=" * 70)
                print("실시간 커서 위치 모니터링 (Ctrl+C로 종료)")
                print("=" * 70)

                print("\n[GetPos]")
                if show_raw:
                    print(f"  원본: {pos}")
                print(f"  list_id:  {pos[0]:3d}  ", end="")
                if pos[0] == 0:
                    print("(본문)")
                elif pos[0] >= 10:
                    print("(테이블/특수영역)")
                else:
                    print("(머리말/꼬리말 등)")

                print(f"  para_id:  {pos[1]:3d}")
                print(f"  char_pos: {pos[2]:3d}")

                print("\n[KeyIndicator]")
                print(f"  튜플 길이: {len(key)}")
                if show_raw:
                    print(f"  원본: {key}")
                print(f"  [0] total_section:   {key[0]:3d}")
                print(f"  [1] current_section: {key[1]:3d}")
                print(f"  [2] page:            {key[2]:3d}")
                print(f"  [3] column_num (단): {key[3]:3d}")
                print(f"  [4] line:            {key[4]:3d}")
                print(f"  [5] column (칸):     {key[5]:3d}")
                print(f"  [6] insert_mode:     {key[6]:3d} ({'수정' if key[6] else '삽입'})")
                if len(key) > 7:
                    print(f"  [7] ctrlname:        {key[7]}")

                try:
                    act = hwp.CreateAction("TableCellBorder")
                    pset = act.GetDefault("TableCellBorder", hwp.HParameterSet.HTableCellBorder.HSet)
                    in_table = act.Execute(pset)
                except:
                    in_table = False

                print("\n[위치 유형]")
                if in_table:
                    print("  테이블 내부")
                elif pos[0] == 0:
                    print("  본문")
                else:
                    print(f"  특수 영역 (list_id={pos[0]})")

                print("\n" + "=" * 70)
                print(f"갱신 주기: {interval}초")

                prev_pos = pos
                prev_key = key

            time.sleep(interval)

    except KeyboardInterrupt:
        print("\n\n모니터링을 종료합니다.")


def print_cursor_position(hwp=None):
    """현재 커서 위치 정보 출력 (단발성)"""
    if hwp is None:
        hwp = get_hwp_instance()
    if not hwp:
        print("한글이 실행 중이 아닙니다")
        return

    cursor = Cursor(hwp)
    cursor.print_info()


# =============================================================================
# 6. CLI 인터페이스
# =============================================================================

def main():
    """CLI 메인 함수"""
    import argparse

    parser = argparse.ArgumentParser(description='한글(HWP) 커서 위치 정보 조회')
    parser.add_argument('-m', '--monitor', action='store_true', help='실시간 모니터링 모드')
    parser.add_argument('-i', '--interval', type=float, default=0.1, help='모니터링 갱신 주기 (초)')
    parser.add_argument('-r', '--raw', action='store_true', help='원본 튜플 표시')

    args = parser.parse_args()

    if args.monitor:
        monitor_position(interval=args.interval, show_raw=args.raw)
    else:
        hwp = get_hwp_instance()
        if not hwp:
            print("한글이 실행 중이 아닙니다")
            return

        cursor = Cursor(hwp)
        cursor.print_info()

        # 추가 정보
        print("=== 범위 정보 ===")
        para_range = get_para_range(hwp)
        print(f"문단 범위: {para_range['start']} ~ {para_range['end']}")

        line_range = get_line_range(hwp)
        print(f"현재 줄 범위: {line_range['start']} ~ {line_range['end']}")

        sentences = get_sentences(hwp)
        print(f"문장 개수: {len(sentences)}")

        cursor_idx = get_cursor_index(hwp)
        if cursor_idx:
            print(f"커서 인덱스: 문장 {cursor_idx['sentence_index']}, 단어 {cursor_idx['word_index']}")


if __name__ == "__main__":
    main()
