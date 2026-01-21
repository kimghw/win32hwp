# -*- coding: utf-8 -*-
"""
위치 정보 조회 모듈

커서 위치, 문단/줄/문장 범위 등 HWP 문서 내 위치 관련 정보를 조회합니다.
"""

from typing import Dict, Tuple, Any, List, Optional


def get_current_pos(hwp) -> Dict[str, Any]:
    """
    현재 커서 위치 정보 반환

    Args:
        hwp: HWP COM 객체

    Returns:
        dict: {
            'list_id': 리스트 ID (0=본문, 10+=셀),
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

    Args:
        hwp: HWP COM 객체

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

    Args:
        hwp: HWP COM 객체

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
        hwp: HWP COM 객체
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
        hwp: HWP COM 객체
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


class KeyIndicatorInfo:
    """상태바 위치 정보 클래스"""

    def __init__(self, hwp):
        self.hwp = hwp

    def get_info(self) -> Dict[str, Any]:
        """
        현재 커서의 상태바 정보 조회

        Returns:
            dict: {
                'total_sections': 전체 구역 수,
                'current_section': 현재 구역,
                'page': 페이지 번호,
                'line': 줄 번호,
                'line_info': 줄 정보,
                'insert_mode': '삽입' 또는 '수정',
                'ctrl_name': 컨트롤 이름
            }
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
