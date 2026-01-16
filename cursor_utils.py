# -*- coding: utf-8 -*-
"""
한글(HWP) 커서 유틸리티

커서 위치 정보 조회 및 범위 계산을 위한 유틸리티 함수 모음.
폴링 루프(모니터링)는 별도 모듈에서 처리.

함수 네이밍 규칙:
- get_hwp_instance: 유틸리티 (snake_case)
- get_*: 위치/범위 정보 조회 함수 (snake_case)

반환값 키 규칙:
- current: 현재 위치 튜플 (list_id, para_id, char_pos)
- start / end: 범위의 시작/끝 pos
- index: 순번 (1부터 시작)
- text: 텍스트 내용

주요 함수:
1. get_hwp_instance() -> 실행 중인 한글 인스턴스 연결
2. get_current_pos(hwp) -> 커서 위치 + 상태바 정보
3. get_para_range(hwp) -> 문단 시작/끝
4. get_line_range(hwp) -> 줄 시작/끝
5. get_sentences(hwp) -> 문장 경계 목록
6. get_cursor_index(hwp) -> 현재 커서의 문장/단어 인덱스
"""

import pythoncom
import win32com.client as win32


def get_hwp_instance():
    """실행 중인 한글 인스턴스에 연결 (ROT 사용)"""
    context = pythoncom.CreateBindCtx(0)
    rot = pythoncom.GetRunningObjectTable()

    for moniker in rot:
        name = moniker.GetDisplayName(context, None)
        if 'HwpObject' in name:
            obj = rot.GetObject(moniker)
            return win32.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))

    return None


def get_current_pos(hwp):
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


def get_para_range(hwp):
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


def get_line_range(hwp):
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


def get_sentences(hwp, include_text=False):
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


def get_cursor_index(hwp, pos=None):
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

    # pos가 속한 문장 찾기
    current_sentence = None
    for s in sentences:
        if s['start'] <= pos <= s['end']:
            current_sentence = s
            break

    if current_sentence is None:
        return None

    # 문장 시작부터 pos까지 공백 개수 = 단어 인덱스
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


if __name__ == "__main__":
    print("=== 커서 유틸리티 테스트 ===\n")

    hwp = get_hwp_instance()
    if hwp is None:
        print("[오류] 한글이 실행 중이지 않습니다. ROT 연결 실패.")
        exit(1)

    print("[OK] 한글 인스턴스 연결 성공\n")

    # 현재 커서 위치
    pos_info = get_current_pos(hwp)
    print(f"현재 위치:")
    print(f"  - 문단 ID: {pos_info['para_id']}")
    print(f"  - 글자 위치: {pos_info['char_pos']}")
    print(f"  - 페이지: {pos_info['page']}, 줄: {pos_info['line']}, 칸: {pos_info['column']}")
    print(f"  - 모드: {pos_info['insert_mode']}\n")

    # 문단 범위
    para_range = get_para_range(hwp)
    print(f"문단 범위: {para_range['start']} ~ {para_range['end']}\n")

    # 줄 범위
    line_range = get_line_range(hwp)
    print(f"현재 줄 범위: {line_range['start']} ~ {line_range['end']}")
    print(f"문단 내 줄 시작점들: {line_range['line_starts']}\n")

    # 문장 목록
    sentences = get_sentences(hwp)
    print(f"문장 개수: {len(sentences)}")
    for s in sentences[:3]:  # 처음 3개만 출력
        print(f"  문장 {s['index']}: pos {s['start']} ~ {s['end']}")
    if len(sentences) > 3:
        print(f"  ... (총 {len(sentences)}개)")
    print()

    # 커서 인덱스
    cursor_idx = get_cursor_index(hwp)
    if cursor_idx:
        print(f"커서 인덱스:")
        print(f"  - 문장 번호: {cursor_idx['sentence_index']}")
        print(f"  - 단어 번호: {cursor_idx['word_index']}")

    print("\n=== 테스트 완료 ===")
