# -*- coding: utf-8 -*-
"""
한글(HWP) 커서 위치 모니터링 - 폴링 방식

주요 구현 로직:

1. 위치 정보 획득
   - GetPos(): (list_id, para_id, char_pos) 반환
   - KeyIndicator(): 페이지, 줄, 칸 등 상태바 정보

2. 문단 범위 (GetParaRange)
   - MoveParaBegin → GetPos() → 문단 시작
   - MoveParaEnd → GetPos() → 문단 끝
   - SetPos()로 원래 위치 복원

3. 줄 범위 (GetLineRange)
   - MoveDown으로 아래로 이동하며 줄 시작 pos 수집
   - para가 바뀌거나 초기 pos보다 커지면 종료
   - 현재 줄 끝 = 다음 줄 시작 - 1 (없으면 문단 끝)

4. 문장 경계 (GetSentenceIndex)
   - 문단 전체 선택 후 GetTextFile("TEXT", "saveblock")로 텍스트 획득
   - Python에서 '.' 위치 찾아 문장 경계 계산
   - '.' 후 3글자 이내 빈칸 아니면 새 문장 시작

5. 폴링 모니터링
   - 0.1초 간격으로 GetPos() 비교
   - 위치 변경 시 콜백 호출
"""

import time
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


def GetCurrentPos(hwp):
    """현재 커서 위치 정보 반환 (GetPos + KeyIndicator)"""
    result = {}
    pos = hwp.GetPos()
    result['list_id'] = pos[0]
    result['para_id'] = pos[1]
    result['char_pos'] = pos[2]

    key = hwp.KeyIndicator()
    result['page'] = key[2]
    result['line'] = key[4]
    result['pos'] = key[5]
    result['insert_mode'] = '수정' if key[6] else '삽입'

    return result


def GetParaRange(hwp):
    """문단 시작/끝 pos 반환"""
    current = hwp.GetPos()
    hwp.HAction.Run("MoveParaBegin")
    para_start = hwp.GetPos()
    hwp.HAction.Run("MoveParaEnd")
    para_end = hwp.GetPos()
    hwp.SetPos(current[0], current[1], current[2])
    return {'current': current, 'para_start': para_start, 'para_end': para_end}


def GetLineRange(hwp):
    """현재 줄의 시작/끝 pos 반환"""
    current = hwp.GetPos()
    init_para, init_pos = current[1], current[2]

    hwp.HAction.Run("MoveParaEnd")
    para_end_pos = hwp.GetPos()[2]
    hwp.HAction.Run("MoveParaBegin")

    line_heads = [0]
    current_line_head = 0
    next_line_head = None

    while True:
        hwp.HAction.Run("MoveDown")
        pos = hwp.GetPos()
        if pos[1] != init_para:
            break
        if pos[2] > init_pos:
            next_line_head = pos[2]
            break
        line_heads.append(pos[2])
        current_line_head = pos[2]

    line_tail_pos = next_line_head - 1 if next_line_head else para_end_pos
    hwp.SetPos(current[0], current[1], current[2])

    return {
        'current_pos': current,
        'line_head_pos': current_line_head,
        'line_tail_pos': line_tail_pos,
        'all_line_heads': line_heads
    }


def GetSentenceIndex(hwp, return_text=False):
    """문단 내 문장 경계 반환 ('.' 기준)"""
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
        return ([], "") if return_text else []

    para_text = para_text.replace('\r\n', '').replace('\r', '').replace('\n', '')

    sentences = []
    sentence_index = 1
    sentence_start = 0
    i = 0

    while i < len(para_text):
        if para_text[i] == '.':
            sentences.append({'index': sentence_index, 'start_pos': sentence_start, 'end_pos': i})

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

    if sentence_start < len(para_text) and (not sentences or sentences[-1]['end_pos'] < len(para_text) - 1):
        sentences.append({'index': sentence_index, 'start_pos': sentence_start, 'end_pos': len(para_text) - 1})

    return (sentences, para_text) if return_text else sentences


def GetTextIndex(hwp, pos=None):
    """
    주어진 pos가 몇 번째 문장의 몇 번째 단어인지 반환

    Args:
        hwp: 한글 인스턴스
        pos: 문단 내 위치 (None이면 현재 커서 위치)

    Returns:
        dict: {
            'sentence_index': 문장 번호 (1부터),
            'word_index': 문장 내 단어 번호 (0부터),
            'sentence_start': 문장 시작 pos,
            'sentence_end': 문장 끝 pos
        }
    """
    if pos is None:
        pos = hwp.GetPos()[2]

    # 문장 목록과 텍스트를 한 번에 가져오기 (블록 1회만 사용)
    sentences, para_text = GetSentenceIndex(hwp, return_text=True)
    if not sentences or not para_text:
        return None

    # pos가 어느 문장에 속하는지 찾기
    current_sentence = None
    for s in sentences:
        if s['start_pos'] <= pos <= s['end_pos']:
            current_sentence = s
            break

    if current_sentence is None:
        return None

    # 문장 내에서 단어 인덱스 계산
    # 문장 시작부터 pos까지 공백 개수 세기
    word_index = 0
    for i in range(current_sentence['start_pos'], pos):
        if i < len(para_text) and para_text[i] == ' ':
            word_index += 1

    return {
        'sentence_index': current_sentence['index'],
        'word_index': word_index,
        'sentence_start': current_sentence['start_pos'],
        'sentence_end': current_sentence['end_pos']
    }


def position_monitor_loop(hwp, interval=0.1, callback=None):
    """폴링 방식 커서 위치 모니터링"""
    last_pos = None
    while True:
        try:
            pos = hwp.GetPos()
            current_key = (pos[0], pos[1], pos[2])
            if current_key != last_pos:
                current = GetCurrentPos(hwp)
                if callback:
                    callback(current)
                else:
                    print(f"문단:{current['para_id']} 글자:{current['char_pos']} 줄:{current['line']}")
                last_pos = current_key
            time.sleep(interval)
        except KeyboardInterrupt:
            break


if __name__ == "__main__":
    hwp = win32.Dispatch('HWPFrame.HwpObject')
    hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')
    hwp.XHwpWindows.Item(0).Visible = True
    hwp.Open(r"d:\hwp_docs\test.hwp", "HWP", "forceopen:true")
    hwp.EditMode = 1

    print("텍스트 인덱스 테스트\n")
    last_pos = None
    while True:
        try:
            pos = hwp.GetPos()
            current_key = (pos[0], pos[1], pos[2])
            if current_key != last_pos:
                result = GetTextIndex(hwp)
                if result:
                    print(f"문장 {result['sentence_index']} / 단어 {result['word_index']} | 문장범위: {result['sentence_start']}~{result['sentence_end']}")
                else:
                    print("위치 없음")
                last_pos = current_key
            time.sleep(0.1)
        except KeyboardInterrupt:
            print("\n종료")
            break
