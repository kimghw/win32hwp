# -*- coding: utf-8 -*-
"""
한글(HWP) 커서 위치 모니터링 - 폴링 방식
GetCurrentPos(): 커서 위치에서 가져올 수 있는 모든 값
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
    """
    현재 커서 위치에서 가져올 수 있는 모든 정보 반환

    Returns:
        dict: 커서 위치 정보
            - list_id: 리스트 ID
            - para_id: 문단 ID
            - char_pos: 문단 내 글자 위치
            - total_sec: 총 구역
            - cur_sec: 현재 구역
            - page: 페이지
            - column: 단
            - line: 줄
            - pos: 칸
            - insert_mode: 삽입/수정 모드
    """
    result = {}

    # 1. GetPos() - 기본 위치 (list, para, pos)
    pos = hwp.GetPos()
    result['list_id'] = pos[0]      # 리스트 ID
    result['para_id'] = pos[1]      # 문단 ID
    result['char_pos'] = pos[2]     # 문단 내 글자 위치

    # 2. KeyIndicator() - 상태바 정보
    key = hwp.KeyIndicator()
    result['total_sec'] = key[0]    # 총 구역
    result['cur_sec'] = key[1]      # 현재 구역
    result['page'] = key[2]         # 페이지
    result['column'] = key[3]       # 단
    result['line'] = key[4]         # 줄
    result['pos'] = key[5]          # 칸
    result['insert_mode'] = '수정' if key[6] else '삽입'

    # 3. GetPosBySet() - ParameterSet으로 위치 정보
    try:
        pos_set = hwp.GetPosBySet()
        result['set_list'] = pos_set.Item("List")
        result['set_para'] = pos_set.Item("Para")
        result['set_pos'] = pos_set.Item("Pos")
    except:
        pass

    return result


def GetParaRange(hwp):
    """
    현재 커서 위치에서 문단의 시작과 끝 pos 값을 추출

    Returns:
        dict: 문단 범위 정보
            - current: 현재 위치 (list, para, pos)
            - para_start: 문단 시작 위치 (list, para, pos)
            - para_end: 문단 끝 위치 (list, para, pos)
    """
    # 현재 위치 저장
    current = hwp.GetPos()

    # 문단 시작으로 이동
    hwp.HAction.Run("MoveParaBegin")
    para_start = hwp.GetPos()

    # 문단 끝으로 이동
    hwp.HAction.Run("MoveParaEnd")
    para_end = hwp.GetPos()

    # 원래 위치로 복원
    hwp.SetPos(current[0], current[1], current[2])

    return {
        'current': current,
        'para_start': para_start,
        'para_end': para_end
    }


def GetLineRange(hwp):
    """
    현재 커서가 위치한 줄의 시작과 끝 pos를 반환

    로직:
    1. 현재 para, pos 저장
    2. 문단 시작(pos=0)으로 이동
    3. MoveDown으로 아래로 이동하며 각 줄의 시작 pos 수집
    4. 초기 pos보다 커지거나 para가 바뀌면 종료
    5. 현재 줄의 시작/끝 pos 반환

    Returns:
        dict: 줄 위치 정보
            - current_pos: 현재 위치 (list, para, pos)
            - line_head_pos: 현재 줄의 시작 pos
            - line_tail_pos: 현재 줄의 끝 pos
            - all_line_heads: 문단 내 모든 줄의 시작 pos 목록
    """
    # 현재 위치 저장
    current = hwp.GetPos()
    init_para = current[1]
    init_pos = current[2]

    # 문단 끝 pos 가져오기
    hwp.HAction.Run("MoveParaEnd")
    para_end_pos = hwp.GetPos()[2]

    # 문단 시작으로 이동
    hwp.HAction.Run("MoveParaBegin")

    # 각 줄의 시작 pos 수집
    line_heads = [0]  # 첫 줄은 항상 0
    current_line_head = 0  # 현재 줄의 시작 pos
    next_line_head = None  # 다음 줄의 시작 pos

    while True:
        # 아래로 이동
        hwp.HAction.Run("MoveDown")
        pos = hwp.GetPos()

        # para가 바뀌면 종료
        if pos[1] != init_para:
            break

        # 현재 pos
        cur_pos = pos[2]

        # 초기 pos보다 크면 종료 (현재 줄을 지나침)
        if cur_pos > init_pos:
            next_line_head = cur_pos
            break

        # 줄 시작 pos 저장
        line_heads.append(cur_pos)
        current_line_head = cur_pos

    # 현재 줄의 끝 pos 계산
    if next_line_head is not None:
        # 다음 줄이 있으면 다음 줄 시작 - 1
        line_tail_pos = next_line_head - 1
    else:
        # 마지막 줄이면 문단 끝
        line_tail_pos = para_end_pos

    # 원래 위치로 복원
    hwp.SetPos(current[0], current[1], current[2])

    return {
        'current_pos': current,
        'line_head_pos': current_line_head,
        'line_tail_pos': line_tail_pos,
        'all_line_heads': line_heads
    }


def GetSentenceIndex(hwp):
    """
    현재 커서가 위치한 문단에서 문장 경계를 찾아 반환

    로직:
    1. 문단 전체 텍스트를 한번에 가져옴
    2. Python에서 '.' 위치를 찾아 문장 경계 계산
    3. '.' 후 3글자 이내에 빈칸 아닌 글자가 있으면 새 문장

    Returns:
        list: 문장 정보 리스트
            - [{'index': 1, 'start_pos': 0, 'end_pos': 15}, ...]
    """
    # 현재 위치 저장
    current = hwp.GetPos()

    # 문단 전체 선택해서 텍스트 가져오기
    hwp.HAction.Run("MoveParaBegin")
    start_pos = hwp.GetPos()

    hwp.HAction.Run("MoveParaEnd")
    end_pos = hwp.GetPos()

    # 문단 전체 선택
    hwp.SetPos(start_pos[0], start_pos[1], start_pos[2])
    hwp.HAction.Run("Select")
    hwp.SetPos(end_pos[0], end_pos[1], end_pos[2])
    para_text = hwp.GetTextFile("TEXT", "saveblock")
    hwp.HAction.Run("Cancel")

    # None 체크 및 줄바꿈 제거
    if para_text is None:
        hwp.SetPos(current[0], current[1], current[2])
        return []
    para_text = para_text.replace('\r\n', '').replace('\r', '').replace('\n', '')

    print(f"문단 텍스트: '{para_text}'")

    # 원래 위치로 복원
    hwp.SetPos(current[0], current[1], current[2])

    # 빈 문단이면 빈 리스트 반환
    if not para_text.strip():
        return []

    # '.' 위치 찾기
    sentences = []
    sentence_index = 1
    sentence_start = 0
    i = 0

    while i < len(para_text):
        if para_text[i] == '.':
            # 문장 끝
            sentences.append({
                'index': sentence_index,
                'start_pos': sentence_start,
                'end_pos': i
            })

            # 다음 문장 시작점 찾기: 3글자 뒤 확인
            next_start = i + 1
            # 공백 건너뛰기 (최대 3글자)
            skip_count = 0
            while next_start < len(para_text) and skip_count < 3:
                if para_text[next_start] == ' ':
                    next_start += 1
                    skip_count += 1
                else:
                    break

            # 3글자 뒤에 빈칸이면 종료
            if next_start >= len(para_text) or (skip_count >= 3 and para_text[next_start] == ' '):
                break

            # 새 문장 시작
            sentence_index += 1
            sentence_start = next_start
            i = next_start
        else:
            i += 1

    # 마지막 문장 (문단 끝까지)
    if sentence_start < len(para_text) and (not sentences or sentences[-1]['end_pos'] < len(para_text) - 1):
        sentences.append({
            'index': sentence_index,
            'start_pos': sentence_start,
            'end_pos': len(para_text) - 1
        })

    return sentences


def position_monitor_loop(hwp, interval=0.1, callback=None):
    """
    커서 위치 모니터링 루프 (폴링 방식)

    Args:
        hwp: 한글 인스턴스
        interval: 감지 주기 (초), 기본 0.1초
        callback: 위치 변경 시 호출할 함수 (pos_info를 인자로 받음)
    """
    last_pos = None

    while True:
        try:
            pos = hwp.GetPos()
            current_key = (pos[0], pos[1], pos[2])

            if current_key != last_pos:
                current = GetCurrentPos(hwp)
                current['changed_time'] = time.time()

                if callback:
                    callback(current)
                else:
                    print(f"[{time.strftime('%H:%M:%S')}] 문단:{current['para_id']:3d} | 글자:{current['char_pos']:3d} | "
                          f"페이지:{current['page']} | 줄:{current['line']:2d} | 칸:{current['pos']:2d} | "
                          f"{current['insert_mode']}")

                last_pos = current_key

            time.sleep(interval)

        except KeyboardInterrupt:
            print("\n종료")
            break
        except Exception as e:
            print(f"오류: {e}")
            break


def para_range_loop(hwp, interval=0.1, callback=None):
    """
    문단 범위 모니터링 루프 (폴링 방식)

    Args:
        hwp: 한글 인스턴스
        interval: 감지 주기 (초), 기본 0.1초
        callback: 위치 변경 시 호출할 함수 (range_info를 인자로 받음)
    """
    last_pos = None

    while True:
        try:
            pos = hwp.GetPos()
            current_key = (pos[0], pos[1], pos[2])

            if current_key != last_pos:
                result = GetParaRange(hwp)

                if callback:
                    callback(result)
                else:
                    print(f"현재 위치: {result['current']}")
                    print(f"문단 시작: {result['para_start']}")
                    print(f"문단 끝  : {result['para_end']}")
                    print("-" * 40)

                last_pos = current_key

            time.sleep(interval)

        except KeyboardInterrupt:
            print("\n종료")
            break
        except Exception as e:
            print(f"오류: {e}")
            break


def main():
    print("문장 경계 탐지 테스트")
    print("-" * 60)

    # 새 한글 인스턴스 생성
    hwp = win32.Dispatch('HWPFrame.HwpObject')
    hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')
    hwp.XHwpWindows.Item(0).Visible = True
    hwp.Open(r"d:\hwp_docs\test.hwp", "HWP", "forceopen:true")
    hwp.EditMode = 1  # 편집 모드 (0:읽기, 1:편집)

    print("한글 연결됨 - 커서를 움직여보세요\n")

    # GetSentenceIndex 테스트
    last_pos = None
    while True:
        try:
            pos = hwp.GetPos()
            current_key = (pos[0], pos[1], pos[2])

            if current_key != last_pos:
                print(f"\n현재 위치: para={pos[1]}, pos={pos[2]}")
                sentences = GetSentenceIndex(hwp)
                print(f"문장 개수: {len(sentences)}")
                for s in sentences:
                    print(f"  문장 {s['index']}: pos {s['start_pos']} ~ {s['end_pos']}")
                print("-" * 40)
                last_pos = current_key

            time.sleep(0.1)

        except KeyboardInterrupt:
            print("\n종료")
            break


if __name__ == "__main__":
    main()
