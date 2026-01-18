# -*- coding: utf-8 -*-
"""
HWP 문서에 각 문단의 위치 정보(list_id, para_id, ctrl)를 파랑색으로 삽입

사용법:
1. 한글에서 1.hwp 파일을 열어둔 상태에서
2. python insert_pos_info.py 실행
"""

import pythoncom
import win32com.client as win32
from datetime import datetime
import os

# 로그 파일 설정
LOG_DIR = r"C:\win32hwp\logs"
os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE = os.path.join(LOG_DIR, f"pos_info_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")


def log(msg):
    """콘솔과 파일에 동시 로깅"""
    print(msg)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(msg + "\n")


def get_hwp_instance():
    """여러 방법으로 HWP 인스턴스 연결 시도"""

    # 방법 1: ROT (Running Object Table)
    try:
        context = pythoncom.CreateBindCtx(0)
        rot = pythoncom.GetRunningObjectTable()

        for moniker in rot:
            name = moniker.GetDisplayName(context, None)
            if 'HwpObject' in name:
                obj = rot.GetObject(moniker)
                hwp = win32.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
                log("ROT 방식으로 연결 성공")
                return hwp
    except Exception as e:
        log(f"ROT 연결 실패: {e}")

    # 방법 2: GetActiveObject
    try:
        hwp = win32.GetActiveObject("HWPFrame.HwpObject")
        log("GetActiveObject 방식으로 연결 성공")
        return hwp
    except Exception as e:
        log(f"GetActiveObject 실패: {e}")

    # 방법 3: 새로 생성 후 1.hwp 파일 열기
    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
        hwp.XHwpWindows.Item(0).Visible = True

        file_path = r"C:\win32hwp\1.hwp"
        hwp.Open(file_path)
        log(f"새 인스턴스에서 {file_path} 열기 성공")
        return hwp
    except Exception as e:
        log(f"새 인스턴스 생성 실패: {e}")

    return None


def get_text_at_cursor(hwp):
    """현재 커서 위치의 문단 텍스트 가져오기 (단순화)"""
    try:
        # 텍스트 가져오기 생략 - 위치 정보만 중요
        return "(텍스트 생략)"
    except Exception as e:
        return f"(오류)"


def get_all_paragraphs_info(hwp):
    """문서의 모든 문단 정보 수집"""
    paragraphs = []

    # 문서 처음으로 이동
    hwp.HAction.Run("MoveDocBegin")

    prev_key = None
    max_iter = 1000  # 무한루프 방지

    for _ in range(max_iter):
        pos = hwp.GetPos()
        list_id = pos[0]
        para_id = pos[1]

        key = (list_id, para_id)

        # 새로운 문단인지 확인
        if key != prev_key:
            # 컨트롤 정보
            ctrl_id = hwp.FindCtrl() or "본문"

            paragraphs.append({
                'list_id': list_id,
                'para_id': para_id,
                'ctrl_id': ctrl_id
            })

            prev_key = key

        # 다음 문단으로 이동
        old_pos = hwp.GetPos()

        result = hwp.HAction.Run("MoveNextParaBegin")
        if not result:
            break

        new_pos = hwp.GetPos()

        # 더 이상 이동하지 않으면 종료
        if old_pos == new_pos:
            break

    return paragraphs


def insert_pos_info_to_document(hwp):
    """각 문단 끝에 위치 정보를 파랑색으로 삽입"""

    # 문서 처음으로
    hwp.HAction.Run("MoveDocBegin")

    processed = set()

    while True:
        pos = hwp.GetPos()
        list_id = pos[0]
        para_id = pos[1]

        key = (list_id, para_id)

        if key not in processed:
            processed.add(key)

            # 컨트롤 정보
            ctrl_id = hwp.FindCtrl() or "본문"

            # 문단 끝으로 이동
            hwp.HAction.Run("MoveParaEnd")

            # 위치 정보 문자열 생성
            info_text = f"  [L:{list_id}, P:{para_id}, C:{ctrl_id}]"

            # 텍스트 삽입
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
            hwp.HParameterSet.HInsertText.Text = info_text
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

            # 방금 삽입한 텍스트 선택
            for _ in range(len(info_text)):
                hwp.HAction.Run("MoveSelLeft")

            # 파랑색 적용 (BGR: 0xFF0000)
            hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
            hwp.HParameterSet.HCharShape.TextColor = 0xFF0000
            hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)

            hwp.HAction.Run("Cancel")

            log(f"본문 삽입: L:{list_id}, P:{para_id}, C:{ctrl_id}")

        # 다음 문단으로 이동
        old_pos = hwp.GetPos()

        if not hwp.HAction.Run("MoveNextParaBegin"):
            break

        new_pos = hwp.GetPos()

        if old_pos == new_pos:
            break

    log(f"총 {len(processed)}개 문단에 위치 정보 삽입 완료")


def scan_all_controls(hwp):
    """문서의 모든 컨트롤 순회하며 정보 출력"""
    log("\n=== 컨트롤 목록 ===")

    ctrl = hwp.HeadCtrl
    count = 0
    tables = []

    while ctrl:
        ctrl_id = ctrl.CtrlID
        user_desc = ctrl.UserDesc if hasattr(ctrl, 'UserDesc') else ""

        log(f"[{count}] CtrlID: {ctrl_id}, 설명: {user_desc}")

        if ctrl_id == "tbl":
            tables.append(ctrl)
            log(f"     -> 표 컨트롤 저장")

        elif ctrl_id == "gso":
            log(f"     -> 그리기 개체")

        ctrl = ctrl.Next
        count += 1

    log(f"\n총 {count}개 컨트롤, 표 {len(tables)}개")
    return tables


def scan_table_cells(hwp, table_index=0):
    """테이블 내부 셀 순회하며 위치 정보 수집"""
    log(f"\n=== 테이블 {table_index} 셀 스캔 ===")

    cells = []

    # 셀 이동 상수
    MOVE_RIGHT_OF_CELL = 101
    MOVE_DOWN_OF_CELL = 103
    MOVE_START_OF_CELL = 104
    MOVE_TOP_OF_CELL = 106

    # 현재 위치 확인
    init_pos = hwp.GetPos()
    log(f"  초기 위치: L:{init_pos[0]}, P:{init_pos[1]}, pos:{init_pos[2]}")

    # 첫 셀로 이동
    hwp.MovePos(MOVE_START_OF_CELL, 0, 0)
    hwp.MovePos(MOVE_TOP_OF_CELL, 0, 0)

    first_pos = hwp.GetPos()
    log(f"  첫 셀 위치: L:{first_pos[0]}, P:{first_pos[1]}, pos:{first_pos[2]}")

    # list_id가 0이면 테이블 안이 아님
    if first_pos[0] == 0:
        log(f"  [경고] list_id=0 - 테이블 내부가 아닙니다!")
        return cells

    row = 0
    while True:
        row_start = hwp.GetPos()
        col = 0

        # 현재 행의 셀들 순회
        while True:
            pos = hwp.GetPos()
            list_id = pos[0]
            para_id = pos[1]

            cells.append({
                'row': row,
                'col': col,
                'list_id': list_id,
                'para_id': para_id
            })

            log(f"  셀({row},{col}): L:{list_id}, P:{para_id}")

            # 오른쪽 셀로
            before = hwp.GetPos()[0]
            result = hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
            after = hwp.GetPos()[0]

            if not result or after == before:
                break
            col += 1

        # 다음 행으로
        hwp.SetPos(row_start[0], row_start[1], row_start[2])
        before = hwp.GetPos()[0]
        result = hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
        after = hwp.GetPos()[0]

        if not result or after == before:
            break
        row += 1

    log(f"  총 {len(cells)}개 셀 발견")
    return cells


def insert_info_to_table_cells(hwp, table_index=0):
    """테이블 셀에 위치 정보 삽입"""
    MOVE_RIGHT_OF_CELL = 101
    MOVE_DOWN_OF_CELL = 103
    MOVE_START_OF_CELL = 104
    MOVE_TOP_OF_CELL = 106

    # 첫 셀로 이동
    hwp.MovePos(MOVE_START_OF_CELL, 0, 0)
    hwp.MovePos(MOVE_TOP_OF_CELL, 0, 0)

    first_pos = hwp.GetPos()
    if first_pos[0] == 0:
        log(f"  [건너뜀] list_id=0 - 테이블 내부가 아님")
        return 0

    processed = set()
    row = 0

    while True:
        row_start = hwp.GetPos()
        col = 0

        while True:
            pos = hwp.GetPos()
            list_id = pos[0]
            para_id = pos[1]
            key = (list_id, para_id)

            if key not in processed:
                processed.add(key)

                # 문단 끝으로
                hwp.HAction.Run("MoveParaEnd")

                # 위치 정보 삽입
                info_text = f" [L:{list_id},P:{para_id},R:{row},C:{col}]"

                hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
                hwp.HParameterSet.HInsertText.Text = info_text
                hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

                # 선택 후 파랑색 적용
                for _ in range(len(info_text)):
                    hwp.HAction.Run("MoveSelLeft")

                hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
                hwp.HParameterSet.HCharShape.TextColor = 0xFF0000
                hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
                hwp.HAction.Run("Cancel")

                log(f"  테이블{table_index} 셀({row},{col}): L:{list_id}, P:{para_id} 삽입")

            # 오른쪽 셀로
            before = hwp.GetPos()[0]
            result = hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
            after = hwp.GetPos()[0]

            if not result or after == before:
                break
            col += 1

        # 다음 행
        hwp.SetPos(row_start[0], row_start[1], row_start[2])
        before = hwp.GetPos()[0]
        result = hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
        after = hwp.GetPos()[0]

        if not result or after == before:
            break
        row += 1

    return len(processed)


def move_to_table_by_ctrl(hwp, table_ctrl):
    """테이블 컨트롤로 이동 후 내부 첫 셀로 진입"""
    try:
        # 방법 1: GetAnchorPos 사용
        anchor = table_ctrl.GetAnchorPos(0)
        list_id = anchor.Item("List")
        para_id = anchor.Item("Para")
        char_pos = anchor.Item("Pos")

        log(f"  테이블 앵커: L:{list_id}, P:{para_id}, pos:{char_pos}")

        # 테이블 앵커 위치로 이동
        hwp.SetPos(list_id, para_id, char_pos)

        # 테이블 선택
        hwp.HAction.Run("SelectCtrlFront")

        # 테이블 안으로 진입 (Enter 키)
        hwp.HAction.Run("TableCellBlock")
        hwp.HAction.Run("Cancel")

        # 현재 위치 확인
        pos = hwp.GetPos()
        log(f"  진입 후 위치: L:{pos[0]}, P:{pos[1]}, pos:{pos[2]}")

        # list_id가 0보다 크면 테이블 내부
        if pos[0] > 0:
            return True

        # 다른 방법 시도: MoveRight
        hwp.SetPos(list_id, para_id, char_pos)
        hwp.HAction.Run("MoveRight")

        pos = hwp.GetPos()
        log(f"  MoveRight 후: L:{pos[0]}, P:{pos[1]}, pos:{pos[2]}")

        return pos[0] > 0

    except Exception as e:
        log(f"  테이블 진입 실패: {e}")
        return False


def main():
    """메인 함수"""
    log(f"=== HWP 위치 정보 삽입 시작 ===")
    log(f"로그 파일: {LOG_FILE}\n")

    log("HWP 인스턴스 연결 중...")
    hwp = get_hwp_instance()

    if not hwp:
        log("ERROR: 한글 연결 실패")
        return

    log("HWP 연결 성공!\n")

    # 1. 컨트롤 목록 출력
    tables = scan_all_controls(hwp)

    # 2. 본문 문단 정보
    log("\n=== 본문 문단 정보 ===")
    paragraphs = get_all_paragraphs_info(hwp)
    for i, p in enumerate(paragraphs):
        log(f"[{i}] L:{p['list_id']}, P:{p['para_id']}, C:{p['ctrl_id']}")

    # 3. 테이블 셀 정보
    for i, tbl in enumerate(tables):
        log(f"\n--- 테이블 {i} 진입 ---")
        if move_to_table_by_ctrl(hwp, tbl):
            scan_table_cells(hwp, i)
        else:
            log(f"  [실패] 테이블 {i} 진입 불가")

    # 4. 요약 및 확인
    log(f"\n{'='*50}")
    log(f"본문: {len(paragraphs)}개 문단")
    log(f"테이블: {len(tables)}개")
    log(f"{'='*50}")

    response = input("위치 정보를 삽입하시겠습니까? (y/n): ").strip().lower()

    if response != 'y':
        log("취소되었습니다.")
        return

    # 5. 본문 위치 정보 삽입
    log("\n=== 본문 위치 정보 삽입 ===")
    insert_pos_info_to_document(hwp)

    # 6. 테이블 셀 위치 정보 삽입
    for i, tbl in enumerate(tables):
        log(f"\n=== 테이블 {i} 위치 정보 삽입 ===")
        if move_to_table_by_ctrl(hwp, tbl):
            count = insert_info_to_table_cells(hwp, i)
            log(f"  {count}개 셀 처리 완료")
        else:
            log(f"  [건너뜀] 테이블 {i}")

    log(f"\n완료! 로그 파일: {LOG_FILE}")


if __name__ == "__main__":
    main()
