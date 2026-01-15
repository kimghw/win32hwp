"""
한글 커서 위치 모니터 (이벤트 기반)
- DocumentChange 이벤트 핸들러 사용
- API 이벤트와 사용자 이벤트 구분
"""
import win32com.client as win32
import pythoncom
from win32com.client import DispatchWithEvents
import time

# 전역 플래그
IS_API_MOVING = False
LAST_EVENT_TIME = 0
EVENT_THRESHOLD = 2  # 2초

hwp = None


class HwpEvents:
    """한글 이벤트 핸들러"""

    def OnDocumentChange(self, doc_id=0):
        global IS_API_MOVING, LAST_EVENT_TIME, hwp
        print(f"[이벤트] DocumentChange (doc_id={doc_id})")

        # API가 커서 움직이는 중이면 무시
        if IS_API_MOVING:
            print("  [무시] API 이동 중")
            return

        current_time = time.time()
        elapsed = current_time - LAST_EVENT_TIME

        # 2초 이내면 무시
        if elapsed < EVENT_THRESHOLD:
            print(f"  [무시] {elapsed:.1f}초 경과 (2초 미만)")
            return

        # 2초 이상 지났으면 사용자 이벤트로 처리
        LAST_EVENT_TIME = current_time
        print(f"\n[사용자 이벤트] {elapsed:.1f}초 경과")

        # 문단 위치 정보 추출
        result = get_pos_within_para()
        if result:
            print(f"  현재 위치: {result['current']}")
            print(f"  문단 시작: {result['para_start']}")
            print(f"  문단 끝  : {result['para_end']}")

    def OnQuit(self):
        print("[이벤트] Quit")

    def OnNewDocument(self, doc_id=0):
        print(f"[이벤트] NewDocument (doc_id={doc_id})")

    def OnDocumentAfterOpen(self, doc_id=0):
        print(f"[이벤트] DocumentAfterOpen (doc_id={doc_id})")

    def OnDocumentBeforeClose(self, doc_id=0):
        print(f"[이벤트] DocumentBeforeClose (doc_id={doc_id})")


def get_pos_within_para():
    """문단 시작/끝 위치 추출"""
    global IS_API_MOVING, hwp

    if not hwp:
        return None

    try:
        IS_API_MOVING = True  # API 이동 시작

        # 현재 위치 저장
        current_pos = hwp.GetPos()
        orig_list, orig_para, orig_pos = current_pos

        # 문단 시작으로 이동
        hwp.HAction.Run("MoveParaBegin")
        para_start = hwp.GetPos()

        # 문단 끝으로 이동
        hwp.HAction.Run("MoveParaEnd")
        para_end = hwp.GetPos()

        # 원래 위치로 복원
        hwp.SetPos(orig_list, orig_para, orig_pos)

        IS_API_MOVING = False  # API 이동 끝

        return {
            'current': current_pos,
            'para_start': para_start,
            'para_end': para_end
        }
    except Exception as e:
        IS_API_MOVING = False
        print(f"  [에러] {e}")
        return None


def main():
    global hwp

    print("=" * 50)
    print("  한글 커서 위치 모니터")
    print("  Ctrl+C로 종료")
    print("=" * 50)
    print()

    try:
        # 이벤트 핸들러와 함께 한글 연결 (새 인스턴스 생성)
        print("한글 인스턴스 생성 중...")
        hwp = DispatchWithEvents("HWPFrame.HwpObject", HwpEvents)
        hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')
        hwp.XHwpWindows.Item(0).Visible = True

        # test.hwp 파일 열기
        hwp.Open("d:/hwp_docs/test.hwp")

        print("연결 완료! test.hwp 열림")
        print("한글에서 커서를 움직여보세요...")
        print("-" * 50)

        # 메시지 루프 (이벤트 대기)
        while True:
            pythoncom.PumpWaitingMessages()
            time.sleep(0.01)

    except KeyboardInterrupt:
        print("\n\n모니터링 종료.")
    except Exception as e:
        print(f"에러: {e}")


if __name__ == "__main__":
    main()
