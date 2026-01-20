"""현재 열려있는 HWP 테이블의 각 셀에 list_id 입력"""
import pythoncom
import win32com.client as win32


def get_hwp_from_rot():
    """ROT에서 실행 중인 한글 인스턴스 찾기"""
    context = pythoncom.CreateBindCtx(0)
    rot = pythoncom.GetRunningObjectTable()

    for moniker in rot:
        name = moniker.GetDisplayName(context, None)
        if "HwpObject" in name:
            obj = rot.GetObject(moniker)
            hwp = win32.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
            return hwp
    return None


def insert_list_ids_to_table(hwp):
    """테이블 순회하면서 각 셀에 list_id 입력 (BFS)"""
    from collections import deque

    # 테이블 안에 있는지 확인
    ctrl = hwp.ParentCtrl
    if not ctrl or ctrl.CtrlID != "tbl":
        print("커서가 테이블 안에 있지 않습니다.")
        return

    # 테이블 첫 셀로 이동
    hwp.HAction.Run("TableColBegin")
    hwp.HAction.Run("TableRowBegin")

    start_list_id, _, _ = hwp.GetPos()

    visited = set()
    queue = deque([start_list_id])

    while queue:
        current = queue.popleft()

        if current in visited:
            continue

        # 해당 셀로 이동
        hwp.SetPos(current, 0, 0)

        # 테이블 안인지 확인
        parent = hwp.ParentCtrl
        if not parent or parent.CtrlID != "tbl":
            continue

        visited.add(current)

        # 셀 내용 삭제 후 list_id 입력
        hwp.HAction.Run("SelectAll")
        hwp.HAction.Run("Delete")
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = str(current)
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

        # 4방향 탐색 (상, 하, 좌, 우)
        directions = ["MoveUp", "MoveDown", "MoveLeft", "MoveRight"]
        for direction in directions:
            hwp.SetPos(current, 0, 0)
            hwp.HAction.Run(direction)

            new_list_id, _, _ = hwp.GetPos()
            parent = hwp.ParentCtrl

            if parent and parent.CtrlID == "tbl" and new_list_id not in visited:
                queue.append(new_list_id)

    print(f"완료: {len(visited)}개 셀에 list_id 입력")


if __name__ == "__main__":
    hwp = get_hwp_from_rot()
    if not hwp:
        print("한글이 실행 중이 아닙니다.")
    else:
        print("한글 인스턴스 연결 성공")
        insert_list_ids_to_table(hwp)
