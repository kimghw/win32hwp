# 테이블 셀 내부 컨트롤 순회 테스트
from cursor_utils import get_hwp_instance
from table_info import TableInfo, MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL

def get_ctrls_in_cell(hwp, target_list_id):
    """특정 셀(list_id)에 속한 컨트롤 찾기"""
    ctrls = []
    ctrl = hwp.HeadCtrl
    while ctrl:
        try:
            anchor = ctrl.GetAnchorPos(0)
            # anchor에서 List 항목으로 list_id 확인
            ctrl_list_id = anchor.Item("List")
            if ctrl_list_id == target_list_id:
                ctrls.append({
                    'ctrl': ctrl,
                    'id': ctrl.CtrlID,
                    'desc': ctrl.UserDesc if hasattr(ctrl, 'UserDesc') else None
                })
        except Exception as e:
            pass
        ctrl = ctrl.Next
    return ctrls

def test_cell_ctrl():
    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return

    table = TableInfo(hwp, debug=False)

    # 테이블 찾기
    tables = table.find_all_tables()
    if not tables:
        print("[오류] 문서에 테이블이 없습니다.")
        return

    print(f"=== 테이블 셀 내부 컨트롤 순회 테스트 ===\n")

    # 먼저 전체 문서의 모든 컨트롤 목록 출력
    print("=== 문서 내 모든 컨트롤 ===")
    ctrl = hwp.HeadCtrl
    ctrl_idx = 0
    while ctrl:
        try:
            anchor = ctrl.GetAnchorPos(0)
            anchor_list = anchor.Item("List")
            user_desc = ctrl.UserDesc if hasattr(ctrl, 'UserDesc') else "N/A"
            print(f"  [{ctrl_idx}] CtrlID={ctrl.CtrlID}, UserDesc={user_desc}, AnchorList={anchor_list}")
        except Exception as e:
            print(f"  [{ctrl_idx}] CtrlID={ctrl.CtrlID}, Error: {e}")
        ctrl_idx += 1
        ctrl = ctrl.Next
    print(f"  총 {ctrl_idx}개 컨트롤\n")

    for t in tables:
        print(f"테이블 #{t['num']}")

        # 테이블로 진입
        table.enter_table(t['num'])
        table.cells.clear()
        cells = table.collect_cells_bfs()

        print(f"셀 수: {len(cells)}개\n")

        # 각 셀 순회
        table.move_to_first_cell()

        visited = set()
        row = 0

        while True:
            row_start_pos = hwp.GetPos()
            col = 0

            while True:
                pos = hwp.GetPos()
                list_id = pos[0]

                if list_id not in visited:
                    visited.add(list_id)

                    print(f"셀({row},{col}) list_id={list_id}")

                    # 방법1: HeadCtrl 순회 + AnchorPos로 필터링
                    print("  [HeadCtrl + AnchorPos 필터링]")
                    cell_ctrls = get_ctrls_in_cell(hwp, list_id)
                    if cell_ctrls:
                        for c in cell_ctrls:
                            print(f"    CtrlID={c['id']}, UserDesc={c['desc']}")
                            # 속성 조회
                            try:
                                props = c['ctrl'].Properties
                                if props:
                                    w = props.Item("Width")
                                    h = props.Item("Height")
                                    treat = props.Item("TreatAsChar")
                                    print(f"      Width={w}, Height={h}, TreatAsChar={treat}")
                            except Exception as e:
                                print(f"      속성 조회 실패: {e}")
                    else:
                        print("    (컨트롤 없음)")

                    # 방법2: FindCtrl로 셀 내 컨트롤 찾기
                    print("  [FindCtrl 순회]")
                    hwp.SetPos(list_id, 0, 0)
                    hwp.HAction.Run("MoveParaBegin")

                    found_ctrls = []
                    max_iter = 100  # 무한루프 방지
                    iter_count = 0

                    while iter_count < max_iter:
                        cur_pos = hwp.GetPos()
                        if cur_pos[0] != list_id:
                            break

                        # 현재 위치에서 컨트롤 찾기
                        found = hwp.FindCtrl()
                        if found and found not in found_ctrls:
                            found_ctrls.append(found)
                            print(f"    FindCtrl: {found}")

                            # 선택된 컨트롤 속성 확인
                            try:
                                sel_ctrl = hwp.CurSelectedCtrl
                                if sel_ctrl:
                                    props = sel_ctrl.Properties
                                    if props and found == "gso":
                                        w = props.Item("Width")
                                        h = props.Item("Height")
                                        treat = props.Item("TreatAsChar")
                                        print(f"      Width={w}, Height={h}, TreatAsChar={treat}")
                                    hwp.HAction.Run("Cancel")
                            except Exception as e:
                                print(f"      속성 조회 실패: {e}")

                        # 다음 위치로 이동
                        before = hwp.GetPos()
                        hwp.HAction.Run("MoveRight")
                        after = hwp.GetPos()
                        if before == after:
                            break
                        iter_count += 1

                    if not found_ctrls:
                        print("    (FindCtrl 없음)")

                    # 방법3: SelectCtrlFront 시도
                    print("  [SelectCtrlFront 시도]")
                    hwp.SetPos(list_id, 0, 0)
                    hwp.HAction.Run("MoveParaBegin")

                    try:
                        result = hwp.HAction.Run("SelectCtrlFront")
                        if result:
                            sel_ctrl = hwp.CurSelectedCtrl
                            if sel_ctrl:
                                print(f"    선택된 컨트롤: CtrlID={sel_ctrl.CtrlID}")
                                try:
                                    user_desc = sel_ctrl.UserDesc
                                    print(f"    UserDesc={user_desc}")
                                except:
                                    pass
                            hwp.HAction.Run("Cancel")
                        else:
                            print("    (SelectCtrlFront 실패)")
                    except Exception as e:
                        print(f"    SelectCtrlFront 오류: {e}")

                    print()

                # 오른쪽 셀로
                before = list_id
                result = hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
                after = hwp.GetPos()[0]

                if not result or after == before:
                    break
                col += 1

            # 다음 행
            hwp.SetPos(row_start_pos[0], row_start_pos[1], row_start_pos[2])
            before = hwp.GetPos()[0]
            result = hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
            after = hwp.GetPos()[0]

            if not result or after == before:
                break
            row += 1

        hwp.HAction.Run("MoveParentList")
        hwp.HAction.Run("Cancel")


if __name__ == "__main__":
    test_cell_ctrl()
