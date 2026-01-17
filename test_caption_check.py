"""캡션 존재 여부 확인 테스트 - ShapeCaption 속성 사용"""
from cursor_utils import get_hwp_instance


def test_shape_caption(hwp):
    """테이블 선택 후 ShapeCaption 속성 확인"""
    ctrl = hwp.HeadCtrl
    table_num = 0

    while ctrl:
        if ctrl.CtrlID == "tbl":
            print(f"\n=== 테이블 {table_num} ===")

            # 테이블 위치로 이동 후 선택
            hwp.SetPosBySet(ctrl.GetAnchorPos(0))
            hwp.HAction.Run("SelectCtrlFront")

            # HShapeObject의 ShapeCaption 조회 시도
            try:
                pset = hwp.HParameterSet.HShapeObject
                hwp.HAction.GetDefault("TablePropertyDialog", pset.HSet)

                # ShapeCaption 항목 확인
                caption_set = pset.HSet.Item("ShapeCaption")
                if caption_set:
                    print(f"  ShapeCaption 존재: {caption_set}")
                    try:
                        side = caption_set.Item("Side")
                        print(f"  Side: {side} (0=왼쪽, 1=오른쪽, 2=위, 3=아래)")
                    except:
                        print("  Side 항목 없음")
                    try:
                        width = caption_set.Item("Width")
                        print(f"  Width: {width}")
                    except:
                        print("  Width 항목 없음")
                    try:
                        gap = caption_set.Item("Gap")
                        print(f"  Gap: {gap}")
                    except:
                        print("  Gap 항목 없음")
                else:
                    print("  ShapeCaption: None (캡션 없음)")
            except Exception as e:
                print(f"  ShapeCaption 조회 실패: {e}")

            # 선택 해제
            hwp.HAction.Run("Cancel")
            table_num += 1

        ctrl = ctrl.Next


if __name__ == "__main__":
    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        exit(1)

    test_shape_caption(hwp)
