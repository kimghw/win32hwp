# gso 컨트롤 속성 전체 덤프
from cursor_utils import get_hwp_instance

hwp = get_hwp_instance()
if not hwp:
    print('한글 없음')
    exit()

# 모든 컨트롤 순회하며 gso 찾아서 선택
ctrl = hwp.HeadCtrl
while ctrl:
    if ctrl.CtrlID == 'gso':
        print('=== gso 컨트롤 발견 ===')
        print(f'UserDesc: {ctrl.UserDesc}')

        # 컨트롤 위치로 이동 후 선택
        hwp.SetPosBySet(ctrl.GetAnchorPos(0))
        hwp.FindCtrl()

        # HShapeObject 파라미터셋으로 조회
        print('\n[HShapeObject 속성]')
        shape = hwp.HParameterSet.HShapeObject
        hwp.HAction.GetDefault("ShapeObjDialog", shape.HSet)

        # 알려진 속성 조회
        prop_names = [
            "Width", "Height",
            "TreatAsChar", "TextWrap", "TextFlow",
            "HorzRelTo", "HorzAlign", "HorzOffset",
            "VertRelTo", "VertAlign", "VertOffset",
            # 여백 관련
            "OutsideMarginLeft", "OutsideMarginRight",
            "OutsideMarginTop", "OutsideMarginBottom",
            "InsideMarginLeft", "InsideMarginRight",
            "InsideMarginTop", "InsideMarginBottom",
            "MarginX", "MarginY",
            "WrapMarginLeft", "WrapMarginRight",
            "WrapMarginTop", "WrapMarginBottom",
        ]

        for name in prop_names:
            try:
                val = getattr(shape, name, None)
                if val is not None:
                    print(f'  {name}: {val}')
            except Exception as e:
                pass

        # HSet에서도 조회
        print('\n[HSet 속성]')
        hset = shape.HSet
        for name in prop_names:
            try:
                val = hset.Item(name)
                if val is not None:
                    print(f'  {name}: {val}')
            except:
                pass

        hwp.HAction.Run("Cancel")
        break
    ctrl = ctrl.Next
