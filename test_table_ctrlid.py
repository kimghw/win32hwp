"""
테이블의 ctrl_id 확인 테스트
"""
from cursor_utils import get_hwp_instance

hwp = get_hwp_instance()
if not hwp:
    print("한글이 실행 중이 아닙니다")
    exit()

hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')

print("=== 테이블 ctrl_id 확인 ===\n")

# 현재 위치
pos = hwp.GetPos()
print(f"현재 위치: list_id={pos[0]}, para_id={pos[1]}, char_pos={pos[2]}")

# 테이블 내부인지 확인
try:
    act = hwp.CreateAction("TableCellBorder")
    pset = act.GetDefault("TableCellBorder", hwp.HParameterSet.HTableCellBorder.HSet)
    in_table = act.Execute(pset)
except:
    in_table = False

if not in_table:
    print("\n테이블 내부에 커서를 위치시켜주세요")
    exit()

print("테이블 내부입니다\n")

# GetFieldList로 컨트롤 정보 가져오기
ctrl = hwp.HeadCtrl
if ctrl:
    print(f"컨트롤 정보:")
    print(f"  CtrlID: {ctrl.GetItem('CtrlID') if hasattr(ctrl, 'GetItem') else 'N/A'}")
    print(f"  UserDesc: {ctrl.UserDesc if hasattr(ctrl, 'UserDesc') else 'N/A'}")

# 다른 방법: GetCurFieldName
field_name = hwp.GetCurFieldName()
print(f"\n현재 필드 이름: '{field_name}'")

# InitScan으로 컨트롤 스캔
print("\n문서 컨트롤 스캔:")
hwp.InitScan()
state = 0
index = 0
while state >= 0 and index < 100:  # 최대 100개까지
    state, ctrl = hwp.GetFieldList()
    if state == 1:  # 컨트롤 발견
        if ctrl.name == 'tbl':  # 테이블
            print(f"  [{index}] 테이블 발견: {ctrl}")
            if hasattr(ctrl, 'CtrlID'):
                print(f"       CtrlID: {ctrl.CtrlID}")
    index += 1
hwp.ReleaseScan()

print("\n=== 완료 ===")
print("여러 테이블이 있다면 각각 이동하면서 테스트해보세요")
