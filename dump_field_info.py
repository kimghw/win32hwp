# 필드 정보 조회 테스트
from cursor_utils import get_hwp_instance

hwp = get_hwp_instance()
if not hwp:
    print('한글 없음')
    exit()

print("=== 필드 정보 조회 ===\n")

# 방법1: GetFieldList로 필드 목록 조회
print("[GetFieldList]")
try:
    # 옵션: 0=일반, 1=셀, 2=누름틀
    for opt in [0, 1, 2]:
        field_list = hwp.GetFieldList(opt, 0)
        if field_list:
            print(f"  옵션 {opt}: {field_list}")
except Exception as e:
    print(f"  오류: {e}")

# 방법2: 컨트롤 순회하며 필드 찾기
print("\n[컨트롤 순회]")
ctrl = hwp.HeadCtrl
ctrl_idx = 0
while ctrl:
    ctrl_id = ctrl.CtrlID
    # 필드 관련 컨트롤 ID
    if ctrl_id in ("field", "fld", "%", "%%", "clickhere", "fn", "en"):
        print(f"  [{ctrl_idx}] CtrlID={ctrl_id}")
        try:
            props = ctrl.Properties
            if props:
                # 알려진 필드 속성
                for name in ["Name", "FieldName", "Direction", "Memo", "Value", "Text"]:
                    try:
                        val = props.Item(name)
                        if val:
                            print(f"    {name}: {val}")
                    except:
                        pass
        except:
            pass
    ctrl_idx += 1
    ctrl = ctrl.Next

# 방법3: 누름틀(ClickHere) 필드 찾기
print("\n[누름틀 필드]")
try:
    # MoveToField로 필드 이동
    field_names = hwp.GetFieldList(2, 0)  # 누름틀
    if field_names:
        for name in field_names.split('\x02'):
            if name:
                print(f"  필드: {name}")
                try:
                    val = hwp.GetFieldText(name)
                    print(f"    값: {val}")
                except:
                    pass
except Exception as e:
    print(f"  오류: {e}")

# 방법4: 셀 필드 찾기
print("\n[셀 필드]")
try:
    field_names = hwp.GetFieldList(1, 0)  # 셀
    if field_names:
        for name in field_names.split('\x02'):
            if name:
                print(f"  필드: {name}")
                try:
                    val = hwp.GetFieldText(name)
                    print(f"    값: {val}")
                except:
                    pass
except Exception as e:
    print(f"  오류: {e}")

# 방법5: 일반 필드
print("\n[일반 필드]")
try:
    field_names = hwp.GetFieldList(0, 0)  # 일반
    if field_names:
        for name in field_names.split('\x02'):
            if name:
                print(f"  필드: {name}")
                try:
                    val = hwp.GetFieldText(name)
                    print(f"    값: {val}")
                except:
                    pass
except Exception as e:
    print(f"  오류: {e}")
