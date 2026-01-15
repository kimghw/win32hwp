# -*- coding: utf-8 -*-
"""한글 커서 위치 실시간 추적"""

import sys
print("스크립트 시작", flush=True)

try:
    import pythoncom
    import win32com.client as win32
    import time
    print("모듈 로드 완료", flush=True)
except Exception as e:
    print(f"모듈 로드 실패: {e}", flush=True)
    sys.exit(1)

# ROT에서 HWP 연결
try:
    context = pythoncom.CreateBindCtx(0)
    rot = pythoncom.GetRunningObjectTable()
    print("ROT 연결 완료", flush=True)
except Exception as e:
    print(f"ROT 연결 실패: {e}", flush=True)
    sys.exit(1)

hwp = None
for moniker in rot:
    name = moniker.GetDisplayName(context, None)
    print(f"  발견: {name}", flush=True)
    if 'HwpObject' in name:
        try:
            obj = rot.GetObject(moniker)
            hwp = win32.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
            print(f"  -> HWP 연결 성공!", flush=True)
            break
        except Exception as e:
            print(f"  -> 연결 실패: {e}", flush=True)

if hwp is None:
    print("실행 중인 한글이 없습니다.", flush=True)
    sys.exit(1)

print("=" * 50, flush=True)
print("  한글 커서 위치 실시간 추적", flush=True)
print("  Ctrl+C로 종료", flush=True)
print("=" * 50, flush=True)
print(flush=True)

last_pos = None
while True:
    try:
        pos = hwp.GetPos()
        current = (pos[0], pos[1], pos[2])

        if current != last_pos:
            parent = hwp.ParentCtrl
            location = "표 안" if parent and parent.CtrlID == "tbl" else "본문"
            key = hwp.KeyIndicator()

            print(f"list={pos[0]:3d} | para={pos[1]:3d} | pos={pos[2]:4d} | "
                  f"page={key[2]} line={key[4]} col={key[5]} | {location}", flush=True)
            last_pos = current

        time.sleep(0.05)

    except KeyboardInterrupt:
        print("\n종료", flush=True)
        break
    except Exception as e:
        print(f"에러: {e}", flush=True)
        time.sleep(1)
