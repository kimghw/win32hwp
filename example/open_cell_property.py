# 셀 속성 대화상자 열기
import win32com.client as win32
import pythoncom

pythoncom.CoInitialize()

# ROT에서 한글 찾기
context = pythoncom.CreateBindCtx(0)
rot = pythoncom.GetRunningObjectTable()
hwp = None

for moniker in rot:
    name = moniker.GetDisplayName(context, None)
    if 'HwpObject' in name:
        obj = rot.GetObject(moniker)
        hwp = win32.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
        break

if hwp:
    # 셀 블록 선택 후 속성 대화상자 열기
    hwp.HAction.Run('TableCellBlock')
    hwp.HAction.Run('TableCellPropertyDialog')
    print('셀 속성 대화상자가 열렸습니다.')
    print('"셀" 탭에서 "필드 이름" 항목을 확인하세요!')
else:
    print('한글을 찾을 수 없습니다.')
