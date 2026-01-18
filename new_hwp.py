# 새 한글 창 생성
import win32com.client

hwp = win32com.client.Dispatch('HWPFrame.HwpObject')
hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')
hwp.XHwpWindows.Item(0).Visible = True
print('새 한글 창 생성 완료')
