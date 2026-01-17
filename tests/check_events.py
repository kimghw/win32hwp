"""한글 COM 객체의 이벤트 인터페이스 확인"""
import win32com.client as win32

# 타입 라이브러리 확인
hwp = win32.Dispatch("HWPFrame.HwpObject")

print("=== 한글 COM 객체 정보 ===")
print(f"객체: {hwp}")
print(f"타입: {type(hwp)}")

# gencache로 타입 정보 가져오기
try:
    import win32com.client.gencache as gencache
    hwp2 = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    print(f"\nEnsureDispatch 객체: {hwp2}")
    print(f"타입: {type(hwp2)}")
except Exception as e:
    print(f"\nEnsureDispatch 에러: {e}")

# 이벤트 클래스 찾기
try:
    from win32com.client import getevents
    events = getevents("HWPFrame.HwpObject")
    print(f"\n이벤트 클래스: {events}")
except Exception as e:
    print(f"\ngetevents 에러: {e}")

hwp.Quit()
