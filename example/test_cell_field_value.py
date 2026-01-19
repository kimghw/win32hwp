# 셀 필드: 셀 값 vs 필드 값 관계 테스트
import time
import win32com.client as win32

hwp = win32.gencache.EnsureDispatch('HWPFrame.HwpObject')
hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')
hwp.XHwpWindows.Item(0).Visible = True

# 테이블 생성
print("[1] 테이블 생성...", flush=True)
hwp.HAction.GetDefault('TableCreate', hwp.HParameterSet.HTableCreation.HSet)
hwp.HParameterSet.HTableCreation.Rows = 1
hwp.HParameterSet.HTableCreation.Cols = 1
hwp.HParameterSet.HTableCreation.WidthType = 2
hwp.HAction.Execute('TableCreate', hwp.HParameterSet.HTableCreation.HSet)

# 셀 필드 설정
print("[2] 셀 필드 설정 (필드명: test_field)...", flush=True)
hwp.HAction.Run('TableCellBlock')
hwp.HAction.Run('Cancel')
hwp.HAction.Run('TableCellBlock')
hwp.SetCurFieldName('test_field', 1, '', '')
hwp.HAction.Run('Cancel')

# PutFieldText로 값 설정
hwp.PutFieldText('test_field', '초기값')

print("\n[3] 현재 상태:", flush=True)
field_val = hwp.GetFieldText('test_field').strip(chr(2))
print(f"  GetFieldText('test_field') = '{field_val}'", flush=True)

print("\n" + "=" * 50, flush=True)
print("셀에 직접 타이핑해서 값을 바꿔보세요!", flush=True)
print("3초마다 필드 값을 출력합니다.", flush=True)
print("종료: Ctrl+C", flush=True)
print("=" * 50 + "\n", flush=True)

last_val = field_val
try:
    while True:
        field_val = hwp.GetFieldText('test_field').strip(chr(2))
        if field_val != last_val:
            print(f">>> 필드 값 변경: '{last_val}' → '{field_val}'", flush=True)
            last_val = field_val
        time.sleep(1)
except KeyboardInterrupt:
    print("\n종료", flush=True)
