# 누름틀 필드 vs 셀 필드 비교
import win32com.client as win32

hwp = win32.gencache.EnsureDispatch('HWPFrame.HwpObject')
hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')
hwp.XHwpWindows.Item(0).Visible = True

# 2x2 테이블 생성
print("[1] 테이블 생성...")
hwp.HAction.GetDefault('TableCreate', hwp.HParameterSet.HTableCreation.HSet)
hwp.HParameterSet.HTableCreation.Rows = 2
hwp.HParameterSet.HTableCreation.Cols = 2
hwp.HParameterSet.HTableCreation.WidthType = 2
hwp.HAction.Execute('TableCreate', hwp.HParameterSet.HTableCreation.HSet)

# 첫 번째 셀 - 셀 필드
print("\n[2] (0,0) 셀 필드 방식...")
hwp.HAction.Run('TableCellBlock')
hwp.HAction.Run('Cancel')
hwp.HAction.Run('TableCellBlock')
hwp.SetCurFieldName('cell_field', 1, '', '')
hwp.HAction.Run('Cancel')
hwp.PutFieldText('cell_field', '셀필드값')
print("  필드명: cell_field")
print("  필드값: " + hwp.GetFieldText('cell_field').strip(chr(2)))

# 두 번째 셀 - 누름틀 필드
print("\n[3] (0,1) 누름틀 방식...")
hwp.MovePos(101, 0, 0)  # 오른쪽 셀로

# 먼저 일반 텍스트 입력
hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
hwp.HParameterSet.HInsertText.Text = "이름: "
hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

# 누름틀 생성
hwp.CreateField("입력하세요", "이름을 입력", "name_field")
print("  필드명: name_field")
print("  필드값: '" + (hwp.GetFieldText('name_field').strip(chr(2)) or "(빈값)") + "'")

# 누름틀에 값 넣기
hwp.PutFieldText('name_field', '홍길동')
print("  값 입력 후: " + hwp.GetFieldText('name_field').strip(chr(2)))

# 비교 출력
print("\n" + "=" * 50)
print("[비교]")
print("=" * 50)
print("셀 필드 (cell_field):")
print("  - 셀 전체가 필드")
print("  - 셀 값 = 필드 값")

print("\n누름틀 필드 (name_field):")
print("  - 셀 안에 '이름: 『홍길동』' 형태")
print("  - 셀 값 ≠ 필드 값 (필드는 『』 안만)")

print("\n한글 문서를 확인하세요!")
