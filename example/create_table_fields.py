# 테이블 생성 후 셀에 필드 설정 테스트
import win32com.client as win32

hwp = win32.gencache.EnsureDispatch('HWPFrame.HwpObject')
hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')
hwp.XHwpWindows.Item(0).Visible = True

# 1. 3x3 테이블 생성
print('[1] 3x3 테이블 생성...')
hwp.HAction.GetDefault('TableCreate', hwp.HParameterSet.HTableCreation.HSet)
hwp.HParameterSet.HTableCreation.Rows = 3
hwp.HParameterSet.HTableCreation.Cols = 3
hwp.HParameterSet.HTableCreation.WidthType = 2  # 단에 맞춤
hwp.HParameterSet.HTableCreation.HeightType = 0  # 자동
hwp.HAction.Execute('TableCreate', hwp.HParameterSet.HTableCreation.HSet)
print('  테이블 생성 완료')

# 2. 첫 번째 셀로 이동 (테이블 안으로)
hwp.HAction.Run('TableCellBlock')
hwp.HAction.Run('Cancel')

# 3. 셀에 필드명 설정 (셀 자체가 필드가 됨)
# SetCurFieldName(fieldname, option, direction, memo)
# option: 0=일반, 1=셀, 2=누름틀
print('\n[2] 셀 필드 생성 (첫 번째 행)...')

# (0, 0) - 셀 블록 선택 후 필드명 설정
hwp.HAction.Run('TableCellBlock')
result = hwp.SetCurFieldName('cell_0_0', 1, '', '')  # option=1 (셀 필드)
hwp.HAction.Run('Cancel')
print(f'  (0,0) cell_0_0 -> {result}')

# (0, 1)
hwp.MovePos(101, 0, 0)  # MOVE_RIGHT_OF_CELL
hwp.HAction.Run('TableCellBlock')
result = hwp.SetCurFieldName('cell_0_1', 1, '', '')
hwp.HAction.Run('Cancel')
print(f'  (0,1) cell_0_1 -> {result}')

# (0, 2)
hwp.MovePos(101, 0, 0)
hwp.HAction.Run('TableCellBlock')
result = hwp.SetCurFieldName('cell_0_2', 1, '', '')
hwp.HAction.Run('Cancel')
print(f'  (0,2) cell_0_2 -> {result}')

# 두 번째 행으로 이동
print('\n[3] 셀 필드 생성 (두 번째 행)...')
hwp.MovePos(103, 0, 0)  # MOVE_DOWN_OF_CELL
hwp.MovePos(104, 0, 0)  # MOVE_START_OF_CELL (행 시작)

hwp.HAction.Run('TableCellBlock')
result = hwp.SetCurFieldName('cell_1_0', 1, '', '')
hwp.HAction.Run('Cancel')
print(f'  (1,0) cell_1_0 -> {result}')

hwp.MovePos(101, 0, 0)
hwp.HAction.Run('TableCellBlock')
result = hwp.SetCurFieldName('cell_1_1', 1, '', '')
hwp.HAction.Run('Cancel')
print(f'  (1,1) cell_1_1 -> {result}')

hwp.MovePos(101, 0, 0)
hwp.HAction.Run('TableCellBlock')
result = hwp.SetCurFieldName('cell_1_2', 1, '', '')
hwp.HAction.Run('Cancel')
print(f'  (1,2) cell_1_2 -> {result}')

# 4. 필드 확인
print('\n[4] 셀 필드 조회...')
field_list = hwp.GetFieldList(1, 1)  # FIELD_NUMBER, FIELD_CELL
if field_list:
    fields = field_list.split(chr(2))
    for f in fields:
        if f:
            print(f'  - {f}')
else:
    print('  필드 없음')

# 5. 필드에 값 넣기
print('\n[5] 필드 값 입력...')
hwp.PutFieldText('cell_0_0', '제목')
hwp.PutFieldText('cell_0_1', '작성자')
hwp.PutFieldText('cell_0_2', '날짜')
hwp.PutFieldText('cell_1_0', '내용1')
hwp.PutFieldText('cell_1_1', '내용2')
hwp.PutFieldText('cell_1_2', '내용3')

# 확인
print('\n[6] 필드 값 확인...')
for row in range(2):
    row_values = []
    for col in range(3):
        name = f'cell_{row}_{col}'
        text = hwp.GetFieldText(name)
        text = text.strip(chr(2)) if text else ''
        row_values.append(f'{name}={text}')
    print(f'  행{row}: {", ".join(row_values)}')

print('\n완료! 한글 문서를 확인하세요.')
