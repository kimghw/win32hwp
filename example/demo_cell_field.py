# 셀 필드 데모: 테이블 생성 + 필드 설정 + 모니터링
import time
import sys
import win32com.client as win32

def p(msg):
    print(msg, flush=True)

# 한글 실행 (새 인스턴스)
p("[1] 한글 실행...")
hwp = win32.gencache.EnsureDispatch('HWPFrame.HwpObject')
hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')
hwp.XHwpWindows.Item(0).Visible = True

# 테이블 생성
p("[2] 3x3 테이블 생성...")
hwp.HAction.GetDefault('TableCreate', hwp.HParameterSet.HTableCreation.HSet)
hwp.HParameterSet.HTableCreation.Rows = 3
hwp.HParameterSet.HTableCreation.Cols = 3
hwp.HParameterSet.HTableCreation.WidthType = 2
hwp.HParameterSet.HTableCreation.HeightType = 0
hwp.HAction.Execute('TableCreate', hwp.HParameterSet.HTableCreation.HSet)

# 첫 셀로 이동
hwp.HAction.Run('TableCellBlock')
hwp.HAction.Run('Cancel')

# 셀 필드 설정
p("[3] 셀 필드 설정...")
fields = [
    (0, 0, 'title', '제목'),
    (0, 1, 'author', '작성자'),
    (0, 2, 'date', '날짜'),
    (1, 0, 'content1', '내용1'),
    (1, 1, 'content2', '내용2'),
    (1, 2, 'content3', '내용3'),
]

for i, (row, col, name, value) in enumerate(fields):
    if i > 0:
        if col == 0:
            hwp.MovePos(103, 0, 0)
            hwp.MovePos(104, 0, 0)
        else:
            hwp.MovePos(101, 0, 0)

    hwp.HAction.Run('TableCellBlock')
    hwp.SetCurFieldName(name, 1, '', '')
    hwp.HAction.Run('Cancel')
    hwp.PutFieldText(name, value)
    p(f"  ({row},{col}) {name} = '{value}'")

p("\n[4] 셀 필드 목록:")
field_list = hwp.GetFieldList(1, 1)
if field_list:
    for f in field_list.split(chr(2)):
        if f:
            p(f"  - {f}")

p("\n" + "=" * 50)
p("셀 필드 모니터링 시작!")
p("테이블 셀을 클릭해보세요.")
p("종료: Ctrl+C")
p("=" * 50 + "\n")

last_list_id = None

try:
    while True:
        pos = hwp.GetPos()
        list_id = pos[0]

        if list_id != last_list_id:
            last_list_id = list_id
            field_name = hwp.GetCurFieldName(1)

            if field_name:
                text = hwp.GetFieldText(field_name)
                text = text.strip(chr(2)) if text else ""
                p(f">>> 필드: {field_name} = '{text}'")
            elif list_id > 1:
                p(f"    (필드 없음) list_id={list_id}")

        time.sleep(0.1)

except KeyboardInterrupt:
    p("\n모니터링 종료")
