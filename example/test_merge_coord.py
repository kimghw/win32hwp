# 병합 셀 대표 좌표 테스트
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import win32com.client as win32
from table.table_info import TableInfo

def p(msg):
    print(msg, flush=True)

# 한글 실행
p("[1] 한글 실행...")
hwp = win32.gencache.EnsureDispatch('HWPFrame.HwpObject')
hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')
hwp.XHwpWindows.Item(0).Visible = True

# 4x4 테이블 생성
p("[2] 4x4 테이블 생성...")
hwp.HAction.GetDefault('TableCreate', hwp.HParameterSet.HTableCreation.HSet)
hwp.HParameterSet.HTableCreation.Rows = 4
hwp.HParameterSet.HTableCreation.Cols = 4
hwp.HParameterSet.HTableCreation.WidthType = 2
hwp.HAction.Execute('TableCreate', hwp.HParameterSet.HTableCreation.HSet)

# 테이블 진입
hwp.HAction.Run('TableCellBlock')
hwp.HAction.Run('Cancel')

# (0,0)~(0,1) 가로 병합: 첫 번째 행, 첫 두 셀
p("[3] 셀 병합...")
p("  (0,0)~(0,1) 가로 병합")
hwp.HAction.Run('TableCellBlock')  # (0,0) 선택
hwp.HAction.Run('TableCellBlockExtend')  # 확장 모드
hwp.MovePos(101, 0, 0)  # 오른쪽으로 확장
hwp.HAction.Run('TableMergeCell')  # 병합

# (2,0)~(3,0) 세로 병합
p("  (2,0)~(3,0) 세로 병합")
hwp.MovePos(103, 0, 0)  # 아래로
hwp.MovePos(103, 0, 0)  # 아래로 (2,0)
hwp.MovePos(104, 0, 0)  # 행 시작
hwp.HAction.Run('TableCellBlock')
hwp.HAction.Run('TableCellBlockExtend')
hwp.MovePos(103, 0, 0)  # 아래로 확장
hwp.HAction.Run('TableMergeCell')

# 좌표 매핑 테스트
p("\n[4] 좌표 매핑 테스트...")
ti = TableInfo(hwp, debug=False)
ti.enter_table(0)
ti.collect_cells_bfs()
coord_map = ti.build_coordinate_map()

ti.print_coordinate_map()

# 대표 좌표 확인
p("\n[5] 대표 좌표 확인...")
for list_id in sorted(ti._representative_coords.keys()):
    rep = ti._representative_coords[list_id]
    coords = ti._cell_coords[list_id]
    if len(coords) > 1:
        p(f"  list_id={list_id}: 대표좌표={rep}, 전체좌표={coords}")
    else:
        p(f"  list_id={list_id}: 좌표={rep}")

p("\n완료!")
