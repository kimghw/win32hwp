# only_table.hwp 파일에서 테이블 내부 정보 추출
# list_id, para_id, ctrl 정보 출력

import sys
from datetime import datetime
from cursor_utils import get_hwp_instance
from table_info import TableInfo, MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL, MOVE_START_OF_CELL, MOVE_TOP_OF_CELL

# 로그 파일 설정
LOG_FILE = "logs/table_info_extract.log"

# 테스트 파일 경로
TEST_FILE = r"C:\win32hwp\1.hwp"


def open_hwp_file(file_path):
    """HWP 파일 열기 (새 인스턴스 생성)"""
    import win32com.client as win32

    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
    hwp.Open(file_path)
    hwp.XHwpWindows.Item(0).Visible = True

    print(f"파일 열기 완료: {file_path}")
    return hwp


def extract_field_info():
    """문서 내 필드 정보 추출"""
    hwp = get_hwp()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return

    print(f"\n{'='*70}")
    print("필드 정보 추출")
    print(f"{'='*70}")

    field_count = 0

    # 필드 유형별 조회
    field_types = [
        (0, "일반 필드"),
        (1, "셀 필드"),
        (2, "누름틀 필드"),
    ]

    for opt, type_name in field_types:
        print(f"\n[{type_name}]")
        try:
            field_list = hwp.GetFieldList(opt, 0)
            if field_list:
                fields = field_list.split('\x02')
                for name in fields:
                    if name:
                        field_count += 1
                        print(f"  필드명: {name}")
                        try:
                            val = hwp.GetFieldText(name)
                            if val:
                                print(f"    값: {val}")
                        except:
                            pass
            else:
                print("  (없음)")
        except Exception as e:
            print(f"  조회 오류: {e}")

    # 컨트롤 순회로 추가 필드 찾기
    print(f"\n[컨트롤 기반 필드 검색]")
    ctrl = hwp.HeadCtrl
    ctrl_fields = []
    while ctrl:
        ctrl_id = ctrl.CtrlID
        # 필드 관련 컨트롤
        if ctrl_id in ("%", "%%", "field", "clickhere"):
            try:
                anchor = ctrl.GetAnchorPos(0)
                list_id = anchor.Item("List")
                para_id = anchor.Item("Para")
                user_desc = ctrl.UserDesc if hasattr(ctrl, 'UserDesc') else ctrl_id
                ctrl_fields.append({
                    'id': ctrl_id,
                    'desc': user_desc,
                    'list_id': list_id,
                    'para_id': para_id
                })
            except:
                pass
        ctrl = ctrl.Next

    if ctrl_fields:
        for f in ctrl_fields:
            print(f"  CtrlID={f['id']}, Desc={f['desc']}, 위치=({f['list_id']},{f['para_id']})")
            field_count += 1
    else:
        print("  (없음)")

    print(f"\n총 {field_count}개 필드 발견")

def extract_table_cell_info():
    """테이블 셀별 list_id, para_id, ctrl 정보 추출"""
    hwp = get_hwp()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return

    table = TableInfo(hwp, debug=False)

    # 테이블 찾기
    tables = table.find_all_tables()
    if not tables:
        print("[오류] 문서에 테이블이 없습니다.")
        return

    print(f"=== 문서 내 테이블 수: {len(tables)}개 ===\n")

    for t in tables:
        print(f"\n{'='*60}")
        print(f"테이블 #{t['num']}")
        print(f"{'='*60}")

        # 테이블로 진입
        table.enter_table(t['num'])

        # 테이블 크기 조회
        table.cells.clear()  # 이전 데이터 클리어
        cells = table.collect_cells_bfs()
        size = table.get_table_size()

        print(f"크기: {size['rows']}행 x {size['cols']}열")
        print(f"셀 수: {len(cells)}개")
        print()

        # 첫 번째 셀로 이동
        table.move_to_first_cell()

        # 셀별 상세 정보 출력
        print(f"{'셀':<10} {'list_id':<10} {'para_id':<10} {'char_pos':<10} {'ctrl_id':<10}")
        print("-" * 60)

        row = 0
        visited = set()

        while True:
            row_start_pos = hwp.GetPos()
            col = 0

            # 현재 행 순회
            while True:
                pos = hwp.GetPos()
                list_id = pos[0]
                para_id = pos[1]
                char_pos = pos[2]

                if list_id not in visited:
                    visited.add(list_id)

                    # ctrl_id 확인 (현재 위치의 컨트롤)
                    ctrl_id = ""
                    try:
                        # 현재 위치의 컨트롤 정보 조회
                        ctrl = hwp.CurSelectedCtrl
                        if ctrl:
                            ctrl_id = ctrl.CtrlID
                    except:
                        pass

                    print(f"({row},{col}){' '*4} {list_id:<10} {para_id:<10} {char_pos:<10} {ctrl_id:<10}")

                # 오른쪽으로 이동
                before = list_id
                result = hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
                after = hwp.GetPos()[0]

                if not result or after == before:
                    break
                col += 1

            # 다음 행으로 이동
            hwp.SetPos(row_start_pos[0], row_start_pos[1], row_start_pos[2])
            before = hwp.GetPos()[0]
            result = hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
            after = hwp.GetPos()[0]

            if not result or after == before:
                break
            row += 1

        # 좌표 매핑 출력
        print("\n[좌표 → list_id 매핑]")
        coord_map = table.build_coordinate_map()

        if coord_map:
            max_row = max(r for r, c in coord_map.keys())
            max_col = max(c for r, c in coord_map.keys())

            for r in range(max_row + 1):
                row_str = f"행 {r}: "
                cells_str = []
                for c in range(max_col + 1):
                    lid = coord_map.get((r, c), '-')
                    cells_str.append(f"({r},{c})→{lid}")
                print(row_str + "  ".join(cells_str))

        # 캡션 확인
        caption = table.get_table_caption(t['num'])
        if caption:
            print(f"\n캡션: {caption}")

        # 테이블 밖으로 나가기
        hwp.HAction.Run("MoveParentList")
        hwp.HAction.Run("Cancel")


def extract_detailed_cell_info():
    """각 셀의 상세 정보 (para 포함) 추출"""
    hwp = get_hwp()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return

    table = TableInfo(hwp, debug=False)

    # 테이블 찾기
    tables = table.find_all_tables()
    if not tables:
        print("[오류] 문서에 테이블이 없습니다.")
        return

    for t in tables:
        print(f"\n{'='*70}")
        print(f"테이블 #{t['num']} - 상세 정보")
        print(f"{'='*70}")

        # 테이블로 진입
        table.enter_table(t['num'])
        table.cells.clear()
        cells = table.collect_cells_bfs()

        print(f"\n{'list_id':<10} {'left':<8} {'right':<8} {'up':<8} {'down':<8} {'width':<10} {'height':<10}")
        print("-" * 70)

        for list_id in sorted(cells.keys()):
            cell = cells[list_id]
            print(f"{cell.list_id:<10} {cell.left:<8} {cell.right:<8} {cell.up:<8} {cell.down:<8} {cell.width:<10} {cell.height:<10}")

        # 각 셀의 문단 정보 출력
        print(f"\n[셀별 문단 (para) 정보]")
        print("-" * 70)

        table.move_to_first_cell()

        visited = set()
        row = 0

        while True:
            row_start_pos = hwp.GetPos()
            col = 0

            while True:
                pos = hwp.GetPos()
                list_id = pos[0]

                if list_id not in visited:
                    visited.add(list_id)

                    # 셀 내 모든 문단 탐색
                    paras = []
                    hwp.SetPos(list_id, 0, 0)  # 셀의 첫 문단으로

                    while True:
                        p = hwp.GetPos()
                        if p[0] != list_id:
                            break
                        paras.append(p[1])

                        # 다음 문단으로 이동
                        before_para = p[1]
                        hwp.HAction.Run("MoveParaEnd")
                        hwp.HAction.Run("MoveRight")
                        after = hwp.GetPos()

                        if after[0] != list_id or after[1] == before_para:
                            break

                    # 셀로 돌아가기
                    hwp.SetPos(list_id, 0, 0)

                    print(f"셀 ({row},{col}) list_id={list_id}: para_ids={paras}")

                # 오른쪽 셀로
                before = list_id
                result = hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
                after = hwp.GetPos()[0]

                if not result or after == before:
                    break
                col += 1

            # 다음 행
            hwp.SetPos(row_start_pos[0], row_start_pos[1], row_start_pos[2])
            before = hwp.GetPos()[0]
            result = hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
            after = hwp.GetPos()[0]

            if not result or after == before:
                break
            row += 1

        hwp.HAction.Run("MoveParentList")
        hwp.HAction.Run("Cancel")


def extract_table_properties():
    """테이블 속성 및 필드 정보 추출"""
    hwp = get_hwp()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return

    print(f"\n{'='*70}")
    print("테이블 속성 및 필드 정보")
    print(f"{'='*70}")

    # 테이블 컨트롤 순회
    ctrl = hwp.HeadCtrl
    table_num = 0

    while ctrl:
        if ctrl.CtrlID == "tbl":
            print(f"\n{'─'*70}")
            print(f"테이블 #{table_num}")
            print(f"{'─'*70}")

            # 테이블 속성 (Properties)
            try:
                props = ctrl.Properties
                if props:
                    print("\n[테이블 속성 (Properties)]")

                    # 주요 속성 목록
                    prop_names = [
                        "TreatAsChar",      # 글자처럼 취급
                        "PageBreak",        # 페이지 경계 처리
                        "RepeatHeader",     # 제목 행 반복
                        "Width",            # 너비
                        "Height",           # 높이
                        "WidthRelTo",       # 너비 기준
                        "HeightRelTo",      # 높이 기준
                        "VertRelTo",        # 세로 위치 기준
                        "VertAlign",        # 세로 정렬
                        "VertOffset",       # 세로 오프셋
                        "HorzRelTo",        # 가로 위치 기준
                        "HorzAlign",        # 가로 정렬
                        "HorzOffset",       # 가로 오프셋
                        "TextWrap",         # 텍스트 배치
                        "TextFlow",         # 배치 방향
                        "CellSpacing",      # 셀 간격
                        "CellMarginLeft",   # 셀 여백 (왼쪽)
                        "CellMarginRight",  # 셀 여백 (오른쪽)
                        "CellMarginTop",    # 셀 여백 (위)
                        "CellMarginBottom", # 셀 여백 (아래)
                        "ProtectSize",      # 크기 보호
                        "FlowWithText",     # 본문 영역 제한
                        "AllowOverlap",     # 겹침 허용
                    ]

                    for name in prop_names:
                        try:
                            value = props.Item(name)
                            if value is not None:
                                print(f"  {name}: {value}")
                        except:
                            pass

                    # 추가 속성 탐색
                    print("\n  [전체 속성 덤프]")
                    try:
                        # ItemCount가 있으면 순회
                        count = props.ItemCount
                        for i in range(count):
                            try:
                                item_id = props.ItemID(i)
                                item_val = props.Item(item_id)
                                print(f"    {item_id}: {item_val}")
                            except:
                                pass
                    except:
                        pass

            except Exception as e:
                print(f"  속성 조회 실패: {e}")

            # 테이블 필드 정보
            print("\n[테이블 필드 정보]")
            try:
                # 테이블 위치로 이동하여 필드 확인
                hwp.SetPosBySet(ctrl.GetAnchorPos(0))
                hwp.HAction.Run("SelectCtrlFront")
                hwp.HAction.Run("ShapeObjTableSelCell")

                # 테이블 셀 순회하며 필드 확인
                table = TableInfo(hwp, debug=False)
                table.cells.clear()
                cells = table.collect_cells_bfs()

                field_count = 0
                table.move_to_first_cell()

                visited = set()
                row = 0

                while True:
                    row_start_pos = hwp.GetPos()
                    col = 0

                    while True:
                        pos = hwp.GetPos()
                        list_id = pos[0]

                        if list_id not in visited:
                            visited.add(list_id)

                            # 셀 내 필드 확인
                            hwp.SetPos(list_id, 0, 0)

                            # 현재 셀의 컨트롤 순회
                            try:
                                para_ctrl = hwp.ParaCtrl
                                while para_ctrl:
                                    ctrl_id = para_ctrl.CtrlID
                                    if ctrl_id:
                                        field_count += 1
                                        print(f"  셀({row},{col}) - CtrlID: {ctrl_id}")

                                        # 필드 속성
                                        try:
                                            field_props = para_ctrl.Properties
                                            if field_props:
                                                print(f"    속성:")
                                                try:
                                                    fc = field_props.ItemCount
                                                    for fi in range(min(fc, 10)):  # 최대 10개
                                                        fid = field_props.ItemID(fi)
                                                        fval = field_props.Item(fid)
                                                        print(f"      {fid}: {fval}")
                                                except:
                                                    pass
                                        except:
                                            pass

                                    para_ctrl = para_ctrl.Next
                            except:
                                pass

                        # 오른쪽 셀로
                        before = list_id
                        result = hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
                        after = hwp.GetPos()[0]

                        if not result or after == before:
                            break
                        col += 1

                    # 다음 행
                    hwp.SetPos(row_start_pos[0], row_start_pos[1], row_start_pos[2])
                    before = hwp.GetPos()[0]
                    result = hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
                    after = hwp.GetPos()[0]

                    if not result or after == before:
                        break
                    row += 1

                if field_count == 0:
                    print("  (필드 없음)")

                hwp.HAction.Run("MoveParentList")
                hwp.HAction.Run("Cancel")

            except Exception as e:
                print(f"  필드 조회 실패: {e}")

            # 셀 속성 정보를 빨간색 텍스트로 셀 마지막에 삽입
            print("\n[셀 속성 정보 → 빨간색 텍스트로 삽입]")
            try:
                hwp.SetPosBySet(ctrl.GetAnchorPos(0))
                hwp.HAction.Run("SelectCtrlFront")
                hwp.HAction.Run("ShapeObjTableSelCell")

                table.move_to_first_cell()

                visited_cells = set()
                row = 0
                insert_count = 0

                while True:
                    row_start_pos = hwp.GetPos()
                    col = 0

                    while True:
                        pos = hwp.GetPos()
                        list_id = pos[0]

                        if list_id not in visited_cells:
                            visited_cells.add(list_id)

                            # CellShape로 셀 속성 조회
                            try:
                                cell_shape = hwp.CellShape
                                if cell_shape:
                                    cell = cell_shape.Item("Cell")
                                    if cell:
                                        width = cell.Item("Width")
                                        height = cell.Item("Height")
                                        header = cell.Item("Header")
                                        protected = cell.Item("Protected")
                                        has_margin = cell.Item("HasMargin")
                                        editable = cell.Item("Editable")
                                        dirty = cell.Item("Dirty")

                                        # 셀 여백 (상하좌우)
                                        margin_top = cell.Item("MarginTop") or 0
                                        margin_bottom = cell.Item("MarginBottom") or 0
                                        margin_left = cell.Item("MarginLeft") or 0
                                        margin_right = cell.Item("MarginRight") or 0

                                        # 셀 필드 정보 조회
                                        cell_field = " 필드없음"
                                        try:
                                            field_list = hwp.GetFieldList(1, 0)  # 옵션 1 = 셀 필드
                                            if field_list:
                                                fields = field_list.split('\x02')
                                                # 현재 셀 위치와 일치하는 필드 찾기
                                                for fname in fields:
                                                    if fname:
                                                        # 필드로 이동해서 위치 확인
                                                        try:
                                                            hwp.MoveToField(fname, True, True, True)
                                                            fpos = hwp.GetPos()
                                                            if fpos[0] == list_id:
                                                                fval = hwp.GetFieldText(fname) or ""
                                                                cell_field = f" 필드:{fname}"
                                                                if fval:
                                                                    cell_field += f"={fval}"
                                                                break
                                                        except:
                                                            pass
                                                # 원래 셀로 복귀
                                                hwp.SetPos(list_id, 0, 0)
                                        except:
                                            pass

                                        # 셀 속성 문자열 생성
                                        props_str = f" {{({row},{col}) list:{list_id} W:{width} H:{height}"
                                        props_str += f" 여백(상:{margin_top} 하:{margin_bottom} 좌:{margin_left} 우:{margin_right})"

                                        flags = []
                                        if header:
                                            flags.append("제목")
                                        if protected:
                                            flags.append("보호")
                                        if has_margin:
                                            flags.append("자체여백")
                                        if editable:
                                            flags.append("편집가능")
                                        if dirty:
                                            flags.append("수정됨")

                                        if flags:
                                            props_str += f" [{','.join(flags)}]"
                                        if cell_field:
                                            props_str += cell_field
                                        props_str += "}"

                                        # 셀 내 마지막 문단 끝으로 이동
                                        hwp.SetPos(list_id, 0, 0)

                                        # 셀 내에서 마지막 문단 찾기
                                        while True:
                                            hwp.HAction.Run("MoveParaEnd")
                                            before_pos = hwp.GetPos()
                                            hwp.HAction.Run("MoveRight")
                                            after_pos = hwp.GetPos()

                                            # 셀 밖으로 나갔거나 더 이상 이동 안되면 중단
                                            if after_pos[0] != list_id or after_pos == before_pos:
                                                hwp.SetPos(before_pos[0], before_pos[1], before_pos[2])
                                                hwp.HAction.Run("MoveParaEnd")
                                                break

                                        # 새 줄 추가 후 속성 텍스트 삽입
                                        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
                                        hwp.HParameterSet.HInsertText.Text = "\r\n" + props_str
                                        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

                                        # 삽입한 텍스트 선택하여 빨간색 적용
                                        for _ in range(len(props_str)):
                                            hwp.HAction.Run("MoveSelLeft")

                                        # 빨간색(RGB: 255, 0, 0) 적용 - BGR 변환
                                        hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
                                        hwp.HParameterSet.HCharShape.TextColor = 0x0000FF  # BGR: 빨간색
                                        hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)

                                        hwp.HAction.Run("Cancel")

                                        insert_count += 1
                                        print(f"  셀({row},{col}) list_id={list_id}:{props_str}")

                            except Exception as e:
                                print(f"  셀({row},{col}) 속성 삽입 실패: {e}")

                        # 오른쪽 셀로
                        before = list_id
                        result = hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
                        after = hwp.GetPos()[0]

                        if not result or after == before:
                            break
                        col += 1

                    # 다음 행
                    hwp.SetPos(row_start_pos[0], row_start_pos[1], row_start_pos[2])
                    before = hwp.GetPos()[0]
                    result = hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
                    after = hwp.GetPos()[0]

                    if not result or after == before:
                        break
                    row += 1

                hwp.HAction.Run("MoveParentList")
                hwp.HAction.Run("Cancel")

                print(f"  → 총 {insert_count}개 셀에 속성 정보 삽입 완료")

            except Exception as e:
                print(f"  셀 속성 삽입 실패: {e}")

            table_num += 1

        ctrl = ctrl.Next

    if table_num == 0:
        print("문서에 테이블이 없습니다.")


def insert_para_info_to_document():
    """각 셀의 문단 앞에 [list_id, para_id] 정보 삽입"""
    hwp = get_hwp()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return

    table = TableInfo(hwp, debug=False)

    # 테이블 찾기
    tables = table.find_all_tables()
    if not tables:
        print("[오류] 문서에 테이블이 없습니다.")
        return

    print(f"\n=== 문단 앞에 [list_id, para_id] 정보 삽입 ===\n")

    for t in tables:
        print(f"테이블 #{t['num']} 처리 중...")

        # 테이블로 진입
        table.enter_table(t['num'])
        table.cells.clear()
        cells = table.collect_cells_bfs()

        # 각 셀 순회하며 문단 정보 삽입
        table.move_to_first_cell()

        visited = set()
        row = 0
        insert_count = 0

        while True:
            row_start_pos = hwp.GetPos()
            col = 0

            while True:
                pos = hwp.GetPos()
                list_id = pos[0]

                if list_id not in visited:
                    visited.add(list_id)

                    # 셀의 첫 문단으로 이동
                    hwp.SetPos(list_id, 0, 0)

                    # 셀 내 모든 문단에 정보 삽입
                    while True:
                        p = hwp.GetPos()
                        if p[0] != list_id:
                            break

                        current_para = p[1]

                        # 문단 내 컨트롤 정보 확인
                        ctrl_info_list = []

                        # HeadCtrl 순회 + GetAnchorPos().Item("List")로 현재 셀 내 컨트롤 찾기
                        try:
                            ctrl = hwp.HeadCtrl
                            while ctrl:
                                try:
                                    anchor = ctrl.GetAnchorPos(0)
                                    ctrl_list_id = anchor.Item("List")
                                    ctrl_para_id = anchor.Item("Para")

                                    # 현재 셀(list_id)이고 현재 문단(para_id)에 속한 컨트롤만
                                    if ctrl_list_id == list_id and ctrl_para_id == current_para:
                                        ctrl_id = ctrl.CtrlID
                                        if ctrl_id and ctrl_id not in ("secd", "cold"):  # 섹션/열 정의 제외
                                            ctrl_detail = ctrl_id
                                            user_desc = ctrl.UserDesc if hasattr(ctrl, 'UserDesc') else ""
                                            if user_desc:
                                                ctrl_detail = user_desc

                                            try:
                                                props = ctrl.Properties
                                                if props:
                                                    if ctrl_id == "tbl":  # 표
                                                        rows = props.Item("RowCount") or ""
                                                        cols = props.Item("ColCount") or ""
                                                        if rows or cols:
                                                            ctrl_detail += f"({rows}x{cols})"
                                                    elif ctrl_id == "gso":  # 그리기 개체 (그림 포함)
                                                        w = props.Item("Width") or 0
                                                        h = props.Item("Height") or 0
                                                        treat = props.Item("TreatAsChar") or 0
                                                        # HWPUNIT → mm 변환 (7200 = 1인치 = 25.4mm)
                                                        w_mm = round(w * 25.4 / 7200, 1)
                                                        h_mm = round(h * 25.4 / 7200, 1)
                                                        ctrl_detail += f"({w_mm}x{h_mm}mm)"

                                                        # 바깥 여백 조회 (HShapeObject로)
                                                        try:
                                                            hwp.SetPosBySet(ctrl.GetAnchorPos(0))
                                                            hwp.FindCtrl()
                                                            shape = hwp.HParameterSet.HShapeObject
                                                            hwp.HAction.GetDefault("ShapeObjDialog", shape.HSet)
                                                            oml = shape.OutsideMarginLeft or 0
                                                            omr = shape.OutsideMarginRight or 0
                                                            omt = shape.OutsideMarginTop or 0
                                                            omb = shape.OutsideMarginBottom or 0
                                                            # HWPUNIT → mm 변환
                                                            oml_mm = round(oml * 25.4 / 7200, 1)
                                                            omr_mm = round(omr * 25.4 / 7200, 1)
                                                            omt_mm = round(omt * 25.4 / 7200, 1)
                                                            omb_mm = round(omb * 25.4 / 7200, 1)
                                                            ctrl_detail += f" 바깥여백:{oml_mm}/{omr_mm}/{omt_mm}/{omb_mm}mm"
                                                            hwp.HAction.Run("Cancel")
                                                        except:
                                                            pass

                                                        if treat:
                                                            ctrl_detail += "[글자취급]"
                                                    elif ctrl_id == "eqed":  # 수식
                                                        ctrl_detail = "수식"
                                            except:
                                                pass
                                            ctrl_info_list.append(ctrl_detail)
                                except:
                                    pass
                                ctrl = ctrl.Next
                        except:
                            pass

                        ctrl_info = ",".join(ctrl_info_list) if ctrl_info_list else ""

                        # 글자 모양 정보 조회 (CharShape 속성 사용)
                        font_name = ""
                        font_size = 0
                        char_spacing = 0
                        try:
                            cs = hwp.CharShape
                            if cs:
                                font_name = cs.Item("FaceNameHangul") or ""
                                height = cs.Item("Height") or 0
                                font_size = height // 100 if height else 0  # HWPUNIT → pt
                                char_spacing = cs.Item("CharSpacing") or 0  # 자간 (%)
                        except:
                            pass

                        # 문단 모양 정보 조회 (ParaShape 속성 사용)
                        align = ""
                        line_spacing = 0
                        align_map = {0: "양쪽", 1: "왼쪽", 2: "오른쪽", 3: "가운데", 4: "배분", 5: "나눔"}
                        try:
                            ps = hwp.ParaShape
                            if ps:
                                align_val = ps.Item("Align")
                                align = align_map.get(align_val, str(align_val))
                                line_spacing = ps.Item("LineSpacing") or 0  # 줄간격
                        except:
                            pass

                        # 문단 시작으로 이동
                        hwp.HAction.Run("MoveParaBegin")

                        # [list_id, para_id] + 스타일 + 컨트롤 정보 삽입
                        info_parts = [f"{list_id}", f"{current_para}"]

                        style_parts = []
                        if font_name:
                            style_parts.append(font_name)
                        if font_size:
                            style_parts.append(f"{font_size}pt")
                        if char_spacing:
                            style_parts.append(f"자간:{char_spacing}%")
                        if align:
                            style_parts.append(align)
                        if line_spacing:
                            style_parts.append(f"줄간:{line_spacing}%")

                        # 정보 문자열 조합
                        info_text = f"[{', '.join(info_parts)}"
                        if style_parts:
                            info_text += f" | {' '.join(style_parts)}"
                        if ctrl_info:
                            info_text += f" | ctrl:{ctrl_info}"
                        info_text += "] "
                        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
                        hwp.HParameterSet.HInsertText.Text = info_text
                        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

                        # 삽입한 텍스트 선택하여 파란색 적용
                        # 커서가 삽입 텍스트 끝에 있으므로 왼쪽으로 선택
                        for _ in range(len(info_text)):
                            hwp.HAction.Run("MoveSelLeft")

                        # 파란색(RGB: 0, 0, 255) 적용 - BGR 변환 필요
                        hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
                        hwp.HParameterSet.HCharShape.TextColor = 0xFF0000  # BGR: 파란색
                        hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)

                        # 선택 해제 후 커서를 텍스트 끝으로
                        hwp.HAction.Run("Cancel")

                        insert_count += 1
                        print(f"  셀({row},{col}) - 삽입: {info_text.strip()}")

                        # 다음 문단으로 이동
                        hwp.HAction.Run("MoveParaEnd")
                        hwp.HAction.Run("MoveRight")
                        after = hwp.GetPos()

                        if after[0] != list_id or after[1] == current_para:
                            break

                    # 셀로 돌아가기
                    hwp.SetPos(list_id, 0, 0)

                # 오른쪽 셀로
                before = list_id
                result = hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)
                after = hwp.GetPos()[0]

                if not result or after == before:
                    break
                col += 1

            # 다음 행
            hwp.SetPos(row_start_pos[0], row_start_pos[1], row_start_pos[2])
            before = hwp.GetPos()[0]
            result = hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)
            after = hwp.GetPos()[0]

            if not result or after == before:
                break
            row += 1

        hwp.HAction.Run("MoveParentList")
        hwp.HAction.Run("Cancel")

        print(f"  → 총 {insert_count}개 문단에 정보 삽입 완료")


class Logger:
    """콘솔과 파일에 동시 출력"""
    def __init__(self, log_path):
        import os
        os.makedirs(os.path.dirname(log_path), exist_ok=True)
        self.terminal = sys.stdout
        self.log = open(log_path, "w", encoding="utf-8")

    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)

    def flush(self):
        self.terminal.flush()
        self.log.flush()

    def close(self):
        self.log.close()


# 전역 hwp 인스턴스
_hwp_instance = None

def get_hwp():
    """전역 hwp 인스턴스 반환 (없으면 ROT에서 가져옴)"""
    global _hwp_instance
    if _hwp_instance is None:
        _hwp_instance = get_hwp_instance()
    return _hwp_instance

def set_hwp(hwp):
    """전역 hwp 인스턴스 설정"""
    global _hwp_instance
    _hwp_instance = hwp


if __name__ == "__main__":
    # 로그 파일로 출력 리다이렉트
    logger = Logger(LOG_FILE)
    sys.stdout = logger

    print(f"실행 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 70)
    print("1.hwp 테이블 정보 추출 테스트")
    print("=" * 70)

    # 1.hwp 파일 자동 열기
    import os
    if os.path.exists(TEST_FILE):
        print(f"\n테스트 파일 열기: {TEST_FILE}")
        hwp = open_hwp_file(TEST_FILE)
        set_hwp(hwp)  # 전역 인스턴스로 설정
        print("파일 열기 성공!\n")
    else:
        print(f"[오류] 테스트 파일이 없습니다: {TEST_FILE}")
        print("기존 열린 한글 문서를 사용합니다.\n")

    extract_table_cell_info()
    print("\n\n")
    extract_detailed_cell_info()
    print("\n\n")

    # 문단 앞에 [list_id, para_id, ctrl_id] 삽입 (먼저 실행)
    insert_para_info_to_document()
    print("\n\n")

    # 테이블 속성 및 필드 정보 추출 + 메모 삽입 (마지막에 실행)
    extract_table_properties()
    print("\n\n")

    # 필드 정보 추출
    extract_field_info()

    print(f"\n\n로그 저장 완료: {LOG_FILE}")
    logger.close()
