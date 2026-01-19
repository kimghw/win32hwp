# -*- coding: utf-8 -*-
"""
테이블 셀 정보 추출 스크립트

테이블 내 list_id, para_id, ctrl 정보를 추출하고 문서에 삽입합니다.
"""

import sys
import os
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor_utils import get_hwp_instance

try:
    from .table_info import TableInfo, MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL
except ImportError:
    from table_info import TableInfo, MOVE_RIGHT_OF_CELL, MOVE_DOWN_OF_CELL

# 로그 파일 설정
LOG_FILE = "logs/table_info_extract.log"
TEST_FILE = r"C:\win32hwp\1.hwp"


# ============================================================
# 공통 유틸리티
# ============================================================

def iterate_table_cells(hwp, callback):
    """
    테이블 셀을 순회하며 callback 함수 호출

    Args:
        hwp: HWP 인스턴스
        callback: func(hwp, row, col, list_id) -> bool
                  False 반환 시 순회 중단

    Returns:
        set: 방문한 list_id 집합
    """
    visited = set()
    row = 0

    while True:
        row_start_pos = hwp.GetPos()
        col = 0

        # 현재 행 순회
        while True:
            pos = hwp.GetPos()
            list_id = pos[0]

            if list_id not in visited:
                visited.add(list_id)

                # 콜백 호출
                if callback(hwp, row, col, list_id) is False:
                    return visited

            # 오른쪽 셀로 이동
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

    return visited


def get_ctrls_in_cell(hwp, target_list_id, target_para_id=None):
    """
    특정 셀(list_id)에 속한 컨트롤 찾기

    Args:
        hwp: HWP 인스턴스
        target_list_id: 대상 list_id
        target_para_id: 특정 문단만 필터링 (None=전체)

    Returns:
        list: 컨트롤 정보 리스트
    """
    ctrls = []
    ctrl = hwp.HeadCtrl

    while ctrl:
        try:
            anchor = ctrl.GetAnchorPos(0)
            ctrl_list_id = anchor.Item("List")
            ctrl_para_id = anchor.Item("Para")

            if ctrl_list_id == target_list_id:
                if target_para_id is None or ctrl_para_id == target_para_id:
                    ctrl_id = ctrl.CtrlID
                    if ctrl_id and ctrl_id not in ("secd", "cold"):  # 섹션/열 정의 제외
                        ctrls.append({
                            'ctrl': ctrl,
                            'id': ctrl_id,
                            'desc': getattr(ctrl, 'UserDesc', ctrl_id),
                            'para': ctrl_para_id
                        })
        except:
            pass
        ctrl = ctrl.Next

    return ctrls


def get_char_shape_info(hwp):
    """현재 위치의 글자 모양 정보 조회"""
    try:
        cs = hwp.CharShape
        if cs:
            font_name = cs.Item("FaceNameHangul") or ""
            height = cs.Item("Height") or 0
            font_size = height // 100 if height else 0
            char_spacing = cs.Item("CharSpacing") or 0
            return font_name, font_size, char_spacing
    except:
        pass
    return "", 0, 0


def get_para_shape_info(hwp):
    """현재 위치의 문단 모양 정보 조회"""
    align_map = {0: "양쪽", 1: "왼쪽", 2: "오른쪽", 3: "가운데", 4: "배분", 5: "나눔"}
    try:
        ps = hwp.ParaShape
        if ps:
            align_val = ps.Item("Align")
            align = align_map.get(align_val, str(align_val))
            line_spacing = ps.Item("LineSpacing") or 0
            return align, line_spacing
    except:
        pass
    return "", 0


def insert_colored_text(hwp, text, color_bgr):
    """색상이 적용된 텍스트 삽입"""
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

    # 삽입한 텍스트 선택
    for _ in range(len(text)):
        hwp.HAction.Run("MoveSelLeft")

    # 색상 적용
    hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
    hwp.HParameterSet.HCharShape.TextColor = color_bgr
    hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)

    hwp.HAction.Run("Cancel")


# ============================================================
# 정보 추출 함수
# ============================================================

def extract_field_info(hwp):
    """문서 내 필드 정보 추출"""
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
        if ctrl_id in ("%", "%%", "field", "clickhere"):
            try:
                anchor = ctrl.GetAnchorPos(0)
                list_id = anchor.Item("List")
                para_id = anchor.Item("Para")
                user_desc = getattr(ctrl, 'UserDesc', ctrl_id)
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


def extract_table_cell_info(hwp):
    """테이블 셀별 list_id, para_id, ctrl 정보 추출"""
    table = TableInfo(hwp, debug=False)

    tables = table.find_all_tables()
    if not tables:
        print("[오류] 문서에 테이블이 없습니다.")
        return

    print(f"=== 문서 내 테이블 수: {len(tables)}개 ===\n")

    for t in tables:
        print(f"\n{'='*60}")
        print(f"테이블 #{t['num']}")
        print(f"{'='*60}")

        table.enter_table(t['num'])
        table.cells.clear()
        cells = table.collect_cells_bfs()
        size = table.get_table_size()

        print(f"크기: {size['rows']}행 x {size['cols']}열")
        print(f"셀 수: {len(cells)}개")
        print()

        table.move_to_first_cell()

        print(f"{'셀':<10} {'list_id':<10} {'para_id':<10} {'char_pos':<10} {'ctrl_id':<10}")
        print("-" * 60)

        def print_cell_info(hwp, row, col, list_id):
            pos = hwp.GetPos()
            para_id = pos[1]
            char_pos = pos[2]

            ctrl_id = ""
            try:
                ctrl = hwp.CurSelectedCtrl
                if ctrl:
                    ctrl_id = ctrl.CtrlID
            except:
                pass

            print(f"({row},{col}){' '*4} {list_id:<10} {para_id:<10} {char_pos:<10} {ctrl_id:<10}")
            return True

        iterate_table_cells(hwp, print_cell_info)

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

        caption = table.get_table_caption(t['num'])
        if caption:
            print(f"\n캡션: {caption}")

        hwp.HAction.Run("MoveParentList")
        hwp.HAction.Run("Cancel")


def extract_detailed_cell_info(hwp):
    """각 셀의 상세 정보 (para 포함) 추출"""
    table = TableInfo(hwp, debug=False)

    tables = table.find_all_tables()
    if not tables:
        print("[오류] 문서에 테이블이 없습니다.")
        return

    for t in tables:
        print(f"\n{'='*70}")
        print(f"테이블 #{t['num']} - 상세 정보")
        print(f"{'='*70}")

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

        def print_para_info(hwp, row, col, list_id):
            paras = []
            hwp.SetPos(list_id, 0, 0)

            while True:
                p = hwp.GetPos()
                if p[0] != list_id:
                    break
                paras.append(p[1])

                before_para = p[1]
                hwp.HAction.Run("MoveParaEnd")
                hwp.HAction.Run("MoveRight")
                after = hwp.GetPos()

                if after[0] != list_id or after[1] == before_para:
                    break

            hwp.SetPos(list_id, 0, 0)
            print(f"셀 ({row},{col}) list_id={list_id}: para_ids={paras}")
            return True

        iterate_table_cells(hwp, print_para_info)

        hwp.HAction.Run("MoveParentList")
        hwp.HAction.Run("Cancel")


def extract_table_properties(hwp):
    """테이블 속성 및 필드 정보 추출"""
    print(f"\n{'='*70}")
    print("테이블 속성 및 필드 정보")
    print(f"{'='*70}")

    ctrl = hwp.HeadCtrl
    table_num = 0
    table_info = TableInfo(hwp, debug=False)

    while ctrl:
        if ctrl.CtrlID == "tbl":
            print(f"\n{'─'*70}")
            print(f"테이블 #{table_num}")
            print(f"{'─'*70}")

            # 테이블 속성 출력
            _print_table_properties(ctrl)

            # 테이블 필드 정보
            print("\n[테이블 필드 정보]")
            try:
                hwp.SetPosBySet(ctrl.GetAnchorPos(0))
                hwp.HAction.Run("SelectCtrlFront")
                hwp.HAction.Run("ShapeObjTableSelCell")

                table_info.cells.clear()
                table_info.collect_cells_bfs()
                table_info.move_to_first_cell()

                field_count = [0]  # mutable for closure

                def check_cell_ctrls(hwp, row, col, list_id):
                    ctrls = get_ctrls_in_cell(hwp, list_id)
                    for c in ctrls:
                        field_count[0] += 1
                        print(f"  셀({row},{col}) - CtrlID: {c['id']}, Desc: {c['desc']}")
                    return True

                iterate_table_cells(hwp, check_cell_ctrls)

                if field_count[0] == 0:
                    print("  (필드 없음)")

                hwp.HAction.Run("MoveParentList")
                hwp.HAction.Run("Cancel")

            except Exception as e:
                print(f"  필드 조회 실패: {e}")

            # 셀 속성 정보 삽입
            _insert_cell_properties(hwp, ctrl, table_info)

            table_num += 1

        ctrl = ctrl.Next

    if table_num == 0:
        print("문서에 테이블이 없습니다.")


def _print_table_properties(ctrl):
    """테이블 속성 출력"""
    try:
        props = ctrl.Properties
        if props:
            print("\n[테이블 속성 (Properties)]")

            prop_names = [
                "TreatAsChar", "PageBreak", "RepeatHeader",
                "Width", "Height", "CellSpacing",
                "CellMarginLeft", "CellMarginRight",
                "CellMarginTop", "CellMarginBottom",
            ]

            for name in prop_names:
                try:
                    value = props.Item(name)
                    if value is not None:
                        print(f"  {name}: {value}")
                except:
                    pass
    except Exception as e:
        print(f"  속성 조회 실패: {e}")


def _insert_cell_properties(hwp, table_ctrl, table_info):
    """셀 속성 정보를 빨간색 텍스트로 삽입"""
    print("\n[셀 속성 정보 → 빨간색 텍스트로 삽입]")

    try:
        hwp.SetPosBySet(table_ctrl.GetAnchorPos(0))
        hwp.HAction.Run("SelectCtrlFront")
        hwp.HAction.Run("ShapeObjTableSelCell")

        table_info.move_to_first_cell()
        insert_count = [0]

        def insert_cell_info(hwp, row, col, list_id):
            try:
                cell_shape = hwp.CellShape
                if cell_shape:
                    cell = cell_shape.Item("Cell")
                    if cell:
                        width = cell.Item("Width")
                        height = cell.Item("Height")
                        header = cell.Item("Header")
                        protected = cell.Item("Protected")

                        # 셀 필드 정보
                        cell_field = _get_cell_field(hwp, list_id)

                        # 속성 문자열 생성
                        props_str = f" {{({row},{col}) list:{list_id} W:{width} H:{height}"

                        flags = []
                        if header:
                            flags.append("제목")
                        if protected:
                            flags.append("보호")
                        if flags:
                            props_str += f" [{','.join(flags)}]"
                        if cell_field:
                            props_str += cell_field
                        props_str += "}"

                        # 셀 끝으로 이동 후 삽입
                        _move_to_cell_end(hwp, list_id)

                        insert_colored_text(hwp, "\r\n" + props_str, 0x0000FF)  # 빨간색

                        insert_count[0] += 1
                        print(f"  셀({row},{col}) list_id={list_id}:{props_str}")

            except Exception as e:
                print(f"  셀({row},{col}) 속성 삽입 실패: {e}")

            return True

        iterate_table_cells(hwp, insert_cell_info)

        hwp.HAction.Run("MoveParentList")
        hwp.HAction.Run("Cancel")

        print(f"  → 총 {insert_count[0]}개 셀에 속성 정보 삽입 완료")

    except Exception as e:
        print(f"  셀 속성 삽입 실패: {e}")


def _get_cell_field(hwp, list_id):
    """셀 내 필드 정보 조회"""
    try:
        field_list = hwp.GetFieldList(1, 0)  # 옵션 1 = 셀 필드
        if field_list:
            fields = field_list.split('\x02')
            for fname in fields:
                if fname:
                    try:
                        hwp.MoveToField(fname, True, True, True)
                        fpos = hwp.GetPos()
                        if fpos[0] == list_id:
                            fval = hwp.GetFieldText(fname) or ""
                            result = f" 필드:{fname}"
                            if fval:
                                result += f"={fval}"
                            hwp.SetPos(list_id, 0, 0)
                            return result
                    except:
                        pass
            hwp.SetPos(list_id, 0, 0)
    except:
        pass
    return " 필드없음"


def _move_to_cell_end(hwp, list_id):
    """셀 내 마지막 문단 끝으로 이동"""
    hwp.SetPos(list_id, 0, 0)
    while True:
        hwp.HAction.Run("MoveParaEnd")
        before_pos = hwp.GetPos()
        hwp.HAction.Run("MoveRight")
        after_pos = hwp.GetPos()

        if after_pos[0] != list_id or after_pos == before_pos:
            hwp.SetPos(before_pos[0], before_pos[1], before_pos[2])
            hwp.HAction.Run("MoveParaEnd")
            break


def insert_para_info_to_document(hwp):
    """각 셀의 문단 앞에 [list_id, para_id] 정보 삽입"""
    table = TableInfo(hwp, debug=False)

    tables = table.find_all_tables()
    if not tables:
        print("[오류] 문서에 테이블이 없습니다.")
        return

    print(f"\n=== 문단 앞에 [list_id, para_id] 정보 삽입 ===\n")

    for t in tables:
        print(f"테이블 #{t['num']} 처리 중...")

        table.enter_table(t['num'])
        table.cells.clear()
        table.collect_cells_bfs()
        table.move_to_first_cell()

        insert_count = [0]

        def insert_para_info(hwp, row, col, list_id):
            hwp.SetPos(list_id, 0, 0)

            while True:
                p = hwp.GetPos()
                if p[0] != list_id:
                    break

                current_para = p[1]

                # 컨트롤 정보 조회
                ctrls = get_ctrls_in_cell(hwp, list_id, current_para)
                ctrl_info = _format_ctrl_info(hwp, ctrls)

                # 스타일 정보 조회
                font_name, font_size, char_spacing = get_char_shape_info(hwp)
                align, line_spacing = get_para_shape_info(hwp)

                # 문단 시작으로 이동
                hwp.HAction.Run("MoveParaBegin")

                # 정보 문자열 조합
                info_text = _build_info_text(
                    list_id, current_para,
                    font_name, font_size, char_spacing,
                    align, line_spacing, ctrl_info
                )

                insert_colored_text(hwp, info_text, 0xFF0000)  # 파란색

                insert_count[0] += 1
                print(f"  셀({row},{col}) - 삽입: {info_text.strip()}")

                # 다음 문단으로 이동
                hwp.HAction.Run("MoveParaEnd")
                hwp.HAction.Run("MoveRight")
                after = hwp.GetPos()

                if after[0] != list_id or after[1] == current_para:
                    break

            hwp.SetPos(list_id, 0, 0)
            return True

        iterate_table_cells(hwp, insert_para_info)

        hwp.HAction.Run("MoveParentList")
        hwp.HAction.Run("Cancel")

        print(f"  → 총 {insert_count[0]}개 문단에 정보 삽입 완료")


def _format_ctrl_info(hwp, ctrls):
    """컨트롤 정보를 문자열로 포맷"""
    ctrl_info_list = []

    for c in ctrls:
        ctrl_detail = c['desc'] or c['id']

        try:
            props = c['ctrl'].Properties
            if props:
                if c['id'] == "tbl":
                    rows = props.Item("RowCount") or ""
                    cols = props.Item("ColCount") or ""
                    if rows or cols:
                        ctrl_detail += f"({rows}x{cols})"

                elif c['id'] == "gso":
                    w = props.Item("Width") or 0
                    h = props.Item("Height") or 0
                    treat = props.Item("TreatAsChar") or 0
                    w_mm = round(w * 25.4 / 7200, 1)
                    h_mm = round(h * 25.4 / 7200, 1)
                    ctrl_detail += f"({w_mm}x{h_mm}mm)"
                    if treat:
                        ctrl_detail += "[글자취급]"

                elif c['id'] == "eqed":
                    ctrl_detail = "수식"
        except:
            pass

        ctrl_info_list.append(ctrl_detail)

    return ",".join(ctrl_info_list) if ctrl_info_list else ""


def _build_info_text(list_id, para_id, font_name, font_size, char_spacing,
                     align, line_spacing, ctrl_info):
    """정보 텍스트 문자열 생성"""
    info_parts = [f"{list_id}", f"{para_id}"]

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

    info_text = f"[{', '.join(info_parts)}"
    if style_parts:
        info_text += f" | {' '.join(style_parts)}"
    if ctrl_info:
        info_text += f" | ctrl:{ctrl_info}"
    info_text += "] "

    return info_text


# ============================================================
# 로거
# ============================================================

class Logger:
    """콘솔과 파일에 동시 출력"""
    def __init__(self, log_path):
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


# ============================================================
# 메인
# ============================================================

if __name__ == "__main__":
    logger = Logger(LOG_FILE)
    sys.stdout = logger

    print(f"실행 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 70)
    print("테이블 정보 추출 테스트")
    print("=" * 70)

    # HWP 인스턴스 연결
    hwp = get_hwp_instance()

    if not hwp:
        # 테스트 파일 열기 시도
        if os.path.exists(TEST_FILE):
            print(f"\n테스트 파일 열기: {TEST_FILE}")
            import win32com.client as win32
            hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
            hwp.Open(TEST_FILE)
            hwp.XHwpWindows.Item(0).Visible = True
            print("파일 열기 성공!\n")
        else:
            print(f"[오류] 한글이 실행 중이 아니고 테스트 파일도 없습니다: {TEST_FILE}")
            sys.exit(1)

    # 정보 추출
    extract_table_cell_info(hwp)
    print("\n\n")

    extract_detailed_cell_info(hwp)
    print("\n\n")

    insert_para_info_to_document(hwp)
    print("\n\n")

    extract_table_properties(hwp)
    print("\n\n")

    extract_field_info(hwp)

    print(f"\n\n로그 저장 완료: {LOG_FILE}")
    logger.close()
