# -*- coding: utf-8 -*-
"""테이블 위에 그리드 라인 그리기 V2 - HAction 기반"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from cursor import get_hwp_instance
from table.cell_position import CellPositionCalculator


def draw_line_haction(hwp, x1, y1, x2, y2, color=0x0000FF):
    """HAction 기반 선 그리기"""
    try:
        # 선 그리기
        hwp.HAction.GetDefault("DrawLine", hwp.HParameterSet.HDrawLine.HSet)
        pset = hwp.HParameterSet.HDrawLine

        # 시작점, 끝점 (HWPUNIT)
        pset.StartX = int(x1)
        pset.StartY = int(y1)
        pset.EndX = int(x2)
        pset.EndY = int(y2)

        # 선 속성
        pset.LineColor = color
        pset.LineWidth = 20  # 선 두께

        hwp.HAction.Execute("DrawLine", hwp.HParameterSet.HDrawLine.HSet)
        return True
    except Exception as e:
        return False


def insert_line_shape(hwp, x1, y1, x2, y2, color=0x0000FF):
    """직선 도형 삽입"""
    try:
        # 먼저 문서 시작으로 이동
        hwp.MovePos(2)  # 문서 시작

        # 직선 삽입
        hwp.HAction.GetDefault("InsertLine", hwp.HParameterSet.HShapeObject.HSet)
        pset = hwp.HParameterSet.HShapeObject

        # 위치 설정
        pset.HSet.SetItem("ShapePosition", 0)  # 절대 위치
        pset.HSet.SetItem("ShapeStartX", int(x1))
        pset.HSet.SetItem("ShapeStartY", int(y1))
        pset.HSet.SetItem("ShapeEndX", int(x2))
        pset.HSet.SetItem("ShapeEndY", int(y2))

        # 선 색상
        pset.HSet.SetItem("LineColor", color)
        pset.HSet.SetItem("LineWidth", 20)

        hwp.HAction.Execute("InsertLine", hwp.HParameterSet.HShapeObject.HSet)
        return True
    except Exception as e:
        return False


def draw_grid_lines():
    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return

    print("=== 테이블 위에 그리드 라인 그리기 V2 ===\n")

    calc = CellPositionCalculator(hwp, debug=False)

    try:
        result = calc.calculate(max_cells=2000)

        print(f"X 레벨: {len(result.x_levels)}개")
        print(f"Y 레벨: {len(result.y_levels)}개")

        # 테이블 크기
        table_width = result.x_levels[-1] if result.x_levels else 0
        table_height = result.y_levels[-1] if result.y_levels else 0

        print(f"테이블 크기: {table_width} x {table_height} HWPUNIT")

        # 테이블 시작 위치 오프셋 (조정 필요할 수 있음)
        # 현재 테이블의 실제 위치를 가져와야 함
        offset_x = 0
        offset_y = 0

        # 테이블 컨트롤 찾기
        ctrl = hwp.ParentCtrl
        if ctrl and ctrl.CtrlID == "tbl":
            try:
                props = ctrl.Properties
                # 테이블 위치 정보 (가능한 경우)
                offset_x = props.Item("StartX") if props else 0
                offset_y = props.Item("StartY") if props else 0
            except:
                pass

        print(f"오프셋: ({offset_x}, {offset_y})")

        # 선 색상 (빨간색)
        line_color = 0x0000FF  # BGR

        # 테스트: 하나의 선만 그려보기
        print("\n테스트: 대각선 하나 그리기...")

        # 방법 1: HAction DrawLine
        success = draw_line_haction(hwp,
                                   offset_x, offset_y,
                                   offset_x + table_width, offset_y + table_height,
                                   line_color)
        if success:
            print("[성공] HAction DrawLine")
        else:
            print("[실패] HAction DrawLine")

            # 방법 2: InsertLine
            success = insert_line_shape(hwp,
                                       offset_x, offset_y,
                                       offset_x + table_width, offset_y + table_height,
                                       line_color)
            if success:
                print("[성공] InsertLine")
            else:
                print("[실패] InsertLine")

        if not success:
            print("\n선 그리기 실패. 다른 방법을 시도합니다...")
            print("\n=== 대안: 셀에 좌표 텍스트 삽입 ===")
            insert_coords = input("각 셀에 좌표를 삽입하시겠습니까? (y/n): ").strip().lower()
            if insert_coords == 'y':
                calc.insert_all_coordinates(result, verbose=True)

    except Exception as e:
        print(f"[오류] {e}")
        import traceback
        traceback.print_exc()


def create_excel_comparison():
    """엑셀로 그리드 매핑 비교 이미지 생성"""
    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return

    print("=== 엑셀로 그리드 매핑 비교 ===\n")

    try:
        from table.table_excel_converter import TableExcelConverter

        converter = TableExcelConverter(hwp, debug=False)

        # 엑셀로 변환
        output_path = "C:\\win32hwp\\grid_comparison.xlsx"
        converter.to_excel(output_path, with_text=True)
        print(f"[완료] {output_path}")
        print("엑셀 파일에서 셀 병합 구조를 확인하세요.")

    except Exception as e:
        print(f"[오류] {e}")


if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1 and sys.argv[1] == "excel":
        create_excel_comparison()
    else:
        draw_grid_lines()
