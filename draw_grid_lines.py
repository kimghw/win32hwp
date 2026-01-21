# -*- coding: utf-8 -*-
"""테이블 위에 그리드 라인 그리기 - 매핑 시각적 검증"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from cursor import get_hwp_instance
from table.cell_position import CellPositionCalculator


def draw_grid_lines():
    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return

    print("=== 테이블 위에 그리드 라인 그리기 ===\n")

    calc = CellPositionCalculator(hwp, debug=False)

    try:
        result = calc.calculate(max_cells=2000)

        print(f"X 레벨: {len(result.x_levels)}개")
        print(f"Y 레벨: {len(result.y_levels)}개")
        print(f"테이블 크기: {result.max_row + 1}행 x {result.max_col + 1}열")

        # 테이블 위치 정보 (첫 번째 셀 기준)
        # 테이블의 물리적 시작 위치를 구해야 함

        # 현재 선택된 테이블의 컨트롤 찾기
        ctrl = hwp.CurSelectedCtrl
        if not ctrl:
            # 테이블 안에 있으면 ParentCtrl로 찾기
            ctrl = hwp.ParentCtrl

        if not ctrl or ctrl.CtrlID != "tbl":
            print("[오류] 테이블을 찾을 수 없습니다. 테이블을 선택하거나 테이블 안에 커서를 두세요.")
            return

        # 테이블 속성에서 위치 정보 가져오기
        props = ctrl.Properties

        # 테이블 전체 너비/높이
        table_width = result.x_levels[-1] if result.x_levels else 0
        table_height = result.y_levels[-1] if result.y_levels else 0

        print(f"\n테이블 물리 크기: {table_width} x {table_height} HWPUNIT")
        print(f"X 레벨: {result.x_levels[:10]}..." if len(result.x_levels) > 10 else f"X 레벨: {result.x_levels}")
        print(f"Y 레벨: {result.y_levels[:10]}..." if len(result.y_levels) > 10 else f"Y 레벨: {result.y_levels}")

        # 그리드 라인 색상 (빨간색)
        line_color = 0x0000FF  # BGR: 빨간색

        print(f"\n세로선 {len(result.x_levels)}개, 가로선 {len(result.y_levels)}개 그리기...")

        # 세로선 그리기 (X 레벨)
        for i, x in enumerate(result.x_levels):
            draw_line(hwp, x, 0, x, table_height, line_color, f"X{i}")

        # 가로선 그리기 (Y 레벨)
        for i, y in enumerate(result.y_levels):
            draw_line(hwp, 0, y, table_width, y, line_color, f"Y{i}")

        print("\n[완료] 그리드 라인이 그려졌습니다.")
        print("※ 테이블과 라인의 위치가 맞지 않으면 테이블 시작 위치 오프셋 조정 필요")

    except Exception as e:
        print(f"[오류] {e}")
        import traceback
        traceback.print_exc()


def draw_line(hwp, x1, y1, x2, y2, color, name=""):
    """HWP에 선 그리기"""
    try:
        # 선 그리기 액션
        act = hwp.CreateAction("DrawLine")
        pset = act.CreateSet()
        act.GetDefault(pset)

        # 시작점, 끝점 설정 (HWPUNIT)
        pset.SetItem("StartX", int(x1))
        pset.SetItem("StartY", int(y1))
        pset.SetItem("EndX", int(x2))
        pset.SetItem("EndY", int(y2))

        # 선 색상
        pset.SetItem("LineColor", color)
        pset.SetItem("LineWidth", 10)  # 선 두께 (HWPUNIT)

        act.Execute(pset)

    except Exception as e:
        # DrawLine이 안되면 다른 방식 시도
        print(f"[경고] 선 그리기 실패: {e}")


def draw_grid_overlay_simple():
    """간단한 방식: 각 셀에 좌표 텍스트 삽입"""
    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return

    print("=== 각 셀에 좌표 표시 ===\n")

    calc = CellPositionCalculator(hwp, debug=False)

    try:
        result = calc.calculate(max_cells=2000)

        print(f"총 {len(result.cells)}개 셀에 좌표 삽입...")
        calc.insert_all_coordinates(result, verbose=False)
        print("[완료]")

    except Exception as e:
        print(f"[오류] {e}")


if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1 and sys.argv[1] == "text":
        # 텍스트 좌표 삽입 모드
        draw_grid_overlay_simple()
    else:
        # 라인 그리기 모드
        draw_grid_lines()
