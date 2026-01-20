# -*- coding: utf-8 -*-
"""
셀 좌표를 한글 표에 파란색으로 매핑

각 셀에 (행, 열) 좌표를 파란색 텍스트로 삽입
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from cursor import get_hwp_instance
from table.table_boundary import TableBoundary


def insert_coordinate_text(hwp, list_id: int, row: int, col: int):
    """셀에 좌표 텍스트를 파란색으로 삽입"""
    # 셀로 이동
    hwp.SetPos(list_id, 0, 0)

    # 셀 끝으로 이동 (리스트 끝)
    hwp.MovePos(5, 0, 0)  # moveBottomOfList

    # 좌표 텍스트 삽입 (줄바꿈 + 좌표)
    coord_text = f"\r({row}, {col})"

    # 텍스트 삽입
    act = hwp.CreateAction("InsertText")
    pset = act.CreateSet()
    act.GetDefault(pset)
    pset.SetItem("Text", coord_text)
    act.Execute(pset)

    # 삽입한 텍스트 선택 (뒤에서부터 좌표 길이만큼)
    coord_len = len(f"({row}, {col})")
    for _ in range(coord_len):
        hwp.HAction.Run("MoveSelPrevChar")

    # 파란색 글자 모양 적용
    act = hwp.CreateAction("CharShape")
    pset = act.CreateSet()
    act.GetDefault(pset)
    pset.SetItem("TextColor", 0xFF0000)  # BGR: 파란색
    act.Execute(pset)

    # 선택 해제
    hwp.HAction.Run("Cancel")


def main():
    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return

    # 테이블 경계 분석
    boundary = TableBoundary(hwp, debug=False)

    # 테이블 내부인지 확인
    if not boundary._is_in_table():
        print("[오류] 커서가 테이블 내부에 있지 않습니다.")
        return

    print("경계 분석 중...")
    result = boundary.check_boundary_table()

    print("좌표 매핑 중 (v3: 서브셀 기반)...")
    coord_result = boundary.map_cell_coordinates_v3(result)

    print(f"총 {len(coord_result.cells)}개 셀에 좌표 삽입 중...")

    # 각 셀에 좌표 삽입
    for list_id, coord in coord_result.cells.items():
        insert_coordinate_text(hwp, list_id, coord.row, coord.col)
        print(f"  list_id={list_id} → ({coord.row}, {coord.col})")

    print("\n완료!")
    print(f"최대 행: {coord_result.max_row}, 최대 열: {coord_result.max_col}")


if __name__ == "__main__":
    main()
