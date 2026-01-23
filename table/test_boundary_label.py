# -*- coding: utf-8 -*-
"""
테이블 경계 셀에 라벨 작성 테스트

각 경계 셀에 해당하는 라벨을 텍스트로 삽입:
- top_border_line: "첫행"
- bottom_border_line: "끝행"
- left_border_line: "첫열"
- right_border_line: "끝열"
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor import get_hwp_instance
from table.table_boundary import TableBoundary


def insert_text(hwp, text: str):
    """현재 커서 위치에 텍스트 삽입"""
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)


def main():
    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return

    boundary = TableBoundary(hwp, debug=True)

    # 테이블 내부인지 확인
    if not boundary._is_in_table():
        print("[오류] 커서가 테이블 내부에 있지 않습니다.")
        return

    # 경계 분석
    result = boundary.check_boundary_table()

    print("\n=== 경계 정보 ===")
    print(f"top_border_line: {result.top_border_line}")
    print(f"bottom_border_line: {result.bottom_border_line}")
    print(f"left_border_line: {result.left_border_line}")
    print(f"right_border_line: {result.right_border_line}")

    # 각 셀에 라벨 작성
    labels = {}

    for cell_id in result.top_border_line:
        labels.setdefault(cell_id, []).append("첫행")

    for cell_id in result.bottom_border_line:
        labels.setdefault(cell_id, []).append("끝행")

    for cell_id in result.left_border_line:
        labels.setdefault(cell_id, []).append("첫열")

    for cell_id in result.right_border_line:
        labels.setdefault(cell_id, []).append("끝열")

    print("\n=== 라벨 작성 ===")
    for cell_id, cell_labels in labels.items():
        label_text = " / ".join(cell_labels)
        print(f"셀 {cell_id}: {label_text}")

        # 셀로 이동 후 텍스트 삽입
        hwp.SetPos(cell_id, 0, 0)
        insert_text(hwp, f"[{label_text}]")

    print("\n완료!")


if __name__ == "__main__":
    main()
