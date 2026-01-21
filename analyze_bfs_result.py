# -*- coding: utf-8 -*-
"""BFS 매핑 결과 분석 - 중복/누락 검증"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from cursor import get_hwp_instance
from table.cell_position import CellPositionCalculator


def analyze_bfs_result():
    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        return

    print("=== BFS 매핑 결과 분석 ===\n")

    calc = CellPositionCalculator(hwp, debug=False)

    try:
        result = calc.calculate_bfs(max_cells=2000)

        print(f"총 셀: {len(result.cells)}개")
        print(f"X 레벨: {len(result.x_levels)}개")
        print(f"Y 레벨: {len(result.y_levels)}개")
        print(f"테이블 크기: {result.max_row + 1}행 x {result.max_col + 1}열")

        # 1. 그리드 위치별 셀 매핑 확인
        print("\n=== 1. 그리드 위치별 셀 점유 분석 ===")
        grid = {}  # (row, col) -> list of list_ids

        for list_id, cell in result.cells.items():
            for r in range(cell.start_row, cell.end_row + 1):
                for c in range(cell.start_col, cell.end_col + 1):
                    key = (r, c)
                    if key not in grid:
                        grid[key] = []
                    grid[key].append(list_id)

        # 중복 점유 확인
        duplicates = {k: v for k, v in grid.items() if len(v) > 1}
        if duplicates:
            print(f"\n[경고] 중복 점유 위치: {len(duplicates)}개")
            for pos, ids in sorted(duplicates.items())[:20]:
                print(f"  ({pos[0]}, {pos[1]}): list_ids = {ids}")
            if len(duplicates) > 20:
                print(f"  ... 외 {len(duplicates) - 20}개")
        else:
            print("중복 점유 없음 [OK]")

        # 2. 빈 위치 확인 (0,0부터 max_row, max_col까지)
        print("\n=== 2. 빈 위치 분석 ===")
        empty_positions = []
        for r in range(result.max_row + 1):
            for c in range(result.max_col + 1):
                if (r, c) not in grid:
                    empty_positions.append((r, c))

        if empty_positions:
            print(f"[경고] 비어있는 위치: {len(empty_positions)}개")
            # 행별로 그룹화
            empty_by_row = {}
            for r, c in empty_positions:
                if r not in empty_by_row:
                    empty_by_row[r] = []
                empty_by_row[r].append(c)

            for row in sorted(empty_by_row.keys())[:30]:
                cols = empty_by_row[row]
                if len(cols) <= 10:
                    print(f"  행 {row}: 열 {cols}")
                else:
                    print(f"  행 {row}: 열 {cols[:5]}...{cols[-5:]} (총 {len(cols)}개)")
            if len(empty_by_row) > 30:
                print(f"  ... 외 {len(empty_by_row) - 30}개 행")
        else:
            print("빈 위치 없음 [OK]")

        # 3. 음수 좌표 확인
        print("\n=== 3. 음수 좌표 확인 ===")
        negative_cells = []
        for list_id, cell in result.cells.items():
            if cell.start_row < 0 or cell.start_col < 0:
                negative_cells.append((list_id, cell))
            if cell.start_x < 0 or cell.start_y < 0:
                negative_cells.append((list_id, cell))

        if negative_cells:
            print(f"[경고] 음수 좌표 셀: {len(negative_cells)}개")
            for list_id, cell in negative_cells[:10]:
                print(f"  list_id={list_id}: row={cell.start_row}, col={cell.start_col}, "
                      f"x={cell.start_x}, y={cell.start_y}")
        else:
            print("음수 좌표 없음 [OK]")

        # 4. 첫 번째 열/행 분석
        print("\n=== 4. 첫 번째 열(col=0) 셀 목록 ===")
        first_col_cells = [(lid, c) for lid, c in result.cells.items() if c.start_col == 0]
        first_col_cells.sort(key=lambda x: x[1].start_row)
        print(f"총 {len(first_col_cells)}개 셀")
        for list_id, cell in first_col_cells[:20]:
            span_info = f"[{cell.rowspan}x{cell.colspan}]" if cell.rowspan > 1 or cell.colspan > 1 else ""
            print(f"  list_id={list_id}: row={cell.start_row} {span_info}")
        if len(first_col_cells) > 20:
            print(f"  ... 외 {len(first_col_cells) - 20}개")

        print("\n=== 5. 첫 번째 행(row=0) 셀 목록 ===")
        first_row_cells = [(lid, c) for lid, c in result.cells.items() if c.start_row == 0]
        first_row_cells.sort(key=lambda x: x[1].start_col)
        print(f"총 {len(first_row_cells)}개 셀")
        for list_id, cell in first_row_cells[:20]:
            span_info = f"[{cell.rowspan}x{cell.colspan}]" if cell.rowspan > 1 or cell.colspan > 1 else ""
            print(f"  list_id={list_id}: col={cell.start_col} {span_info}")
        if len(first_row_cells) > 20:
            print(f"  ... 외 {len(first_row_cells) - 20}개")

        # 5. X/Y 레벨 분포 확인
        print("\n=== 6. X 레벨 분포 (물리 좌표) ===")
        print(f"X 레벨 ({len(result.x_levels)}개):")
        for i, x in enumerate(result.x_levels[:15]):
            print(f"  [{i}] {x}")
        if len(result.x_levels) > 15:
            print(f"  ... (중간 생략)")
            for i, x in enumerate(result.x_levels[-5:], len(result.x_levels) - 5):
                print(f"  [{i}] {x}")

        print(f"\nY 레벨 ({len(result.y_levels)}개):")
        for i, y in enumerate(result.y_levels[:15]):
            print(f"  [{i}] {y}")
        if len(result.y_levels) > 15:
            print(f"  ... (중간 생략)")
            for i, y in enumerate(result.y_levels[-5:], len(result.y_levels) - 5):
                print(f"  [{i}] {y}")

        # 6. 셀 크기 통계
        print("\n=== 7. 셀 크기 통계 ===")
        widths = [c.end_x - c.start_x for c in result.cells.values()]
        heights = [c.end_y - c.start_y for c in result.cells.values()]
        print(f"너비: 최소={min(widths)}, 최대={max(widths)}, 평균={sum(widths)//len(widths)}")
        print(f"높이: 최소={min(heights)}, 최대={max(heights)}, 평균={sum(heights)//len(heights)}")

        # 7. 병합 셀 분석
        print("\n=== 8. 병합 셀 분석 ===")
        merged = calc.get_merged_cells(result)
        print(f"병합 셀: {len(merged)}개")

        # rowspan/colspan 분포
        rowspan_dist = {}
        colspan_dist = {}
        for cell in result.cells.values():
            rs = cell.rowspan
            cs = cell.colspan
            rowspan_dist[rs] = rowspan_dist.get(rs, 0) + 1
            colspan_dist[cs] = colspan_dist.get(cs, 0) + 1

        print(f"Rowspan 분포: {dict(sorted(rowspan_dist.items()))}")
        print(f"Colspan 분포: {dict(sorted(colspan_dist.items()))}")

    except ValueError as e:
        print(f"[오류] {e}")


if __name__ == "__main__":
    analyze_bfs_result()
