# -*- coding: utf-8 -*-
"""
테이블 누름틀 자동 입력 스크립트

실행 중인 한글 문서의 모든 테이블에 누름틀을 자동으로 삽입합니다.
"""

from table_manager import TableManager
from cursor_utils import get_hwp_instance
import yaml


def auto_insert_fields():
    """
    실행 중인 한글 문서의 모든 테이블에 누름틀 자동 삽입

    Returns:
        list: 처리된 테이블 결과 목록
    """
    print("=" * 60)
    print("  테이블 누름틀 자동 입력")
    print("=" * 60)
    print()

    # 1. 한글 인스턴스 연결
    print("[1단계] 한글 인스턴스 연결 중...")
    hwp = get_hwp_instance()
    if not hwp:
        print("✗ 오류: 한글이 실행 중이지 않습니다.")
        print("  → 한글을 먼저 실행하고 문서를 열어주세요.")
        return None

    print("✓ 한글 연결 성공\n")

    # 2. 테이블 스캔
    print("[2단계] 문서 내 테이블 스캔 중...")
    manager = TableManager(hwp)
    tables = manager.scan_tables()

    if not tables:
        print("✗ 문서에 테이블이 없습니다.")
        return []

    print(f"✓ {len(tables)}개의 테이블 발견\n")

    # 3. 테이블 목록 출력
    print("[3단계] 발견된 테이블 목록:")
    for i, t in enumerate(tables):
        print(f"  {i}. 테이블 (Para={t['para']})")
    print()

    # 4. 작성자 정보 입력
    print("[4단계] 작성자 정보 입력 (선택사항, Enter=건너뛰기)")
    created_by = input("  작성자명: ").strip()
    description = input("  문서 설명: ").strip()

    author_info = None
    if created_by or description:
        author_info = {
            'created_by': created_by or 'Unknown',
            'description': description
        }
        print(f"✓ 작성자: {author_info['created_by']}")
        if description:
            print(f"✓ 설명: {description}")
    print()

    # 5. 모든 테이블에 누름틀 삽입
    print("[5단계] 모든 테이블에 누름틀 삽입 시작\n")

    results = []

    for i, table in enumerate(tables):
        # 인덱스로 테이블 컨트롤 가져오기
        table_ctrl = manager.get_table_by_index(i)
        table_id = f"TBL{i+1:03d}"  # TBL001, TBL002, ...

        print(f"{'─' * 60}")
        print(f"테이블 {i+1}/{len(tables)} 처리 중 (ID: {table_id})")
        print(f"{'─' * 60}")

        # 헤더/푸터 행 수 입력
        # 테이블 크기 가져오기 (셀 탐색 방식)
        manager.move_to_cell(table_ctrl, 0, 0)

        # 행 개수 세기
        total_rows = 1
        while True:
            prev_para = hwp.GetPos()[1]
            hwp.HAction.Run("TableLowerCell")
            curr_para = hwp.GetPos()[1]

            # 문단이 달라지면 다음 셀로 이동한 것
            if curr_para != prev_para:
                total_rows += 1
            else:
                # 더 이상 이동 안 되면 마지막 행
                break

        print(f"총 행 수: {total_rows}")
        header_input = input(f"  헤더 행 수 (0~{total_rows}, 기본값=1): ").strip()
        header_rows = int(header_input) if header_input else 1

        remaining = total_rows - header_rows
        footer_input = input(f"  푸터 행 수 (0~{remaining}, 기본값=0): ").strip()
        footer_rows = int(footer_input) if footer_input else 0

        if header_rows + footer_rows >= total_rows:
            print("⚠ 경고: 헤더+푸터가 전체 행 이상입니다. 푸터를 0으로 설정합니다.")
            footer_rows = 0

        # 누름틀 삽입
        result = manager.insert_structured_fields(
            table_ctrl,
            table_id,
            header_rows,
            footer_rows,
            author_info
        )

        # 결과 저장
        results.append({
            'table_index': i,
            'table_id': table_id,
            'result': result
        })

        # YAML 파일로 저장
        yaml_path = f"table_{table_id}_structure.yaml"
        save_data = {
            'author_info': result['author_info'],
            'metadata': result['metadata'],
            'header': {str(k): v for k, v in result['header'].items()},
            'body': {str(k): v for k, v in result['body'].items()},
            'footer': {str(k): v for k, v in result['footer'].items()}
        }

        with open(yaml_path, 'w', encoding='utf-8') as f:
            yaml.dump(save_data, f, allow_unicode=True, default_flow_style=False)

        print(f"✓ 구조 정보 저장: {yaml_path}")
        print()

    # 6. 전체 요약
    print("=" * 60)
    print("  완료 요약")
    print("=" * 60)
    print(f"✓ 처리된 테이블: {len(results)}개")

    total_fields = 0
    for r in results:
        table_result = r['result']
        count = len(table_result['header']) + len(table_result['body']) + len(table_result['footer'])
        total_fields += count
        print(f"  - {r['table_id']}: {count}개 누름틀 삽입")

    print(f"\n✓ 총 삽입된 누름틀: {total_fields}개")
    print()

    return results


def quick_insert_all(header_rows=1, footer_rows=0, author_name="", description=""):
    """
    빠른 일괄 삽입 (모든 테이블에 동일한 설정 적용)

    Args:
        header_rows: 헤더 행 수 (기본값: 1)
        footer_rows: 푸터 행 수 (기본값: 0)
        author_name: 작성자명 (선택)
        description: 설명 (선택)

    Returns:
        list: 처리된 테이블 결과 목록
    """
    print("=" * 60)
    print("  빠른 일괄 누름틀 삽입")
    print("=" * 60)
    print(f"설정: 헤더={header_rows}행, 푸터={footer_rows}행")
    if author_name:
        print(f"작성자: {author_name}")
    print()

    # 한글 연결
    hwp = get_hwp_instance()
    if not hwp:
        print("✗ 한글이 실행 중이지 않습니다.")
        return None

    manager = TableManager(hwp)
    tables = manager.scan_tables()

    if not tables:
        print("✗ 문서에 테이블이 없습니다.")
        return []

    print(f"✓ {len(tables)}개 테이블 발견\n")

    # 작성자 정보
    author_info = None
    if author_name or description:
        author_info = {
            'created_by': author_name or 'Unknown',
            'description': description
        }

    results = []

    for i, table in enumerate(tables):
        # 인덱스로 테이블 컨트롤 가져오기
        table_ctrl = manager.get_table_by_index(i)
        table_id = f"TBL{i+1:03d}"

        print(f"[{i+1}/{len(tables)}] {table_id} 처리 중...")

        result = manager.insert_structured_fields(
            table_ctrl,
            table_id,
            header_rows,
            footer_rows,
            author_info
        )

        results.append({
            'table_index': i,
            'table_id': table_id,
            'result': result
        })

        # YAML 저장
        yaml_path = f"table_{table_id}_structure.yaml"
        save_data = {
            'author_info': result['author_info'],
            'metadata': result['metadata'],
            'header': {str(k): v for k, v in result['header'].items()},
            'body': {str(k): v for k, v in result['body'].items()},
            'footer': {str(k): v for k, v in result['footer'].items()}
        }

        with open(yaml_path, 'w', encoding='utf-8') as f:
            yaml.dump(save_data, f, allow_unicode=True, default_flow_style=False)

        print(f"  → 완료: {yaml_path}")

    print(f"\n✓ 총 {len(results)}개 테이블 처리 완료\n")
    return results


if __name__ == "__main__":
    import sys

    print()
    print("실행 모드 선택:")
    print("1. 대화형 모드 (각 테이블마다 헤더/푸터 행 수 입력)")
    print("2. 빠른 일괄 모드 (모든 테이블에 동일 설정 적용)")
    print()

    choice = input("선택 (1 또는 2, 기본값=1): ").strip()

    if choice == "2":
        # 빠른 모드
        print("\n[빠른 일괄 모드]")
        header = input("모든 테이블 헤더 행 수 (기본값=1): ").strip()
        header_rows = int(header) if header else 1

        footer = input("모든 테이블 푸터 행 수 (기본값=0): ").strip()
        footer_rows = int(footer) if footer else 0

        author = input("작성자명 (선택): ").strip()
        desc = input("문서 설명 (선택): ").strip()

        print()
        quick_insert_all(header_rows, footer_rows, author, desc)
    else:
        # 대화형 모드
        auto_insert_fields()

    print("프로그램 종료")
