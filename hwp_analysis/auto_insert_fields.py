# -*- coding: utf-8 -*-
"""
테이블 누름틀 자동 입력 스크립트

실행 중인 한글 문서의 모든 테이블에 누름틀을 자동으로 삽입합니다.

실행 모드:
    1. 대화형 모드: 각 테이블마다 헤더/푸터 행 수 입력
    2. 빠른 일괄 모드: 모든 테이블에 동일 설정 적용
    3. 배치 모드: 대화 없이 자동 실행 + 로그 저장
"""

import sys
import os
from datetime import datetime
import yaml

# 프로젝트 루트를 path에 추가
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from table_manager import TableManager
from cursor import get_hwp_instance


def _count_table_rows(hwp, manager, table_ctrl):
    """테이블 행 개수 세기"""
    manager.move_to_cell(table_ctrl, 0, 0)
    total_rows = 1
    while True:
        prev_para = hwp.GetPos()[1]
        hwp.HAction.Run("TableLowerCell")
        curr_para = hwp.GetPos()[1]
        if curr_para != prev_para:
            total_rows += 1
        else:
            break
    return total_rows


def _save_result_yaml(result, table_id):
    """결과를 YAML 파일로 저장"""
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
    return yaml_path


def _print_summary(results):
    """처리 결과 요약 출력"""
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


def auto_insert_fields():
    """
    대화형 모드: 실행 중인 한글 문서의 모든 테이블에 누름틀 자동 삽입

    Returns:
        list: 처리된 테이블 결과 목록
    """
    print("=" * 60)
    print("  테이블 누름틀 자동 입력 (대화형 모드)")
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
        table_ctrl = manager.get_table_by_index(i)
        table_id = f"TBL{i+1:03d}"

        print(f"{'─' * 60}")
        print(f"테이블 {i+1}/{len(tables)} 처리 중 (ID: {table_id})")
        print(f"{'─' * 60}")

        # 행 개수 세기
        total_rows = _count_table_rows(hwp, manager, table_ctrl)
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
            table_ctrl, table_id, header_rows, footer_rows, author_info
        )

        results.append({
            'table_index': i,
            'table_id': table_id,
            'result': result
        })

        yaml_path = _save_result_yaml(result, table_id)
        print(f"✓ 구조 정보 저장: {yaml_path}")
        print()

    _print_summary(results)
    return results


def quick_insert_all(header_rows=1, footer_rows=0, author_name="", description=""):
    """
    빠른 일괄 삽입: 모든 테이블에 동일한 설정 적용

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

    author_info = None
    if author_name or description:
        author_info = {
            'created_by': author_name or 'Unknown',
            'description': description
        }

    results = []

    for i, table in enumerate(tables):
        table_ctrl = manager.get_table_by_index(i)
        table_id = f"TBL{i+1:03d}"

        print(f"[{i+1}/{len(tables)}] {table_id} 처리 중...")

        result = manager.insert_structured_fields(
            table_ctrl, table_id, header_rows, footer_rows, author_info
        )

        results.append({
            'table_index': i,
            'table_id': table_id,
            'result': result
        })

        yaml_path = _save_result_yaml(result, table_id)
        print(f"  → 완료: {yaml_path}")

    print(f"\n✓ 총 {len(results)}개 테이블 처리 완료\n")
    return results


def batch_insert(header_rows=1, footer_rows=0, author_name="KIM"):
    """
    배치 모드: 대화 없이 자동 실행 + 로그 파일 저장

    Args:
        header_rows: 헤더 행 수 (기본값: 1)
        footer_rows: 푸터 행 수 (기본값: 0)
        author_name: 작성자명 (기본값: KIM)

    Returns:
        list: 처리된 테이블 결과 목록
    """
    # 로그 파일 설정
    log_dir = "debugs/logs"
    os.makedirs(log_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(log_dir, f"auto_insert_{timestamp}.log")

    log = open(log_file, 'w', encoding='utf-8')

    def log_print(msg=""):
        """콘솔과 파일에 동시 출력"""
        print(msg)
        log.write(msg + '\n')
        log.flush()

    log_print("=" * 60)
    log_print("  테이블 누름틀 자동 입력 (배치 모드)")
    log_print("=" * 60)
    log_print(f"실행 시간: {timestamp}")
    log_print(f"로그 파일: {log_file}")
    log_print()

    log_print("[설정]")
    log_print(f"  - 헤더 행: {header_rows}")
    log_print(f"  - 푸터 행: {footer_rows}")
    log_print(f"  - 작성자: {author_name}")
    log_print()

    # 한글 인스턴스 연결
    log_print("[1단계] 한글 인스턴스 연결 중...")
    hwp = get_hwp_instance()
    if not hwp:
        log_print("✗ 오류: 한글이 실행 중이지 않습니다.")
        log_print("  → 한글을 먼저 실행하고 문서를 열어주세요.")
        log.close()
        return None

    log_print("✓ 한글 연결 성공\n")

    # 테이블 스캔
    log_print("[2단계] 문서 내 테이블 스캔 중...")
    manager = TableManager(hwp)
    tables = manager.scan_tables()

    if not tables:
        log_print("✗ 문서에 테이블이 없습니다.")
        log.close()
        return []

    log_print(f"✓ {len(tables)}개의 테이블 발견\n")

    # 테이블 목록 출력
    log_print("[3단계] 발견된 테이블 목록:")
    for i, t in enumerate(tables):
        log_print(f"  {i}. 테이블 (Para={t['para']})")
    log_print()

    author_info = {
        'created_by': author_name,
        'description': 'Auto-generated by batch script'
    }

    # 모든 테이블에 누름틀 삽입
    log_print("[4단계] 모든 테이블에 누름틀 삽입 시작\n")

    results = []

    for i, table in enumerate(tables):
        table_ctrl = manager.get_table_by_index(i)
        table_id = f"TBL{i+1:03d}"

        log_print(f"{'─' * 60}")
        log_print(f"테이블 {i+1}/{len(tables)} 처리 중 (ID: {table_id})")
        log_print(f"{'─' * 60}")

        # stdout을 로그 파일로 임시 리다이렉트
        old_stdout = sys.stdout
        sys.stdout = log

        result = manager.insert_structured_fields(
            table_ctrl, table_id, header_rows, footer_rows, author_info
        )

        sys.stdout = old_stdout

        results.append({
            'table_index': i,
            'table_id': table_id,
            'result': result
        })

        yaml_path = _save_result_yaml(result, table_id)
        log_print(f"✓ 구조 정보 저장: {yaml_path}")
        log_print()

    # 요약
    log_print("=" * 60)
    log_print("  완료 요약")
    log_print("=" * 60)
    log_print(f"✓ 처리된 테이블: {len(results)}개")

    total_fields = 0
    for r in results:
        table_result = r['result']
        count = len(table_result['header']) + len(table_result['body']) + len(table_result['footer'])
        total_fields += count
        log_print(f"  - {r['table_id']}: {count}개 누름틀 삽입")

    log_print(f"\n✓ 총 삽입된 누름틀: {total_fields}개")
    log_print(f"\n로그 저장 위치: {log_file}")
    log_print()

    log.close()
    return results


if __name__ == "__main__":
    print()
    print("실행 모드 선택:")
    print("1. 대화형 모드 (각 테이블마다 헤더/푸터 행 수 입력)")
    print("2. 빠른 일괄 모드 (모든 테이블에 동일 설정 적용)")
    print("3. 배치 모드 (대화 없이 자동 실행 + 로그 저장)")
    print()

    choice = input("선택 (1, 2, 또는 3, 기본값=1): ").strip()

    if choice == "2":
        print("\n[빠른 일괄 모드]")
        header = input("모든 테이블 헤더 행 수 (기본값=1): ").strip()
        header_rows = int(header) if header else 1

        footer = input("모든 테이블 푸터 행 수 (기본값=0): ").strip()
        footer_rows = int(footer) if footer else 0

        author = input("작성자명 (선택): ").strip()
        desc = input("문서 설명 (선택): ").strip()

        print()
        quick_insert_all(header_rows, footer_rows, author, desc)

    elif choice == "3":
        print("\n[배치 모드]")
        header = input("모든 테이블 헤더 행 수 (기본값=1): ").strip()
        header_rows = int(header) if header else 1

        footer = input("모든 테이블 푸터 행 수 (기본값=0): ").strip()
        footer_rows = int(footer) if footer else 0

        author = input("작성자명 (기본값=KIM): ").strip() or "KIM"

        print()
        batch_insert(header_rows, footer_rows, author)

    else:
        auto_insert_fields()

    print("프로그램 종료")
