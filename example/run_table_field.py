# 테이블 필드 관리 테스트
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor import get_hwp_instance
from table import TableField


def test_field_operations():
    """필드 CRUD 작업 테스트"""
    print("TableField 테스트")
    print("=" * 60)

    hwp = get_hwp_instance()
    if not hwp:
        print("한글이 실행 중이 아닙니다.")
        return

    tf = TableField(hwp, debug=True)

    # 1. 테이블 진입 시도
    print("\n[1] 테이블 진입 시도...")
    if not tf.enter_table(0):
        print("  테이블이 없습니다. 필드 조회만 수행합니다.")
        print("\n[문서 내 모든 필드 조회]")
        all_fields = tf.get_all_fields()
        if all_fields:
            for f in all_fields:
                print(f"  - {f.name}: list_id={f.list_id}, text='{f.text}'")
        else:
            print("  필드가 없습니다.")
        return

    # 2. 테이블 정보
    print("\n[2] 테이블 정보")
    size = tf.get_table_size()
    print(f"  크기: {size['rows']}행 x {size['cols']}열")

    # 3. 필드 조회
    print("\n[3] 필드 조회")
    summary = tf.get_field_summary()
    print(f"  총 필드: {summary['total']}개")
    print(f"  테이블 내: {summary['in_table']}개")

    if summary['fields']:
        for f in summary['fields']:
            coord = f"({f.row}, {f.col})" if f.row >= 0 else "(문서)"
            print(f"  - {f.name}: {coord}, text='{f.text[:20]}...' " if len(f.text) > 20 else f"  - {f.name}: {coord}, text='{f.text}'")

    # 4. 필드 맵 출력
    print("\n[4] 필드 맵")
    tf.print_field_map()

    # 5. 셀 필드 vs 누름틀 필드
    print("\n[5] 필드 타입별 조회")
    cell_fields = tf.get_cell_fields()
    clickhere_fields = tf.get_clickhere_fields()
    print(f"  셀 필드: {len(cell_fields)}개")
    print(f"  누름틀 필드: {len(clickhere_fields)}개")


def test_field_create():
    """필드 생성 테스트"""
    print("\n필드 생성 테스트")
    print("=" * 60)

    hwp = get_hwp_instance()
    if not hwp:
        print("한글이 실행 중이 아닙니다.")
        return

    tf = TableField(hwp, debug=True)

    if not tf.enter_table(0):
        print("테이블이 없습니다.")
        return

    # (0, 0) 셀에 필드 생성
    print("\n[1] (0, 0) 셀에 필드 생성...")
    result = tf.create_field_at_coord(0, 0, "test_field", "입력하세요", "테스트 필드입니다")
    print(f"  결과: {'성공' if result else '실패'}")

    # 필드 확인
    print("\n[2] 필드 확인...")
    if tf.field_exists("test_field"):
        field = tf.get_field_by_name("test_field")
        if field:
            print(f"  필드 발견: {field.name} at ({field.row}, {field.col})")


def test_field_update():
    """필드 수정 테스트"""
    print("\n필드 수정 테스트")
    print("=" * 60)

    hwp = get_hwp_instance()
    if not hwp:
        print("한글이 실행 중이 아닙니다.")
        return

    tf = TableField(hwp, debug=True)

    # 모든 필드 조회
    all_fields = tf.get_all_fields()
    if not all_fields:
        print("필드가 없습니다.")
        return

    first_field = all_fields[0]
    print(f"\n첫 번째 필드: {first_field.name}")
    print(f"  현재 텍스트: '{first_field.text}'")

    # 텍스트 변경
    new_text = "테스트 입력값"
    print(f"\n[1] 텍스트 변경: '{new_text}'")
    tf.put_field_text(first_field.name, new_text)

    # 확인
    updated = tf.get_field_by_name(first_field.name)
    if updated:
        print(f"  변경 후: '{updated.text}'")


if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1:
        cmd = sys.argv[1]
        if cmd == "create":
            test_field_create()
        elif cmd == "update":
            test_field_update()
        else:
            test_field_operations()
    else:
        test_field_operations()
