"""문서 내 모든 테이블 순회 및 캡션 출력 테스트"""
from table_info import TableInfo
from cursor_utils import get_hwp_instance


if __name__ == "__main__":
    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        exit(1)

    table = TableInfo(hwp, debug=True)

    print("=== 문서 내 모든 테이블 찾기 ===\n")
    tables = table.find_all_tables()

    print(f"\n총 {len(tables)}개 테이블 발견")

    if tables:
        print("\n=== 캡션 확인 ===")
        for t in tables:
            has_cap = table.has_caption(t['num'])
            caption = table.get_table_caption(t['num']) if has_cap else ""
            print(f"테이블 {t['num']}: list_id={t['first_cell_list_id']}, 캡션={has_cap}, '{caption}'")
