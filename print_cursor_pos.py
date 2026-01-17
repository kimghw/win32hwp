"""
현재 커서 위치 정보를 출력하는 유틸리티
"""
from cursor_utils import get_hwp_instance


def print_cursor_position(hwp=None):
    """
    현재 커서의 위치 정보를 출력

    Args:
        hwp: HWP 인스턴스 (None이면 자동으로 가져옴)
    """
    if hwp is None:
        hwp = get_hwp_instance()
        if not hwp:
            print("한글이 실행 중이 아닙니다")
            return None

    # 보안 모듈 등록
    hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')

    # 기본 위치 정보 (GetPos)
    pos = hwp.GetPos()
    list_id, para_id, char_pos = pos

    # 상세 위치 정보 (KeyIndicator)
    # 반환값: (총구역, 현재구역, 페이지, 단, 줄, 칸, 수정모드)
    key = hwp.KeyIndicator()
    total_section = key[0]
    current_section = key[1]
    page = key[2]
    column_num = key[3]  # 단
    line = key[4]
    column = key[5]  # 칸
    insert_mode = '수정' if key[6] else '삽입'

    # 테이블 확인
    try:
        act = hwp.CreateAction("TableCellBorder")
        pset = act.GetDefault("TableCellBorder", hwp.HParameterSet.HTableCellBorder.HSet)
        in_table = act.Execute(pset)
    except:
        in_table = False

    # 결과 딕셔너리
    pos_info = {
        'list_id': list_id,
        'para_id': para_id,
        'char_pos': char_pos,
        'page': page,
        'line': line,
        'column': column,
        'total_section': total_section,
        'current_section': current_section,
        'column_num': column_num,
        'insert_mode': insert_mode,
        'in_table': in_table
    }

    # 출력
    print("=" * 60)
    print("현재 커서 위치 정보")
    print("=" * 60)
    print(f"list_id:         {list_id}")
    print(f"para_id:         {para_id}")
    print(f"char_pos:        {char_pos}")
    print(f"page:            {page}")
    print(f"line:            {line}")
    print(f"column (칸):     {column}")
    print(f"column_num (단): {column_num}")
    print(f"section:         {current_section} / {total_section}")
    print(f"insert_mode:     {insert_mode}")
    print("-" * 60)

    if in_table:
        print("위치 유형:  테이블 내부")
    elif list_id == 0:
        print("위치 유형:  본문")
    elif list_id > 0:
        print(f"위치 유형:  특수 영역 (머리말/꼬리말/텍스트박스 등, list_id={list_id})")

    print("=" * 60)

    return pos_info


if __name__ == "__main__":
    print_cursor_position()
