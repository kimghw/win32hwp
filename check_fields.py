# 필드 확인
from cursor_utils import get_hwp_instance
hwp = get_hwp_instance()
if hwp:
    # 셀 필드 조회
    field_list = hwp.GetFieldList(1, 0)
    print(f'셀 필드: {repr(field_list)}')

    # 일반 필드
    field_list = hwp.GetFieldList(0, 0)
    print(f'일반 필드: {repr(field_list)}')

    # 누름틀 필드
    field_list = hwp.GetFieldList(2, 0)
    print(f'누름틀 필드: {repr(field_list)}')
else:
    print('한글 없음')
