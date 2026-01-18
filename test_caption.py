# 캡션 기능 테스트
from table_cell_generate import insert_picture, get_hwp

hwp = get_hwp()
if hwp:
    insert_picture(hwp, '/mnt/c/win32hwp/test.jpg')
else:
    print('HWP 연결 실패')
