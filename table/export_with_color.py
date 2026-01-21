# -*- coding: utf-8 -*-
"""한글 표를 배경색 포함하여 엑셀로 변환"""

import sys
import os
import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from cursor import get_hwp_instance
from table.table_excel_converter import TableExcelConverter

def main():
    hwp = get_hwp_instance()
    if not hwp:
        print('[오류] 한글이 실행 중이지 않습니다.')
        return

    converter = TableExcelConverter(hwp, debug=False)

    # 텍스트와 배경색 포함하여 엑셀로 저장
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    filepath = f'C:\\win32hwp\\table_with_color_{timestamp}.xlsx'

    print("한글 표를 엑셀로 변환 중...")
    converter.to_excel(filepath, with_text=True, show_cell_info=False)
    print(f'\n파일 저장 완료: {filepath}')
    print('엑셀에서 열어 배경색을 확인하세요.')

if __name__ == "__main__":
    main()
