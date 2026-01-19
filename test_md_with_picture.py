# md_content.md 파일을 한글에 적용
from table.table_md_2_hwp import markdown_to_hwp, get_hwp

hwp = get_hwp()
if not hwp:
    print('HWP 연결 실패')
    exit(1)

# md_content.md 읽기
with open('C:/win32hwp/table/md_content.md', 'r', encoding='utf-8') as f:
    content = f.read()

# 한글에 적용
markdown_to_hwp(content)
