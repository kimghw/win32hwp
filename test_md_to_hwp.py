# 마크다운 → 한글 문서 변환 테스트
import os
import win32com.client as win32
from table_cell_generate import markdown_to_hwp, set_hwp

# 마크다운 파일 경로
MD_FILE = os.path.join(os.path.dirname(__file__), "test_content.md")

if __name__ == "__main__":
    print("=== 마크다운 → 한글 문서 삽입 테스트 ===\n")

    # 마크다운 파일 읽기
    print(f"마크다운 파일 읽기: {MD_FILE}")
    with open(MD_FILE, "r", encoding="utf-8") as f:
        test_markdown = f.read()

    # 한글 새 문서 열기
    print("한글 새 문서 열기...")
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
    hwp.XHwpWindows.Item(0).Visible = True
    set_hwp(hwp)

    print("[삽입할 마크다운]")
    print(test_markdown)
    print("\n" + "="*50)
    print("한글 문서에 삽입 중...\n")

    result = markdown_to_hwp(test_markdown)

    if result:
        print("\n삽입 완료!")
    else:
        print("\n삽입 실패!")
