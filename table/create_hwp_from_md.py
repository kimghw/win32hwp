# -*- coding: utf-8 -*-
"""
마크다운 파일을 읽어서 한글 문서로 저장하는 스크립트
"""

import os
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import win32com.client as win32
from table_md_2_hwp import markdown_to_hwp, set_hwp

# 파일 경로
MD_FILE = os.path.join(os.path.dirname(__file__), "md_content.md")
OUTPUT_FILE = r"C:\win32hwp\table\output.hwp"


def main():
    print("=== 마크다운 → 한글 문서 생성 ===\n")

    # 1. 마크다운 파일 읽기
    print(f"[1] 마크다운 파일 읽기: {MD_FILE}")
    with open(MD_FILE, "r", encoding="utf-8") as f:
        markdown_text = f.read()
    print(f"    {len(markdown_text.splitlines())}줄 읽음\n")

    # 2. 한글 새 문서 열기
    print("[2] 한글 새 문서 열기...")
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
    hwp.XHwpWindows.Item(0).Visible = True
    set_hwp(hwp)
    print("    한글 연결 완료\n")

    # 3. 마크다운 삽입
    print("[3] 마크다운 내용 삽입 중...")
    result = markdown_to_hwp(markdown_text)

    if not result:
        print("    삽입 실패!")
        return
    print()

    # 4. 파일 저장
    print(f"[4] 파일 저장: {OUTPUT_FILE}")
    hwp.SaveAs(OUTPUT_FILE, "HWP", "")
    print("    저장 완료!\n")

    print("=" * 50)
    print(f"생성된 파일: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
