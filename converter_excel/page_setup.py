# -*- coding: utf-8 -*-
"""엑셀 페이지 설정 추출 모듈"""

from openpyxl import load_workbook
from dataclasses import dataclass
from typing import Optional


@dataclass
class PageSettings:
    """페이지 설정 정보"""
    fit_to_page: bool          # 페이지 맞춤 사용 여부
    fit_to_width: Optional[int]   # 페이지 폭 (1 = 1페이지)
    fit_to_height: Optional[int]  # 페이지 높이 (0 = 자동)
    scale: Optional[int]          # 배율 (%)
    orientation: str              # portrait/landscape
    paper_size: Optional[int]     # 용지 크기


def get_page_settings(file_path: str, sheet_name: str = None) -> PageSettings:
    """
    엑셀 파일의 페이지 설정 정보 추출

    Args:
        file_path: 엑셀 파일 경로
        sheet_name: 시트 이름 (None이면 활성 시트)

    Returns:
        PageSettings 객체
    """
    wb = load_workbook(file_path)

    if sheet_name:
        ws = wb[sheet_name]
    else:
        ws = wb.active

    ps = ws.page_setup
    sp = ws.sheet_properties.pageSetUpPr

    fit_to_page = sp.fitToPage if sp else False

    settings = PageSettings(
        fit_to_page=fit_to_page,
        fit_to_width=ps.fitToWidth,
        fit_to_height=ps.fitToHeight,
        scale=ps.scale,
        orientation=ps.orientation or 'portrait',
        paper_size=ps.paperSize
    )

    wb.close()
    return settings


def is_fit_to_one_page_wide(file_path: str, sheet_name: str = None) -> bool:
    """
    1페이지 폭에 맞춤 설정 여부 확인

    Args:
        file_path: 엑셀 파일 경로
        sheet_name: 시트 이름 (None이면 활성 시트)

    Returns:
        True면 1페이지 폭에 맞춤 설정됨
    """
    settings = get_page_settings(file_path, sheet_name)
    return settings.fit_to_page and settings.fit_to_width == 1


if __name__ == "__main__":
    # 테스트
    import sys

    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        file_path = r"C:\win32hwp\test_output.xlsx"

    settings = get_page_settings(file_path)

    print(f"파일: {file_path}")
    print(f"\n=== 페이지 설정 ===")
    print(f"  fit_to_page: {settings.fit_to_page}")
    print(f"  fit_to_width: {settings.fit_to_width}")
    print(f"  fit_to_height: {settings.fit_to_height}")
    print(f"  scale: {settings.scale}")
    print(f"  orientation: {settings.orientation}")
    print(f"  paper_size: {settings.paper_size}")

    print(f"\n=== 분석 ===")
    if is_fit_to_one_page_wide(file_path):
        print("  --> [1페이지 폭에 맞춤] 설정됨")
    elif settings.fit_to_page:
        print(f"  --> [{settings.fit_to_width}페이지 폭] 설정")
    else:
        print(f"  --> 배율 {settings.scale}% 사용")
