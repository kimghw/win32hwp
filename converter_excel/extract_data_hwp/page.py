# -*- coding: utf-8 -*-
"""한글 페이지 정보 추출 모듈

페이지 여백, 용지 크기, 방향 등의 정보를 추출합니다.
Excel/openpyxl 관련 코드는 포함하지 않습니다.
"""

from dataclasses import dataclass
from typing import Optional

from ..page_meta import HwpPageMeta


@dataclass
class PageMatchResult:
    """페이지 매칭 결과"""
    page_meta: HwpPageMeta = None
    success: bool = False
    error: str = None


def extract_page_info(hwp) -> PageMatchResult:
    """한글에서 페이지 정보 추출

    Args:
        hwp: HWP 객체

    Returns:
        PageMatchResult 객체
    """
    result = PageMatchResult()

    try:
        meta = HwpPageMeta()

        act = hwp.CreateAction("PageSetup")
        pset = act.CreateSet()
        act.GetDefault(pset)
        page_def = pset.Item("PageDef")

        # 용지 크기
        meta.page_size.width = page_def.Item("PaperWidth")
        meta.page_size.height = page_def.Item("PaperHeight")
        meta.page_size.orientation = 'landscape' if page_def.Item("Landscape") else 'portrait'

        # 여백
        meta.margin.left = page_def.Item("LeftMargin")
        meta.margin.right = page_def.Item("RightMargin")
        meta.margin.top = page_def.Item("TopMargin")
        meta.margin.bottom = page_def.Item("BottomMargin")
        meta.margin.header = page_def.Item("HeaderLen")
        meta.margin.footer = page_def.Item("FooterLen")
        meta.margin.gutter = page_def.Item("GutterLen")

        # 본문 영역 계산
        meta.calculate_content_area()

        result.page_meta = meta
        result.success = True

    except Exception as e:
        result.error = str(e)
        result.success = False

    return result


# ============================================================
# 테스트
# ============================================================

if __name__ == "__main__":
    import sys
    import os
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

    from cursor import get_hwp_instance
    from converter_excel.page_meta import Unit

    hwp = get_hwp_instance()
    if not hwp:
        print("[오류] 한글이 실행 중이지 않습니다.")
        exit(1)

    print("=== 페이지 정보 추출 테스트 ===\n")

    result = extract_page_info(hwp)

    if result.success:
        meta = result.page_meta
        print(f"용지 크기: {meta.page_size.width} x {meta.page_size.height} HWPUNIT")
        print(f"          {Unit.hwpunit_to_cm(meta.page_size.width):.1f} x {Unit.hwpunit_to_cm(meta.page_size.height):.1f} cm")
        print(f"용지 방향: {meta.page_size.orientation}")
        print(f"\n여백 (cm):")
        print(f"  왼쪽: {Unit.hwpunit_to_cm(meta.margin.left):.2f}")
        print(f"  오른쪽: {Unit.hwpunit_to_cm(meta.margin.right):.2f}")
        print(f"  위쪽: {Unit.hwpunit_to_cm(meta.margin.top):.2f}")
        print(f"  아래쪽: {Unit.hwpunit_to_cm(meta.margin.bottom):.2f}")
        print(f"\n본문 영역: {Unit.hwpunit_to_cm(meta.content_width):.1f} x {Unit.hwpunit_to_cm(meta.content_height):.1f} cm")
    else:
        print(f"[오류] {result.error}")
