"""
HWP 페이지/문단 정보 추출 모듈

현재 페이지의 문단 정보를 수집합니다.

사용법:
    from text_align_page import TextAlignPage, get_hwp_instance

    hwp = get_hwp_instance()
    helper = TextAlignPage(hwp)

    # 현재 페이지 문단 정보
    result = helper.get_page_paragraph_count()
    print(f"문단 수: {result['paragraph_count']}")
"""

import json
from datetime import datetime
from pathlib import Path
from typing import Dict, Tuple

from cursor_position_monitor import get_hwp_instance
from text_align import TextAlign


class TextAlignPage:
    """HWP 페이지/문단 정보 추출 클래스"""

    # 클래스 변수: 문단-페이지 매핑 저장소
    para_page_map = {}  # {para_id: {'start_page': int, 'end_page': int}}

    def __init__(self, hwp, log_dir: str = "debugs/logs"):
        """
        Args:
            hwp: HWP 객체
            log_dir: 로그 파일 저장 디렉토리
        """
        self.hwp = hwp
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(parents=True, exist_ok=True)

    def ParaAlignWords(self) -> Dict:
        """
        모든 문단의 페이지 정보를 수집하여 클래스 변수에 저장

        Returns:
            {
                para_id: {'start_page': int, 'end_page': int},
                ...
            }
        """
        saved_pos = self.hwp.GetPos()
        TextAlignPage.para_page_map = {}

        try:
            # 문서 처음으로 이동
            self.hwp.MovePos(2)  # moveDocBegin

            # 모든 문단 순회
            while True:
                para_id = self._get_current_para_id()
                start_page, end_page = self._get_para_page_range(para_id)

                # 클래스 변수에 저장
                TextAlignPage.para_page_map[para_id] = {
                    'start_page': start_page,
                    'end_page': end_page
                }

                # 다음 문단으로 이동
                before_para = para_id
                self.hwp.HAction.Run("MoveNextParaBegin")
                after_para = self._get_current_para_id()

                if before_para == after_para:
                    break

            # 원래 위치 복원
            self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])

            return TextAlignPage.para_page_map

        except Exception as e:
            self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])
            return TextAlignPage.para_page_map

    def _get_current_page(self) -> int:
        """현재 커서 위치의 페이지 번호 반환"""
        key_info = self.hwp.KeyIndicator()
        # KeyIndicator 반환: (BOOL, seccnt, secno, prnpageno, colno, line, pos, over, ctrlname)
        return key_info[3]

    def _get_current_para_id(self) -> int:
        """현재 커서 위치의 문단 ID 반환"""
        pos = self.hwp.GetPos()
        return pos[1]

    def _get_para_page_range(self, para_id: int) -> Tuple[int, int]:
        """
        문단의 시작/끝 페이지 번호 반환

        Args:
            para_id: 문단 ID

        Returns:
            (시작 페이지, 끝 페이지)
        """
        saved_pos = self.hwp.GetPos()

        # 문단 시작으로 이동
        self.hwp.HAction.Run("MoveParaBegin")
        start_page = self._get_current_page()

        # 문단 끝으로 이동
        self.hwp.HAction.Run("MoveParaEnd")
        end_page = self._get_current_page()

        # 원래 위치로 복원
        self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])

        return (start_page, end_page)

    def get_page_paragraph_count(self, target_page: int = None) -> Dict:
        """
        특정 페이지의 문단 개수 및 정보 반환

        Args:
            target_page: 대상 페이지 번호 (None이면 현재 페이지)

        Returns:
            {
                'page': int,              # 페이지 번호
                'paragraph_count': int,   # 문단 개수
                'paragraphs': [{'para_id': int, 'start_page': int, 'end_page': int}, ...]
            }
        """
        saved_pos = self.hwp.GetPos()

        if target_page is None:
            target_page = self._get_current_page()

        paragraphs = []

        try:
            # 문서 처음으로 이동
            self.hwp.MovePos(2)  # moveDocBegin

            # 문단 순회
            while True:
                para_id = self._get_current_para_id()
                start_page, end_page = self._get_para_page_range(para_id)

                # 문단 시작이 target_page 이후면 종료
                if start_page > target_page:
                    break

                # 시작 페이지 기준: start_page == target_page 인 문단만 포함
                if start_page == target_page:
                    paragraphs.append({
                        'para_id': para_id,
                        'start_page': start_page,
                        'end_page': end_page
                    })

                # 다음 문단으로 이동
                before_para = para_id
                self.hwp.HAction.Run("MoveNextParaBegin")
                after_para = self._get_current_para_id()

                if before_para == after_para:
                    break

            # 원래 위치 복원
            self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])

            return {
                'page': target_page,
                'paragraph_count': len(paragraphs),
                'paragraphs': paragraphs
            }

        except Exception as e:
            self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])
            return {
                'page': target_page,
                'paragraph_count': 0,
                'paragraphs': [],
                'error': str(e)
            }

    def align_paragraph(self, para_id: int, spacing_step: float = -1.0,
                        min_spacing: float = -100, max_iterations: int = 100) -> Dict:
        """
        특정 문단으로 이동해서 text_align 실행

        Args:
            para_id: 문단 ID
            spacing_step: 자간 감소 단위
            min_spacing: 최소 자간 값
            max_iterations: 최대 반복 횟수

        Returns:
            text_align.align_paragraph() 결과
        """
        saved_pos = self.hwp.GetPos()

        # 해당 문단으로 이동
        list_id = saved_pos[0]
        self.hwp.SetPos(list_id, para_id, 0)
        self.hwp.HAction.Run("MoveParaBegin")

        # text_align 실행
        align = TextAlign(self.hwp, debug=False)
        result = align.align_paragraph(
            spacing_step=spacing_step,
            min_spacing=min_spacing,
            max_iterations=max_iterations
        )

        return result

    def get_page_paragraphs(self, page: int) -> list:
        """
        페이지에 해당하는 문단 목록 반환 (para_page_map 최신 업데이트)

        Args:
            page: 페이지 번호

        Returns:
            [para_id, para_id, ...] 해당 페이지에서 시작하는 문단 ID 목록
        """
        # para_page_map 최신 업데이트
        self.ParaAlignWords()

        # 해당 페이지의 문단 필터링
        result = []
        for para_id, info in TextAlignPage.para_page_map.items():
            if info['start_page'] == page:
                result.append(para_id)

        return result

    def align_page(self, page: int, spacing_step: float = -1.0,
                   min_spacing: float = -100, max_iterations: int = 100) -> Dict:
        """
        페이지의 모든 문단에 대해 text_align 실행

        Args:
            page: 페이지 번호
            spacing_step: 자간 감소 단위
            min_spacing: 최소 자간 값
            max_iterations: 문단당 최대 반복 횟수

        Returns:
            {
                'page': int,
                'total_paragraphs': int,
                'total_adjusted': int,
                'total_skipped': int,
                'total_failed': int,
                'results': [...]
            }
        """
        # 페이지의 문단 목록 가져오기
        para_ids = self.get_page_paragraphs(page)

        total_adjusted = 0
        total_skipped = 0
        total_failed = 0
        results = []

        # 각 문단에 대해 align_paragraph 실행
        for para_id in para_ids:
            result = self.align_paragraph(
                para_id,
                spacing_step=spacing_step,
                min_spacing=min_spacing,
                max_iterations=max_iterations
            )

            total_adjusted += result.get('adjusted_lines', 0)
            total_skipped += result.get('skipped_lines', 0)
            total_failed += result.get('failed_lines', 0)

            results.append({
                'para_id': para_id,
                'adjusted': result.get('adjusted_lines', 0),
                'skipped': result.get('skipped_lines', 0),
                'failed': result.get('failed_lines', 0)
            })

        return {
            'page': page,
            'total_paragraphs': len(para_ids),
            'total_adjusted': total_adjusted,
            'total_skipped': total_skipped,
            'total_failed': total_failed,
            'results': results
        }


def main():
    """CLI 실행 함수"""
    import argparse

    parser = argparse.ArgumentParser(description='HWP 페이지/문단 정보 추출')
    parser.add_argument('--page', type=int, default=None,
                       help='대상 페이지 번호 (기본: 현재 페이지)')
    parser.add_argument('--all', action='store_true',
                       help='전체 문단-페이지 매핑 (ParaAlignWords)')

    args = parser.parse_args()

    # HWP 인스턴스 연결
    hwp = get_hwp_instance()
    if not hwp:
        print("[ERROR] 실행 중인 한글을 찾을 수 없습니다.")
        return

    print("[OK] 한글에 연결되었습니다.")

    helper = TextAlignPage(hwp)

    # --all: 전체 문단-페이지 매핑
    if args.all:
        para_map = helper.ParaAlignWords()
        print(f"\n전체 문단 수: {len(para_map)}개")
        print("-" * 50)
        for para_id, info in para_map.items():
            span = " (걸침)" if info['start_page'] != info['end_page'] else ""
            print(f"  para_id={para_id}, page={info['start_page']}~{info['end_page']}{span}")
        return

    # 기본: 특정 페이지 문단 조회
    result = helper.get_page_paragraph_count(args.page)

    print(f"\n페이지 {result['page']}: {result['paragraph_count']}개 문단")
    print("-" * 50)

    for i, p in enumerate(result['paragraphs'], 1):
        span = " (걸침)" if p['start_page'] != p['end_page'] else ""
        print(f"  {i}. para_id={p['para_id']}, page={p['start_page']}~{p['end_page']}{span}")


if __name__ == '__main__':
    main()
