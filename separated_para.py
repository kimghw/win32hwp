"""
분리된 문단 처리 - 문단이 두 페이지에 걸쳐있을 때 한 페이지로 맞춤

사용법:
    from separated_para import SeparatedPara, get_hwp_instance

    hwp = get_hwp_instance()
    helper = SeparatedPara(hwp)

    # 현재 페이지 문단 정보
    result = helper.get_page_paragraph_count()
    print(f"문단 수: {result['paragraph_count']}")
"""

import json
from datetime import datetime
from pathlib import Path
from typing import Dict, Tuple

from cursor import get_hwp_instance
from separated_word import SeparatedWord


class SeparatedPara:
    """분리된 문단 처리 클래스 - 페이지 걸침 문단 해소"""

    # 클래스 변수: 문단-페이지 매핑 저장소
    para_page_map = {}  # {para_id: {'start_page': int, 'end_page': int, 'is_empty': bool}}

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
                para_id: {'start_page': int, 'end_page': int, 'is_empty': bool},
                ...
            }
        """
        saved_pos = self.hwp.GetPos()
        SeparatedPara.para_page_map = {}

        try:
            # 문서 처음으로 이동
            self.hwp.MovePos(2)  # moveDocBegin

            # 모든 문단 순회
            while True:
                para_id = self._get_current_para_id()
                start_page, end_page = self._get_para_page_range(para_id)
                is_empty = self._is_empty_paragraph()

                # 클래스 변수에 저장
                SeparatedPara.para_page_map[para_id] = {
                    'start_page': start_page,
                    'end_page': end_page,
                    'is_empty': is_empty
                }

                # 다음 문단으로 이동
                before_para = para_id
                self.hwp.HAction.Run("MoveNextParaBegin")
                after_para = self._get_current_para_id()

                if before_para == after_para:
                    break

            # 원래 위치 복원
            self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])

            return SeparatedPara.para_page_map

        except Exception as e:
            self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])
            return SeparatedPara.para_page_map

    def _is_empty_paragraph(self) -> bool:
        """현재 문단이 빈 문단인지 확인"""
        saved_pos = self.hwp.GetPos()

        self.hwp.HAction.Run("MoveParaBegin")
        self.hwp.HAction.Run("MoveParaEnd")
        para_end = self.hwp.GetPos()[2]

        self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])

        # para_end가 1 이하면 빈 문단 (줄바꿈만 있음)
        return para_end <= 1

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

    def fix_word_in_paragraph(self, para_id: int, spacing_step: float = -1.0,
                        min_spacing: float = -100, max_iterations: int = 100) -> Dict:
        """
        특정 문단으로 이동해서 분리단어 처리 실행

        Args:
            para_id: 문단 ID
            spacing_step: 자간 감소 단위
            min_spacing: 최소 자간 값
            max_iterations: 최대 반복 횟수

        Returns:
            SeparatedWord.align_paragraph() 결과
        """
        saved_pos = self.hwp.GetPos()

        # 해당 문단으로 이동
        list_id = saved_pos[0]
        self.hwp.SetPos(list_id, para_id, 0)
        self.hwp.HAction.Run("MoveParaBegin")

        # 분리단어 처리 실행
        align = SeparatedWord(self.hwp, debug=False)
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
        for para_id, info in SeparatedPara.para_page_map.items():
            if info['start_page'] == page:
                result.append(para_id)

        return result

    def fix_paragraph(self, para_id: int, min_font_size: int = 4) -> Dict:
        """
        페이지 걸친 문단 앞의 빈 문단 글자 크기를 줄여서 공간 확보

        Args:
            para_id: 걸친 문단 ID
            min_font_size: 빈 문단의 최소 글자 크기 (pt 단위, 기본 4pt)

        Returns:
            {
                'success': bool,
                'para_id': int,
                'start_page': int,
                'original_end_page': int,
                'final_end_page': int,
                'empty_paras_reduced': list,  # 줄인 빈 문단 ID 목록
                'total_reduction': int  # 총 줄인 pt
            }
        """
        saved_pos = self.hwp.GetPos()
        list_id = saved_pos[0]

        # 해당 문단으로 이동
        self.hwp.SetPos(list_id, para_id, 0)

        # 문단 시작/끝 페이지 확인
        start_page, end_page = self._get_para_page_range(para_id)
        original_end_page = end_page

        # 이미 같은 페이지면 처리 불필요
        if start_page == end_page:
            self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])
            return {
                'success': True,
                'para_id': para_id,
                'start_page': start_page,
                'original_end_page': original_end_page,
                'final_end_page': end_page,
                'empty_paras_reduced': [],
                'total_reduction': 0
            }

        # 같은 페이지(start_page)에 있는 빈 문단 찾기
        empty_paras = []
        for pid, info in SeparatedPara.para_page_map.items():
            if info.get('is_empty', False) and info['start_page'] == start_page:
                if pid < para_id:  # 걸친 문단보다 앞에 있는 빈 문단만
                    empty_paras.append(pid)

        empty_paras.sort()  # 문단 순서대로

        if not empty_paras:
            self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])
            return {
                'success': False,
                'para_id': para_id,
                'start_page': start_page,
                'original_end_page': original_end_page,
                'final_end_page': end_page,
                'empty_paras_reduced': [],
                'total_reduction': 0,
                'error': '빈 문단 없음'
            }

        min_font_hwp = min_font_size * 100  # HWP 단위
        empty_paras_reduced = []
        total_reduction = 0

        # 빈 문단들의 글자 크기를 줄이면서 확인
        for empty_para_id in empty_paras:
            if start_page == end_page:
                break

            # 빈 문단으로 이동
            self.hwp.SetPos(list_id, empty_para_id, 0)
            self.hwp.HAction.Run("MoveParaBegin")

            # 빈 문단 전체 선택 (줄바꿈 문자까지 포함하려면 다음 문단 시작까지 선택)
            self.hwp.HAction.Run("MoveSelNextParaBegin")

            # 현재 글자 크기 가져오기
            pset = self.hwp.HParameterSet.HCharShape
            self.hwp.HAction.GetDefault("CharShape", pset.HSet)
            current_height = pset.Height

            # 글자 크기가 0이면 기본값 1000 (10pt)으로 설정
            if current_height == 0:
                current_height = 1000

            reduced_this_para = 0

            # 글자 크기 줄이기 반복 (시작/끝 페이지가 같아질 때까지)
            while start_page != end_page and current_height > min_font_hwp:
                current_height -= 100  # 1pt 줄이기
                reduced_this_para += 1

                # 빈 문단 전체 선택 후 글자 크기 적용 (줄바꿈 문자까지 포함)
                self.hwp.SetPos(list_id, empty_para_id, 0)
                self.hwp.HAction.Run("MoveParaBegin")
                self.hwp.HAction.Run("MoveSelNextParaBegin")
                pset.Height = current_height
                self.hwp.HAction.Execute("CharShape", pset.HSet)
                self.hwp.HAction.Run("Cancel")

                # 걸친 문단 페이지 다시 확인
                self.hwp.SetPos(list_id, para_id, 0)
                start_page, end_page = self._get_para_page_range(para_id)

            if reduced_this_para > 0:
                empty_paras_reduced.append(empty_para_id)
                total_reduction += reduced_this_para

        # 원래 위치 복원
        self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])

        return {
            'success': start_page == end_page,
            'para_id': para_id,
            'start_page': start_page,
            'original_end_page': original_end_page,
            'final_end_page': end_page,
            'empty_paras_reduced': empty_paras_reduced,
            'total_reduction': total_reduction
        }

    def fix_all_paragraphs(self, page: int = None, min_font_size: int = 4,
                                      max_rounds: int = 50) -> Dict:
        """
        걸친 문단이 없을 때까지 반복 처리

        Args:
            page: 대상 페이지 (None이면 전체)
            min_font_size: 빈 문단 최소 글자 크기
            max_rounds: 최대 반복 횟수

        Returns:
            {
                'rounds': int,  # 반복 횟수
                'processed': int,
                'success': int,
                'failed': int,
                'remaining_spanning': int,  # 남은 걸친 문단 수
                'results': [...]
            }
        """
        processed = 0
        success = 0
        failed = 0
        results = []
        rounds = 0
        failed_paras = set()  # 실패한 문단 기록 (무한 루프 방지)

        while rounds < max_rounds:
            rounds += 1

            # para_page_map 최신 업데이트
            self.ParaAlignWords()

            # 페이지 걸친 문단 찾기 (실패한 것 제외)
            spanning_paras = []
            for para_id, info in SeparatedPara.para_page_map.items():
                if info['start_page'] != info['end_page'] and not info.get('is_empty', False):
                    if page is None or info['start_page'] == page:
                        if para_id not in failed_paras:
                            spanning_paras.append(para_id)

            # 걸친 문단이 없으면 종료
            if not spanning_paras:
                break

            # 첫 번째 걸친 문단 처리
            para_id = spanning_paras[0]
            result = self.fix_paragraph(para_id, min_font_size)
            processed += 1
            results.append(result)

            if result['success']:
                success += 1
            else:
                failed += 1
                failed_paras.add(para_id)  # 실패한 문단 기록

        # 최종 상태 확인
        self.ParaAlignWords()
        remaining = 0
        for para_id, info in SeparatedPara.para_page_map.items():
            if info['start_page'] != info['end_page'] and not info.get('is_empty', False):
                if page is None or info['start_page'] == page:
                    remaining += 1

        return {
            'rounds': rounds,
            'processed': processed,
            'success': success,
            'failed': failed,
            'remaining_spanning': remaining,
            'results': results
        }

    def fix_all_words_in_page(self, page: int, spacing_step: float = -1.0,
                   min_spacing: float = -100, max_iterations: int = 100) -> Dict:
        """
        페이지의 모든 문단에 대해 분리단어 처리 실행

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

        # 각 문단에 대해 fix_word_in_paragraph 실행
        for para_id in para_ids:
            # 빈 문단 스킵
            para_info = SeparatedPara.para_page_map.get(para_id, {})
            if para_info.get('is_empty', False):
                results.append({
                    'para_id': para_id,
                    'adjusted': 0,
                    'skipped': 0,
                    'failed': 0,
                    'empty': True
                })
                continue

            result = self.fix_word_in_paragraph(
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
                'failed': result.get('failed_lines', 0),
                'empty': False
            })

        return {
            'page': page,
            'total_paragraphs': len(para_ids),
            'total_adjusted': total_adjusted,
            'total_skipped': total_skipped,
            'total_failed': total_failed,
            'results': results
        }

    def get_spanning_lines(self, para_id: int) -> Dict:
        """
        걸친 문단의 페이지별 줄 수 분석

        Args:
            para_id: 문단 ID

        Returns:
            {
                'para_id': int,
                'start_page': int,
                'end_page': int,
                'is_spanning': bool,
                'lines_per_page': {page: line_count, ...},
                'total_lines': int
            }
        """
        saved_pos = self.hwp.GetPos()
        list_id = saved_pos[0]

        # 해당 문단으로 이동
        self.hwp.SetPos(list_id, para_id, 0)

        # 문단의 시작/끝 페이지 확인
        start_page, end_page = self._get_para_page_range(para_id)
        is_spanning = (start_page != end_page)

        lines_per_page = {}
        total_lines = 0

        # 문단의 각 줄을 순회하며 페이지 확인
        self.hwp.HAction.Run("MoveParaBegin")

        while True:
            current_page = self._get_current_page()
            lines_per_page[current_page] = lines_per_page.get(current_page, 0) + 1
            total_lines += 1

            # 다음 줄로 이동
            before_pos = self.hwp.GetPos()
            self.hwp.HAction.Run("MoveDown")
            after_pos = self.hwp.GetPos()

            # 문단을 벗어났는지 확인
            if after_pos[1] != para_id:
                break

            # 더 이상 이동하지 않으면 종료
            if before_pos == after_pos:
                break

        # 원래 위치 복원
        self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])

        return {
            'para_id': para_id,
            'start_page': start_page,
            'end_page': end_page,
            'is_spanning': is_spanning,
            'lines_per_page': lines_per_page,
            'total_lines': total_lines
        }

    def get_page_last_spanning_para(self, page: int) -> Dict:
        """
        페이지의 마지막 문단이 걸쳐있는지 확인

        Args:
            page: 페이지 번호

        Returns:
            {
                'has_spanning': bool,
                'para_id': int or None,
                'lines_info': get_spanning_lines 결과 or None
            }
        """
        # para_page_map 최신화
        self.ParaAlignWords()

        # 해당 페이지에서 시작하는 문단 중 마지막 문단 찾기
        last_para_id = None
        for para_id, info in SeparatedPara.para_page_map.items():
            if info['start_page'] == page:
                if last_para_id is None or para_id > last_para_id:
                    last_para_id = para_id

        if last_para_id is None:
            return {
                'has_spanning': False,
                'para_id': None,
                'lines_info': None
            }

        # 마지막 문단의 정보 확인
        para_info = SeparatedPara.para_page_map[last_para_id]

        if para_info['start_page'] == para_info['end_page']:
            return {
                'has_spanning': False,
                'para_id': last_para_id,
                'lines_info': None
            }

        # 걸친 문단의 줄 분포 분석
        lines_info = self.get_spanning_lines(last_para_id)

        return {
            'has_spanning': True,
            'para_id': last_para_id,
            'lines_info': lines_info
        }

    def reduce_empty_para_font_size(self, page: int, para_id: int, min_font_size: int = 4) -> Dict:
        """
        걸친 문단 앞의 빈 문단 글자 크기를 1pt씩 줄이기

        Args:
            page: 걸친 문단이 있는 페이지
            para_id: 걸친 문단 ID
            min_font_size: 최소 글자 크기 (pt)

        Returns:
            {
                'reduced_paras': [para_id, ...],
                'total_reduction': int  # 총 줄인 pt 수
            }
        """
        saved_pos = self.hwp.GetPos()
        list_id = saved_pos[0]

        # 같은 페이지의 빈 문단 찾기 (걸친 문단보다 앞)
        empty_paras = []
        for pid, info in SeparatedPara.para_page_map.items():
            if info.get('is_empty', False) and info['start_page'] == page:
                if pid < para_id:
                    empty_paras.append(pid)

        empty_paras.sort()

        if not empty_paras:
            self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])
            return {'reduced_paras': [], 'total_reduction': 0}

        min_font_hwp = min_font_size * 100
        reduced_paras = []
        total_reduction = 0

        # 각 빈 문단의 글자 크기를 1pt씩 줄이기
        for empty_para_id in empty_paras:
            self.hwp.SetPos(list_id, empty_para_id, 0)
            self.hwp.HAction.Run("MoveParaBegin")

            # 빈 문단은 줄바꿈 문자만 있으므로 다음 문단 시작까지 선택해야 함
            # MoveSelNextParaBegin 또는 MoveRight로 줄바꿈 문자까지 포함
            self.hwp.HAction.Run("MoveSelNextParaBegin")

            pset = self.hwp.HParameterSet.HCharShape
            self.hwp.HAction.GetDefault("CharShape", pset.HSet)
            current_height = pset.Height

            if current_height == 0:
                current_height = 1000  # 기본값 10pt

            if current_height > min_font_hwp:
                pset.Height = current_height - 100  # 1pt 감소
                self.hwp.HAction.Execute("CharShape", pset.HSet)
                reduced_paras.append(empty_para_id)
                total_reduction += 1

            self.hwp.HAction.Run("Cancel")

        self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])

        return {
            'reduced_paras': reduced_paras,
            'total_reduction': total_reduction
        }

    def get_last_line_remaining_chars(self, para_id: int) -> Dict:
        """
        걸친 문단의 마지막 줄(다음 페이지에 있는 부분)의 글자 수 확인

        Args:
            para_id: 문단 ID

        Returns:
            {
                'para_id': int,
                'total_lines': int,
                'remaining_chars': int,  # 다음 페이지에 있는 글자 수
                'should_try_spacing': bool  # remaining_chars < (total_lines * 2 + 1) 여부
            }
        """
        saved_pos = self.hwp.GetPos()
        list_id = saved_pos[0]

        # 문단으로 이동
        self.hwp.SetPos(list_id, para_id, 0)

        # 문단 시작/끝 페이지 확인
        start_page, end_page = self._get_para_page_range(para_id)

        if start_page == end_page:
            self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])
            return {
                'para_id': para_id,
                'total_lines': 0,
                'remaining_chars': 0,
                'should_try_spacing': False
            }

        # 줄 분포 분석
        lines_info = self.get_spanning_lines(para_id)
        total_lines = lines_info['total_lines']

        # 다음 페이지의 텍스트 길이 계산
        self.hwp.SetPos(list_id, para_id, 0)
        self.hwp.HAction.Run("MoveParaBegin")

        # 다음 페이지로 넘어가는 지점 찾기
        remaining_chars = 0
        found_next_page = False

        while True:
            current_page = self._get_current_page()

            if current_page > start_page and not found_next_page:
                found_next_page = True
                # 여기서부터 문단 끝까지 선택
                self.hwp.HAction.Run("MoveSelParaEnd")
                text = self.hwp.GetTextFile("TEXT", "saveblock")
                self.hwp.HAction.Run("Cancel")
                if text:
                    remaining_chars = len(text.strip())
                break

            before_pos = self.hwp.GetPos()
            self.hwp.HAction.Run("MoveRight")
            after_pos = self.hwp.GetPos()

            # 문단을 벗어났거나 더 이상 이동 안되면
            if after_pos[1] != para_id or before_pos == after_pos:
                break

        self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])

        # 조건 확인: remaining_chars < (total_lines * 2 + 1)
        threshold = total_lines * 2 + 1
        should_try = remaining_chars > 0 and remaining_chars < threshold

        return {
            'para_id': para_id,
            'total_lines': total_lines,
            'remaining_chars': remaining_chars,
            'threshold': threshold,
            'should_try_spacing': should_try
        }

    def try_char_spacing_align(self, para_id: int, max_attempts: int = 2) -> Dict:
        """
        글자 간격을 줄이고 분리단어 처리 적용 (우선 전략)

        마지막 줄의 남은 글자 수가 (줄 수 * 2 + 1)보다 작으면
        글자 간격을 1 줄이고 분리단어 처리를 적용. 최대 2회 반복.

        Args:
            para_id: 문단 ID
            max_attempts: 최대 시도 횟수 (기본 2)

        Returns:
            {
                'applied': bool,
                'attempts': int,
                'success': bool  # 걸침 해소 여부
            }
        """
        saved_pos = self.hwp.GetPos()
        list_id = saved_pos[0]

        # 먼저 조건 확인
        check_result = self.get_last_line_remaining_chars(para_id)

        if not check_result['should_try_spacing']:
            self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])
            return {
                'applied': False,
                'attempts': 0,
                'success': False,
                'reason': f"조건 미충족: remaining={check_result['remaining_chars']}, threshold={check_result.get('threshold', 0)}"
            }

        attempts = 0
        success = False

        for i in range(max_attempts):
            attempts += 1

            # 문단으로 이동하여 전체 선택
            self.hwp.SetPos(list_id, para_id, 0)
            self.hwp.HAction.Run("MoveParaBegin")
            self.hwp.HAction.Run("MoveSelParaEnd")

            # 현재 자간 가져오기
            pset = self.hwp.HParameterSet.HCharShape
            self.hwp.HAction.GetDefault("CharShape", pset.HSet)

            # 자간 1 줄이기
            pset.SpacingHangul = pset.SpacingHangul - 1
            pset.SpacingLatin = pset.SpacingLatin - 1
            self.hwp.HAction.Execute("CharShape", pset.HSet)
            self.hwp.HAction.Run("Cancel")

            # 분리단어 처리 적용
            self.fix_word_in_paragraph(para_id)

            # 걸침 해소 확인
            start_page, end_page = self._get_para_page_range(para_id)

            if start_page == end_page:
                success = True
                break

        self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])

        return {
            'applied': True,
            'attempts': attempts,
            'success': success
        }

    def remove_empty_line_at_page_start(self, page: int) -> Dict:
        """
        지정된 페이지 시작 부분의 빈 줄(빈 문단) 제거

        Args:
            page: 대상 페이지 번호

        Returns:
            {
                'removed': bool,
                'para_id': int or None,
                'message': str
            }
        """
        saved_pos = self.hwp.GetPos()

        # 해당 페이지로 이동
        # MovePos로 특정 페이지로 이동하기 어려우므로 문단 순회로 찾기
        self.ParaAlignWords()

        # 해당 페이지에서 시작하는 첫 번째 문단 찾기
        first_para_id = None
        for para_id, info in sorted(SeparatedPara.para_page_map.items()):
            if info['start_page'] == page:
                first_para_id = para_id
                break

        if first_para_id is None:
            self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])
            return {
                'removed': False,
                'para_id': None,
                'message': f'페이지 {page}에 문단 없음'
            }

        # 첫 번째 문단이 빈 문단인지 확인
        para_info = SeparatedPara.para_page_map.get(first_para_id, {})

        if not para_info.get('is_empty', False):
            self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])
            return {
                'removed': False,
                'para_id': first_para_id,
                'message': '첫 문단이 빈 문단 아님'
            }

        # 빈 문단 삭제
        list_id = saved_pos[0]
        self.hwp.SetPos(list_id, first_para_id, 0)
        self.hwp.HAction.Run("MoveParaBegin")

        # DeleteLine 액션으로 한 줄 삭제 (빈 문단 전체 삭제)
        self.hwp.HAction.Run("DeleteLine")

        self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])

        return {
            'removed': True,
            'para_id': first_para_id,
            'message': f'빈 문단 para_id={first_para_id} 삭제됨'
        }

    def fix_page(self, page: int, max_iterations: int = 50,
                          strategy: str = 'empty_font', log_callback=None) -> Dict:
        """
        페이지 걸친 문단 처리 (메인 함수)

        전략 우선순위:
        1. char_spacing_align: 남은 글자가 적으면 자간 줄이고 분리단어 처리 (2회)
        2. empty_font: 빈 문단 글자 크기 줄이기
        3. line_spacing: 줄 간격 줄이기 (미구현)
        4. font_size: 문단 글자 크기 줄이기 (미구현)

        Args:
            page: 대상 페이지
            max_iterations: 최대 반복 횟수
            strategy: 전략 ('empty_font', 'line_spacing', 'font_size', 'char_spacing')
            log_callback: 로그 출력 함수 (선택)

        Returns:
            {
                'success': bool,
                'iterations': int,
                'para_id': int,
                'original_lines_info': dict,
                'final_lines_info': dict
            }
        """
        def log(msg):
            if log_callback:
                log_callback(msg)

        # 1. para_page_map 조회
        self.ParaAlignWords()
        log(f"[1] para_page_map 조회 완료: {len(SeparatedPara.para_page_map)}개 문단")

        # 2. 페이지 마지막 걸친 문단 확인
        spanning_info = self.get_page_last_spanning_para(page)
        log(f"[2] 페이지 {page} 마지막 문단 확인: para_id={spanning_info['para_id']}, 걸침={spanning_info['has_spanning']}")

        if not spanning_info['has_spanning']:
            return {
                'success': True,
                'iterations': 0,
                'para_id': spanning_info['para_id'],
                'message': '걸친 문단 없음'
            }

        para_id = spanning_info['para_id']
        original_lines_info = spanning_info['lines_info']
        log(f"[3] 걸친 문단: para_id={para_id}, 줄분포={original_lines_info['lines_per_page']}")

        # 우선 전략: char_spacing_align 시도
        log(f"[3.5] 우선 전략 'char_spacing_align' 시도")
        spacing_result = self.try_char_spacing_align(para_id, max_attempts=2)

        if spacing_result['applied']:
            log(f"    자간+분리단어처리 적용: {spacing_result['attempts']}회 시도, 성공={spacing_result['success']}")

            if spacing_result['success']:
                # 다음 페이지 첫 줄 빈 문단 제거
                next_page = page + 1
                remove_result = self.remove_empty_line_at_page_start(next_page)
                if remove_result['removed']:
                    log(f"    [후처리] 페이지 {next_page} 첫 빈줄 제거: {remove_result['message']}")

                final_lines_info = self.get_spanning_lines(para_id)
                return {
                    'success': True,
                    'iterations': spacing_result['attempts'],
                    'para_id': para_id,
                    'original_lines_info': original_lines_info,
                    'final_lines_info': final_lines_info,
                    'strategy_used': 'char_spacing_align',
                    'empty_line_removed': remove_result['removed']
                }
        else:
            log(f"    자간 전략 스킵: {spacing_result.get('reason', '')}")

        # 4. 선택된 전략으로 반복 처리
        log(f"[4] 전략 '{strategy}'로 반복 처리 시작 (최대 {max_iterations}회)")

        iterations = 0
        success = False

        for i in range(max_iterations):
            iterations += 1

            # 현재 줄 분포 확인
            current_lines_info = self.get_spanning_lines(para_id)
            log(f"  [반복 {iterations}] 현재 줄분포: {current_lines_info['lines_per_page']}")

            # 걸침 해소 확인
            if not current_lines_info['is_spanning']:
                success = True
                log(f"  -> 성공! 더 이상 걸쳐있지 않음")
                break

            # 전략별 처리
            if strategy == 'empty_font':
                result = self.reduce_empty_para_font_size(page, para_id)
                if result['total_reduction'] > 0:
                    log(f"    빈문단 글자크기 감소: {result['reduced_paras']}, 총 {result['total_reduction']}pt 감소")
                else:
                    log(f"    빈문단 글자크기 더 이상 줄일 수 없음")
                    break
            else:
                log(f"    [WARNING] 미지원 전략: {strategy}")
                break

        # 최종 상태 확인
        final_lines_info = self.get_spanning_lines(para_id)
        log(f"[5] 최종: 성공={success}, 반복={iterations}회, 줄분포={final_lines_info['lines_per_page']}")

        # 성공 시 다음 페이지 첫 줄 빈 문단 제거
        empty_line_removed = False
        if success:
            next_page = page + 1
            remove_result = self.remove_empty_line_at_page_start(next_page)
            empty_line_removed = remove_result['removed']
            if empty_line_removed:
                log(f"[6] 페이지 {next_page} 첫 빈줄 제거: {remove_result['message']}")

        return {
            'success': success,
            'iterations': iterations,
            'para_id': para_id,
            'original_lines_info': original_lines_info,
            'final_lines_info': final_lines_info,
            'strategy_used': strategy,
            'empty_line_removed': empty_line_removed
        }


# ============================================================
# 페이지 걸침 문단 처리 전략 (fix_page 함수)
# ============================================================
#
# 처리 순서:
# 1. para_page_map 조회 (ParaAlignWords)
# 2. 페이지 마지막 문단의 걸침 여부 확인 (get_page_last_spanning_para)
# 3. 걸친 문단이 없으면 종료
#
# 전략 우선순위 (구현 완료):
# ----------------------------------------------------------
# [우선 전략] char_spacing_align (try_char_spacing_align)
#   - 조건: 다음 페이지에 있는 글자 수 < (전체 줄 수 * 2 + 1)
#   - 동작: 문단 전체 자간을 1 줄이고 분리단어 처리 적용
#   - 반복: 최대 2회
#   - 함수: get_last_line_remaining_chars(), try_char_spacing_align()
#
# [전략 1] empty_font (reduce_empty_para_font_size)
#   - 같은 페이지의 빈 문단 글자 크기를 1pt씩 줄임
#   - 최소 4pt까지
#   - 함수: reduce_empty_para_font_size()
#
# [전략 2] line_spacing (미구현)
#   - 문단의 줄 간격을 줄임
#
# [전략 3] font_size (미구현)
#   - 걸친 문단의 글자 크기를 줄임
#
# [후처리] 빈 줄 제거 (remove_empty_line_at_page_start)
#   - 걸침 해소 후 다음 페이지 첫 줄이 빈 문단이면 삭제
#   - HAction.Run("DeleteLine") 사용
#
# 관련 함수:
# - get_spanning_lines(para_id): 문단의 페이지별 줄 수 분석
# - get_page_last_spanning_para(page): 페이지 마지막 걸친 문단 확인
# - remove_empty_line_at_page_start(page): 페이지 첫 빈줄 제거
# - fix_page(page, strategy): 메인 처리 함수
# ============================================================
