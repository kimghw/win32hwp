"""
HWP 페이지 단위 텍스트 정렬 모듈

현재 페이지의 모든 문단을 순회하면서 텍스트 정렬을 수행합니다.
문단이 페이지를 걸쳐있으면 다음 페이지의 연결된 문단까지만 처리하고 종료합니다.

사용법:
    from text_align_page import TextAlignPage, get_hwp_instance

    hwp = get_hwp_instance()
    aligner = TextAlignPage(hwp, debug=True)

    # 현재 페이지 정렬
    result = aligner.align_current_page()
    print(f"처리된 문단 수: {result['processed_paragraphs']}")
"""

import time
import json
import traceback
import sys
import os
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from text_align import TextAlign
from cursor_position_monitor import get_hwp_instance

# 모듈 실행 정보 로그
_MODULE_INFO = {
    'file': os.path.abspath(__file__),
    'name': __name__,
    'loaded_at': datetime.now().isoformat()
}
print(f"[MODULE LOAD] {_MODULE_INFO['file']}")
print(f"[MODULE LOAD] __name__={_MODULE_INFO['name']}, loaded_at={_MODULE_INFO['loaded_at']}")


class TextAlignPage:
    """HWP 페이지 단위 텍스트 정렬 클래스"""

    def __init__(self, hwp, debug: bool = False, log_dir: str = "debugs/logs"):
        """
        Args:
            hwp: HWP 객체
            debug: 디버그 모드 (True시 상세 로그 출력)
            log_dir: 로그 파일 저장 디렉토리
        """
        self.hwp = hwp
        self.debug = debug
        self.text_align = TextAlign(hwp, debug=debug, log_dir=log_dir)
        self.log_messages = []
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(parents=True, exist_ok=True)

        # 세션 정보
        self.session_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.session_start = datetime.now()

        # 초기화 로그
        self._log(f"{'=' * 70}")
        self._log(f"[INIT] TextAlignPage 인스턴스 생성")
        self._log(f"[INIT] 실행 파일: {_MODULE_INFO['file']}")
        self._log(f"[INIT] 세션 ID: {self.session_id}")
        self._log(f"[INIT] 디버그 모드: {self.debug}")
        self._log(f"{'=' * 70}")

    def _log(self, message: str, level: str = "INFO"):
        """로그 메시지 출력 및 저장"""
        timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        msg = f"[{timestamp}] [{level}] {message}"
        if self.debug:
            print(msg)
        self.log_messages.append({
            'timestamp': timestamp,
            'level': level,
            'message': message
        })

    def _get_current_page(self) -> int:
        """현재 커서 위치의 페이지 번호 반환"""
        try:
            key_info = self.hwp.KeyIndicator()
            # KeyIndicator 반환: (BOOL, seccnt, secno, prnpageno, colno, line, pos, over, ctrlname)
            # index:              [0]   [1]     [2]    [3]        [4]    [5]   [6]  [7]   [8]
            self._log(f"[PAGE] KeyIndicator: page={key_info[3]}, line={key_info[5]}")
            page = key_info[3]
            return page
        except Exception as e:
            self._log(f"[PAGE] 페이지 정보 가져오기 실패: {e}", "ERROR")
            return 0

    def _get_current_para_id(self) -> int:
        """현재 커서 위치의 문단 ID 반환"""
        pos = self.hwp.GetPos()
        return pos[1]

    def _move_to_para_begin(self):
        """현재 문단의 시작으로 이동"""
        self.hwp.HAction.Run("MoveParaBegin")
        self._log(f"[MOVE] 문단 시작으로 이동")

    def _move_to_para_end(self):
        """현재 문단의 끝으로 이동"""
        self.hwp.HAction.Run("MoveParaEnd")
        self._log(f"[MOVE] 문단 끝으로 이동")

    def _move_to_next_para(self) -> bool:
        """
        다음 문단으로 이동

        Returns:
            True: 이동 성공, False: 더 이상 문단 없음
        """
        before_para = self._get_current_para_id()
        before_page = self._get_current_page()
        self.hwp.HAction.Run("MoveNextParaBegin")
        after_para = self._get_current_para_id()
        after_page = self._get_current_page()

        if before_para == after_para:
            self._log(f"[MOVE] 다음 문단 없음 (마지막 문단, page={before_page})")
            return False

        self._log(f"[MOVE] 다음 문단: para {before_para} -> {after_para}, page {before_page} -> {after_page}")
        return True

    def _move_to_prev_para(self) -> bool:
        """
        이전 문단으로 이동

        Returns:
            True: 이동 성공, False: 더 이상 문단 없음
        """
        before_para = self._get_current_para_id()
        before_page = self._get_current_page()
        self.hwp.HAction.Run("MovePrevParaBegin")
        after_para = self._get_current_para_id()
        after_page = self._get_current_page()

        if before_para == after_para:
            self._log(f"[MOVE] 이전 문단 없음 (첫 문단, page={before_page})")
            return False

        self._log(f"[MOVE] 이전 문단: para {before_para} -> {after_para}, page {before_page} -> {after_page}")
        return True

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
        self._move_to_para_begin()
        start_page = self._get_current_page()

        # 문단 끝으로 이동
        self._move_to_para_end()
        end_page = self._get_current_page()

        # 원래 위치로 복원
        self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])

        self._log(f"[PARA] 문단 {para_id} 페이지 범위: {start_page} ~ {end_page}")
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
                'paragraphs': [{'para_id': int, 'page': int}, ...]
            }
        """
        saved_pos = self.hwp.GetPos()

        if target_page is None:
            target_page = self._get_current_page()

        self._log(f"[COUNT] 페이지 {target_page} 문단 개수 조회")

        paragraphs = []

        try:
            # 문서 처음으로 이동
            self.hwp.MovePos(2)  # moveDocBegin
            self._log(f"[COUNT] 문서 처음으로 이동")

            # 문단 순회
            while True:
                para_id = self._get_current_para_id()
                start_page, end_page = self._get_para_page_range(para_id)

                self._log(f"[COUNT] para_id={para_id}, page={start_page}~{end_page}")

                # 문단 시작이 target_page 이후면 종료
                if start_page > target_page:
                    self._log(f"[COUNT] 종료: 문단 시작 {start_page} > {target_page}")
                    break

                # 문단이 target_page에 포함되는지 (start_page <= target_page <= end_page)
                if end_page >= target_page:
                    paragraphs.append({'para_id': para_id, 'start_page': start_page, 'end_page': end_page})

                # 다음 문단으로 이동
                before_para = para_id
                self.hwp.HAction.Run("MoveNextParaBegin")
                after_para = self._get_current_para_id()

                if before_para == after_para:
                    self._log(f"[COUNT] 마지막 문단")
                    break

            # 원래 위치 복원
            self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])

            self._log(f"[COUNT] 결과: {len(paragraphs)}개 문단")

            return {
                'page': target_page,
                'paragraph_count': len(paragraphs),
                'paragraphs': paragraphs
            }

        except Exception as e:
            self._log(f"[COUNT] 오류: {e}", "ERROR")
            self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])
            return {
                'page': target_page,
                'paragraph_count': 0,
                'paragraphs': [],
                'error': str(e)
            }

    def align_current_page(
        self,
        spacing_step: float = -1.0,
        min_spacing: float = -100,
        max_iterations: int = 100
    ) -> Dict:
        """
        현재 페이지의 모든 문단 정렬

        Args:
            spacing_step: 자간 감소 단위 (음수, 기본 -1.0)
            min_spacing: 최소 자간 값 (기본 -100)
            max_iterations: 문단당 최대 반복 횟수 (기본 100)

        Returns:
            {
                'success': True/False,
                'target_page': 1,              # 대상 페이지
                'processed_paragraphs': 3,     # 처리된 문단 수
                'total_adjusted_lines': 5,     # 총 조정된 줄 수
                'total_skipped_lines': 2,      # 총 건너뛴 줄 수
                'total_failed_lines': 0,       # 총 실패한 줄 수
                'paragraph_results': [...],    # 각 문단별 결과
                'message': '...'
            }
        """
        self.log_messages = []

        # 현재 위치 저장
        saved_pos = self.hwp.GetPos()
        list_id, para_id, char_pos = saved_pos

        # 현재 페이지 확인
        target_page = self._get_current_page()

        self._log(f"페이지 {target_page} 정렬 시작")

        # 결과 집계
        processed_paragraphs = 0
        total_adjusted = 0
        total_skipped = 0
        total_failed = 0
        paragraph_results = []

        try:
            # 문서 처음으로 이동
            self.hwp.MovePos(2)  # moveDocBegin
            self._log(f"문서 처음으로 이동")

            # 문단 순회
            while True:
                current_para_id = self._get_current_para_id()
                start_page, end_page = self._get_para_page_range(current_para_id)

                self._log(f"--- 문단 #{processed_paragraphs + 1} (para_id={current_para_id}, page={start_page}~{end_page}) ---")

                # 페이지 체크: 문단이 target_page에 걸쳐있는지
                if start_page > target_page:
                    # 문단 시작이 target_page 이후면 종료
                    self._log(f"[종료] 문단 시작 페이지 {start_page} > {target_page}")
                    break

                if end_page < target_page:
                    # 문단 끝이 target_page 이전이면 다음 문단으로
                    self._log(f"[SKIP] 문단 끝 페이지 {end_page} < {target_page}")
                    if not self._move_to_next_para():
                        break
                    continue

                # target_page에 포함되는 문단 (start_page <= target_page <= end_page)
                # TextAlign으로 문단 정렬 수행
                result = self.text_align.align_paragraph(
                    spacing_step=spacing_step,
                    min_spacing=min_spacing,
                    max_iterations=max_iterations
                )

                # 결과 집계
                processed_paragraphs += 1
                total_adjusted += result['adjusted_lines']
                total_skipped += result['skipped_lines']
                total_failed += result['failed_lines']

                paragraph_results.append({
                    'para_id': current_para_id,
                    'page': current_page,
                    'adjusted_lines': result['adjusted_lines'],
                    'skipped_lines': result['skipped_lines'],
                    'failed_lines': result['failed_lines'],
                    'success': result['success']
                })

                self._log(f"[ALIGN] 결과: 조정={result['adjusted_lines']}, 건너뜀={result['skipped_lines']}, 실패={result['failed_lines']}")

                # 다음 문단으로 이동
                if not self._move_to_next_para():
                    self._log(f"[종료] 마지막 문단")
                    break

            # 최종 결과
            result = {
                'success': total_failed == 0,
                'target_page': target_page,
                'processed_paragraphs': processed_paragraphs,
                'total_adjusted_lines': total_adjusted,
                'total_skipped_lines': total_skipped,
                'total_failed_lines': total_failed,
                'paragraph_results': paragraph_results,
                'message': f"페이지 {target_page}: {processed_paragraphs}개 문단 처리, 조정 {total_adjusted}줄"
            }

            self._log(f"완료: {processed_paragraphs}개 문단, 조정 {total_adjusted}줄")

            return result

        except Exception as e:
            self._log(f"예외 발생: {e}", "ERROR")
            self._log(f"traceback: {traceback.format_exc()}", "ERROR")

            # 원래 위치 복원 시도
            try:
                self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])
            except Exception as restore_err:
                self._log(f"위치 복원 실패: {restore_err}", "ERROR")

            return {
                'success': False,
                'target_page': target_page,
                'processed_paragraphs': 0,
                'total_adjusted_lines': 0,
                'total_skipped_lines': 0,
                'total_failed_lines': 0,
                'paragraph_results': [],
                'message': f"오류 발생: {e}"
            }

    def save_debug_log(self, result: Dict) -> str:
        """디버그 로그를 파일로 저장"""
        log_filename = f"text_align_page_{self.session_id}.json"
        log_filepath = self.log_dir / log_filename

        log_data = {
            'session_id': self.session_id,
            'module_info': _MODULE_INFO,
            'start_time': self.session_start.isoformat(),
            'end_time': datetime.now().isoformat(),
            'duration_seconds': (datetime.now() - self.session_start).total_seconds(),
            'result': result,
            'logs': self.log_messages
        }

        with open(log_filepath, 'w', encoding='utf-8') as f:
            json.dump(log_data, f, ensure_ascii=False, indent=2)

        return str(log_filepath)

    def save_text_log(self, result: Dict) -> str:
        """텍스트 로그 저장"""
        log_filename = f"text_align_page_{self.session_id}.txt"
        log_filepath = self.log_dir / log_filename

        lines = []
        lines.append("=" * 80)
        lines.append("HWP 페이지 텍스트 정렬 디버그 로그")
        lines.append("=" * 80)
        lines.append(f"실행 파일: {_MODULE_INFO['file']}")
        lines.append(f"세션 ID: {self.session_id}")
        lines.append(f"시작 시간: {self.session_start.strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append(f"종료 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append(f"소요 시간: {(datetime.now() - self.session_start).total_seconds():.2f}초")
        lines.append("")
        lines.append("-" * 80)
        lines.append("작업 결과")
        lines.append("-" * 80)
        lines.append(f"대상 페이지: {result.get('target_page', 'N/A')}")
        lines.append(f"처리된 문단: {result.get('processed_paragraphs', 0)}개")
        lines.append(f"총 조정: {result.get('total_adjusted_lines', 0)}줄")
        lines.append(f"총 건너뜀: {result.get('total_skipped_lines', 0)}줄")
        lines.append(f"총 실패: {result.get('total_failed_lines', 0)}줄")
        lines.append(f"메시지: {result.get('message', '')}")
        lines.append("")
        lines.append("-" * 80)
        lines.append("문단별 결과")
        lines.append("-" * 80)

        for pr in result.get('paragraph_results', []):
            lines.append(f"  문단 {pr['para_id']} (page {pr['page']}): "
                        f"조정={pr['adjusted_lines']}, 건너뜀={pr['skipped_lines']}, 실패={pr['failed_lines']}")

        lines.append("")
        lines.append("-" * 80)
        lines.append("상세 로그")
        lines.append("-" * 80)

        for log_entry in self.log_messages:
            lines.append(f"[{log_entry['timestamp']}] [{log_entry['level']:7s}] {log_entry['message']}")

        lines.append("")
        lines.append("=" * 80)

        with open(log_filepath, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines))

        return str(log_filepath)


def main():
    """CLI 실행 함수"""
    import argparse

    parser = argparse.ArgumentParser(description='HWP 페이지 단위 텍스트 정렬 도구')
    parser.add_argument('--spacing-step', type=float, default=-1.0,
                       help='자간 감소 단위 (기본: -1.0)')
    parser.add_argument('--min-spacing', type=float, default=-100,
                       help='최소 자간 값 (기본: -100)')
    parser.add_argument('--max-iterations', type=int, default=100,
                       help='문단당 최대 반복 횟수 (기본: 100)')
    parser.add_argument('--debug', action='store_true',
                       help='디버그 모드 활성화')
    parser.add_argument('--count-only', action='store_true',
                       help='문단 개수만 조회 (정렬 수행 안함)')
    parser.add_argument('--page', type=int, default=None,
                       help='대상 페이지 번호 (기본: 현재 페이지)')

    args = parser.parse_args()

    # HWP 인스턴스 연결
    hwp = get_hwp_instance()
    if not hwp:
        print("[ERROR] 실행 중인 한글을 찾을 수 없습니다.")
        print("한글을 먼저 실행하고 문서를 열어주세요.")
        return

    print("[OK] 한글에 연결되었습니다.")

    # TextAlignPage 객체 생성
    aligner = TextAlignPage(hwp, debug=args.debug)

    # 문단 개수만 조회
    if args.count_only:
        result = aligner.get_page_paragraph_count(args.page)

        print(f"페이지 {result['page']}: {result['paragraph_count']}개 문단")

        # 로그 저장
        try:
            log_filename = f"text_align_page_count_{aligner.session_id}.txt"
            log_filepath = aligner.log_dir / log_filename

            lines = []
            lines.append(f"페이지 {result['page']} 문단 개수 조회")
            lines.append(f"문단 개수: {result['paragraph_count']}개")
            lines.append("")
            for i, p in enumerate(result['paragraphs'], 1):
                span = f" (걸침)" if p['start_page'] != p['end_page'] else ""
                lines.append(f"  {i}. para_id={p['para_id']}, page={p['start_page']}~{p['end_page']}{span}")
            lines.append("")
            lines.append("-" * 50)
            lines.append("상세 로그")
            lines.append("-" * 50)
            for log_entry in aligner.log_messages:
                lines.append(f"[{log_entry['timestamp']}] {log_entry['message']}")

            with open(log_filepath, 'w', encoding='utf-8') as f:
                f.write('\n'.join(lines))

            print(f"로그: {log_filepath}")
        except Exception as e:
            print(f"로그 저장 실패: {e}")
        return

    # 현재 페이지 정렬
    print("\n[RUN] 현재 페이지 정렬 시작...")
    result = aligner.align_current_page(
        spacing_step=args.spacing_step,
        min_spacing=args.min_spacing,
        max_iterations=args.max_iterations
    )

    # 결과 출력
    print(f"\n{'='*50}")
    print(f"[OK] 완료" if result['success'] else "[WARN] 일부 실패")
    print(f"{'='*50}")
    print(f"대상 페이지: {result['target_page']}")
    print(f"처리된 문단: {result['processed_paragraphs']}개")
    print(f"총 조정: {result['total_adjusted_lines']}줄")
    print(f"총 건너뜀: {result['total_skipped_lines']}줄")
    print(f"총 실패: {result['total_failed_lines']}줄")
    print(f"메시지: {result['message']}")

    # 로그 저장
    try:
        json_path = aligner.save_debug_log(result)
        print(f"\n[LOG] JSON 로그 저장: {json_path}")

        text_path = aligner.save_text_log(result)
        print(f"[LOG] 텍스트 로그 저장: {text_path}")
    except Exception as e:
        print(f"[WARN] 로그 저장 실패: {e}")


if __name__ == '__main__':
    main()
