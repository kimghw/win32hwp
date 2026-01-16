"""
HWP 텍스트 정렬 모듈

한 단어가 두 줄에 걸쳐 분리된 경우, 이전 줄의 자간을 줄여서
분리된 부분을 이전 줄로 이동시켜 단어를 한 줄에 합치는 기능.

사용법:
    from text_align import TextAlign, get_hwp_instance

    hwp = get_hwp_instance()
    align = TextAlign(hwp, debug=True)

    # 현재 문단 정렬
    result = align.align_paragraph()
    print(f"조정된 줄 수: {result['adjusted_lines']}")
"""

import time
import re
import json
import traceback
import sys
import os
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from custom_block import CustomBlock
from cursor_position_monitor import get_hwp_instance

# 모듈 실행 정보 로그
_MODULE_INFO = {
    'file': os.path.abspath(__file__),
    'name': __name__,
    'loaded_at': datetime.now().isoformat()
}
print(f"[MODULE LOAD] {_MODULE_INFO['file']}")
print(f"[MODULE LOAD] __name__={_MODULE_INFO['name']}, loaded_at={_MODULE_INFO['loaded_at']}")


class TextAlign:
    """HWP 텍스트 정렬 클래스"""

    def __init__(self, hwp, debug: bool = False, log_dir: str = "debugs/logs"):
        """
        Args:
            hwp: HWP 객체
            debug: 디버그 모드 (True시 상세 로그 출력)
            log_dir: 로그 파일 저장 디렉토리
        """
        self.hwp = hwp
        self.debug = debug
        self.block = CustomBlock(hwp)
        self.log_messages = []
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(parents=True, exist_ok=True)

        # 세션 정보
        self.session_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.session_start = datetime.now()

        # 파라미터 정보 저장
        self.current_params = None

        # 커서 이력 추적
        self.cursor_history = []

        # 초기화 로그
        self._log(f"{'=' * 70}")
        self._log(f"[INIT] TextAlign 인스턴스 생성")
        self._log(f"[INIT] 실행 파일: {_MODULE_INFO['file']}")
        self._log(f"[INIT] Python 버전: {sys.version}")
        self._log(f"[INIT] 세션 ID: {self.session_id}")
        self._log(f"[INIT] 디버그 모드: {self.debug}")
        self._log(f"[INIT] 로그 디렉토리: {self.log_dir}")
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

    def _track_cursor(self, action: str, context: str = "") -> Tuple[int, int, int]:
        """
        커서 위치를 추적하고 이력에 기록

        Args:
            action: 수행한 작업 (예: "GetPos", "SetPos", "SelectText")
            context: 추가 컨텍스트 정보

        Returns:
            현재 커서 위치 (list_id, para_id, pos)
        """
        try:
            pos = self.hwp.GetPos()
            entry = {
                'timestamp': datetime.now().strftime("%H:%M:%S.%f")[:-3],
                'action': action,
                'context': context,
                'list_id': pos[0],
                'para_id': pos[1],
                'pos': pos[2]
            }
            self.cursor_history.append(entry)
            self._log(f"[CURSOR] {action}: list={pos[0]}, para={pos[1]}, pos={pos[2]} | {context}", "CURSOR")
            return pos
        except Exception as e:
            self._log(f"[CURSOR] {action} 실패: {e}", "ERROR")
            return (0, 0, 0)

    def _log_selection(self, action: str, start_para: int, start_pos: int,
                       end_para: int, end_pos: int, result_text: str = None):
        """
        선택/복사 작업 상세 로그

        Args:
            action: 수행한 작업 (예: "SelectText", "GetTextFile")
            start_para, start_pos: 선택 시작 위치
            end_para, end_pos: 선택 끝 위치
            result_text: 선택된 텍스트 (있는 경우)
        """
        self._log(f"[SELECT] {action}", "SELECT")
        self._log(f"   범위: para {start_para}:{start_pos} ~ para {end_para}:{end_pos}", "SELECT")
        self._log(f"   문자 수: {end_pos - start_pos if start_para == end_para else '(다중 문단)'}", "SELECT")
        if result_text is not None:
            preview = result_text[:30] + "..." if len(result_text) > 30 else result_text
            self._log(f"   결과 텍스트: '{preview}' (길이: {len(result_text)})", "SELECT")
            self._log(f"   결과 repr: {repr(result_text[:50]) if len(result_text) > 50 else repr(result_text)}", "SELECT")

    def _log_cursor_change(self, before: Tuple, after: Tuple, operation: str):
        """
        커서 위치 변경 상세 로그

        Args:
            before: 이전 위치 (list_id, para_id, pos)
            after: 이후 위치 (list_id, para_id, pos)
            operation: 수행한 작업명
        """
        if before == after:
            self._log(f"[CURSOR_CHG] {operation}: 위치 변경 없음 (para={before[1]}, pos={before[2]})", "CURSOR")
        else:
            self._log(f"[CURSOR_CHG] {operation}:", "CURSOR")
            self._log(f"   이전: list={before[0]}, para={before[1]}, pos={before[2]}", "CURSOR")
            self._log(f"   이후: list={after[0]}, para={after[1]}, pos={after[2]}", "CURSOR")
            # 변경 분석
            if before[1] != after[1]:
                self._log(f"   >> 문단 변경: {before[1]} -> {after[1]}", "CURSOR")
            if before[2] != after[2]:
                self._log(f"   >> pos 변경: {before[2]} -> {after[2]} (차이: {after[2] - before[2]})", "CURSOR")

    def get_cursor_history(self) -> List[Dict]:
        """커서 이력 반환"""
        return self.cursor_history.copy()

    def _get_line_info(self, para_id: int) -> Dict:
        """
        문단의 줄 정보 수집

        Args:
            para_id: 문단 ID

        Returns:
            {
                'line_starts': [0, 30, 66, ...],  # 각 줄 시작 pos
                'para_end': 150,                   # 문단 끝 pos
                'line_count': 5                    # 총 줄 수
            }
        """
        self._log(f"")
        self._log(f"[_get_line_info] 문단 줄 정보 수집 시작: para_id={para_id}")

        # 현재 커서 위치 저장
        saved_pos = self.hwp.GetPos()
        self._log(f"   [1] 현재 커서 위치 저장: list={saved_pos[0]}, para={saved_pos[1]}, pos={saved_pos[2]}")

        # CustomBlock의 _get_line_starts 활용
        self._log(f"   [2] CustomBlock._get_line_starts({para_id}) 호출")
        line_starts, para_end = self.block._get_line_starts(para_id)
        self._log(f"   [3] 결과: line_starts={line_starts}, para_end={para_end}")

        # 커서 위치 확인 (CustomBlock 내부에서 변경되었을 수 있음)
        current_pos = self.hwp.GetPos()
        self._log(f"   [4] _get_line_starts 후 커서 위치: list={current_pos[0]}, para={current_pos[1]}, pos={current_pos[2]}")

        # 커서 위치가 변경되었으면 복원
        if current_pos != saved_pos:
            self._log(f"   [5] 커서 위치 변경 감지! 원래 위치로 복원: SetPos({saved_pos[0]}, {saved_pos[1]}, {saved_pos[2]})")
            self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])
            restored_pos = self.hwp.GetPos()
            self._log(f"   [6] 복원 후 커서 위치: list={restored_pos[0]}, para={restored_pos[1]}, pos={restored_pos[2]}")
        else:
            self._log(f"   [5] 커서 위치 변경 없음")

        result = {
            'line_starts': line_starts,
            'para_end': para_end,
            'line_count': len(line_starts)
        }
        self._log(f"[_get_line_info] 완료: line_count={len(line_starts)}")
        return result

    def _get_line_text(self, para_id: int, line_index: int, line_info: Dict) -> str:
        """
        특정 줄의 텍스트 추출

        Args:
            para_id: 문단 ID
            line_index: 줄 번호 (0부터 시작)
            line_info: _get_line_info() 반환값

        Returns:
            줄 텍스트
        """
        line_starts = line_info['line_starts']
        para_end = line_info['para_end']

        if line_index >= len(line_starts):
            self._log(f"_get_line_text: line_index({line_index}) >= len(line_starts)({len(line_starts)})", "WARNING")
            return ""

        start_pos = line_starts[line_index]

        # 마지막 줄인 경우
        if line_index == len(line_starts) - 1:
            end_pos = para_end
        else:
            end_pos = line_starts[line_index + 1]

        self._log(f"_get_line_text: para_id={para_id}, line_index={line_index}, start_pos={start_pos}, end_pos={end_pos}")

        # 범위 선택 및 텍스트 추출
        try:
            # 현재 위치 저장
            saved_pos = self._track_cursor("SAVE", f"_get_line_text 시작 (line_index={line_index})")
            list_id = saved_pos[0]

            # 줄의 시작 위치로 이동
            self._log(f"   [2] 줄 시작 위치로 이동: SetPos({list_id}, {para_id}, {start_pos})")
            self.hwp.SetPos(list_id, para_id, start_pos)
            actual_pos = self._track_cursor("MOVE", f"SetPos({list_id}, {para_id}, {start_pos}) 후")
            self._log_cursor_change(saved_pos, actual_pos, "SetPos (줄 시작)")

            # 문단 내 범위 선택
            self._log(f"   [4] 범위 선택: SelectText({para_id}, {start_pos}, {para_id}, {end_pos})")
            before_select = self.hwp.GetPos()
            self.hwp.SelectText(para_id, start_pos, para_id, end_pos)
            after_select = self._track_cursor("SELECT", f"SelectText({para_id}, {start_pos}, {para_id}, {end_pos}) 후")
            self._log_cursor_change(before_select, after_select, "SelectText")

            # 선택된 텍스트 가져오기
            self._log(f"   [5] GetTextFile('TEXT', 'saveblock') 호출")
            text = self.hwp.GetTextFile("TEXT", "saveblock")
            self._log_selection("GetTextFile", para_id, start_pos, para_id, end_pos, text)

            # 선택 해제
            self._log(f"   [7] 선택 해제: Cancel")
            before_cancel = self.hwp.GetPos()
            self.hwp.HAction.Run("Cancel")
            after_cancel = self._track_cursor("CANCEL", "HAction.Run('Cancel') 후")
            self._log_cursor_change(before_cancel, after_cancel, "Cancel")

            # 원래 위치 복원
            self._log(f"   [8] 원래 위치 복원: SetPos({saved_pos[0]}, {saved_pos[1]}, {saved_pos[2]})")
            self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])
            restored_pos = self._track_cursor("RESTORE", "원래 위치 복원 후")
            self._log_cursor_change(after_cancel, restored_pos, "SetPos (복원)")

            # 복원 검증
            if restored_pos != saved_pos:
                self._log(f"   [WARNING] 위치 복원 불일치! 예상: {saved_pos}, 실제: {restored_pos}", "WARNING")

            if text:
                # 개행 문자만 제거 (공백은 유지)
                self._log(f"   [10] 개행 문자 처리 전: {repr(text)}")
                text = text.replace('\r\n', ' ').replace('\r', ' ').replace('\n', ' ')
                # 연속된 공백을 하나로
                text = re.sub(r' +', ' ', text)
                self._log(f"   [11] 개행 문자 처리 후: {repr(text)}")
            else:
                self._log(f"   [10] 텍스트 없음 (None 또는 빈 문자열)")
                text = ""

            self._log(f"_get_line_text: 최종 텍스트 = '{text}' (길이: {len(text)})")
            return text
        except Exception as e:
            self._log(f"_get_line_text: 텍스트 추출 실패 - {e}", "ERROR")
            self._log(f"_get_line_text: traceback = {traceback.format_exc()}", "ERROR")

            # 원래 위치 복원 시도
            try:
                self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])
                self._log(f"_get_line_text: 예외 후 위치 복원 완료")
            except Exception as restore_err:
                self._log(f"_get_line_text: 예외 후 위치 복원 실패: {restore_err}", "ERROR")

            return ""

    def _needs_alignment(self, text: str) -> bool:
        """
        줄이 정렬 대상인지 판단

        조건: "1~2글자 + 공백"으로 시작하는 경우

        Args:
            text: 줄 텍스트

        Returns:
            True: 정렬 필요, False: 정렬 불필요
        """
        if not text or len(text) == 0:
            return False

        # 공백 위치 찾기
        first_space_idx = text.find(' ')

        if first_space_idx == -1:
            return False

        # 공백이 위치 0 또는 1에 있으면 정렬 대상
        # 위치 0: "적 과정..." (1글자 + 공백)
        # 위치 1: "며, 제원..." (2글자 + 공백)
        return first_space_idx in [0, 1]

    def _line_ends_with_space(self, text: str) -> bool:
        """줄이 공백으로 끝나는지 확인"""
        return len(text) > 0 and text[-1] == ' '

    def _adjust_line_spacing(
        self,
        para_id: int,
        line_index: int,
        spacing: int
    ) -> bool:
        """
        특정 줄의 자간 조정

        Args:
            para_id: 문단 ID
            line_index: 줄 번호
            spacing: 자간 값 (HWPUNIT, 음수 가능)

        Returns:
            True: 성공, False: 실패
        """
        self._log(f"")
        self._log(f"[_adjust_line_spacing] 자간 조정 시작")
        self._log(f"   입력: para_id={para_id}, line_index={line_index}, spacing={spacing}")

        try:
            # 현재 커서 위치 저장
            saved_pos = self._track_cursor("SAVE", f"_adjust_line_spacing 시작 (line_index={line_index})")

            # 줄 선택
            self._log(f"   [2] 블록 선택 시작: block.select_line_by_index({para_id}, {line_index})")
            before_select = self.hwp.GetPos()
            self.block.select_line_by_index(para_id, line_index)
            after_select_pos = self._track_cursor("SELECT", f"block.select_line_by_index({para_id}, {line_index}) 후")
            self._log_cursor_change(before_select, after_select_pos, "select_line_by_index")

            # 선택된 텍스트 확인 (디버그용)
            selected_text = self.hwp.GetTextFile("TEXT", "saveblock")
            if selected_text:
                preview = selected_text[:50] + "..." if len(selected_text) > 50 else selected_text
                self._log(f"   [4] 선택된 텍스트 (미리보기): '{preview}'")
                self._log(f"       선택된 텍스트 길이: {len(selected_text)}")
                self._log(f"       선택된 텍스트 repr: {repr(selected_text[:80]) if len(selected_text) > 80 else repr(selected_text)}")
            else:
                self._log(f"   [4] 선택된 텍스트: (없음 또는 빈 문자열)", "WARNING")

            # 자간 설정
            self._log(f"   [5] HParameterSet.HCharShape 가져오기")
            pset = self.hwp.HParameterSet.HCharShape

            self._log(f"   [6] HAction.GetDefault('CharShape', pset.HSet) 호출")
            self.hwp.HAction.GetDefault("CharShape", pset.HSet)

            self._log(f"   [7] 자간 값 설정:")
            self._log(f"       - SpacingHangul = {spacing}")
            self._log(f"       - SpacingLatin = {spacing}")
            pset.SpacingHangul = spacing
            pset.SpacingLatin = spacing

            self._log(f"   [8] HAction.Execute('CharShape', pset.HSet) 호출")
            before_execute = self.hwp.GetPos()
            self.hwp.HAction.Execute("CharShape", pset.HSet)
            after_action_pos = self._track_cursor("EXECUTE", "HAction.Execute('CharShape') 후")
            self._log_cursor_change(before_execute, after_action_pos, "CharShape Execute")

            # 선택 해제
            self._log(f"   [10] 블록 선택 해제: block.cancel()")
            before_cancel = self.hwp.GetPos()
            self.block.cancel()
            after_cancel_pos = self._track_cursor("CANCEL", "block.cancel() 후")
            self._log_cursor_change(before_cancel, after_cancel_pos, "block.cancel")

            # 레이아웃 재계산 대기
            self._log(f"   [12] 레이아웃 재계산 대기 (0.025초)")
            time.sleep(0.025)

            # 최종 위치 확인
            final_pos = self._track_cursor("FINAL", "_adjust_line_spacing 완료")
            self._log_cursor_change(saved_pos, final_pos, "_adjust_line_spacing 전체")

            self._log(f"[_adjust_line_spacing] 완료: 성공")
            return True

        except Exception as e:
            self._log(f"[_adjust_line_spacing] 예외 발생: {e}", "ERROR")
            self._log(f"   traceback: {traceback.format_exc()}", "ERROR")
            return False

    def save_debug_log(self, result: Dict, extra_info: Dict = None) -> str:
        """
        디버그 로그를 파일로 저장

        Args:
            result: align_paragraph() 반환값
            extra_info: 추가 정보 (선택사항)

        Returns:
            저장된 파일 경로
        """
        # 로그 파일명
        log_filename = f"text_align_{self.session_id}.json"
        log_filepath = self.log_dir / log_filename

        # 저장할 데이터
        log_data = {
            'session_id': self.session_id,
            'module_info': _MODULE_INFO,
            'start_time': self.session_start.isoformat(),
            'end_time': datetime.now().isoformat(),
            'duration_seconds': (datetime.now() - self.session_start).total_seconds(),
            'result': {
                'success': result['success'],
                'adjusted_lines': result['adjusted_lines'],
                'skipped_lines': result['skipped_lines'],
                'failed_lines': result['failed_lines'],
                'total_lines': result['total_lines'],
                'message': result['message']
            },
            'cursor_history': self.cursor_history,
            'cursor_history_count': len(self.cursor_history),
            'logs': self.log_messages
        }

        # 추가 정보가 있으면 포함
        if extra_info:
            log_data['extra_info'] = extra_info

        # JSON으로 저장
        with open(log_filepath, 'w', encoding='utf-8') as f:
            json.dump(log_data, f, ensure_ascii=False, indent=2)

        return str(log_filepath)

    def save_text_log(self, result: Dict) -> str:
        """
        사람이 읽기 쉬운 텍스트 로그 저장

        Args:
            result: align_paragraph() 반환값

        Returns:
            저장된 파일 경로
        """
        # 로그 파일명
        log_filename = f"text_align_{self.session_id}.txt"
        log_filepath = self.log_dir / log_filename

        # 텍스트 로그 작성
        lines = []
        lines.append("=" * 80)
        lines.append(f"HWP 텍스트 정렬 디버그 로그")
        lines.append("=" * 80)
        lines.append(f"실행 파일: {_MODULE_INFO['file']}")
        lines.append(f"모듈 로드 시간: {_MODULE_INFO['loaded_at']}")
        lines.append(f"세션 ID: {self.session_id}")
        lines.append(f"시작 시간: {self.session_start.strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append(f"종료 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append(f"소요 시간: {(datetime.now() - self.session_start).total_seconds():.2f}초")
        lines.append("")
        lines.append("-" * 80)
        lines.append("작업 결과")
        lines.append("-" * 80)
        lines.append(f"성공 여부: {'✅ 성공' if result['success'] else '⚠️  실패'}")
        lines.append(f"조정된 줄 수: {result['adjusted_lines']}")
        lines.append(f"건너뛴 줄 수: {result['skipped_lines']}")
        lines.append(f"실패한 줄 수: {result['failed_lines']}")
        lines.append(f"전체 줄 수: {result['total_lines']}")
        lines.append(f"메시지: {result['message']}")
        lines.append("")
        lines.append("-" * 80)
        lines.append("상세 로그")
        lines.append("-" * 80)

        # 로그 메시지 출력
        for log_entry in self.log_messages:
            timestamp = log_entry['timestamp']
            level = log_entry['level']
            message = log_entry['message']
            lines.append(f"[{timestamp}] [{level:7s}] {message}")

        lines.append("")
        lines.append("-" * 80)
        lines.append("커서 이력 요약")
        lines.append("-" * 80)
        lines.append(f"총 커서 추적 횟수: {len(self.cursor_history)}")

        # 커서 이력 통계
        if self.cursor_history:
            actions = {}
            for entry in self.cursor_history:
                action = entry['action']
                actions[action] = actions.get(action, 0) + 1
            lines.append("액션별 횟수:")
            for action, count in sorted(actions.items(), key=lambda x: -x[1]):
                lines.append(f"   {action}: {count}회")

            # 마지막 10개 커서 이력
            lines.append("")
            lines.append("최근 커서 이력 (마지막 10개):")
            for entry in self.cursor_history[-10:]:
                lines.append(f"   [{entry['timestamp']}] {entry['action']}: "
                           f"para={entry['para_id']}, pos={entry['pos']} | {entry['context']}")

        lines.append("")
        lines.append("=" * 80)

        # 파일로 저장
        with open(log_filepath, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines))

        return str(log_filepath)

    def align_paragraph(
        self,
        spacing_step: float = -1.0,
        min_spacing: float = -100,
        max_iterations: int = 100
    ) -> Dict:
        """
        현재 커서가 위치한 문단의 모든 줄 정렬

        Args:
            spacing_step: 자간 감소 단위 (음수, 기본 -1.0)
            min_spacing: 최소 자간 값 (기본 -100)
            max_iterations: 최대 반복 횟수 (기본 100)

        Returns:
            {
                'success': True/False,
                'adjusted_lines': 3,           # 조정된 줄 수
                'skipped_lines': 2,            # 건너뛴 줄 수
                'failed_lines': 0,             # 실패한 줄 수
                'total_lines': 5,              # 전체 줄 수
                'message': '...',
                'log': [...]                   # 로그 메시지
            }
        """
        self.log_messages = []

        # 파라미터 저장
        self.current_params = {
            'spacing_step': spacing_step,
            'min_spacing': min_spacing,
            'max_iterations': max_iterations
        }

        # 현재 커서 위치 저장
        list_id, para_id, char_pos = self.hwp.GetPos()
        self._log(f"=" * 60)
        self._log(f"문단 정렬 시작")
        self._log(f"=" * 60)
        self._log(f"📍 위치: para_id={para_id}, char_pos={char_pos}")
        self._log(f"⚙️  파라미터:")
        self._log(f"   - spacing_step: {spacing_step} (자간 감소 단위)")
        self._log(f"   - min_spacing: {min_spacing} (최소 자간값)")
        self._log(f"   - max_iterations: {max_iterations} (최대 반복)")
        self._log(f"-" * 60)

        try:
            # 줄 정보 수집
            line_info = self._get_line_info(para_id)
            total_lines = line_info['line_count']

            self._log(f"전체 줄 수: {total_lines}")
            self._log(f"줄 시작 위치: {line_info['line_starts']}")
            self._log(f"문단 끝 위치: {line_info['para_end']}")

            if total_lines < 2:
                return {
                    'success': True,
                    'adjusted_lines': 0,
                    'skipped_lines': 0,
                    'failed_lines': 0,
                    'total_lines': total_lines,
                    'message': '줄이 1개 이하로 정렬 불필요',
                    'log': self.log_messages
                }

            adjusted_count = 0
            skipped_count = 0
            failed_count = 0
            iteration_count = 0

            # 2번째 줄부터 검사 (첫 줄은 검사 불필요)
            line_idx = 1

            while line_idx < total_lines and iteration_count < max_iterations:
                iteration_count += 1

                self._log(f"")
                self._log(f"{'#' * 70}")
                self._log(f"# 메인 루프 반복 #{iteration_count}")
                self._log(f"# line_idx={line_idx}, total_lines={total_lines}")
                self._log(f"{'#' * 70}")

                # 현재 커서 위치 확인
                loop_pos = self.hwp.GetPos()
                self._log(f"[루프 시작] 현재 커서 위치: list={loop_pos[0]}, para={loop_pos[1]}, pos={loop_pos[2]}")

                # line_info는 이미 갱신되어 있음 (초기 또는 자간 조정 후)
                current_total_lines = line_info['line_count']
                self._log(f"[루프] 현재 줄 수: {current_total_lines}")

                # 줄 수가 줄어든 경우 (정렬 성공으로 인한 줄 병합)
                if line_idx >= current_total_lines:
                    self._log(f"[루프 종료] 줄 {line_idx}가 병합됨 (전체 줄 수: {current_total_lines})")
                    break

                # 현재 줄 텍스트
                current_text = self._get_line_text(para_id, line_idx, line_info)
                self._log(f"\n--- 줄 {line_idx + 1}/{current_total_lines} ---")
                self._log(f"텍스트 전체: '{current_text}'")
                self._log(f"텍스트 길이: {len(current_text)}")

                # 공백 위치 디버깅
                first_space_idx = current_text.find(' ')
                if first_space_idx >= 0:
                    self._log(f"🔍 첫 공백 위치: {first_space_idx}")
                    self._log(f"   공백 앞 텍스트: '{current_text[:first_space_idx]}' (길이: {first_space_idx})")

                    # 분리 패턴 분석 (공백 앞 글자 수 기준)
                    if first_space_idx == 0:
                        self._log(f"   📌 패턴: 공백 앞 0글자 (정렬 대상)")
                    elif first_space_idx == 1:
                        self._log(f"   📌 패턴: 공백 앞 1글자 (정렬 대상)")
                    else:
                        self._log(f"   ❌ 패턴: 공백 앞 {first_space_idx}글자 (정렬 대상: 0~1글자만)")
                else:
                    self._log(f"❌ 공백 없음 (정렬 불가)")

                # 정렬 필요 여부 확인
                needs_align = self._needs_alignment(current_text)
                if not needs_align:
                    self._log(f"⏭️  건너뜀: 정렬 패턴 불일치")
                    skipped_count += 1
                    line_idx += 1
                    continue

                self._log(f"✅ 정렬 대상으로 판단됨")

                # 이전 줄 텍스트
                prev_text = self._get_line_text(para_id, line_idx - 1, line_info)
                self._log(f"이전 줄 텍스트: '{prev_text}'")
                self._log(f"이전 줄 끝 문자: '{prev_text[-1] if prev_text else ''}'")
                self._log(f"이전 줄 끝이 공백? {self._line_ends_with_space(prev_text)}")

                # 이미 이전 줄이 공백으로 끝나면 성공
                if self._line_ends_with_space(prev_text):
                    self._log(f"이미 정렬됨: 이전 줄 끝이 공백")
                    line_idx += 1
                    continue

                self._log(f"🎯 정렬 시작")
                self._log(f"   이전 줄 끝 10자: '{prev_text[-10:] if len(prev_text) >= 10 else prev_text}'")

                # 자간 조정 시작
                current_spacing = 0
                line_adjusted = False
                same_line_attempts = 0
                max_same_line_attempts = 10

                self._log(f"-" * 60)
                self._log(f"자간 조정 루프 시작:")
                self._log(f"   초기 자간: {current_spacing}")
                self._log(f"   자간 감소 단위: {spacing_step}")
                self._log(f"   최소 자간: {min_spacing}")
                self._log(f"   최대 시도: {max_same_line_attempts}")
                self._log(f"-" * 60)

                # 현재 줄의 처음 단어 길이 확인 (몇 글자를 올려야 하는지)
                first_space_idx = current_text.find(' ')
                target_chars = first_space_idx + 1  # 공백 포함
                self._log(f"🎯 올려야 할 글자 수: {target_chars}글자 (공백 포함)")

                while current_spacing > min_spacing and same_line_attempts < max_same_line_attempts:
                    same_line_attempts += 1
                    current_spacing += spacing_step

                    self._log(f"")
                    self._log(f"🔧 시도 #{same_line_attempts}")
                    self._log(f"   현재 자간: {current_spacing}")
                    self._log(f"   남은 여유: {current_spacing - min_spacing} (최소값까지)")

                    # 이전 줄 자간 조정
                    if not self._adjust_line_spacing(para_id, line_idx - 1, current_spacing):
                        self._log("자간 조정 실패", "ERROR")
                        break

                    # 줄 정보 재수집
                    line_info = self._get_line_info(para_id)

                    # 줄 수 변경 확인
                    new_total_lines = line_info['line_count']
                    if new_total_lines != current_total_lines:
                        self._log(f"줄 수 변경: {current_total_lines} -> {new_total_lines}")
                        current_total_lines = new_total_lines

                        # 현재 줄이 사라진 경우 (성공)
                        if line_idx >= new_total_lines:
                            self._log("✅ 성공: 줄 병합됨")
                            line_adjusted = True
                            break

                    # 이전 줄 다시 확인
                    if line_idx - 1 >= len(line_info['line_starts']):
                        self._log("이전 줄 인덱스 오류", "ERROR")
                        break

                    prev_text = self._get_line_text(para_id, line_idx - 1, line_info)
                    self._log(f"자간 조정 후 이전 줄: '{prev_text[-20:]}'")
                    self._log(f"이전 줄 끝 문자 (repr): {repr(prev_text[-1]) if prev_text else 'None'}")

                    # 이전 줄이 공백으로 끝나면 성공
                    if self._line_ends_with_space(prev_text):
                        self._log(f"✅ 성공: 이전 줄 끝이 공백")
                        line_adjusted = True
                        adjusted_count += 1
                        break
                    else:
                        self._log(f"아직 공백 아님, 계속 시도...")

                if line_adjusted:
                    # 성공했으므로 현재 줄을 다시 검사 (줄 번호 유지)
                    # 왜냐하면 자간 조정으로 줄 구조가 변경되었을 수 있음
                    self._log(f"")
                    self._log(f"✨ 줄 {line_idx + 1} 정렬 성공!")
                    self._log(f"   최종 자간: {current_spacing}")
                    self._log(f"   시도 횟수: {same_line_attempts}")
                    total_lines = line_info['line_count']
                    continue
                else:
                    self._log(f"")
                    self._log(f"❌ 줄 {line_idx + 1} 정렬 실패", "WARNING")
                    self._log(f"   최종 자간: {current_spacing}")
                    self._log(f"   시도 횟수: {same_line_attempts}/{max_same_line_attempts}")
                    self._log(f"   실패 이유: ", "WARNING")
                    if current_spacing <= min_spacing:
                        self._log(f"      - 최소 자간({min_spacing}) 도달", "WARNING")
                    if same_line_attempts >= max_same_line_attempts:
                        self._log(f"      - 최대 시도 횟수({max_same_line_attempts}) 초과", "WARNING")
                    failed_count += 1
                    line_idx += 1

            # 커서를 문단 끝으로 이동
            para_end_pos = line_info['para_end']
            before_final = self.hwp.GetPos()
            self.hwp.SetPos(list_id, para_id, para_end_pos)
            after_final = self._track_cursor("FINAL_MOVE", f"문단 끝으로 이동: SetPos({list_id}, {para_id}, {para_end_pos})")
            self._log_cursor_change(before_final, after_final, "문단 끝으로 커서 이동")

            result = {
                'success': failed_count == 0,
                'adjusted_lines': adjusted_count,
                'skipped_lines': skipped_count,
                'failed_lines': failed_count,
                'total_lines': total_lines,
                'message': f"조정: {adjusted_count}, 건너뜀: {skipped_count}, 실패: {failed_count}",
                'log': self.log_messages
            }

            self._log(f"")
            self._log(f"=" * 60)
            self._log(f"작업 완료")
            self._log(f"=" * 60)
            self._log(f"📊 결과 요약:")
            self._log(f"   전체 줄 수: {total_lines}")
            self._log(f"   조정 성공: {adjusted_count} 줄")
            self._log(f"   건너뜀: {skipped_count} 줄")
            self._log(f"   실패: {failed_count} 줄")
            self._log(f"   반복 횟수: {iteration_count}/{max_iterations}")
            if failed_count == 0:
                self._log(f"✅ 모든 줄 처리 완료!")
            else:
                self._log(f"⚠️  일부 줄 처리 실패", "WARNING")
            self._log(f"=" * 60)

            return result

        except Exception as e:
            self._log(f"예외 발생: {e}", "ERROR")
            self._log(f"   traceback: {traceback.format_exc()}", "ERROR")
            # 커서 위치 복원 시도
            try:
                self.hwp.SetPos(list_id, para_id, char_pos)
                self._log(f"예외 후 커서 위치 복원 성공")
            except Exception as restore_err:
                self._log(f"예외 후 커서 위치 복원 실패: {restore_err}", "ERROR")

            return {
                'success': False,
                'adjusted_lines': 0,
                'skipped_lines': 0,
                'failed_lines': 0,
                'total_lines': 0,
                'message': f"오류 발생: {e}",
                'log': self.log_messages
            }


def main():
    """CLI 실행 함수"""
    import argparse

    parser = argparse.ArgumentParser(description='HWP 텍스트 정렬 도구')
    parser.add_argument('--spacing-step', type=float, default=-1.0,
                       help='자간 감소 단위 (기본: -1.0)')
    parser.add_argument('--min-spacing', type=float, default=-100,
                       help='최소 자간 값 (기본: -100)')
    parser.add_argument('--max-iterations', type=int, default=100,
                       help='최대 반복 횟수 (기본: 100)')
    parser.add_argument('--debug', action='store_true',
                       help='디버그 모드 활성화')
    parser.add_argument('--save-log', action='store_true',
                       help='로그를 파일로 저장')

    args = parser.parse_args()

    # HWP 인스턴스 연결
    hwp = get_hwp_instance()
    if not hwp:
        print("[ERROR] 실행 중인 한글을 찾을 수 없습니다.")
        print("한글을 먼저 실행하고 문서를 열어주세요.")
        return

    print("[OK] 한글에 연결되었습니다.")

    # TextAlign 객체 생성
    align = TextAlign(hwp, debug=args.debug)

    # 현재 문단 정렬
    print("\n[RUN] 현재 문단 정렬 시작...")
    result = align.align_paragraph(
        spacing_step=args.spacing_step,
        min_spacing=args.min_spacing,
        max_iterations=args.max_iterations
    )

    # 결과 출력
    print(f"\n{'='*50}")
    print(f"[OK] 완료" if result['success'] else "[WARN] 일부 실패")
    print(f"{'='*50}")
    print(f"조정된 줄 수: {result['adjusted_lines']}")
    print(f"건너뛴 줄 수: {result['skipped_lines']}")
    print(f"실패한 줄 수: {result['failed_lines']}")
    print(f"전체 줄 수: {result['total_lines']}")
    print(f"메시지: {result['message']}")

    # 로그 저장 (항상 저장)
    try:
        # JSON 로그 저장
        json_path = align.save_debug_log(result)
        print(f"\n[LOG] JSON 로그 저장: {json_path}")

        # 텍스트 로그 저장
        text_path = align.save_text_log(result)
        print(f"[LOG] 텍스트 로그 저장: {text_path}")
    except Exception as e:
        print(f"[WARN] 로그 저장 실패: {e}")


if __name__ == '__main__':
    main()
