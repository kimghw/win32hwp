# -*- coding: utf-8 -*-
"""
한글(HWP) 블록 선택 클래스 (BlockSelector)

블록 선택 방법:
1. select_para(para) - 문단 전체 선택
2. select_line_by_index(para, line_index) - 문단 내 n번째 줄 선택
3. select_line_by_pos(para, pos) - 문단 내 pos가 속한 줄 선택
4. select_lines_range(para, start, end) - n번째~m번째 줄 선택 (없으면 마지막까지)
5. select_sentence(para, sentence_index) - 문단 내 n번째 문장 선택
6. select_sentences_range(para, start, end) - n번째~m번째 문장 선택 (없으면 마지막까지)
7. select_sentence_in_line(para, pos) - pos가 속한 줄 내 문장만 선택 (줄 넘김 X)
"""

import win32com.client as win32


class BlockSelector:
    """한글 블록 선택 클래스"""

    def __init__(self, hwp):
        """
        Args:
            hwp: 한글 인스턴스
        """
        self.hwp = hwp

    def _save_pos(self):
        """현재 위치 저장"""
        return self.hwp.GetPos()

    def _restore_pos(self, pos):
        """위치 복원"""
        self.hwp.SetPos(pos[0], pos[1], pos[2])

    def _get_line_starts(self, para_id):
        """문단 내 모든 줄의 시작 pos 목록 반환"""
        saved = self._save_pos()

        # 해당 문단으로 이동
        self.hwp.SetPos(saved[0], para_id, 0)

        # 문단 끝 pos
        self.hwp.HAction.Run("MoveParaEnd")
        para_end = self.hwp.GetPos()[2]

        # 문단 시작으로
        self.hwp.HAction.Run("MoveParaBegin")

        line_starts = [0]
        while True:
            self.hwp.HAction.Run("MoveLineDown")
            pos = self.hwp.GetPos()
            if pos[1] != para_id:
                break
            if pos[2] in line_starts:
                break
            line_starts.append(pos[2])
            if pos[2] >= para_end:
                break

        self._restore_pos(saved)
        return sorted(line_starts), para_end

    def _get_pos_map(self, para_id):
        """문단 내 각 문자의 실제 HWP pos 매핑 반환"""
        saved = self._save_pos()
        self.hwp.SetPos(saved[0], para_id, 0)

        self.hwp.HAction.Run("MoveParaBegin")
        start = self.hwp.GetPos()
        self.hwp.HAction.Run("MoveParaEnd")
        para_end_pos = self.hwp.GetPos()[2]

        # 문단 텍스트 가져오기
        self.hwp.SelectText(start[1], start[2], start[1], para_end_pos)
        para_text = self.hwp.GetTextFile("TEXT", "saveblock")
        self.hwp.HAction.Run("Cancel")

        if not para_text:
            self._restore_pos(saved)
            return [], "", []

        para_text = para_text.replace('\r\n', '').replace('\r', '').replace('\n', '')

        # 각 문자의 HWP pos 수집
        self.hwp.HAction.Run("MoveParaBegin")
        pos_list = []
        for _ in range(len(para_text)):
            pos_list.append(self.hwp.GetPos()[2])
            self.hwp.HAction.Run("MoveRight")

        self._restore_pos(saved)
        return pos_list, para_text, para_end_pos

    def _get_sentences(self, para_id):
        """문단 내 문장 경계 목록 반환 (HWP pos 기반)"""
        pos_list, para_text, para_end_pos = self._get_pos_map(para_id)

        if not para_text:
            return [], ""

        sentences = []
        sentence_index = 1
        sentence_start_idx = 0
        i = 0

        while i < len(para_text):
            if para_text[i] == '.':
                sentences.append({
                    'index': sentence_index,
                    'start': pos_list[sentence_start_idx],
                    'end': pos_list[i]
                })

                next_idx = i + 1
                skip_count = 0
                while next_idx < len(para_text) and skip_count < 3:
                    if para_text[next_idx] == ' ':
                        next_idx += 1
                        skip_count += 1
                    else:
                        break

                if next_idx >= len(para_text) or (skip_count >= 3 and para_text[next_idx] == ' '):
                    break

                sentence_index += 1
                sentence_start_idx = next_idx
                i = next_idx
            else:
                i += 1

        if sentence_start_idx < len(para_text) and (not sentences or sentences[-1]['end'] < pos_list[-1]):
            sentences.append({
                'index': sentence_index,
                'start': pos_list[sentence_start_idx],
                'end': para_end_pos
            })

        return sentences, para_text

    def select_line_by_index(self, para_id, line_index):
        """
        문단 내 n번째 줄 선택

        Args:
            para_id: 문단 번호
            line_index: 줄 번호 (0부터 시작)

        Returns:
            bool: 성공 여부
        """
        line_starts, para_end = self._get_line_starts(para_id)

        if line_index < 0 or line_index >= len(line_starts):
            return False

        start = line_starts[line_index]
        if line_index + 1 < len(line_starts):
            end = line_starts[line_index + 1] - 1
        else:
            end = para_end

        self.hwp.SelectText(para_id, start, para_id, end + 1)
        return True

    def select_line_by_pos(self, para_id, pos):
        """
        문단 내 pos가 속한 줄 선택

        Args:
            para_id: 문단 번호
            pos: 문단 내 위치

        Returns:
            bool: 성공 여부
        """
        line_starts, para_end = self._get_line_starts(para_id)

        # pos가 속한 줄 찾기
        current_line_index = 0
        for i, line_start in enumerate(line_starts):
            if line_start <= pos:
                current_line_index = i
            else:
                break

        return self.select_line_by_index(para_id, current_line_index)

    def select_sentence(self, para_id, sentence_index):
        """
        문단 내 n번째 문장 선택

        Args:
            para_id: 문단 번호
            sentence_index: 문장 번호 (1부터 시작)

        Returns:
            bool: 성공 여부
        """
        sentences, _ = self._get_sentences(para_id)

        for s in sentences:
            if s['index'] == sentence_index:
                self.hwp.SelectText(para_id, s['start'], para_id, s['end'] + 1)
                return True

        return False

    def select_sentence_in_line(self, para_id, pos):
        """
        pos가 속한 문장 선택 (줄을 넘어가지 않음)

        Args:
            para_id: 문단 번호
            pos: 문단 내 위치

        Returns:
            bool: 성공 여부
        """
        sentences, _ = self._get_sentences(para_id)
        line_starts, para_end = self._get_line_starts(para_id)

        # pos가 속한 문장 찾기
        current_sentence = None
        for s in sentences:
            if s['start'] <= pos <= s['end']:
                current_sentence = s
                break

        if current_sentence is None:
            return False

        # pos가 속한 줄 범위 찾기
        line_start = 0
        line_end = para_end
        for i, ls in enumerate(line_starts):
            if ls <= pos:
                line_start = ls
                if i + 1 < len(line_starts):
                    line_end = line_starts[i + 1] - 1
                else:
                    line_end = para_end

        # 문장과 줄의 교집합
        select_start = max(current_sentence['start'], line_start)
        select_end = min(current_sentence['end'], line_end)

        self.hwp.SelectText(para_id, select_start, para_id, select_end + 1)
        return True

    def select_para(self, para_id):
        """
        문단 전체 선택

        Args:
            para_id: 문단 번호

        Returns:
            bool: 성공 여부
        """
        saved = self._save_pos()
        self.hwp.SetPos(saved[0], para_id, 0)

        self.hwp.HAction.Run("MoveParaBegin")
        start = self.hwp.GetPos()[2]
        self.hwp.HAction.Run("MoveParaEnd")
        end = self.hwp.GetPos()[2]

        self._restore_pos(saved)
        self.hwp.SelectText(para_id, start, para_id, end)
        return True

    def select_lines_range(self, para_id, start_line, end_line):
        """
        n번째 줄부터 m번째 줄까지 선택
        end_line이 존재하지 않으면 마지막 줄까지 선택

        Args:
            para_id: 문단 번호
            start_line: 시작 줄 번호 (0부터)
            end_line: 끝 줄 번호 (0부터, 포함)

        Returns:
            bool: 성공 여부
        """
        line_starts, para_end = self._get_line_starts(para_id)

        if start_line < 0 or start_line >= len(line_starts):
            return False

        # end_line이 범위를 벗어나면 마지막 줄까지
        if end_line >= len(line_starts):
            end_line = len(line_starts) - 1

        start_pos = line_starts[start_line]
        if end_line + 1 < len(line_starts):
            end_pos = line_starts[end_line + 1] - 1
        else:
            end_pos = para_end

        self.hwp.SelectText(para_id, start_pos, para_id, end_pos + 1)
        return True

    def select_sentences_range(self, para_id, start_sentence, end_sentence):
        """
        n번째 문장부터 m번째 문장까지 선택
        end_sentence가 존재하지 않으면 마지막 문장까지 선택

        Args:
            para_id: 문단 번호
            start_sentence: 시작 문장 번호 (1부터)
            end_sentence: 끝 문장 번호 (1부터, 포함)

        Returns:
            bool: 성공 여부
        """
        sentences, _ = self._get_sentences(para_id)

        if not sentences:
            return False

        # 유효한 문장 찾기
        start_s = None
        end_s = None

        for s in sentences:
            if s['index'] == start_sentence:
                start_s = s
            if s['index'] <= end_sentence:
                end_s = s

        if start_s is None:
            return False

        # end_sentence가 없으면 마지막 문장
        if end_s is None:
            end_s = sentences[-1]

        self.hwp.SelectText(para_id, start_s['start'], para_id, end_s['end'] + 1)
        return True

    def cancel(self):
        """블록 선택 해제"""
        self.hwp.HAction.Run("Cancel")

    def get_selected_text(self):
        """선택된 텍스트 반환"""
        text = self.hwp.GetTextFile("TEXT", "saveblock")
        if text:
            return text.replace('\r\n', '').replace('\r', '').replace('\n', '')
        return ""
