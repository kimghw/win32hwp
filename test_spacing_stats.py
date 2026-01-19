# 분리단어 처리 - 글자 수별 자간 조정 횟수 통계
import win32com.client as win32
from table.table_md_2_hwp import markdown_to_hwp, set_hwp
from block_selector import BlockSelector
import re

# 테스트 마크다운 - md_content.md 파일에서 읽기
import os
MD_FILE = os.path.join(os.path.dirname(__file__), "table", "md_content.md")
with open(MD_FILE, "r", encoding="utf-8") as f:
    test_markdown = f.read()


class SpacingStatsCollector:
    """자간 조정 통계 수집"""

    def __init__(self, hwp):
        self.hwp = hwp
        self.block = BlockSelector(hwp)
        self.stats = {
            'reduce_1': [],  # 1글자 올릴 때 (자간 줄임)
            'reduce_2': [],  # 2글자 올릴 때 (자간 줄임)
            'expand_3': [],  # 3글자 (자간 늘림, 1글자 내림)
        }

    def _get_line_starts(self, para_id):
        """문단의 줄 시작 위치 목록"""
        return self.block._get_line_starts(para_id)

    def _get_line_text(self, para_id, line_idx, line_starts, para_end):
        """특정 줄 텍스트 가져오기"""
        if line_idx >= len(line_starts):
            return ""

        start_pos = line_starts[line_idx]
        end_pos = line_starts[line_idx + 1] if line_idx < len(line_starts) - 1 else para_end

        saved_pos = self.hwp.GetPos()
        list_id = saved_pos[0]

        self.hwp.SelectText(para_id, start_pos, para_id, end_pos)
        text = self.hwp.GetTextFile("TEXT", "saveblock")
        self.hwp.HAction.Run("Cancel")
        self.hwp.SetPos(saved_pos[0], saved_pos[1], saved_pos[2])

        if text:
            text = text.replace('\r\n', ' ').replace('\r', ' ').replace('\n', ' ')
            text = re.sub(r' +', ' ', text)
        return text or ""

    def _adjust_spacing(self, para_id, line_idx, spacing):
        """줄 자간 조정"""
        self.block.select_line_by_index(para_id, line_idx)

        pset = self.hwp.HParameterSet.HCharShape
        self.hwp.HAction.GetDefault("CharShape", pset.HSet)
        pset.SpacingHangul = spacing
        pset.SpacingLatin = spacing
        self.hwp.HAction.Execute("CharShape", pset.HSet)
        self.block.cancel()

    def process_paragraph(self, para_id):
        """문단 처리하며 통계 수집"""
        line_starts, para_end = self._get_line_starts(para_id)
        total_lines = len(line_starts)

        if total_lines < 2:
            return

        line_idx = 1
        while line_idx < total_lines:
            # 현재 줄 정보 갱신
            line_starts, para_end = self._get_line_starts(para_id)
            if line_idx >= len(line_starts):
                break

            current_text = self._get_line_text(para_id, line_idx, line_starts, para_end)

            # 첫 공백 위치 확인
            first_space_idx = current_text.find(' ')
            if first_space_idx not in [1, 2, 3]:
                line_idx += 1
                continue

            # 이전 줄 텍스트
            prev_text = self._get_line_text(para_id, line_idx - 1, line_starts, para_end)

            # 정렬 방식 결정
            if first_space_idx in [1, 2]:
                align_type = 'reduce'
            elif first_space_idx == 3 and prev_text and prev_text[-1] != ' ':
                align_type = 'expand'
            else:
                line_idx += 1
                continue

            # 이전 줄이 이미 공백으로 끝나면 스킵 (reduce만)
            if align_type == 'reduce' and prev_text and prev_text[-1] == ' ':
                line_idx += 1
                continue

            # 자간 조정 시작
            iterations = 0
            success = False

            if align_type == 'reduce':
                # 1글자: -2, 2글자: -3으로 시작
                if first_space_idx == 1:
                    current_spacing = -2
                else:
                    current_spacing = -3

                # 첫 시도
                iterations = 1
                self._adjust_spacing(para_id, line_idx - 1, current_spacing)
                line_starts, para_end = self._get_line_starts(para_id)
                prev_text = self._get_line_text(para_id, line_idx - 1, line_starts, para_end)

                if prev_text and prev_text[-1] == ' ':
                    success = True
                else:
                    # 추가 시도
                    while current_spacing > -100 and iterations < 50:
                        iterations += 1
                        current_spacing -= 1
                        self._adjust_spacing(para_id, line_idx - 1, current_spacing)
                        line_starts, para_end = self._get_line_starts(para_id)

                        if line_idx - 1 >= len(line_starts):
                            break

                        prev_text = self._get_line_text(para_id, line_idx - 1, line_starts, para_end)
                        if prev_text and prev_text[-1] == ' ':
                            success = True
                            break

                if success:
                    stat_key = f'reduce_{first_space_idx}'
                    self.stats[stat_key].append(iterations)
                    print(f"  [성공] reduce {first_space_idx}글자: {iterations}회 (자간 {current_spacing})")
                else:
                    print(f"  [실패] reduce {first_space_idx}글자: {iterations}회 시도")

            else:  # expand
                current_spacing = 1
                while current_spacing <= 10 and iterations < 50:
                    iterations += 1
                    self._adjust_spacing(para_id, line_idx - 1, current_spacing)
                    line_starts, para_end = self._get_line_starts(para_id)

                    if line_idx - 1 >= len(line_starts):
                        break

                    prev_text = self._get_line_text(para_id, line_idx - 1, line_starts, para_end)
                    if prev_text and prev_text[-1] == ' ':
                        success = True
                        break

                    current_spacing += 1

                if success:
                    self.stats['expand_3'].append(iterations)
                    print(f"  [성공] expand 3글자: {iterations}회 (자간 +{current_spacing})")
                else:
                    print(f"  [실패] expand 3글자: {iterations}회 시도")

            # 다음 줄로 (줄 구조가 바뀌었을 수 있으므로 같은 인덱스 유지 또는 증가)
            line_starts, para_end = self._get_line_starts(para_id)
            total_lines = len(line_starts)
            line_idx += 1

    def print_stats(self):
        """통계 출력"""
        print("\n" + "=" * 60)
        print("자간 조정 통계")
        print("=" * 60)

        labels = {
            'reduce_1': '1글자 올림 (자간 줄임, 초기 -2)',
            'reduce_2': '2글자 올림 (자간 줄임, 초기 -3)',
            'expand_3': '3글자 (자간 늘림, 1글자 내림)',
        }

        for key, iterations_list in self.stats.items():
            label = labels.get(key, key)
            if iterations_list:
                avg = sum(iterations_list) / len(iterations_list)
                min_val = min(iterations_list)
                max_val = max(iterations_list)
                print(f"\n{label}:")
                print(f"  처리 횟수: {len(iterations_list)}회")
                print(f"  평균 반복: {avg:.1f}회")
                print(f"  최소 반복: {min_val}회")
                print(f"  최대 반복: {max_val}회")
                print(f"  상세: {iterations_list}")
            else:
                print(f"\n{label}: 처리된 케이스 없음")


if __name__ == "__main__":
    print("=" * 60)
    print("분리단어 처리 - 글자 수별 자간 조정 횟수 통계")
    print("=" * 60)

    # 한글 새 문서 열기
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
    hwp.XHwpWindows.Item(0).Visible = True
    set_hwp(hwp)

    print("\n[1] 마크다운 삽입...")
    markdown_to_hwp(test_markdown)

    print("\n[2] 분리단어 처리 및 통계 수집...")
    collector = SpacingStatsCollector(hwp)

    # 문서 처음으로 이동
    hwp.MovePos(2)

    # 모든 문단 처리
    para_count = 0
    while True:
        pos = hwp.GetPos()
        para_id = pos[1]

        print(f"\n문단 {para_id}:")
        collector.process_paragraph(para_id)
        para_count += 1

        # 다음 문단으로
        before_para = para_id
        hwp.HAction.Run("MoveNextParaBegin")
        after_pos = hwp.GetPos()

        if after_pos[1] == before_para:
            break

    print(f"\n총 {para_count}개 문단 처리")

    # 통계 출력
    collector.print_stats()
