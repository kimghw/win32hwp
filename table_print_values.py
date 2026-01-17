"""
테이블 셀을 순회하며 내부 값 출력
"""
import win32com.client as win32
from cursor_utils import get_hwp_instance
import logging
from datetime import datetime
import os


def print_table_values():
    """
    커서 위치의 테이블을 첫 번째 셀부터 오른쪽으로 순회하며 값 출력
    """
    # 로그 설정
    log_dir = "debugs/logs"
    os.makedirs(log_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = f"{log_dir}/table_values_{timestamp}.log"

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    logger = logging.getLogger(__name__)

    hwp = get_hwp_instance()
    if not hwp:
        logger.error("한글이 실행되지 않았습니다")
        return

    # 보안 모듈 등록
    hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')

    # 원래 위치 저장
    original_pos = hwp.GetPos()

    # 첫 번째 셀로 이동
    hwp.MovePos(106, 0, 0)  # MOVE_TOP_OF_CELL (첫 행)
    hwp.MovePos(104, 0, 0)  # MOVE_START_OF_CELL (첫 열)

    # 첫 번째 셀 위치 확인
    first_cell_pos = hwp.GetPos()

    logger.info("=== 테이블 값 출력 시작 ===")
    logger.info(f"첫 번째 셀 위치: {first_cell_pos}")

    # 테이블 셀의 list_id 범위 저장
    table_list_ids = set()
    cell_index = 0

    # 한 번만 순회
    while True:
        current_pos = hwp.GetPos()
        current_list_id = current_pos[0]

        # 현재 list_id를 테이블 범위에 추가
        table_list_ids.add(current_list_id)

        # 현재 셀 값 가져오기
        hwp.HAction.Run("TableCellBlock")
        current_text = hwp.GetTextFile("TEXT", "saveblock")
        hwp.HAction.Run("Cancel")
        hwp.SetPos(current_pos[0], current_pos[1], current_pos[2])
        # 개행 문자를 공백으로 변환하고 trim
        current_text = current_text.replace('\r', ' ').replace('\n', ' ').strip() if current_text else ""

        # 위 셀 list_id
        hwp.MovePos(102, 0, 0)  # MOVE_UP_OF_CELL
        up_pos = hwp.GetPos()
        up_id = 0 if up_pos == current_pos else up_pos[0]
        hwp.SetPos(current_pos[0], current_pos[1], current_pos[2])

        # 아래 셀 list_id
        hwp.MovePos(103, 0, 0)  # MOVE_DOWN_OF_CELL
        down_pos = hwp.GetPos()
        down_id = 0 if down_pos == current_pos else down_pos[0]
        hwp.SetPos(current_pos[0], current_pos[1], current_pos[2])

        # 좌 셀 list_id
        hwp.MovePos(100, 0, 0)  # MOVE_LEFT_OF_CELL
        left_pos = hwp.GetPos()
        left_id = 0 if left_pos == current_pos else left_pos[0]
        hwp.SetPos(current_pos[0], current_pos[1], current_pos[2])

        # 우 셀 list_id
        hwp.MovePos(101, 0, 0)  # MOVE_RIGHT_OF_CELL
        right_pos = hwp.GetPos()
        right_id = 0 if right_pos == current_pos else right_pos[0]
        hwp.SetPos(current_pos[0], current_pos[1], current_pos[2])

        # 현재 셀 정보 로그
        logger.info(f"{current_text}[{up_id}, {down_id}, {left_id}, {right_id}]")

        # 오른쪽으로 이동 시도
        hwp.MovePos(101, 0, 0)  # MOVE_RIGHT_OF_CELL
        next_pos = hwp.GetPos()

        # 더 이상 오른쪽으로 못 가면
        if next_pos == current_pos:
            # 아래로 이동 시도
            hwp.MovePos(103, 0, 0)  # MOVE_DOWN_OF_CELL
            down_move_pos = hwp.GetPos()

            # 아래로도 못 가면 종료
            if down_move_pos == current_pos:
                logger.info(f"\n총 {cell_index + 1}개 셀 출력 완료")
                logger.info(f"테이블 list_id 범위: {sorted(table_list_ids)}")
                logger.info(f"로그 파일: {log_file}")
                break

            # 아래로 이동했으면 맨 왼쪽으로
            hwp.MovePos(104, 0, 0)  # MOVE_START_OF_CELL

        cell_index += 1

    # 원래 위치로 복원
    hwp.SetPos(original_pos[0], original_pos[1], original_pos[2])


if __name__ == "__main__":
    print_table_values()
