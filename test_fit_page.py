"""페이지 걸친 문단 처리 테스트"""

from cursor_position_monitor import get_hwp_instance
from text_align_page import TextAlignPage

# 로그 파일 설정
log_file = open("debugs/logs/fit_page.log", "w", encoding="utf-8")

def log(msg):
    print(msg)
    log_file.write(msg + "\n")
    log_file.flush()

hwp = get_hwp_instance()
if not hwp:
    log("[ERROR] 한글을 찾을 수 없습니다.")
    exit()

log("[OK] 한글 연결됨")

helper = TextAlignPage(hwp)

# 1. 전체 문단-페이지 매핑 조회
log("\n" + "="*60)
log("1. 문단-페이지 매핑 조회")
log("="*60)

para_map = helper.ParaAlignWords()
log(f"전체 문단 수: {len(para_map)}개")

# 걸친 문단 찾기
spanning_count = 0
for para_id, info in para_map.items():
    if info['start_page'] != info['end_page']:
        spanning_count += 1
        log(f"  걸침: para_id={para_id}, page {info['start_page']}~{info['end_page']}, empty={info.get('is_empty', False)}")

log(f"걸친 문단: {spanning_count}개")

if spanning_count == 0:
    log("\n걸친 문단이 없습니다.")
    log_file.close()
    exit()

# 2. 1페이지 처리
log("\n" + "="*60)
log("2. 1페이지 걸친 문단 처리 (빈문단 글자크기 전략)")
log("="*60)

# log_callback 전달하여 실시간 로그 출력
result = helper.fit_page_spanning(1, max_iterations=100, strategy='empty_font', log_callback=log)

log("\n" + "-"*40)
log("최종 결과:")
log(f"  성공: {result['success']}")
log(f"  반복 횟수: {result['iterations']}")
log(f"  대상 문단: para_id={result['para_id']}")

if 'original_lines_info' in result:
    log(f"  원래 줄분포: {result['original_lines_info']['lines_per_page']}")
if 'final_lines_info' in result:
    log(f"  최종 줄분포: {result['final_lines_info']['lines_per_page']}")

log("\n테스트 완료")
log_file.close()
