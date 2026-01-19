"""HWP API 문서 검색 에이전트 (5개 프로세스 병렬)

subprocess로 5개 프로세스를 동시에 실행합니다.

사용법:
    python3 hwp_api_search_agent.py "테이블 생성"
    python3 hwp_api_search_agent.py --interactive
"""
import asyncio
import subprocess
import sys
import time
import json
import os

DOCS_PATH = "/mnt/c/win32hwp/win32"
CWD_PATH = "/mnt/c/win32hwp"

# 5개 에이전트 설정
AGENTS = {
    "action-1": {
        "files": ["ActionTable_2504_part01.md", "ActionTable_2504_part02.md", "ActionTable_2504_part03.md"],
        "desc": "ActionTable part01~03"
    },
    "action-2": {
        "files": ["ActionTable_2504_part04.md", "ActionTable_2504_part05.md", "ActionTable_2504_part06.md"],
        "desc": "ActionTable part04~06"
    },
    "param-1": {
        "files": [f"ParameterSetTable_2504_part0{i}.md" for i in range(1, 10)],
        "desc": "ParameterSetTable part01~09"
    },
    "param-2": {
        "files": [f"ParameterSetTable_2504_part{i}.md" for i in range(10, 19)],
        "desc": "ParameterSetTable part10~18"
    },
    "automation": {
        "files": [f"HwpAutomation_2504_part0{i}.md" for i in range(1, 8)],
        "desc": "HwpAutomation part01~07"
    },
}


def run_single_agent(name: str, config: dict, user_query: str) -> dict:
    """단일 에이전트 실행 (subprocess)"""
    files_str = "\n".join([f"{DOCS_PATH}/{f}" for f in config['files']])

    prompt = f'''"{user_query}" 키워드로 검색하세요.

담당 파일:
{files_str}

Grep으로 키워드 검색 후 Read로 내용 읽기.
관련 없으면 "{config["desc"]}: 없음" 답변.'''

    # claude CLI 직접 호출
    cmd = [
        "claude",
        "-p", prompt,
        "--allowedTools", "Read,Grep",
        "--output-format", "text"
    ]

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120,
            cwd=CWD_PATH
        )
        return {"name": name, "result": result.stdout.strip(), "error": result.stderr}
    except subprocess.TimeoutExpired:
        return {"name": name, "result": f"{config['desc']}: 타임아웃", "error": ""}
    except Exception as e:
        return {"name": name, "result": f"{config['desc']}: 오류 - {e}", "error": str(e)}


async def run_agent_async(name: str, config: dict, user_query: str) -> dict:
    """비동기로 subprocess 실행"""
    files_str = "\n".join([f"{DOCS_PATH}/{f}" for f in config['files']])

    prompt = f'''"{user_query}" 키워드로 검색하세요.

담당 파일:
{files_str}

Grep으로 키워드 검색 후 Read로 내용 읽기.
관련 없으면 "{config["desc"]}: 없음" 답변.'''

    cmd = [
        "claude",
        "-p", prompt,
        "--allowedTools", "Read,Grep",
        "--output-format", "text"
    ]

    try:
        proc = await asyncio.create_subprocess_exec(
            *cmd,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
            cwd=CWD_PATH
        )
        stdout, stderr = await asyncio.wait_for(proc.communicate(), timeout=120)
        return {"name": name, "result": stdout.decode().strip()}
    except asyncio.TimeoutError:
        return {"name": name, "result": f"{config['desc']}: 타임아웃"}
    except Exception as e:
        return {"name": name, "result": f"{config['desc']}: 오류 - {e}"}


async def parallel_search(user_query: str):
    """5개 프로세스 동시 실행"""
    start = time.time()
    print(f"[5개 프로세스 병렬] 검색: {user_query}")
    print("=" * 60)

    # 5개 프로세스 동시 시작
    tasks = [
        run_agent_async(name, config, user_query)
        for name, config in AGENTS.items()
    ]

    results = await asyncio.gather(*tasks)

    # 결과 출력
    print("\n## 검색 결과 종합\n")
    for r in results:
        print(f"### [{r['name']}]")
        print(r['result'][:2000] if len(r['result']) > 2000 else r['result'])
        print("-" * 40)

    elapsed = time.time() - start
    print(f"\n[5개 프로세스 병렬] 소요 시간: {elapsed:.2f}초")
    return elapsed


async def interactive_mode():
    """대화형 모드"""
    print("=" * 60)
    print("HWP API 문서 검색 (5개 프로세스 병렬)")
    print("=" * 60)
    print("명령어: 'exit' 종료")
    print("-" * 60)

    while True:
        try:
            user_input = input("\n검색> ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n종료합니다.")
            break

        if user_input.lower() in ('exit', 'quit', 'q'):
            break

        if not user_input:
            continue

        await parallel_search(user_input)


def main():
    if len(sys.argv) > 1:
        if sys.argv[1] in ('--interactive', '-i'):
            asyncio.run(interactive_mode())
        else:
            user_query = " ".join(sys.argv[1:])
            asyncio.run(parallel_search(user_query))
    else:
        asyncio.run(interactive_mode())


if __name__ == "__main__":
    main()
