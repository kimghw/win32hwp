"""단일 에이전트 버전 (비교용)"""
import asyncio
import sys
import time
from claude_agent_sdk import query, ClaudeAgentOptions, AgentDefinition

DOCS_PATH = "/mnt/c/win32hwp/win32"
CWD_PATH = "/mnt/c/win32hwp"

single_agent = AgentDefinition(
    description="HWP API 문서 검색 전문가",
    prompt=f"""당신은 HWP API 문서 검색 전문가입니다.

**검색 대상:** {DOCS_PATH}

**파일 종류:**
1. ActionTable_2504_part01~06.md - 액션 테이블
2. ParameterSetTable_2504_part01~18.md - 파라미터셋
3. HwpAutomation_2504_part01~07.md - API

Grep으로 키워드 찾고, Read로 내용 읽어서 정리하세요.""",
    tools=["Read", "Grep", "Glob"],
    model="sonnet"
)

async def search(user_query: str):
    start = time.time()
    print(f"[단일 에이전트] 검색: {user_query}")
    print("=" * 60)

    async for message in query(
        prompt=f"HWP API 문서에서 '{user_query}' 관련 내용을 검색하세요. ActionTable, ParameterSetTable, HwpAutomation 모든 파일을 검색하세요.",
        options=ClaudeAgentOptions(
            allowed_tools=["Read", "Grep", "Glob", "Task"],
            agents={"searcher": single_agent},
            cwd=CWD_PATH
        )
    ):
        if hasattr(message, "result"):
            print(message.result)

    elapsed = time.time() - start
    print(f"\n[단일 에이전트] 소요 시간: {elapsed:.2f}초")
    return elapsed

if __name__ == "__main__":
    query_text = " ".join(sys.argv[1:]) if len(sys.argv) > 1 else "커서 이동"
    asyncio.run(search(query_text))
