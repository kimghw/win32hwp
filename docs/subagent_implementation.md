# HWP API 문서 검색 서브에이전트 구현 가이드

## 개요

HWP API 문서(`/mnt/c/win32hwp/win32/`)를 검색하는 서브에이전트를 구현했습니다.
메인 컨텍스트를 절약하면서 병렬 검색으로 속도를 향상시키는 것이 목표입니다.

---

## 1. YAML 에이전트 파일 방식

`.claude/agents/` 디렉토리에 YAML 파일을 생성하여 서브에이전트를 등록합니다.

### 디렉토리 구조

```
.claude/
└── agents/
    ├── action-searcher-1.yaml    # ActionTable part01~03
    ├── action-searcher-2.yaml    # ActionTable part04~06
    ├── param-searcher-1.yaml     # ParameterSetTable part01~09
    ├── param-searcher-2.yaml     # ParameterSetTable part10~18
    └── automation-searcher.yaml  # HwpAutomation part01~07
```

### YAML 파일 예시

```yaml
# .claude/agents/action-searcher-1.yaml
name: action-searcher-1
description: HWP 액션 테이블 검색 (ActionTable part01~03)

model: haiku

tools:
  - Read
  - Grep

prompt: |
  HWP 액션 테이블 검색 전문가입니다.

  **담당 파일 (이 파일들만 검색하세요):**
  - /mnt/c/win32hwp/win32/ActionTable_2504_part01.md
  - /mnt/c/win32hwp/win32/ActionTable_2504_part02.md
  - /mnt/c/win32hwp/win32/ActionTable_2504_part03.md

  Grep으로 담당 파일에서만 키워드 검색 후 Read로 내용 읽기.
  관련 내용 없으면 "part01-03에서 없음" 답변.
```

### 특징

- `model: haiku` - 빠른 응답을 위해 haiku 모델 사용
- `tools` - 필요한 도구만 허용 (Read, Grep)
- 담당 파일을 명시하여 검색 범위 제한

---

## 2. subprocess 병렬 실행 방식 (권장)

Claude CLI를 subprocess로 실행하여 **진짜 병렬 처리**를 구현합니다.

### 핵심 코드

```python
# hwp_api_search_agent.py
import asyncio
import subprocess

DOCS_PATH = "/mnt/c/win32hwp/win32"
CWD_PATH = "/mnt/c/win32hwp"

# 5개 에이전트 설정 (파일을 1/5로 분할)
AGENTS = {
    "action-1": {
        "files": ["ActionTable_2504_part01.md", "ActionTable_2504_part02.md",
                  "ActionTable_2504_part03.md"],
        "desc": "ActionTable part01~03"
    },
    "action-2": {
        "files": ["ActionTable_2504_part04.md", "ActionTable_2504_part05.md",
                  "ActionTable_2504_part06.md"],
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
    print(f"[5개 프로세스 병렬] 검색: {user_query}")

    # 5개 프로세스 동시 시작
    tasks = [
        run_agent_async(name, config, user_query)
        for name, config in AGENTS.items()
    ]

    results = await asyncio.gather(*tasks)

    # 결과 출력
    for r in results:
        print(f"### [{r['name']}]")
        print(r['result'][:2000])
```

### 사용법

```bash
# 단일 검색
python3 hwp_api_search_agent.py "테이블 생성"

# 대화형 모드
python3 hwp_api_search_agent.py --interactive
```

---

## 3. 성능 비교

| 구성 | 소요 시간 | 비고 |
|------|----------|------|
| 단일 에이전트 (sonnet) | 75.94초 | 모든 파일 순차 검색 |
| 5개 병렬 - SDK Task 도구 | 108.49초 | 오히려 느림 |
| 5개 병렬 - SDK asyncio.gather | 78.08초 | 내부적으로 순차 처리 |
| **5개 프로세스 병렬 - subprocess** | **38.37초** | **가장 빠름 (2배 향상)** |

### 결론

- SDK의 `asyncio.gather()`는 내부적으로 순차 처리되어 병렬 효과가 없음
- **subprocess로 claude CLI를 직접 호출**해야 진짜 병렬 실행 가능
- 파일을 분할하여 각 에이전트가 담당 파일만 검색하도록 구성

---

## 4. Claude CLI 옵션

```bash
claude -p "프롬프트" \
    --allowedTools "Read,Grep" \
    --output-format text
```

| 옵션 | 설명 |
|------|------|
| `-p` | 프롬프트 전달 (비대화형 모드) |
| `--allowedTools` | 허용할 도구 목록 (쉼표 구분) |
| `--output-format` | 출력 형식 (`text`, `json`) |

---

## 5. 파일 분할 전략

총 31개 파일을 5개 그룹으로 분할:

| 에이전트 | 담당 파일 | 파일 수 |
|----------|----------|---------|
| action-1 | ActionTable part01~03 | 3 |
| action-2 | ActionTable part04~06 | 3 |
| param-1 | ParameterSetTable part01~09 | 9 |
| param-2 | ParameterSetTable part10~18 | 9 |
| automation | HwpAutomation part01~07 | 7 |

---

## 6. 컨텍스트 절약 효과

- 메인 에이전트: 검색 쿼리만 전달
- 서브 에이전트: 문서 내용을 직접 읽고 처리
- 결과: 요약된 검색 결과만 메인으로 반환
- 효과: 메인 컨텍스트에 대용량 문서가 로드되지 않음
