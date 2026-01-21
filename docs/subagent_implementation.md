# HWP API 문서 검색 서브에이전트 구현 가이드

## 개요

HWP API 문서(`/mnt/c/win32hwp/win32/`)를 검색하는 서브에이전트를 구현했습니다.
메인 컨텍스트를 절약하면서 병렬 검색으로 속도를 향상시키는 것이 목표입니다.

---

## 프로젝트 구조

```
/mnt/c/win32hwp/
├── cursor.py                    # HWP ROT 연결, 커서/위치/상태 통합 모듈
├── table/
│   ├── __init__.py
│   ├── table_info.py            # 셀 BFS 탐색, 좌표 매핑 (TableInfo 클래스)
│   ├── table_boundary.py        # 테이블 경계 판별, 인접 관계 (TableBoundaryResult)
│   ├── table_cell_info.py       # 셀 순회 유틸리티, 컨트롤 탐색
│   └── table_field.py           # 필드 CRUD, 셀 좌표 연동 (TableField 클래스)
├── style/
│   ├── style_format.py          # 글자 모양 (CharShape) 조회/설정
│   ├── style_para.py            # 문단 모양 (ParaShape), 스타일 관리 클래스
│   └── style_numb.py            # 번호/글머리 기호 스타일
├── example/
│   ├── run_table_cell_info.py   # 테이블 셀 정보 실행 예제
│   ├── run_table_field.py       # 테이블 필드 실행 예제
│   ├── run_md_to_hwp.py         # MD→HWP 변환 실행 예제
│   └── create_hwp_from_md.py    # MD→HWP 생성 예제
├── hwp_analysis/
│   └── auto_insert_fields.py    # 자동 필드 삽입 분석
├── hwp_api_search_agent.py      # HWP API 병렬 검색 에이전트 (subprocess)
├── hwp_api_search_single.py     # HWP API 단일 검색
├── md_to_hwp.py                 # 마크다운→HWP 변환
├── map_coordinates_to_table.py  # 셀 좌표 디버그 스크립트
├── measure_cell_pos.py          # 셀 위치 측정 스크립트
├── separated_para.py            # 분리된 문단 처리
├── separated_word.py            # 분리된 단어 처리
└── block_selector.py            # 블록 선택 유틸리티
```

### 핵심 모듈 설명

| 모듈 | 주요 클래스/함수 | 설명 |
|------|-----------------|------|
| `cursor.py` | `get_hwp_instance()`, `Cursor` | ROT에서 HWP 인스턴스 연결, 커서 위치 관리 |
| `table/table_info.py` | `TableInfo`, `CellInfo` | BFS로 테이블 구조 탐지, (row, col)→list_id 매핑 |
| `table/table_boundary.py` | `TableBoundaryResult`, `SubTableResult` | 테이블 경계(첫/마지막 행열) 판별 |
| `table/table_cell_info.py` | `iterate_table_cells()` | 테이블 셀 순회 콜백 패턴 |
| `table/table_field.py` | `TableField`, `FieldInfo` | 필드 등록/조회/삭제/변경, 셀 좌표 연동 |
| `style/style_format.py` | - | 글자 모양(굵게, 기울임, 색상 등) 조회/설정 |
| `style/style_para.py` | `StylePara` | 문단 모양(정렬, 줄간격 등), 모양 복사/붙여넣기 |

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
