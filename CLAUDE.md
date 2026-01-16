# 한글(HWP) 자동화 프로젝트 규칙

## 프로젝트 구조

```
hwp_docs/
├── separated_word.py      # 분리된 단어 처리 (SeparatedWord)
├── separated_para.py      # 분리된 문단 처리 (SeparatedPara)
├── style_para.py       # 스타일 관리 (StylePara)
├── style_numb.py   # 문단 번호 관리 (StyleNumb)
├── block_selector.py      # 블록 선택 유틸리티 (BlockSelector)
├── cursor_utils.py        # 커서 유틸리티 함수
├── styles.yaml            # 스타일 프리셋 정의
├── style_format.py        # 서식 관리 (StyleFormat)
├── table_manager.py       # 테이블 관리 (TableManager)
├── run.py                 # 메인 실행 스크립트
└── debugs/logs/           # 디버그 로그
```

## 핵심 모듈

### 1. separated_word.py - 분리된 단어 처리

한 단어가 두 줄에 걸쳐 분리된 경우 자간을 줄여서 한 줄로 합침

```python
from separated_word import SeparatedWord
from cursor_utils import get_hwp_instance

hwp = get_hwp_instance()
fixer = SeparatedWord(hwp, debug=True)
result = fixer.fix_paragraph()  # 현재 문단 처리
```

**독립 실행:** `python separated_word.py`

| 메서드 | 설명 |
|--------|------|
| `fix_paragraph()` | 현재 문단의 분리된 단어 처리 |
| `_adjust_spacing()` | 자간 조정 |

### 2. separated_para.py - 분리된 문단 처리

문단이 두 페이지에 걸쳐있을 때 한 페이지로 맞춤

```python
from separated_para import SeparatedPara
from cursor_utils import get_hwp_instance

hwp = get_hwp_instance()
fixer = SeparatedPara(hwp)
result = fixer.fix_page(page=1)  # 1페이지 처리
```

**독립 실행:** `python separated_para.py`

| 메서드 | 설명 |
|--------|------|
| `fix_page(page)` | 페이지의 분리된 문단 처리 (메인) |
| `fix_paragraph(para_id)` | 단일 문단 처리 |
| `fix_all_paragraphs(page)` | 모든 걸친 문단 처리 |
| `get_spanning_lines(para_id)` | 걸친 문단 줄 분석 |
| `remove_empty_line_at_page_start(page)` | 페이지 첫 빈줄 제거 |

**처리 전략 우선순위:**
1. `char_spacing_align` - 자간 줄이고 text_align (2회)
2. `empty_font` - 빈 문단 글자 크기 줄이기

### 3. style_para.py - 스타일 관리

글자 모양(CharShape), 문단 모양(ParaShape) 설정 및 조회

```python
from style_manager import StylePara
from cursor_utils import get_hwp_instance

hwp = get_hwp_instance()
style = StylePara(hwp)

# 글자 모양 설정
style.set_bold(True)
style.set_font_size(12)
style.set_text_color(255, 0, 0)  # RGB

# 문단 모양 설정
style.set_align('center')
style.set_line_spacing(160)
```

**독립 실행:** `python style_para.py`

| 메서드 | 설명 |
|--------|------|
| `set_bold(enabled)` | 굵게 설정 |
| `set_italic(enabled)` | 기울임 설정 |
| `set_underline(enabled)` | 밑줄 설정 |
| `set_font_size(pt)` | 글자 크기 (pt) |
| `set_font(hangul, latin)` | 글꼴 설정 |
| `set_text_color(r, g, b)` | 글자 색상 (RGB) |
| `get_char_shape()` | 현재 글자 모양 조회 |
| `set_align(type)` | 문단 정렬 (left/center/right/justify) |
| `set_line_spacing(value)` | 줄간격 설정 |
| `set_para_margin(left, right, indent)` | 문단 여백 |
| `get_para_shape()` | 현재 문단 모양 조회 |
| `copy_char_shape()` | 글자 모양 복사 |
| `copy_para_shape()` | 문단 모양 복사 |
| `paste_shape()` | 모양 붙여넣기 |

**유틸리티:**
- `StylePara.rgb_to_bgr(r, g, b)` - RGB → BGR 변환
- `StylePara.pt_to_hwpunit(pt)` - pt → HWPUNIT 변환 (1pt = 100)

### 4. style_numb.py - 문단 번호 관리

번호 매기기, 글머리표, 개요 번호 설정, 마크다운 헤딩 변환

```python
from style_numb import StyleNumb
from cursor_utils import get_hwp_instance

hwp = get_hwp_instance()
numbering = StyleNumb(hwp)

# 번호 매기기
numbering.set_numbering('number')    # 1. 2. 3.
numbering.set_numbering('hangul')    # 가. 나. 다.

# 글머리표
numbering.set_bullet('circle')       # ●
numbering.set_bullet('square')       # ■

# 개요 수준
numbering.set_outline_level(1)       # 1수준
numbering.set_outline_level(2)       # 2수준

# 마크다운 헤딩 처리 (styles.yaml의 heading_levels 설정 사용)
text = '''
# 1장 서론
## 1.1 배경
### 1.1.1 세부 내용
# 2장 본론
## 2.1 방법
'''
numbering.process_markdown_text(text)

# 해제
numbering.remove()
```

**독립 실행:** `python style_numb.py`

| 메서드 | 설명 |
|--------|------|
| `set_numbering(format)` | 번호 매기기 (number/hangul/circled/roman) |
| `set_bullet(type)` | 글머리표 (circle/square/diamond/dash) |
| `set_custom_bullet(char)` | 사용자 정의 글머리표 |
| `set_outline_level(level)` | 개요 수준 설정 (1~7) |
| `increase_level()` | 수준 증가 (들여쓰기) |
| `decrease_level()` | 수준 감소 (내어쓰기) |
| `remove()` | 번호/글머리표/개요 해제 |
| `parse_heading_level(text)` | 마크다운 헤딩 레벨 파싱 |
| `insert_heading(text)` | 마크다운 헤딩 삽입 |
| `process_markdown_text(text)` | 마크다운 텍스트 전체 처리 (새로 입력) |
| `apply_heading_to_document()` | **문서 전체**에서 # 찾아 개요 적용 |
| `apply_heading_to_selection()` | **선택 영역**에서 # 찾아 개요 적용 |
| `apply_heading_to_current_para()` | **현재 문단**에 개요 적용 |
| `get_heading_config(level)` | YAML 헤딩 설정 조회 |
| `set_outline_format(style)` | **개요 번호 양식 설정** (1.1.1 형식 등) |
| `scan_headings()` | 문서 전체 마크다운 헤딩 스캔 및 계층 번호 생성 |
| `get_heading_number_string(numbering)` | 번호 배열을 문자열로 변환 |
| `apply_heading_numbering()` | 헤딩 스캔 후 개요 수준 일괄 적용 |

**번호 형식:** `number`, `hangul`, `circled`, `roman_upper`, `roman_lower`, `alpha_upper`, `alpha_lower`, `parenthesis`

**글머리표 형식:** `circle`, `square`, `diamond`, `arrow`, `check`, `dash`, `asterisk`

**마크다운 헤딩 매핑 (styles.yaml):**
- `#` → 개요 1수준 (h1)
- `##` → 개요 2수준 (h2)
- `###` → 개요 3수준 (h3)
- `####` → 개요 4수준 (h4)
- `#####` → 개요 5수준 (h5)
- `######` → 개요 6수준 (h6)
- `#######` → 개요 7수준 (h7)

### 개요 번호 양식 설정 (set_outline_format)

HWP의 개요 번호 양식을 API로 설정합니다. 기본 개요 양식(1., 가., 1)...)을 법률 형식(1.1.1) 등으로 변경할 수 있습니다.

```python
from style_numb import StyleNumb
from cursor_utils import get_hwp_instance

hwp = get_hwp_instance()
numb = StyleNumb(hwp)

# 개요 양식 설정 (문서에 개요 적용 전에 호출)
numb.set_outline_format('legal')  # 1. → 1.1 → 1.1.1 형식

# 이후 마크다운 헤딩 변환
text = '''
# 서론
## 배경
### 세부 내용
'''
numb.process_markdown_text(text)
```

**지원 스타일:**

| 스타일 | 1수준 | 2수준 | 3수준 | 4수준 | 설명 |
|--------|-------|-------|-------|-------|------|
| `legal` | 1. | 1.1 | 1.1.1 | 1.1.1.1 | 법률/논문 형식 |
| `decimal` | 1 | 1.1 | 1.1.1 | 1.1.1.1 | 숫자만 (마침표 없음) |
| `korean` | 1. | 가. | 1) | 가) | 한글 기본 형식 |
| `article` | 제1장 | 제1절 | 제1조 | ① | 조항 형식 |

**API 상세 (내부 구현):**

```python
# OutlineNumber 액션 + SecDef 파라미터셋 사용
act = hwp.CreateAction("OutlineNumber")
sec_set = act.CreateSet("SecDef")
outline_set = sec_set.Item("OutlineShape")

# 각 수준별 형식 문자열 설정
outline_set.SetItem("StrFormatLevel0", "^1.")      # 1수준: "1."
outline_set.SetItem("StrFormatLevel1", "^1.^2")    # 2수준: "1.1"
outline_set.SetItem("StrFormatLevel2", "^1.^2.^3") # 3수준: "1.1.1"

# 번호 형식 설정 (숫자/가나다/원문자 등)
outline_set.SetItem("NumFormatLevel0", 0)  # 0=1,2,3
outline_set.SetItem("NumFormatLevel1", 0)  # 8=가,나,다
outline_set.SetItem("NumFormatLevel2", 0)  # 1=①②③

act.Execute(sec_set)
```

**형식 문자열 (StrFormatLevelN):**
- `^1` = 1수준 번호
- `^2` = 2수준 번호
- `^3` = 3수준 번호
- 예: `"^1.^2.^3"` → "1.2.3" 형식

**번호 형식 (NumFormatLevelN):**
| 값 | 형식 | 예시 |
|----|------|------|
| 0 | 아라비아 숫자 | 1, 2, 3 |
| 1 | 원문자 | ①, ②, ③ |
| 8 | 가나다 | 가, 나, 다 |

### 개요 번호 vs 스타일 - 우선순위와 역할 분담

개요 번호와 스타일은 모두 '문단 모양'과 '글자 모양' 정보를 포함하고 있어서 설정이 충돌할 수 있습니다.

**적용 우선순위:**
- **글자 모양:** 스타일 설정 우선
- **문단 모양 (여백/들여쓰기):** **개요 번호 설정이 스타일보다 우선**
  - 예: 스타일에 '왼쪽 여백 0'이어도 개요 번호에서 '1수준 왼쪽 여백 10pt'면 실제로 10pt 적용

**각 기능의 역할:**

| 구분 | 개요 번호 (기능 중심) | 스타일 (디자인 중심) |
|------|----------------------|---------------------|
| **핵심 정보** | 번호의 체계와 수준(Level) | 글꼴, 크기, 색상, 정렬 등 외형 |
| **데이터** | 1, 가, (1) 등 번호 모양 및 자동 증가 | 명조/고딕, 12pt, 진하게 등 시각 효과 |
| **위계** | 상위/하위 번호 관계(목차 구조) | 문서 전체의 일관된 테마 |

**개요 번호에 문단 모양이 있는 이유:**
- 번호가 '1.'일 때와 '제 1장'일 때 본문 시작 위치가 달라야 함
- 번호 모양에 최적화된 위치 정보를 개요 번호가 직접 관리

**권장 사용 원칙:**
1. **위치(여백)는 개요 번호에서:** 들여쓰기, 번호 간격은 [개요 번호 모양]에서 관리
2. **외형(글꼴)은 스타일에서:** 글자 크기, 색상은 [스타일]에서 관리
3. **연결해서 사용:** [스타일 편집] → [글자 문단 번호/개요] 버튼으로 스타일과 개요 수준을 묶어서 사용

**여백이 의도대로 안 될 때:** `Ctrl + K, O`로 개요 번호 모양의 '사용자 정의'에서 여백 설정 확인

---

## 참고 문서 (필수!)

HWP API 관련 작업 시 **반드시** 아래 문서를 먼저 확인:

### 로컬 문서 (D:\hwp_docs\win32\)
- `ActionTable_2504_part01~06.md` - 액션 테이블 (사용 가능한 액션 목록)
- `ParameterSetTable_2504_part01~18.md` - 파라미터셋 (HCharShape, HTableCreation 등)
- `HwpAutomation_2504_part01~07.md` - 메서드/속성 API

### 웹 검색
API 사용법이 불확실하거나 로컬 문서에 없으면 웹 검색으로 확인

---

## 필수: 보안 모듈 등록

한글 자동화 코드 작성 시 **반드시** 보안 모듈을 등록해야 합니다.

```python
import win32com.client as win32

hwp = win32.Dispatch('HWPFrame.HwpObject')
hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')  # 필수!
```

## 기존 인스턴스 연결 (ROT)

이미 실행 중인 한글에 연결하려면 **Running Object Table(ROT)** 사용:

```python
from cursor_utils import get_hwp_instance

hwp = get_hwp_instance()
if not hwp:
    print("한글이 실행 중이 아닙니다")
```

**ROT 방식의 특징:**
- 이미 실행 중인 한글 프로세스에 연결
- 새 인스턴스를 생성하지 않음
- 한글이 먼저 열려 있어야 함

## 주요 API 패턴

### 테이블 생성
- 액션 이름: `TableCreate` (TableCreation 아님)
- 파라미터 전달 시 `.HSet` 사용 필수

### 텍스트 삽입
```python
hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
hwp.HParameterSet.HInsertText.Text = "텍스트"
hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
```

### 서식 변경 (필수: 블록 지정 먼저!)
```python
hwp.HAction.Run("SelectAll")  # 또는 hwp.SelectText(para, pos, para, pos)
# 서식 변경 후
hwp.HAction.Run("Cancel")  # 선택 해제
```

## 커서 위치 관련

### 위치 정보 API
- `hwp.GetPos()` → (list_id, para_id, char_pos)
- `hwp.SetPos(list, para, pos)` → 커서 이동
- `hwp.KeyIndicator()` → (총구역, 현재구역, 페이지, 단, 줄, 칸, 수정모드)

### 커서 이동 액션
- `MoveParaBegin` / `MoveParaEnd` - 문단 시작/끝
- `MoveRight` / `MoveLeft` - 한 글자 이동
- `MoveDown` / `MoveUp` - 한 줄 이동
- `MoveSelRight` - 선택하며 이동

### 텍스트 가져오기
```python
hwp.HAction.Run("MoveSelRight")  # 선택
text = hwp.GetTextFile("TEXT", "saveblock")  # 선택 영역만
hwp.HAction.Run("Cancel")  # 선택 해제
```

## cursor_utils.py 함수 요약

| 함수 | 인자 | 반환값 |
|------|------|--------|
| `get_hwp_instance()` | 없음 | hwp 객체 또는 None |
| `get_current_pos(hwp)` | hwp | `{list_id, para_id, char_pos, page, line, column, insert_mode}` |
| `get_para_range(hwp)` | hwp | `{current: (l,p,c), start: int, end: int}` |
| `get_line_range(hwp)` | hwp | `{current: (l,p,c), start: int, end: int, line_starts: []}` |
| `get_sentences(hwp, include_text=False)` | hwp, bool | `[{index, start, end}, ...]` 또는 `(list, text)` |
| `get_cursor_index(hwp, pos=None)` | hwp, int/None | `{sentence_index, word_index, sentence_start, sentence_end}` |

## block_selector.py - BlockSelector 클래스

블록(범위) 선택 전용 클래스. HWP pos와 텍스트 인덱스가 다름에 주의 (한글=pos 2, 영문/공백=pos 1).

### 메서드 요약

| 메서드 | 인자 | 설명 |
|--------|------|------|
| `select_para(para_id)` | 문단ID | 문단 전체 선택 |
| `select_line_by_index(para_id, line_index)` | 문단ID, 줄번호(0~) | n번째 줄 선택 |
| `select_line_by_pos(para_id, pos)` | 문단ID, pos | pos가 속한 줄 선택 |
| `select_lines_range(para_id, start, end)` | 문단ID, 시작줄, 끝줄 | n~m번째 줄 선택 |
| `select_sentence(para_id, sentence_index)` | 문단ID, 문장번호(1~) | n번째 문장 선택 |
| `select_sentences_range(para_id, start, end)` | 문단ID, 시작, 끝 | n~m번째 문장 선택 |
| `select_sentence_in_line(para_id, pos)` | 문단ID, pos | pos가 속한 문장 선택 |
| `cancel()` | 없음 | 블록 선택 해제 |
| `get_selected_text()` | 없음 | 선택된 텍스트 반환 |

### 사용 예시
```python
from block_selector import BlockSelector

block = BlockSelector(hwp)
pos = hwp.GetPos()
para_id = pos[1]

block.select_sentence(para_id, 1)  # 첫 번째 문장 선택
text = block.get_selected_text()
block.cancel()
```

## 주의사항

### 이벤트 기반 vs 폴링
- `DocumentChange` 이벤트: API 커서 이동도 발생 → 무한 루프 위험
- **폴링 방식 권장**: 0.1초 간격으로 `GetPos()` 비교

### 위치 정보 읽을 때
- 문단/줄 범위를 알려면 커서를 이동해야 함
- 이동 후 반드시 원래 위치로 복원: `hwp.SetPos(list, para, pos)`

## 모듈 마이그레이션 가이드

| 기존 | 새 이름 |
|------|---------|
| `text_align.py` / `TextAlign` | `separated_word.py` / `SeparatedWord` |
| `text_align_page.py` / `TextAlignPage` | `separated_para.py` / `SeparatedPara` |
| `custom_block.py` / `CustomBlock` | `block_selector.py` / `BlockSelector` |
| `cursor_position_monitor.py` | `cursor_utils.py` |

```python
# 기존 import
from text_align import TextAlign
from text_align_page import TextAlignPage
from custom_block import CustomBlock
from cursor_position_monitor import get_hwp_instance

# 새 import
from separated_word import SeparatedWord
from separated_para import SeparatedPara
from block_selector import BlockSelector
from cursor_utils import get_hwp_instance
```
