# 한글(HWP) 자동화 프로젝트 규칙

## 프로젝트 구조

```
hwp_docs/
├── separated_word.py      # 분리된 단어 처리 (SeparatedWord)
├── separated_para.py      # 분리된 문단 처리 (SeparatedPara)
├── style_para.py          # 스타일 관리 (StylePara)
├── style_numb.py          # 문단 번호 관리 (StyleNumb)
├── block_selector.py      # 블록 선택 유틸리티 (BlockSelector)
├── cursor_wrapper.py      # 커서 조작 통합 래퍼 (CursorWrapper) ⭐ 권장
├── cursor_utils.py        # 커서 유틸리티 함수 (저수준)
├── get_current_info.py    # 커서 위치 정보 조회 (KeyIndicatorInfo, PosInfo, CtrlInfo)
├── table_info.py          # 테이블 정보 추출 (TableInfo)
├── styles.yaml            # 스타일 프리셋 정의
├── style_format.py        # 서식 관리 (StyleFormat)
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

## 커서 조작 및 위치 정보 - cursor_wrapper.py ⭐

**커서 관련 작업은 `CursorWrapper` 클래스 사용을 권장합니다.**

```python
from cursor_wrapper import CursorWrapper

cursor = CursorWrapper()  # hwp 자동 연결

# 위치 정보 조회
page = cursor.get_page()           # 페이지 번호
line = cursor.get_line()           # 줄 번호
pos = cursor.get_pos()             # (list_id, para_id, char_pos)
ctrl_id = cursor.get_ctrl_id()     # 컨트롤 ID
cursor.is_in_table()               # 표 안에 있는지 확인

# 커서 이동
cursor.move_line_end()             # 줄 끝으로 (End)
cursor.move_line_begin()           # 줄 시작으로 (Home)
cursor.move_para_end()             # 문단 끝
cursor.move_para_begin()           # 문단 시작
cursor.move_doc_begin()            # 문서 처음
cursor.move_doc_end()              # 문서 끝

# 위치 저장/복원
saved_pos = cursor.save_pos()
cursor.move_doc_end()
cursor.restore_pos(saved_pos)

# 텍스트 선택
cursor.select_line()               # 현재 줄 선택
cursor.select_para()               # 현재 문단 선택
cursor.select_all()                # 전체 선택

# 텍스트 조작
text = cursor.get_selected_text()  # 선택된 텍스트
char = cursor.get_char_at_cursor() # 현재 위치 글자
cursor.insert_text("안녕하세요")
cursor.delete_char(forward=True)   # Delete

# 전체 정보 출력
cursor.print_info()
```

**독립 실행:** `python cursor_wrapper.py`

### 주요 메서드

| 카테고리 | 메서드 | 설명 |
|----------|--------|------|
| **정보 조회** | `get_pos()` | (list_id, para_id, char_pos) |
| | `get_page()` | 페이지 번호 |
| | `get_line()` | 줄 번호 |
| | `get_ctrl_id()` | 컨트롤 ID |
| | `is_in_table()` | 표 안에 있는지 |
| | `is_in_ctrl(ctrl_id)` | 특정 컨트롤 안에 있는지 |
| **이동 (문단)** | `move_para_begin()` | 문단 시작 |
| | `move_para_end()` | 문단 끝 |
| **이동 (줄)** | `move_line_begin()` | 줄 시작 (Home) |
| | `move_line_end()` | 줄 끝 (End) |
| **이동 (문서)** | `move_doc_begin()` | 문서 처음 |
| | `move_doc_end()` | 문서 끝 |
| **이동 (글자)** | `move_left(count)` | 왼쪽으로 n글자 |
| | `move_right(count)` | 오른쪽으로 n글자 |
| | `move_up(count)` | 위로 n줄 |
| | `move_down(count)` | 아래로 n줄 |
| **위치 관리** | `save_pos()` | 현재 위치 저장 |
| | `restore_pos(pos)` | 위치 복원 |
| | `set_pos(list, para, char)` | 특정 위치로 이동 |
| **선택** | `select_all()` | 전체 선택 |
| | `select_line()` | 현재 줄 선택 |
| | `select_para()` | 현재 문단 선택 |
| | `cancel_selection()` | 선택 해제 |
| **텍스트** | `get_selected_text()` | 선택된 텍스트 |
| | `get_char_at_cursor()` | 커서 위치 글자 |
| | `insert_text(text)` | 텍스트 삽입 |
| | `delete_char(forward)` | 글자 삭제 |

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

### 5. table_info.py - 테이블 정보 추출

테이블 크기(행/열 개수) 및 셀 위치 정보 추출

```python
from table_info import TableInfo
from cursor_utils import get_hwp_instance

hwp = get_hwp_instance()
table = TableInfo(hwp, debug=True)

# 테이블 크기 조회
size = table.get_table_size()
print(f"행: {size['rows']}, 열: {size['cols']}")

# 모든 셀 위치 추출
cells = table.get_all_cell_positions()
for cell in cells:
    print(f"셀({cell['row']}, {cell['col']}): {cell['pos']}")

# 특정 셀로 이동
table.move_to_cell(1, 2)  # 2행 3열 (0-based)

# 셀 텍스트 가져오기
text = table.get_cell_text(0, 0)

# 전체 테이블 정보
info = table.get_table_info()  # 크기 + 위치 + 텍스트
```

**독립 실행:** `python table_info.py`

| 메서드 | 설명 |
|--------|------|
| `is_in_table()` | 커서가 테이블 내부에 있는지 확인 |
| `get_table_size()` | 테이블 크기 (행/열 개수) 반환 |
| `get_cell_pos(row, col)` | 특정 셀의 위치 (list_id, para_id, char_pos) |
| `get_all_cell_positions()` | 모든 셀 위치 정보 리스트 |
| `move_to_cell(row, col)` | 특정 셀로 커서 이동 |
| `select_cell(row, col)` | 특정 셀 선택 (전체 블록 지정) |
| `get_cell_text(row, col)` | 특정 셀의 텍스트 내용 |
| `get_table_info()` | 전체 테이블 정보 (크기+위치+텍스트) |
| `is_merged_cell(row, col)` | 셀 병합 여부 확인 |

**MovePos 셀 이동 상수:**
- `MOVE_LEFT_OF_CELL (100)` - 왼쪽 셀로 이동
- `MOVE_RIGHT_OF_CELL (101)` - 오른쪽 셀로 이동
- `MOVE_UP_OF_CELL (102)` - 위쪽 셀로 이동
- `MOVE_DOWN_OF_CELL (103)` - 아래쪽 셀로 이동
- `MOVE_START_OF_CELL (104)` - 행의 시작 셀
- `MOVE_END_OF_CELL (105)` - 행의 끝 셀
- `MOVE_TOP_OF_CELL (106)` - 열의 시작 셀
- `MOVE_BOTTOM_OF_CELL (107)` - 열의 끝 셀

**주의사항:**
- 행/열 번호는 모두 0-based (0부터 시작)
- 테이블 외부에서 호출 시 None 반환
- 모든 메서드는 커서 위치를 원래대로 복원

#### 셀 병합 감지 (is_merged_cell)

HWP 테이블에서 병합된 셀을 감지하는 방법 (상세 내용: [table_info.md](table_info.md))

```python
# 셀 병합 확인
result = table.is_merged_cell(5, 0)

if result['is_merged']:
    print(f"병합 유형: {result['merge_type']}")  # 'vertical', 'horizontal', 'both'
    print(f"시작 셀: {result['master_cell']}")   # (row, col)
```

**반환값:**
```python
{
    'is_merged': bool,           # 병합 여부
    'merge_type': str | None,    # 'vertical', 'horizontal', 'both', None
    'master_cell': tuple | None  # 병합 시작 셀 (row, col)
}
```

**병합 감지 원리:**

HWP의 각 셀은 고유한 `sublist` (list_id) 값을 가지며, 병합된 셀은 특정 패턴을 보입니다.

1. **세로 병합 감지**: 오른쪽 열의 sublist 비교
   ```
   ┌──────┬────┐
   │      │ B1 │  ← (4,0) 병합 시작, (4,1) sublist=15
   │  A   ├────┤
   │      │ B2 │  ← (5,0) 접근불가,  (5,1) sublist=14
   └──────┴────┘

   검증: (5,1).sublist == (4,1).sublist - 1
   → 14 == 15 - 1 ✓ → (4,0)과 (5,0) 세로 병합
   ```

2. **가로 병합 감지**: 아래쪽 행의 sublist 비교
   ```
   ┌────┬────┬────┐
   │ A1 │  B      │  ← (2,0), (2,1) 병합, (2,2) 접근불가
   ├────┼────┼────┤
   │ A2 │ C1 │ C2 │  ← (3,1) sublist=20, (3,2) sublist=21
   └────┴────┴────┘

   검증: (3,2).sublist == (3,1).sublist + 1
   → 21 == 20 + 1 ✓ → (2,1)과 (2,2) 가로 병합
   ```

**패턴 요약:**
- **세로 병합**: 오른쪽 열 sublist가 연속 감소 (`-1`)
- **가로 병합**: 아래쪽 행 sublist가 연속 증가 (`+1`)

## 저수준 커서 API (참고용)

대부분의 경우 `CursorWrapper`를 사용하면 됩니다. 아래는 내부 동작을 이해하기 위한 참고 정보입니다.

### GetPos() / SetPos()

```python
# 내부 위치 정보 (list_id, para_id, char_pos)
pos = hwp.GetPos()
list_id = pos[0]   # 0=본문, 10+=테이블
para_id = pos[1]   # 문단 ID
char_pos = pos[2]  # 문자 위치 (한글=2, 영문=1씩 증가)

# 위치 이동
hwp.SetPos(list_id, para_id, char_pos)
```

### KeyIndicator()

```python
# 상태바 정보 (공식 문서와 실제 반환값이 다름!)
key = hwp.KeyIndicator()
page = key[3]      # 페이지 번호 (문서에는 '단 번호'라고 잘못 기재)
line = key[4]      # 줄 번호
insert_mode = key[6]  # 0=삽입, 1=수정
```

**주의:** KeyIndicator는 API 버그가 많습니다. 대신 `CursorWrapper`나 `get_current_info.py`를 사용하세요.

### 상세 정보 조회

```python
from get_current_info import CurrentInfo

info = CurrentInfo()
info.print_all()  # 모든 위치 정보 출력
```

## 주의사항

### 커서 위치 복원
커서를 이동한 후에는 반드시 원래 위치로 복원하세요.

```python
cursor = CursorWrapper()

# 위치 저장
saved = cursor.save_pos()

# 작업 수행
cursor.move_doc_end()
# ... 작업 ...

# 위치 복원
cursor.restore_pos(saved)
```

### 이벤트 기반 vs 폴링
- `DocumentChange` 이벤트: API 커서 이동도 발생 → 무한 루프 위험
- **폴링 방식 권장**: 0.1초 간격으로 위치 비교

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
