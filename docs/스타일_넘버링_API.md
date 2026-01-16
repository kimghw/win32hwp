# HWP 개요 번호/문단 번호/글머리표 API 정리

한글(HWP) 오토메이션 API를 사용한 개요 번호, 문단 번호, 글머리표 관련 기능을 정리한 문서입니다.

## 목차

1. [개요](#개요)
2. [주요 액션](#주요-액션)
3. [파라미터셋](#파라미터셋)
4. [ParagraphShape 속성](#paragraphshape-속성)
5. [GetHeadingString 메서드](#getheadingstring-메서드)
6. [마크다운 헤딩 변환](#마크다운-헤딩-변환)
7. [코드 예제](#코드-예제)
8. [제한사항 및 주의사항](#제한사항-및-주의사항)

---

## 개요

HWP에서 번호를 매기는 방법은 크게 세 가지가 있습니다:

| 종류 | 설명 | 특징 |
|------|------|------|
| **개요 번호** | 문서 전체에 일관성 있게 번호를 매기고 문서 구조를 표현 | 구역(SecDef) 단위로 관리, 1~7수준 |
| **문단 번호** | 원하는 위치에서 원하는 모양으로 번호 매기기 | 문단(ParaShape) 단위로 관리, 1~7수준 |
| **글머리표(불릿)** | 기호를 사용한 목록 | 문단(ParaShape) 단위로 관리 |

---

## 주요 액션

### 1. PutOutlineNumber (개요 번호 달기)

현재 문단에 개요 번호를 적용합니다. 기본적으로 1수준 개요 번호가 적용됩니다.

```python
# 개요 번호 달기 (1수준으로 시작)
hwp.HAction.Run("PutOutlineNumber")
```

| 항목 | 내용 |
|------|------|
| Action ID | `PutOutlineNumber` |
| ParameterSet | `ParaShape*` |
| 설명 | 개요번호 달기 |

### 2. ParaNumberBulletLevelDown/Up (수준 변경)

개요/문단 번호의 수준을 변경합니다.

```python
# 한 수준 아래로 (예: 1수준 → 2수준)
hwp.HAction.Run("ParaNumberBulletLevelDown")

# 한 수준 위로 (예: 2수준 → 1수준)
hwp.HAction.Run("ParaNumberBulletLevelUp")
```

| Action ID | 설명 |
|-----------|------|
| `ParaNumberBullet` | 문단번호/글머리표 한 수준 위로 |
| `ParaNumberBulletLevelDown` | 문단번호/글머리표 한 수준 아래로 |
| `ParaNumberBulletLevelUp` | 문단번호/글머리표 한 수준 위로 |

### 3. OutlineNumber (개요 번호 양식 설정)

개요 번호의 양식(형식)을 설정합니다. `SecDef` 파라미터셋을 사용합니다.

```python
# 개요 번호 양식 대화상자 열기
hwp.HAction.Run("OutlineNumber")
```

| 항목 | 내용 |
|------|------|
| Action ID | `OutlineNumber` |
| ParameterSet | `SecDef` |
| 설명 | 개요번호 양식 설정 |

---

## 파라미터셋

### SecDef (구역 속성)

개요 번호 양식은 구역(Section) 단위로 관리됩니다.

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| OutlineShape | PIT_SET | NumberingShape | 개요 번호 형태 |
| TextDirection | PIT_UI1 | | 글자 방향 |
| PageDef | PIT_SET | PageDef | 용지 설정 정보 |
| TabStop | PIT_I4 | | 기본 탭 간격 |

### NumberingShape (번호 모양)

번호의 형태와 형식을 정의합니다.

| Item ID | Type | Description |
|---------|------|-------------|
| NumberFormat | PIT_UI1 | 번호 모양 (아래 표 참고) |
| UserChar | PIT_UI2 | 사용자 기호 (WCHAR) |
| PrefixChar | PIT_UI2 | 앞 장식 문자 |
| SuffixChar | PIT_UI2 | 뒤 장식 문자 |

**NumberFormat 값:**

| 값 | 형식 | 예시 |
|----|------|------|
| 0 | 아라비아 숫자 | 1, 2, 3 |
| 1 | 원문자 | (1), (2), (3) |
| 2 | 로마 숫자 대문자 | I, II, III |
| 3 | 로마 숫자 소문자 | i, ii, iii |
| 4 | 영문 대문자 | A, B, C |
| 8 | 가나다 | 가, 나, 다 |
| 13 | 한자 숫자 | 一, 二, 三 |
| 15 | 천간 (한글) | 갑, 을, 병 |
| 16 | 천간 (한자) | 甲, 乙, 丙 |

### BulletShape (글머리표 모양)

| Item ID | Type | Description |
|---------|------|-------------|
| HasCharShape | PIT_UI1 | 자체 글자 모양 사용 여부 |
| CharShape | PIT_SET | 글자 모양 (HasCharShape=1일 때) |
| BulletChar | PIT_UI2 | 불릿 문자 코드 |
| WidthAdjust | PIT_I | 번호 너비 보정 값 (HWPUNIT) |
| TextOffset | PIT_I | 본문과의 거리 |
| Alignment | PIT_UI1 | 번호 정렬 (0=왼쪽, 1=가운데, 2=오른쪽) |
| AutoIndent | PIT_UI1 | 자동 들여쓰기 여부 |

---

## ParagraphShape 속성

### HeadingType (문단 머리 모양)

문단에 적용된 번호/글머리표 종류를 나타냅니다.

| 값 | 의미 |
|----|------|
| 0 | 없음 |
| 1 | 개요 (Outline) |
| 2 | 번호 (Numbering) |
| 3 | 불릿 (Bullet/글머리표) |

```python
# HeadingType 조회
hwp.HAction.GetDefault("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)
heading_type = hwp.HParameterSet.HParaShape.HeadingType
print(f"HeadingType: {heading_type}")  # 0=없음, 1=개요, 2=번호, 3=불릿
```

### Level (단계)

개요/번호의 수준을 나타냅니다. (0~6, 즉 1~7수준)

```python
# Level 조회
hwp.HAction.GetDefault("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)
level = hwp.HParameterSet.HParaShape.Level
print(f"Level: {level}")  # 0=1수준, 1=2수준, ..., 6=7수준
```

### 제한사항

> **중요:** HWP API에서 `HParaShape.HeadingType`과 `HParaShape.Level`을 직접 설정해도 개요 번호가 올바르게 적용되지 않을 수 있습니다.

```python
# 이 방법은 권장되지 않음 (동작 불안정)
hwp.HAction.GetDefault("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)
hwp.HParameterSet.HParaShape.HeadingType = 1  # 개요
hwp.HParameterSet.HParaShape.Level = 2        # 3수준
hwp.HAction.Execute("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)
```

**권장 방법:** `PutOutlineNumber` + `ParaNumberBulletLevelDown` 조합 사용

```python
# 권장: PutOutlineNumber로 개요 달고, LevelDown으로 수준 조정
hwp.HAction.Run("PutOutlineNumber")  # 1수준 개요 적용
hwp.HAction.Run("ParaNumberBulletLevelDown")  # 2수준으로
hwp.HAction.Run("ParaNumberBulletLevelDown")  # 3수준으로
```

---

## GetHeadingString 메서드

현재 커서 위치의 글머리표/문단번호/개요번호 문자열을 반환합니다.

```python
# 개요 번호 문자열 조회
heading_str = hwp.GetHeadingString()
print(heading_str)  # 예: "1.", "1.1", "가.", "" (없으면 빈 문자열)
```

| 항목 | 내용 |
|------|------|
| 메서드 | `GetHeadingString()` |
| 반환값 | 문자열 (글머리표/문단번호/개요번호) |
| 비고 | 번호가 없으면 빈 문자열("") 반환 |

**활용 예시:**

```python
def get_outline_level(hwp):
    """현재 문단의 개요 수준 조회"""
    heading_str = hwp.GetHeadingString()
    return {
        'heading_string': heading_str,
        'has_outline': bool(heading_str)
    }

# 사용
info = get_outline_level(hwp)
if info['has_outline']:
    print(f"개요 번호: {info['heading_string']}")
else:
    print("개요 없음")
```

---

## 마크다운 헤딩 변환

마크다운 형식의 헤딩(`#`, `##`, `###` 등)을 HWP 개요 수준으로 변환하는 방법입니다.

### 매핑 규칙

| 마크다운 | HWP 개요 수준 |
|----------|---------------|
| `#` | 1수준 |
| `##` | 2수준 |
| `###` | 3수준 |
| `####` | 4수준 |
| `#####` | 5수준 |
| `######` | 6수준 |
| `#######` | 7수준 |

### 헤딩 파싱 함수

```python
@staticmethod
def parse_heading_level(text):
    """
    마크다운 헤딩 파싱

    Args:
        text: 텍스트 (예: "## 소제목")

    Returns:
        tuple: (level, clean_text)
    """
    if not text or not text.startswith('#'):
        return 0, text

    level = 0
    for char in text:
        if char == '#':
            level += 1
        else:
            break

    level = min(level, 7)  # 최대 7수준
    clean_text = text[level:].lstrip()

    return level, clean_text
```

### 개요 수준 적용 함수

```python
def set_outline_level(hwp, level, debug=False):
    """
    개요 수준 설정 (PutOutlineNumber + LevelDown 방식)

    Args:
        hwp: 한글 인스턴스
        level: 개요 수준 (1~7)
        debug: True면 결과 출력

    Returns:
        bool: 성공 여부
    """
    if not 1 <= level <= 7:
        raise ValueError(f"개요 수준은 1~7 사이여야 합니다: {level}")

    # 1. 개요번호 달기 (1수준으로 시작)
    result = hwp.HAction.Run("PutOutlineNumber")

    # 2. 필요한 만큼 수준 내리기
    for _ in range(level - 1):
        hwp.HAction.Run("ParaNumberBulletLevelDown")

    if debug:
        print(f"  set_outline_level({level}) Run: {result}")
    return result
```

### 문서 전체 헤딩 스캔 및 적용

```python
def scan_and_apply_headings(hwp, remove_marker=True, debug=False):
    """
    문서에서 마크다운 헤딩을 스캔하고 개요 수준 적용

    Args:
        hwp: 한글 인스턴스
        remove_marker: True면 # 마커 제거
        debug: 디버그 출력
    """
    headings = []
    processed_paras = set()

    # 문서 처음으로 이동
    hwp.HAction.Run("MoveDocBegin")

    while True:
        hwp.HAction.Run("MoveParaBegin")
        pos = hwp.GetPos()
        para_id = pos[1]

        if para_id in processed_paras:
            break
        processed_paras.add(para_id)

        # 현재 문단 텍스트 가져오기
        hwp.HAction.Run("MoveSelParaEnd")
        text = hwp.GetTextFile("TEXT", "saveblock")
        hwp.HAction.Run("Cancel")

        if text:
            text = text.strip()
            level, clean_text = parse_heading_level(text)

            if level > 0:
                headings.append({
                    'list_id': pos[0],
                    'para_id': para_id,
                    'level': level,
                    'text': text,
                    'clean_text': clean_text
                })

        # 다음 문단으로 이동
        hwp.HAction.Run("MoveParaEnd")
        hwp.HAction.Run("MoveRight")

        new_pos = hwp.GetPos()
        if new_pos[1] == para_id:
            break

    # 역순으로 처리 (para_id 유지)
    for h in reversed(headings):
        hwp.SetPos(h['list_id'], h['para_id'], 0)

        # 텍스트 교체
        hwp.HAction.Run("MoveSelParaEnd")
        new_text = h['clean_text'] if remove_marker else h['text']

        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = new_text
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

        # 개요 수준 적용
        hwp.HAction.Run("MoveParaBegin")
        hwp.HAction.Run("MoveSelParaEnd")
        set_outline_level(hwp, h['level'], debug=debug)
        hwp.HAction.Run("Cancel")

        if debug:
            print(f"  Level {h['level']}: {h['clean_text'][:20]}")

    return headings
```

---

## 코드 예제

### 예제 1: 기본 개요 번호 적용

```python
from cursor_utils import get_hwp_instance

hwp = get_hwp_instance()

# 새 문서
hwp.HAction.Run("FileNew")

# 텍스트 입력
def insert_text(text):
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HAction.Run("BreakPara")

# 1수준 개요
insert_text("서론")
hwp.HAction.Run("MoveParaBegin")
hwp.HAction.Run("PutOutlineNumber")
hwp.HAction.Run("MoveParaEnd")
hwp.HAction.Run("MoveRight")

# 2수준 개요
insert_text("배경")
hwp.HAction.Run("MoveParaBegin")
hwp.HAction.Run("PutOutlineNumber")
hwp.HAction.Run("ParaNumberBulletLevelDown")
hwp.HAction.Run("MoveParaEnd")
hwp.HAction.Run("MoveRight")

# 3수준 개요
insert_text("세부 내용")
hwp.HAction.Run("MoveParaBegin")
hwp.HAction.Run("PutOutlineNumber")
hwp.HAction.Run("ParaNumberBulletLevelDown")
hwp.HAction.Run("ParaNumberBulletLevelDown")
```

### 예제 2: StyleNumb 클래스 사용

```python
from style_numb import StyleNumb
from cursor_utils import get_hwp_instance

hwp = get_hwp_instance()
numb = StyleNumb(hwp)

# 새 문서 열기
numb.새문서()

# 마크다운 텍스트 입력
numb.텍스트입력('''
# 1장 서론
## 1.1 배경
### 1.1.1 세부 내용
# 2장 본론
## 2.1 방법
### 2.1.1 실험 설계
### 2.1.2 데이터 수집
## 2.2 결과
# 3장 결론
''')

# 개요 수준 정의 (마크다운 → HWP 개요)
result = numb.개요수준정의(remove_marker=True, debug=True)
print(f"처리된 헤딩: {result['processed']}개")
```

### 예제 3: 개요 해제

```python
def remove_outline(hwp):
    """개요 해제"""
    hwp.HAction.GetDefault("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)
    hwp.HParameterSet.HParaShape.HeadingType = 0  # 없음
    hwp.HAction.Execute("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)

# 사용
hwp.HAction.Run("MoveParaBegin")
hwp.HAction.Run("MoveSelParaEnd")
remove_outline(hwp)
hwp.HAction.Run("Cancel")
```

### 예제 4: 개요 번호 검증

```python
from style_numb import StyleNumb
from cursor_utils import get_hwp_instance

hwp = get_hwp_instance()
numb = StyleNumb(hwp)

# 문서의 각 문단 순회하며 개요 번호 확인
hwp.HAction.Run("MoveDocBegin")
processed = set()

while True:
    pos = hwp.GetPos()
    para_id = pos[1]

    if para_id in processed:
        break
    processed.add(para_id)

    # 개요 번호 조회
    outline = numb.get_outline_level()
    if outline['has_outline']:
        hwp.HAction.Run("MoveSelParaEnd")
        text = hwp.GetTextFile("TEXT", "saveblock")
        hwp.HAction.Run("Cancel")
        print(f"  [{para_id}] '{outline['heading_string']}' {text[:30]}")

    hwp.HAction.Run("MoveParaEnd")
    hwp.HAction.Run("MoveRight")

    if hwp.GetPos()[1] == para_id:
        break
```

---

## 제한사항 및 주의사항

### 1. 개요 번호는 구역 단위

- 개요 번호 모양은 구역(Section) 단위로 관리됩니다.
- 같은 구역 내에서는 동일한 개요 번호 양식을 사용해야 합니다.
- 다른 양식을 원하면 새 구역을 만들어야 합니다.

### 2. HeadingType/Level 직접 설정의 한계

- `HParaShape.HeadingType`과 `Level`을 직접 설정하는 방법은 불안정합니다.
- `PutOutlineNumber` + `ParaNumberBulletLevelDown` 조합을 권장합니다.

### 3. 개요 수준 범위

- 개요 수준은 1~7까지 지원됩니다.
- API에서 Level 값은 0~6으로 표현됩니다 (0=1수준, 6=7수준).

### 4. 문단 위치 저장 및 복원

- 문단을 순회하면서 처리할 때 `para_id`가 변경될 수 있습니다.
- 역순 처리(reversed)하거나, 위치를 저장/복원하는 것이 안전합니다.

### 5. 텍스트 선택 해제

- 블록 선택 후 작업이 끝나면 반드시 `hwp.HAction.Run("Cancel")`로 해제하세요.

---

## 참고 자료

### 로컬 문서

- `win32/ActionTable_2504_part*.md` - 액션 테이블
- `win32/ParameterSetTable_2504_part*.md` - 파라미터셋 정의
- `win32/HwpAutomation_2504_part*.md` - API 메서드/속성

### 관련 코드

- `style_numb.py` - StyleNumb 클래스 구현
- `styles.yaml` - 헤딩 레벨 매핑 설정 (heading_levels)

### 웹 리소스

- [한컴디벨로퍼 공식 사이트](https://developer.hancom.com/hwpautomation)
- [한컴디벨로퍼 포럼](https://forum.developer.hancom.com/c/hwp-automation/52)
- [pyhwpx Cookbook (WikiDocs)](https://wikidocs.net/book/8956)
- [한글 도움말 - 개요 번호](https://help.hancom.com/hoffice/multi/ko_kr/hwp/format/outline/outline_numbering(paragraph_number).htm)
