# HWP API 텍스트 삽입 및 서식 적용 가이드

## 개요

이 문서는 HWP API를 사용하여 텍스트를 삽입하고 글자/문단 서식을 적용하는 방법을 정리합니다.

---

## 1. InsertText - 텍스트 삽입

### 기본 패턴

```python
# 텍스트 삽입
hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
hwp.HParameterSet.HInsertText.Text = "삽입할 텍스트"
hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
```

### ParameterSet: InsertText

| Item ID | Type | Description |
|---------|------|-------------|
| Text | PIT_BSTR | 삽입할 텍스트 |

### 사용 예시

```python
def insert_text(hwp, text):
    """텍스트를 현재 캐럿 위치에 삽입"""
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
```

---

## 2. CharShape 액션 - 글자 서식 적용

### 기본 패턴

```python
# 1. 블록 선택 (필수!)
hwp.HAction.Run("SelectAll")  # 또는 다른 선택 방법

# 2. CharShape 설정
hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)

# 3. 원하는 속성 설정
hwp.HParameterSet.HCharShape.TextColor = 0x0000FF  # 빨간색 (BGR 순서)
hwp.HParameterSet.HCharShape.Bold = 1
hwp.HParameterSet.HCharShape.Height = hwp.PointToHwpUnit(12)  # 12pt

# 4. 적용
hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)

# 5. 선택 해제
hwp.HAction.Run("Cancel")
```

### ParameterSet: CharShape (글자 모양)

#### 글꼴 관련

| Item ID | Type | Description |
|---------|------|-------------|
| FaceNameHangul | PIT_BSTR | 글꼴 이름 (한글) |
| FaceNameLatin | PIT_BSTR | 글꼴 이름 (영문) |
| FaceNameHanja | PIT_BSTR | 글꼴 이름 (한자) |
| FaceNameJapanese | PIT_BSTR | 글꼴 이름 (일본어) |
| FaceNameOther | PIT_BSTR | 글꼴 이름 (기타) |
| FaceNameSymbol | PIT_BSTR | 글꼴 이름 (심벌) |
| FaceNameUser | PIT_BSTR | 글꼴 이름 (사용자) |
| FontTypeHangul | PIT_UI1 | 폰트 종류 (한글): 0=don't care, 1=TTF, 2=HFT |
| FontTypeLatin | PIT_UI1 | 폰트 종류 (영문): 0=don't care, 1=TTF, 2=HFT |
| Height | PIT_I4 | 글자 크기 (HWPUNIT) |

#### 크기/비율 관련

| Item ID | Type | Description |
|---------|------|-------------|
| SizeHangul | PIT_UI1 | 각 언어별 크기 비율 (한글) 10% - 250% |
| SizeLatin | PIT_UI1 | 각 언어별 크기 비율 (영문) 10% - 250% |
| RatioHangul | PIT_UI1 | 각 언어별 장평 비율 (한글) 50% - 200% |
| RatioLatin | PIT_UI1 | 각 언어별 장평 비율 (영문) 50% - 200% |
| SpacingHangul | PIT_I1 | 각 언어별 자간 (한글) -50% - 50% |
| SpacingLatin | PIT_I1 | 각 언어별 자간 (영문) -50% - 50% |
| OffsetHangul | PIT_I1 | 각 언어별 오프셋 (한글) -100% - 100% |
| OffsetLatin | PIT_I1 | 각 언어별 오프셋 (영문) -100% - 100% |

#### 스타일 관련

| Item ID | Type | Description |
|---------|------|-------------|
| Bold | PIT_UI1 | Bold: 0=off, 1=on |
| Italic | PIT_UI1 | Italic: 0=off, 1=on |
| SmallCaps | PIT_UI1 | Small Caps: 0=off, 1=on |
| Emboss | PIT_UI1 | Emboss: 0=off, 1=on |
| Engrave | PIT_UI1 | Engrave: 0=off, 1=on |
| SuperScript | PIT_UI1 | 위 첨자: 0=off, 1=on |
| SubScript | PIT_UI1 | 아래 첨자: 0=off, 1=on |

#### 밑줄/외곽선/그림자

| Item ID | Type | Description |
|---------|------|-------------|
| UnderlineType | PIT_UI1 | 밑줄 종류: 0=none, 1=bottom, 2=center, 3=top |
| UnderlineShape | PIT_UI1 | 밑줄 모양 (선 종류) |
| UnderlineColor | PIT_UI4 | 밑줄 색 (COLORREF: 0x00BBGGRR) |
| OutlineType | PIT_UI1 | 외곽선 종류: 0=none, 1=solid, 2=dot, 3=thick, 4=dash, 5=dashdot, 6=dashdotdot |
| ShadowType | PIT_UI1 | 그림자 종류: 0=none, 1=drop, 2=continuous |
| ShadowOffsetX | PIT_I1 | 그림자 간격 (X 방향) -100% - 100% |
| ShadowOffsetY | PIT_I1 | 그림자 간격 (Y 방향) -100% - 100% |
| ShadowColor | PIT_UI4 | 그림자 색 (COLORREF) |

#### 색상 관련

| Item ID | Type | Description |
|---------|------|-------------|
| TextColor | PIT_UI4 | 글자색 (COLORREF: 0x00BBGGRR) |
| ShadeColor | PIT_UI4 | 음영색 (COLORREF) |

#### 취소선/강조점

| Item ID | Type | Description |
|---------|------|-------------|
| StrikeOutType | PIT_UI1 | 취소선 종류: 0=none, 1=red single, 2=red double, 3=text single, 4=text double |
| StrikeOutShape | PIT_UI1 | 취소선 모양 |
| StrikeOutColor | PIT_UI4 | 취소선색 (COLORREF) |
| DiacSymMark | PIT_UI1 | 강조점 종류: 0=none, 1=검정 동그라미, 2=속 빈 동그라미 |

#### 기타

| Item ID | Type | Description |
|---------|------|-------------|
| UseFontSpace | PIT_UI1 | 글꼴에 어울리는 빈칸: 0=off, 1=on |
| UseKerning | PIT_UI1 | 커닝: 0=off, 1=on |

### 색상 변환 (RGB to COLORREF)

HWP에서 색상은 **BGR 순서**로 저장됩니다.

```python
def rgb_to_colorref(r, g, b):
    """RGB를 COLORREF(BGR)로 변환"""
    return b << 16 | g << 8 | r

# 사용 예시
red = rgb_to_colorref(255, 0, 0)      # 0x0000FF
blue = rgb_to_colorref(0, 0, 255)      # 0xFF0000
green = rgb_to_colorref(0, 255, 0)     # 0x00FF00
```

또는 hwp.RGBColor() 메서드 사용:

```python
red_color = hwp.RGBColor(255, 0, 0)
```

### 글자 크기 변환

```python
# pt를 HWPUNIT으로 변환
# 1 pt = 100 HWPUNIT
hwpunit = hwp.PointToHwpUnit(12)  # 12pt -> 1200

# 직접 계산
hwpunit = pt_size * 100
```

---

## 3. ParaShape 액션 - 문단 서식 적용

### 기본 패턴

```python
# 1. 블록 선택 (필수!)
hwp.HAction.Run("SelectAll")

# 2. ParaShape 설정
hwp.HAction.GetDefault("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)

# 3. 원하는 속성 설정
hwp.HParameterSet.HParaShape.AlignType = 3  # 가운데 정렬
hwp.HParameterSet.HParaShape.LineSpacing = 160  # 160%

# 4. 적용
hwp.HAction.Execute("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)

# 5. 선택 해제
hwp.HAction.Run("Cancel")
```

### ParameterSet: ParaShape (문단 모양)

#### 여백/들여쓰기

| Item ID | Type | Description |
|---------|------|-------------|
| LeftMargin | PIT_I4 | 왼쪽 여백 (HWPUNIT) |
| RightMargin | PIT_I4 | 오른쪽 여백 (HWPUNIT) |
| Indentation | PIT_I4 | 들여쓰기/내어 쓰기 (HWPUNIT) |

#### 문단 간격

| Item ID | Type | Description |
|---------|------|-------------|
| PrevSpacing | PIT_I4 | 문단 간격 위 (HWPUNIT) |
| NextSpacing | PIT_I4 | 문단 간격 아래 (HWPUNIT) |

#### 줄 간격

| Item ID | Type | Description |
|---------|------|-------------|
| LineSpacingType | PIT_UI1 | 줄 간격 종류: 0=글자에 따라, 1=고정 값, 2=여백만 지정 |
| LineSpacing | PIT_I4 | 줄 간격 값 (종류에 따라 % 또는 HWPUNIT) |

#### 정렬

| Item ID | Type | Description |
|---------|------|-------------|
| AlignType | PIT_UI1 | 정렬 방식: 0=양쪽, 1=왼쪽, 2=오른쪽, 3=가운데, 4=배분, 5=나눔 |
| TextAlignment | PIT_UI1 | 세로 정렬: 0=글꼴기준, 1=위, 2=가운데, 3=아래 |

#### 줄 나눔

| Item ID | Type | Description |
|---------|------|-------------|
| BreakLatinWord | PIT_UI1 | 줄 나눔 단위 (라틴): 0=단어, 1=하이픈, 2=글자 |
| BreakNonLatinWord | PIT_UI1 | 단위 (비 라틴): TRUE=글자, FALSE=어절 |
| SnapToGrid | PIT_UI1 | 편집 용지의 줄 격자 사용 (on/off) |
| Condense | PIT_UI1 | 공백 최소값 (0 - 75%) |

#### 문단 보호

| Item ID | Type | Description |
|---------|------|-------------|
| WidowOrphan | PIT_UI1 | 외톨이줄 보호 (on/off) |
| KeepWithNext | PIT_UI1 | 다음 문단과 함께 (on/off) |
| KeepLinesTogether | PIT_UI1 | 문단 보호 (on/off) |
| PagebreakBefore | PIT_UI1 | 문단 앞에서 항상 쪽 나눔 (on/off) |
| FontLineHeight | PIT_UI1 | 글꼴에 어울리는 줄 높이 (on/off) |

#### 문단 머리

| Item ID | Type | Description |
|---------|------|-------------|
| HeadingType | PIT_UI1 | 문단 머리 모양: 0=없음, 1=개요, 2=번호, 3=불릿 |
| Level | PIT_UI1 | 단계 (0 - 6) |

#### 테두리/배경

| Item ID | Type | Description |
|---------|------|-------------|
| BorderConnect | PIT_UI1 | 문단 테두리/배경 - 테두리 연결 (on/off) |
| BorderText | PIT_UI1 | 문단 테두리/배경 - 여백 무시 (0=단, 1=텍스트) |
| BorderOffsetLeft | PIT_I | 문단 테두리/배경 - 왼쪽 간격 (HWPUNIT) |
| BorderOffsetRight | PIT_I | 문단 테두리/배경 - 오른쪽 간격 (HWPUNIT) |
| BorderOffsetTop | PIT_I | 문단 테두리/배경 - 위 간격 (HWPUNIT) |
| BorderOffsetBottom | PIT_I | 문단 테두리/배경 - 아래 간격 (HWPUNIT) |
| BorderFill | PIT_SET | 테두리/배경 |

---

## 4. 셀에 텍스트 삽입 후 서식 적용

### 방법 1: 셀로 이동 후 삽입

```python
def insert_text_in_cell(hwp, row, col, text, apply_format=None):
    """특정 셀에 텍스트 삽입 및 서식 적용"""

    # 1. 셀로 이동 (표 내부에 캐럿이 있어야 함)
    for _ in range(row):
        hwp.HAction.Run("TableDownCell")
    for _ in range(col):
        hwp.HAction.Run("TableRightCell")

    # 2. 셀 내용 선택 후 삭제 (기존 내용 제거)
    hwp.HAction.Run("SelectAll")
    hwp.HAction.Run("Delete")

    # 3. 텍스트 삽입
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

    # 4. 서식 적용 (선택 후 적용)
    if apply_format:
        hwp.HAction.Run("MoveSelListBegin")  # 셀 시작부터 선택
        apply_format(hwp)
        hwp.HAction.Run("Cancel")
```

### 방법 2: 필드 사용

```python
def insert_text_by_field(hwp, field_name, text):
    """필드명으로 셀에 텍스트 삽입"""
    hwp.PutFieldText(field_name, text)
```

### 방법 3: 셀 직접 탐색

```python
def move_to_cell(hwp, target_row, target_col):
    """표 내 특정 셀로 이동"""
    # 현재 표의 첫 번째 셀로 이동
    hwp.MovePos(104)  # moveStartOfCell - 행 시작
    hwp.MovePos(106)  # moveTopOfCell - 열 시작

    # 목표 셀로 이동
    for _ in range(target_row):
        hwp.MovePos(103)  # moveDownOfCell
    for _ in range(target_col):
        hwp.MovePos(101)  # moveRightOfCell
```

### 셀 이동 관련 MovePos ID

| ID | 값 | 설명 |
|----|---|------|
| moveLeftOfCell | 100 | 현재 캐럿이 위치한 셀의 왼쪽 |
| moveRightOfCell | 101 | 현재 캐럿이 위치한 셀의 오른쪽 |
| moveUpOfCell | 102 | 현재 캐럿이 위치한 셀의 위쪽 |
| moveDownOfCell | 103 | 현재 캐럿이 위치한 셀의 아래쪽 |
| moveStartOfCell | 104 | 현재 캐럿이 위치한 셀에서 행(row)의 시작 |
| moveEndOfCell | 105 | 현재 캐럿이 위치한 셀에서 행(row)의 끝 |
| moveTopOfCell | 106 | 현재 캐럿이 위치한 셀에서 열(column)의 시작 |
| moveBottomOfCell | 107 | 현재 캐럿이 위치한 셀에서 열(column)의 끝 |

---

## 5. 블록 선택 후 서식 적용 패턴

### 필수 원칙

**서식을 적용하려면 반드시 먼저 블록을 선택해야 합니다!**

```python
# 잘못된 방법 (서식이 적용되지 않음)
hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
hwp.HParameterSet.HCharShape.Bold = 1
hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)

# 올바른 방법
hwp.HAction.Run("SelectAll")  # 선택 먼저!
hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
hwp.HParameterSet.HCharShape.Bold = 1
hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
hwp.HAction.Run("Cancel")  # 선택 해제
```

### 블록 선택 액션 목록

#### 전체/기본 선택

| Action ID | Description |
|-----------|-------------|
| SelectAll | 모두 선택 |
| Select | 선택 (F3 Key를 누른 효과) |
| SelectColumn | 칸 블록 선택 (F4 Key를 누른 효과) |
| Cancel | 선택 해제 (ESC) |

#### 이동하며 선택 (MoveSel 계열)

| Action ID | Description |
|-----------|-------------|
| MoveSelDocBegin | 셀렉션: 문서 처음 |
| MoveSelDocEnd | 셀렉션: 문서 끝 |
| MoveSelDown | 셀렉션: 아래로 이동 |
| MoveSelUp | 셀렉션: 위로 이동 |
| MoveSelLeft | 셀렉션: 왼쪽으로 이동 |
| MoveSelRight | 셀렉션: 오른쪽으로 이동 |
| MoveSelLineBegin | 셀렉션: 줄 처음 |
| MoveSelLineEnd | 셀렉션: 줄 끝 |
| MoveSelLineDown | 셀렉션: 한줄 아래 |
| MoveSelLineUp | 셀렉션: 한줄 위 |
| MoveSelListBegin | 셀렉션: 리스트 처음 |
| MoveSelListEnd | 셀렉션: 리스트 끝 |
| MoveSelNextChar | 셀렉션: 다음 글자 |
| MoveSelPrevChar | 셀렉션: 이전 글자 |
| MoveSelNextWord | 셀렉션: 다음 단어 |
| MoveSelPrevWord | 셀렉션: 이전 단어 |
| MoveSelParaBegin | 셀렉션: 문단 처음 |
| MoveSelParaEnd | 셀렉션: 문단 끝 |
| MoveSelNextParaBegin | 셀렉션: 다음 문단 처음 |
| MoveSelPrevParaBegin | 셀렉션: 이전 문단 시작 |
| MoveSelPrevParaEnd | 셀렉션: 이전 문단 끝 |
| MoveSelWordBegin | 셀렉션: 단어 처음 |
| MoveSelWordEnd | 셀렉션: 단어 끝 |
| MoveSelPageDown | 셀렉션: 페이지다운 |
| MoveSelPageUp | 셀렉션: 페이지 업 |
| MoveSelTopLevelBegin | 셀렉션: 처음 |
| MoveSelTopLevelEnd | 셀렉션: 끝 |

### SelectText 메서드 사용

```python
def select_text_range(hwp, start_para, start_pos, end_para, end_pos):
    """특정 범위의 텍스트 선택"""
    hwp.SelectText(start_para, start_pos, end_para, end_pos)
```

| Parameter | Description |
|-----------|-------------|
| spara | 블록 시작 위치의 문단 번호 |
| spos | 블록 시작 위치의 문단 중에서 문자의 위치 |
| epara | 블록 끝 위치의 문단 번호 |
| epos | 블록 끝 위치의 문단 중에서 문자의 위치 (포함되지 않음) |

---

## 6. 실용 예제

### 예제 1: 텍스트 삽입 후 서식 적용

```python
def insert_formatted_text(hwp, text, font_name="맑은 고딕", font_size=12,
                          bold=False, text_color=None):
    """서식이 적용된 텍스트 삽입"""

    # 현재 위치 저장
    pos = hwp.GetPos()

    # 텍스트 삽입
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

    # 삽입한 텍스트 선택 (시작 위치부터 현재 위치까지)
    current_pos = hwp.GetPos()
    hwp.SelectText(pos[1], pos[2], current_pos[1], current_pos[2])

    # 서식 적용
    hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
    hwp.HParameterSet.HCharShape.FaceNameHangul = font_name
    hwp.HParameterSet.HCharShape.FaceNameLatin = font_name
    hwp.HParameterSet.HCharShape.Height = font_size * 100  # pt to HWPUNIT
    hwp.HParameterSet.HCharShape.Bold = 1 if bold else 0
    if text_color:
        hwp.HParameterSet.HCharShape.TextColor = text_color
    hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)

    # 선택 해제
    hwp.HAction.Run("Cancel")
```

### 예제 2: 셀에 정렬된 텍스트 삽입

```python
def set_cell_text_with_align(hwp, text, align="center"):
    """현재 셀에 정렬된 텍스트 삽입"""

    align_map = {
        "justify": 0,
        "left": 1,
        "right": 2,
        "center": 3,
        "distribute": 4,
        "divide": 5
    }

    # 셀 내용 선택 및 삭제
    hwp.HAction.Run("SelectAll")
    hwp.HAction.Run("Delete")

    # 텍스트 삽입
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

    # 전체 선택 후 정렬 적용
    hwp.HAction.Run("SelectAll")
    hwp.HAction.GetDefault("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)
    hwp.HParameterSet.HParaShape.AlignType = align_map.get(align, 3)
    hwp.HAction.Execute("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)
    hwp.HAction.Run("Cancel")
```

### 예제 3: 전체 문서 서식 일괄 적용

```python
def apply_document_style(hwp, font_name="맑은 고딕", font_size=10, line_spacing=160):
    """문서 전체에 스타일 적용"""

    # 문서 시작으로 이동
    hwp.MovePos(2)  # moveTopOfFile

    # 전체 선택
    hwp.HAction.Run("SelectAll")

    # 글자 모양 적용
    hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
    hwp.HParameterSet.HCharShape.FaceNameHangul = font_name
    hwp.HParameterSet.HCharShape.FaceNameLatin = font_name
    hwp.HParameterSet.HCharShape.Height = font_size * 100
    hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)

    # 문단 모양 적용
    hwp.HAction.GetDefault("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)
    hwp.HParameterSet.HParaShape.LineSpacingType = 0  # 글자에 따라
    hwp.HParameterSet.HParaShape.LineSpacing = line_spacing
    hwp.HAction.Execute("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)

    # 선택 해제
    hwp.HAction.Run("Cancel")
```

---

## 7. 주요 액션 ID 정리

### 텍스트 관련

| Action ID | ParameterSet ID | Description |
|-----------|-----------------|-------------|
| InsertText | InsertText | 텍스트 삽입 |
| InsertSpace | - | 공백 삽입 |
| InsertTab | - | 탭 삽입 |
| BreakPara | - | 문단 나누기 |
| BreakLine | - | 줄 나누기 |

### 서식 관련

| Action ID | ParameterSet ID | Description |
|-----------|-----------------|-------------|
| CharShape | CharShape | 글자 모양 적용 |
| ParagraphShape | ParaShape | 문단 모양 적용 |
| ShapeCopyPaste | ShapeCopyPaste | 모양 복사 |

### 선택/이동 관련

| Action ID | Description |
|-----------|-------------|
| SelectAll | 모두 선택 |
| Cancel | 선택 해제 (ESC) |
| Select | 선택 모드 (F3) |
| SelectColumn | 칸 블록 선택 (F4) |

---

## 8. 참고 사항

### HWPUNIT 단위 변환

| 단위 | HWPUNIT 값 | 변환 공식 |
|------|-----------|----------|
| 1 pt | 100 | `hwpunit / 100 = pt` |
| 1 inch | 7,200 | `hwpunit / 7200 = inch` |
| 1 cm | 2,834.6 | `hwpunit / 2834.6 = cm` |
| 1 mm | 283.46 | `hwpunit / 283.46 = mm` |

### HAction 사용 패턴

```python
# 패턴 1: Run - 파라미터 없는 단순 액션
hwp.HAction.Run("SelectAll")

# 패턴 2: GetDefault + Execute - 파라미터 설정 필요한 액션
hwp.HAction.GetDefault("ActionID", hwp.HParameterSet.HXxx.HSet)
# ... 파라미터 설정 ...
hwp.HAction.Execute("ActionID", hwp.HParameterSet.HXxx.HSet)
```

### 속성 직접 접근 (읽기 전용)

```python
# 현재 선택 영역의 글자 모양 읽기
char_shape = hwp.CharShape
font_height = char_shape.Item("Height")

# 현재 선택 영역의 문단 모양 읽기
para_shape = hwp.ParaShape
align_type = para_shape.Item("AlignType")
```
