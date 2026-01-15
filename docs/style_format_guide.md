# HWP 서식(Style/Format) 자동화 가이드

## 개요

HWP에서 서식은 크게 세 가지로 구분됩니다:
1. **글자 모양(CharShape)**: 글꼴, 크기, 색상, 밑줄, 기울임 등
2. **문단 모양(ParaShape)**: 여백, 정렬, 줄간격, 탭 등
3. **스타일(Style)**: 글자 모양 + 문단 모양의 조합

---

## 1. 글자 모양(CharShape)

### 1.1 주요 속성

| 속성명 | 타입 | 설명 |
|--------|------|------|
| FaceNameHangul | BSTR | 한글 글꼴명 |
| FaceNameLatin | BSTR | 영문 글꼴명 |
| Height | I4 | 글자 크기 (HWPUNIT, 1pt = 100) |
| Bold | UI1 | 굵게 (0=off, 1=on) |
| Italic | UI1 | 기울임 (0=off, 1=on) |
| TextColor | UI4 | 글자색 (0x00BBGGRR) |
| UnderlineType | UI1 | 밑줄 (0=없음, 1=하단, 2=중앙, 3=상단) |
| UnderlineShape | UI1 | 밑줄 모양 (선 종류) |
| UnderlineColor | UI4 | 밑줄 색상 |
| StrikeOutType | UI1 | 취소선 (0=없음, 1=빨강단일, 2=빨강이중, 3=글자색단일, 4=글자색이중) |
| SuperScript | UI1 | 위첨자 |
| SubScript | UI1 | 아래첨자 |
| OutlineType | UI1 | 외곽선 (0=없음, 1=실선, 2=점선, 3=굵은선 등) |
| ShadowType | UI1 | 그림자 (0=없음, 1=drop, 2=continuous) |
| ShadowOffsetX/Y | I1 | 그림자 간격 (-100%~100%) |
| ShadowColor | UI4 | 그림자 색상 |
| Emboss | UI1 | 양각 |
| Engrave | UI1 | 음각 |
| RatioHangul | UI1 | 장평 (50%~200%) |
| SpacingHangul | I1 | 자간 (-50%~50%) |
| ShadeColor | UI4 | 음영색 |
| DiacSymMark | UI1 | 강조점 (0=없음, 1=검정동그라미, 2=빈동그라미) |
| UseKerning | UI1 | 커닝 |
| UseFontSpace | UI1 | 글꼴에 어울리는 빈칸 |

### 1.2 글자 모양 변경 예제

```python
# 방법 1: HAction 사용
hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
hwp.HParameterSet.HCharShape.FaceNameHangul = "맑은 고딕"
hwp.HParameterSet.HCharShape.FaceNameLatin = "Arial"
hwp.HParameterSet.HCharShape.Height = 1200  # 12pt
hwp.HParameterSet.HCharShape.Bold = 1
hwp.HParameterSet.HCharShape.TextColor = 0x0000FF  # 빨강 (BGR)
hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)

# 방법 2: CreateAction 사용
act = hwp.CreateAction("CharShape")
pset = act.CreateSet()
act.GetDefault(pset)
pset.SetItem("FaceNameHangul", "나눔고딕")
pset.SetItem("Height", 1000)  # 10pt
pset.SetItem("Italic", 1)
act.Execute(pset)
```

### 1.3 글자 모양 단축 액션

| 액션 ID | 설명 | 단축키 |
|---------|------|--------|
| CharShapeBold | 굵게 | Alt+L |
| CharShapeItalic | 기울임 | Alt+Shift+I |
| CharShapeUnderline | 밑줄 | Alt+Shift+U |
| CharShapeNormal | 보통모양 | Alt+Shift+C |
| CharShapeSuperscript | 위첨자 | Alt+Shift+P |
| CharShapeSubscript | 아래첨자 | Alt+Shift+S |
| CharShapeHeightIncrease | 크기 크게 | Alt+Shift+E |
| CharShapeHeightDecrease | 크기 작게 | Alt+Shift+R |
| CharShapeWidthIncrease | 장평 넓게 | Alt+Shift+K |
| CharShapeWidthDecrease | 장평 좁게 | Alt+Shift+J |
| CharShapeSpacingIncrease | 자간 넓게 | Alt+Shift+W |
| CharShapeSpacingDecrease | 자간 좁게 | Alt+Shift+N |

```python
# 단축 액션 사용
hwp.HAction.Run("CharShapeBold")  # 굵게 토글
hwp.HAction.Run("CharShapeItalic")  # 기울임 토글
```

---

## 2. 문단 모양(ParaShape)

### 2.1 주요 속성

| 속성명 | 타입 | 설명 |
|--------|------|------|
| LeftMargin | I4 | 왼쪽 여백 (HWPUNIT) |
| RightMargin | I4 | 오른쪽 여백 (HWPUNIT) |
| Indentation | I4 | 들여쓰기/내어쓰기 (HWPUNIT) |
| PrevSpacing | I4 | 문단 위 간격 (HWPUNIT) |
| NextSpacing | I4 | 문단 아래 간격 (HWPUNIT) |
| LineSpacingType | UI1 | 줄간격 종류 (0=글자에따라, 1=고정값, 2=여백만) |
| LineSpacing | I4 | 줄간격 값 |
| AlignType | UI1 | 정렬 (0=양쪽, 1=왼쪽, 2=오른쪽, 3=가운데, 4=배분, 5=나눔) |
| BreakLatinWord | UI1 | 영문 줄나눔 (0=단어, 1=하이픈, 2=글자) |
| SnapToGrid | UI1 | 줄격자 사용 |
| WidowOrphan | UI1 | 외톨이줄 보호 |
| KeepWithNext | UI1 | 다음 문단과 함께 |
| KeepLinesTogether | UI1 | 문단 보호 |
| PagebreakBefore | UI1 | 문단 앞 쪽 나눔 |
| HeadingType | UI1 | 문단 머리 (0=없음, 1=개요, 2=번호, 3=불릿) |
| Level | UI1 | 단계 (0~6) |

### 2.2 문단 모양 변경 예제

```python
hwp.HAction.GetDefault("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)
hwp.HParameterSet.HParaShape.AlignType = 3  # 가운데 정렬
hwp.HParameterSet.HParaShape.LineSpacingType = 0  # 글자에 따라
hwp.HParameterSet.HParaShape.LineSpacing = 160  # 160%
hwp.HParameterSet.HParaShape.LeftMargin = 2000  # 왼쪽 여백
hwp.HAction.Execute("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)
```

### 2.3 문단 모양 단축 액션

| 액션 ID | 설명 |
|---------|------|
| ParagraphShapeAlignLeft | 왼쪽 정렬 |
| ParagraphShapeAlignCenter | 가운데 정렬 |
| ParagraphShapeAlignRight | 오른쪽 정렬 |
| ParagraphShapeAlignJustify | 양쪽 정렬 |
| ParagraphShapeIncreaseLeftMargin | 왼쪽 여백 키우기 |
| ParagraphShapeDecreaseLeftMargin | 왼쪽 여백 줄이기 |
| ParagraphShapeIncreaseLineSpacing | 줄간격 넓히기 |
| ParagraphShapeDecreaseLineSpacing | 줄간격 좁히기 |

```python
hwp.HAction.Run("ParagraphShapeAlignCenter")  # 가운데 정렬
```

---

## 3. 스타일(Style)

### 3.1 스타일 관련 액션

| 액션 ID | 파라미터셋 | 설명 |
|---------|-----------|------|
| Style | Style | 스타일 적용 (글2005이하) |
| StyleEx | Style | 스타일 적용 (글2007+) |
| StyleAdd | Style | 스타일 추가 (대화상자) |
| StyleEdit | Style | 스타일 편집 |
| StyleDelete | StyleDelete | 스타일 삭제 |
| StyleChangeToCurrentShape | StyleItem | 현재 모양으로 스타일 변경 |
| StyleClearCharStyle | - | 글자 스타일 해제 |
| StyleShortcut1~10 | - | 스타일 단축키 (Ctrl+1~0) |

### 3.2 Style 파라미터셋

```
| Item ID | Type | 설명 |
|---------|------|------|
| Apply   | PIT_I | 적용할 스타일 인덱스 |
```

### 3.3 StyleDelete 파라미터셋

```
| Item ID     | Type | 설명 |
|-------------|------|------|
| Target      | PIT_I | 삭제할 스타일 인덱스 |
| Alternation | PIT_I | 대체할 스타일 인덱스 |
```

### 3.4 StyleItem 파라미터셋 (스타일 정의)

```
| Item ID   | Type     | SubType   | 설명 |
|-----------|----------|-----------|------|
| Type      | PIT_UI1  |           | 스타일 종류 |
| NameLocal | PIT_BSTR |           | 스타일 이름(로컬) |
| NameEng   | PIT_BSTR |           | 스타일 이름(영문) |
| Next      | PIT_I    |           | 다음 스타일 인덱스 |
| LockForm  | PIT_UI1  |           | 양식스타일 보호 |
| CharShape | PIT_SET  | CharShape | 글자 모양 |
| ParaShape | PIT_SET  | ParaShape | 문단 모양 |
```

### 3.5 스타일 적용 예제

```python
# 스타일 인덱스로 적용
hwp.HAction.GetDefault("StyleEx", hwp.HParameterSet.HStyle.HSet)
hwp.HParameterSet.HStyle.Apply = 0  # 첫 번째 스타일
hwp.HAction.Execute("StyleEx", hwp.HParameterSet.HStyle.HSet)

# 스타일 단축키 사용
hwp.HAction.Run("StyleShortcut1")  # Ctrl+1 스타일 적용
```

### 3.6 스타일 내보내기/가져오기

```python
# ExportStyle - 스타일을 파일로 내보내기
# ImportStyle - 파일에서 스타일 가져오기

# HStyleTemplate 파라미터셋 사용
# | FileName | PIT_BSTR | 파일 이름 |
```

---

## 4. 모양 복사(ShapeCopyPaste)

### 4.1 ShapeCopyPaste 파라미터셋

| Item ID | Type | 설명 |
|---------|------|------|
| Type | PIT_I | 복사 종류 (0=글자모양, 1=문단모양, 2=둘다, 3=글자스타일, 4=문단스타일) |
| CellAttr | PIT_UI1 | 셀 모양 복사 |
| CellBorder | PIT_UI1 | 선 모양 복사 |
| CellFill | PIT_UI1 | 셀 배경 복사 |
| TypeBodyAndCellOnly | PIT_I | 본문+셀 모양 복사 or 셀만 복사 |

### 4.2 모양 복사/붙여넣기 예제

```python
# 1. 원본 위치에서 모양 복사
hwp.HAction.GetDefault("ShapeCopyPaste", hwp.HParameterSet.HShapeCopyPaste.HSet)
hwp.HParameterSet.HShapeCopyPaste.Type = 0  # 글자 모양만
hwp.HAction.Execute("ShapeCopyPaste", hwp.HParameterSet.HShapeCopyPaste.HSet)

# 2. 대상 위치로 이동 후 모양 붙여넣기
# (HWP 내부 클립보드에 복사됨, 이후 적용 대상 선택 후 붙여넣기)
```

### 4.3 Alt+C (모양 복사) 시뮬레이션

```python
# 모양 복사 (Alt+C와 동일)
act = hwp.CreateAction("ShapeCopyPaste")
pset = act.CreateSet()
act.GetDefault(pset)
pset.SetItem("Type", 2)  # 글자 + 문단 모양 모두 복사
act.Execute(pset)
```

---

## 5. 색상 값 참고

HWP에서 색상은 **BGR** 형식 (0x00BBGGRR)을 사용합니다.

| 색상 | BGR 값 |
|------|--------|
| 검정 | 0x000000 |
| 빨강 | 0x0000FF |
| 녹색 | 0x00FF00 |
| 파랑 | 0xFF0000 |
| 노랑 | 0x00FFFF |
| 흰색 | 0xFFFFFF |

```python
# RGB를 BGR로 변환
def rgb_to_bgr(r, g, b):
    return (b << 16) | (g << 8) | r

# 사용 예
red = rgb_to_bgr(255, 0, 0)  # 0x0000FF
```

---

## 6. HWPUNIT 단위 참고

- 1pt = 100 HWPUNIT
- 1mm = 283.465 HWPUNIT (약 283)
- 1cm = 2834.65 HWPUNIT (약 2835)
- 1inch = 7200 HWPUNIT

```python
def pt_to_hwpunit(pt):
    return int(pt * 100)

def mm_to_hwpunit(mm):
    return int(mm * 283.465)
```

---

## 참고 자료

- [pyhwpx PyPI](https://pypi.org/project/pyhwpx/)
- [hwpapi PyPI](https://pypi.org/project/hwpapi/)
- [한컴디벨로퍼 포럼](https://forum.developer.hancom.com/)
- [pyhwpx Cookbook - WikiDocs](https://wikidocs.net/book/8956)
