# HWP 서식(Style) API 가이드

HWP 자동화에서 서식(스타일) 관련 기능을 정리한 문서입니다.

---

## 목차

1. [개요](#1-개요)
2. [글자 모양 (CharShape)](#2-글자-모양-charshape)
3. [문단 모양 (ParaShape)](#3-문단-모양-parashape)
4. [스타일 (Style)](#4-스타일-style)
5. [모양 복사/붙여넣기](#5-모양-복사붙여넣기)
6. [스타일 내보내기/가져오기](#6-스타일-내보내기가져오기)
7. [찾기/바꾸기에서 서식 활용](#7-찾기바꾸기에서-서식-활용)

---

## 1. 개요

HWP에서 서식은 크게 세 가지로 구분됩니다:

| 구분 | 설명 | 파라미터셋 |
|------|------|-----------|
| **글자 모양** | 글꼴, 크기, 색상, 밑줄 등 | `HCharShape` |
| **문단 모양** | 여백, 정렬, 줄간격, 번호/글머리표 등 | `HParaShape` |
| **스타일** | 글자+문단 모양의 조합을 이름으로 저장 | `HStyle`, `HStyleItem` |

### 기본 사용 패턴

```python
# 1. 기본값 가져오기
hwp.HAction.GetDefault("ActionName", hwp.HParameterSet.XXX.HSet)

# 2. 속성 설정
hwp.HParameterSet.XXX.PropertyName = value

# 3. 실행
hwp.HAction.Execute("ActionName", hwp.HParameterSet.XXX.HSet)
```

---

## 2. 글자 모양 (CharShape)

### 2.1 액션 목록

| 액션 | 설명 | 단축키 |
|------|------|--------|
| `CharShape` | 글자 모양 적용 (코드용) | - |
| `CharShapeDialog` | 글자 모양 대화상자 | Alt+L |
| `CharShapeBold` | 굵게 토글 | Ctrl+B |
| `CharShapeItalic` | 기울임 토글 | Ctrl+I |
| `CharShapeUnderline` | 밑줄 토글 | Ctrl+U |
| `CharShapeStrikeOut` | 취소선 토글 | Ctrl+D |
| `CharShapeSuperScript` | 위첨자 토글 | - |
| `CharShapeSubScript` | 아래첨자 토글 | - |
| `CharShapeNormal` | 보통 모양으로 | Alt+Shift+C |

### 2.2 HCharShape 파라미터셋 속성

#### 글꼴 관련

| 속성 | 타입 | 설명 | 기본값 |
|------|------|------|--------|
| `FaceNameHangul` | BSTR | 한글 글꼴 | "함초롬돋움" |
| `FaceNameLatin` | BSTR | 영문 글꼴 | "함초롬돋움" |
| `FaceNameHanja` | BSTR | 한자 글꼴 | - |
| `FaceNameJapanese` | BSTR | 일본어 글꼴 | - |
| `FaceNameOther` | BSTR | 기타 글꼴 | - |
| `FaceNameSymbol` | BSTR | 기호 글꼴 | - |
| `FaceNameUser` | BSTR | 사용자 글꼴 | - |

#### 크기 관련

| 속성 | 타입 | 설명 | 범위 |
|------|------|------|------|
| `Height` | LONG | 글자 크기 (1/100pt) | 100-409600 |
| `SizeHangul` | SHORT | 한글 비율 (%) | 10-250 |
| `SizeLatin` | SHORT | 영문 비율 (%) | 10-250 |
| `RatioHangul` | SHORT | 한글 장평 (%) | 50-200 |
| `RatioLatin` | SHORT | 영문 장평 (%) | 50-200 |
| `SpacingHangul` | SHORT | 한글 자간 (%) | -50 ~ 50 |
| `SpacingLatin` | SHORT | 영문 자간 (%) | -50 ~ 50 |

#### 서식 관련

| 속성 | 타입 | 설명 | 값 |
|------|------|------|-----|
| `Bold` | BOOL | 굵게 | 0/1 |
| `Italic` | BOOL | 기울임 | 0/1 |
| `SmallCaps` | BOOL | 작은 대문자 | 0/1 |
| `Emboss` | BOOL | 양각 | 0/1 |
| `Engrave` | BOOL | 음각 | 0/1 |
| `SuperScript` | BOOL | 위첨자 | 0/1 |
| `SubScript` | BOOL | 아래첨자 | 0/1 |

#### 밑줄 관련

| 속성 | 타입 | 설명 | 값 |
|------|------|------|-----|
| `UnderlineType` | BYTE | 밑줄 위치 | 0=없음, 1=글자아래, 2=어절아래, 3=위쪽 |
| `UnderlineShape` | BYTE | 밑줄 모양 | 0=실선, 1=점선, 2=굵은실선, 등 |
| `UnderlineColor` | COLORREF | 밑줄 색상 | RGB (0x00BBGGRR) |

#### 외곽선/그림자

| 속성 | 타입 | 설명 | 값 |
|------|------|------|-----|
| `OutlineType` | BYTE | 외곽선 종류 | 0=없음, 1=실선, 2=점선, 3=굵은실선, 4=파선, 5=일점쇄선, 6=이점쇄선 |
| `ShadowType` | BYTE | 그림자 종류 | 0=없음, 1=비연속, 2=연속 |
| `ShadowOffsetX` | SHORT | 그림자 X오프셋 (%) | -100 ~ 100 |
| `ShadowOffsetY` | SHORT | 그림자 Y오프셋 (%) | -100 ~ 100 |
| `ShadowColor` | COLORREF | 그림자 색상 | RGB |

#### 색상

| 속성 | 타입 | 설명 |
|------|------|------|
| `TextColor` | COLORREF | 글자 색상 (0x00BBGGRR) |
| `ShadeColor` | COLORREF | 음영 색상 |

#### 취소선/강조점

| 속성 | 타입 | 설명 |
|------|------|------|
| `StrikeOutType` | BYTE | 취소선 종류 (0=없음, 1~7=다양한 스타일) |
| `StrikeOutShape` | BYTE | 취소선 모양 |
| `StrikeOutColor` | COLORREF | 취소선 색상 |
| `DiacSymMark` | BYTE | 강조점 (0=없음, 1=점, 2=원, 등) |

### 2.3 사용 예제

```python
# 글자 모양 변경 (선택 영역에 적용)
hwp.HAction.Run("SelectAll")  # 또는 특정 영역 선택

hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
hwp.HParameterSet.HCharShape.FaceNameHangul = "맑은 고딕"
hwp.HParameterSet.HCharShape.FaceNameLatin = "Arial"
hwp.HParameterSet.HCharShape.Height = 1200  # 12pt
hwp.HParameterSet.HCharShape.Bold = 1
hwp.HParameterSet.HCharShape.TextColor = 0x00FF0000  # 빨강 (BGR)
hwp.HParameterSet.HCharShape.UnderlineType = 1
hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)

hwp.HAction.Run("Cancel")  # 선택 해제
```

---

## 3. 문단 모양 (ParaShape)

### 3.1 액션 목록

| 액션 | 설명 | 단축키 |
|------|------|--------|
| `ParagraphShape` | 문단 모양 적용 (코드용) | - |
| `ParaShapeDialog` | 문단 모양 대화상자 | Alt+T |
| `ParaNumberBullet` | 문단번호/글머리표 | - |
| `BulletDlg` | 글머리표 대화상자 | - |

### 3.2 HParaShape 파라미터셋 속성

#### 여백/들여쓰기

| 속성 | 타입 | 설명 | 단위 |
|------|------|------|------|
| `LeftMargin` | LONG | 왼쪽 여백 | 1/100pt |
| `RightMargin` | LONG | 오른쪽 여백 | 1/100pt |
| `Indentation` | LONG | 첫줄 들여쓰기 | 1/100pt |

#### 간격

| 속성 | 타입 | 설명 | 단위 |
|------|------|------|------|
| `PrevSpacing` | LONG | 문단 위 간격 | 1/100pt |
| `NextSpacing` | LONG | 문단 아래 간격 | 1/100pt |
| `LineSpacing` | LONG | 줄 간격 | 1/100pt 또는 % |
| `LineSpacingType` | BYTE | 줄간격 종류 | 0=글자에따라, 1=고정값, 2=여백만지정, 3=최소 |

#### 정렬

| 속성 | 타입 | 설명 | 값 |
|------|------|------|-----|
| `AlignType` | BYTE | 정렬 방식 | 0=양쪽, 1=왼쪽, 2=오른쪽, 3=가운데, 4=배분, 5=나눔 |
| `TextAlignment` | BYTE | 세로 정렬 | 0=글꼴기준, 1=위, 2=가운데, 3=아래 |
| `TextDir` | BYTE | 텍스트 방향 | 0=자동, 1=오른쪽→왼쪽, 2=왼쪽→오른쪽 |

#### 문단 옵션

| 속성 | 타입 | 설명 |
|------|------|------|
| `BreakLatinWord` | BOOL | 영어 단어 나눔 |
| `BreakNonLatinWord` | BOOL | 비라틴 단어 나눔 |
| `SnapToGrid` | BOOL | 격자에 맞춤 |
| `Condense` | BYTE | 한 줄로 압축 |
| `WidowOrphan` | BOOL | 외톨이 줄 방지 |
| `KeepWithNext` | BOOL | 다음 문단과 함께 |
| `KeepLinesTogether` | BOOL | 문단 보호 |
| `PagebreakBefore` | BOOL | 문단 앞에서 항상 쪽 나눔 |

#### 머리/꼬리 모양 (번호/글머리표)

| 속성 | 타입 | 설명 | 값 |
|------|------|------|-----|
| `HeadingType` | BYTE | 머리 모양 종류 | 0=없음, 1=개요, 2=번호, 3=불릿 |
| `Level` | BYTE | 수준 (1~10) | - |
| `Numbering` | SET | 번호 모양 | NumberingShape |
| `Bullet` | SET | 글머리표 모양 | BulletShape |

### 3.3 사용 예제

```python
# 문단 모양 변경
hwp.HAction.Run("SelectAll")

hwp.HAction.GetDefault("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)
hwp.HParameterSet.HParaShape.AlignType = 3  # 가운데 정렬
hwp.HParameterSet.HParaShape.LineSpacingType = 0  # 글자에 따라
hwp.HParameterSet.HParaShape.LineSpacing = 160  # 160%
hwp.HParameterSet.HParaShape.LeftMargin = 500  # 5pt 왼쪽 여백
hwp.HParameterSet.HParaShape.Indentation = 1000  # 10pt 첫줄 들여쓰기
hwp.HAction.Execute("ParagraphShape", hwp.HParameterSet.HParaShape.HSet)

hwp.HAction.Run("Cancel")
```

---

## 4. 스타일 (Style)

### 4.1 스타일 액션 목록

| 액션 | 설명 | 파라미터셋 |
|------|------|-----------|
| `Style` | 스타일 적용 (글2005 이하) | HStyle |
| `StyleEx` | 스타일 적용 (글2007 이상) | HStyle |
| `StyleAdd` | 스타일 추가 (대화상자) | HStyle |
| `StyleEdit` | 스타일 편집 | HStyle |
| `StyleDelete` | 스타일 삭제 | HStyleDelete |
| `StyleChangeToCurrentShape` | 현재 모양으로 스타일 변경 | HStyleItem |
| `StyleClearCharStyle` | 글자 스타일 해제 | - |
| `StyleShortcut1~10` | 스타일 단축키 | - |

### 4.2 스타일 관련 파라미터셋

#### HStyle

| 속성 | 타입 | 설명 |
|------|------|------|
| `Apply` | INT | 적용할 스타일 인덱스 |

#### HStyleDelete

| 속성 | 타입 | 설명 |
|------|------|------|
| `Target` | INT | 삭제할 스타일 인덱스 |
| `Alternation` | INT | 대체할 스타일 인덱스 |

#### HStyleItem (스타일 세부 편집)

| 속성 | 타입 | 설명 |
|------|------|------|
| `Type` | BYTE | 스타일 종류 |
| `NameLocal` | BSTR | 스타일 이름 (한글) |
| `NameEng` | BSTR | 스타일 이름 (영문) |
| `Next` | INT | 다음 스타일 인덱스 |
| `LockForm` | BOOL | 양식 스타일 보호 |
| `CharShape` | SET | 글자 모양 (HCharShape) |
| `ParaShape` | SET | 문단 모양 (HParaShape) |

### 4.3 스타일 사용 예제

```python
# 스타일 적용 (인덱스로)
hwp.HAction.GetDefault("Style", hwp.HParameterSet.HStyle.HSet)
hwp.HParameterSet.HStyle.Apply = 1  # 첫 번째 스타일
hwp.HAction.Execute("Style", hwp.HParameterSet.HStyle.HSet)

# 스타일 단축키 사용
hwp.HAction.Run("StyleShortcut1")  # Ctrl+1
hwp.HAction.Run("StyleShortcut2")  # Ctrl+2
```

---

## 5. 모양 복사/붙여넣기

### 5.1 ShapeCopyPaste 액션

글자/문단 모양을 복사하여 다른 텍스트에 적용합니다.

#### HShapeCopyPaste 파라미터셋

| 속성 | 타입 | 설명 | 값 |
|------|------|------|-----|
| `Type` | BYTE | 복사 종류 | 0=글자모양, 1=문단모양, 2=글자+문단, 3=글자스타일, 4=문단스타일 |
| `CellAttr` | BOOL | 셀 모양 복사 | 0/1 |
| `CellBorder` | BOOL | 셀 테두리 복사 | 0/1 |
| `CellFill` | BOOL | 셀 배경 복사 | 0/1 |
| `TypeBodyAndCellOnly` | BOOL | 본문+셀 또는 셀만 | 0/1 |

### 5.2 사용 예제

```python
# 1단계: 원본 텍스트 선택 후 모양 복사
hwp.SetPos(list_id, para1, pos1)
hwp.HAction.Run("MoveSelRight")  # 원본 선택

hwp.HAction.GetDefault("ShapeCopyPaste", hwp.HParameterSet.HShapeCopyPaste.HSet)
hwp.HParameterSet.HShapeCopyPaste.Type = 0  # 글자 모양 복사
hwp.HAction.Execute("ShapeCopyPaste", hwp.HParameterSet.HShapeCopyPaste.HSet)

# 2단계: 대상 텍스트 선택 후 모양 붙여넣기
hwp.SetPos(list_id, para2, pos2)
hwp.HAction.Run("SelectAll")  # 또는 특정 영역 선택

hwp.HAction.GetDefault("ShapeCopyPaste", hwp.HParameterSet.HShapeCopyPaste.HSet)
hwp.HParameterSet.HShapeCopyPaste.Type = 0
hwp.HAction.Execute("ShapeCopyPaste", hwp.HParameterSet.HShapeCopyPaste.HSet)
```

### 5.3 그리기 개체 모양 복사 (ShapeObjectCopyPaste)

| 속성 | 타입 | 설명 |
|------|------|------|
| `Type` | BYTE | 복사 종류 |
| `ShapeObjectLine` | BOOL | 선 모양 복사 |
| `ShapeObjectFill` | BOOL | 채우기 복사 |
| `ShapeObjectSize` | BOOL | 크기 복사 |
| `ShapeObjectShadow` | BOOL | 그림자 복사 |
| `ShapeObjectPicEffect` | BOOL | 그림 효과 복사 |

---

## 6. 스타일 내보내기/가져오기

### 6.1 ExportStyle / ImportStyle 메서드

스타일을 파일로 내보내거나 가져옵니다.

```python
# 스타일 내보내기
hwp.HParameterSet.HStyleTemplate.FileName = "C:\\styles\\my_style.hwt"
hwp.ExportStyle(hwp.HParameterSet.HStyleTemplate)

# 스타일 가져오기
hwp.HParameterSet.HStyleTemplate.FileName = "C:\\styles\\my_style.hwt"
hwp.ImportStyle(hwp.HParameterSet.HStyleTemplate)
```

### 6.2 StyleTemplate 액션

스타일 마당 대화상자를 통해 스타일을 관리합니다.

```python
hwp.HAction.GetDefault("StyleTemplate", hwp.HParameterSet.HStyleTemplate.HSet)
hwp.HParameterSet.HStyleTemplate.FileName = "template.hwt"
hwp.HAction.Execute("StyleTemplate", hwp.HParameterSet.HStyleTemplate.HSet)
```

---

## 7. 찾기/바꾸기에서 서식 활용

### 7.1 FindReplace 파라미터셋의 서식 관련 속성

| 속성 | 타입 | 설명 |
|------|------|------|
| `FindCharShape` | SET | 찾을 글자 모양 |
| `FindParaShape` | SET | 찾을 문단 모양 |
| `ReplaceCharShape` | SET | 바꿀 글자 모양 |
| `ReplaceParaShape` | SET | 바꿀 문단 모양 |
| `FindStyle` | BSTR | 찾을 스타일 이름 |
| `ReplaceStyle` | BSTR | 바꿀 스타일 이름 |

### 7.2 서식으로 찾아 바꾸기 예제

```python
# 빨간색 글자를 파란색으로 변경
hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
hwp.HParameterSet.HFindReplace.Direction = 0
hwp.HParameterSet.HFindReplace.FindString = ""  # 텍스트 무시, 서식만
hwp.HParameterSet.HFindReplace.ReplaceString = ""

# 찾을 글자 모양 설정
hwp.HParameterSet.HFindReplace.FindCharShape.TextColor = 0x000000FF  # 빨강

# 바꿀 글자 모양 설정
hwp.HParameterSet.HFindReplace.ReplaceCharShape.TextColor = 0x00FF0000  # 파랑

hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
```

---

## 8. 색상 값 참고

HWP에서 색상은 **COLORREF** 형식을 사용합니다: `0x00BBGGRR`

| 색상 | COLORREF 값 |
|------|-------------|
| 검정 | 0x00000000 |
| 빨강 | 0x000000FF |
| 초록 | 0x0000FF00 |
| 파랑 | 0x00FF0000 |
| 노랑 | 0x0000FFFF |
| 흰색 | 0x00FFFFFF |

```python
# RGB를 COLORREF로 변환
def rgb_to_colorref(r, g, b):
    return b << 16 | g << 8 | r

# 예: 빨강(255, 0, 0)
red = rgb_to_colorref(255, 0, 0)  # 0x000000FF
```

---

## 9. 주의사항

1. **블록 지정 필수**: 서식 변경은 반드시 텍스트를 선택한 후 적용해야 합니다.
2. **선택 해제**: 서식 적용 후 `hwp.HAction.Run("Cancel")`로 선택 해제
3. **스타일 인덱스**: 스타일은 0부터 시작하는 인덱스로 관리됩니다.
4. **단위**: 크기 관련 값은 대부분 1/100pt 단위입니다 (1000 = 10pt)
