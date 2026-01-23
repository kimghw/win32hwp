# HWP 서식 정보 조회 API 문서

최종 수정일: 2026-01-23

이 문서는 HWP API에서 서식 정보를 조회하는 방법을 정리합니다.

---

## 1. CharShape - 글자 모양

### 1.1 속성 접근 방법

```python
# 현재 Selection의 글자 모양을 조회
char_shape = hwp.CharShape

# Item() 메서드로 개별 속성 조회
font_name = char_shape.Item("FaceNameHangul")  # 한글 폰트명
font_size = char_shape.Item("Height")          # 글자 크기 (HWPUNIT)
text_color = char_shape.Item("TextColor")      # 글자색 (COLORREF)
```

### 1.2 주요 Item 키 목록

#### 폰트 이름 (FaceName)

| Item ID | Type | 설명 |
|---------|------|------|
| FaceNameHangul | PIT_BSTR | 글꼴 이름 (한글) |
| FaceNameLatin | PIT_BSTR | 글꼴 이름 (영문) |
| FaceNameHanja | PIT_BSTR | 글꼴 이름 (한자) |
| FaceNameJapanese | PIT_BSTR | 글꼴 이름 (일본어) |
| FaceNameOther | PIT_BSTR | 글꼴 이름 (기타) |
| FaceNameSymbol | PIT_BSTR | 글꼴 이름 (심벌) |
| FaceNameUser | PIT_BSTR | 글꼴 이름 (사용자) |

#### 폰트 종류 (FontType)

| Item ID | Type | 설명 |
|---------|------|------|
| FontTypeHangul | PIT_UI1 | 폰트 종류 (한글): 0=don't care, 1=TTF, 2=HFT |
| FontTypeLatin | PIT_UI1 | 폰트 종류 (영문) |
| FontTypeHanja | PIT_UI1 | 폰트 종류 (한자) |
| FontTypeJapanese | PIT_UI1 | 폰트 종류 (일본어) |
| FontTypeOther | PIT_UI1 | 폰트 종류 (기타) |
| FontTypeSymbol | PIT_UI1 | 폰트 종류 (심벌) |
| FontTypeUser | PIT_UI1 | 폰트 종류 (사용자) |

#### 글자 크기 및 비율

| Item ID | Type | 설명 |
|---------|------|------|
| Height | PIT_I4 | 글자 크기 (HWPUNIT, 100 = 1pt) |
| SizeHangul | PIT_UI1 | 한글 크기 비율 (10% - 250%) |
| SizeLatin | PIT_UI1 | 영문 크기 비율 |
| SizeHanja | PIT_UI1 | 한자 크기 비율 |
| SizeJapanese | PIT_UI1 | 일본어 크기 비율 |
| SizeOther | PIT_UI1 | 기타 크기 비율 |
| SizeSymbol | PIT_UI1 | 심벌 크기 비율 |
| SizeUser | PIT_UI1 | 사용자 크기 비율 |

#### 장평 비율 (Ratio)

| Item ID | Type | 설명 |
|---------|------|------|
| RatioHangul | PIT_UI1 | 한글 장평 비율 (50% - 200%) |
| RatioLatin | PIT_UI1 | 영문 장평 비율 |
| RatioHanja | PIT_UI1 | 한자 장평 비율 |
| RatioJapanese | PIT_UI1 | 일본어 장평 비율 |
| RatioOther | PIT_UI1 | 기타 장평 비율 |
| RatioSymbol | PIT_UI1 | 심벌 장평 비율 |
| RatioUser | PIT_UI1 | 사용자 장평 비율 |

#### 자간 (Spacing)

| Item ID | Type | 설명 |
|---------|------|------|
| SpacingHangul | PIT_I1 | 한글 자간 (-50% - 50%) |
| SpacingLatin | PIT_I1 | 영문 자간 |
| SpacingHanja | PIT_I1 | 한자 자간 |
| SpacingJapanese | PIT_I1 | 일본어 자간 |
| SpacingOther | PIT_I1 | 기타 자간 |
| SpacingSymbol | PIT_I1 | 심벌 자간 |
| SpacingUser | PIT_I1 | 사용자 자간 |

#### 오프셋 (Offset)

| Item ID | Type | 설명 |
|---------|------|------|
| OffsetHangul | PIT_I1 | 한글 오프셋 (-100% - 100%) |
| OffsetLatin | PIT_I1 | 영문 오프셋 |
| OffsetHanja | PIT_I1 | 한자 오프셋 |
| OffsetJapanese | PIT_I1 | 일본어 오프셋 |
| OffsetOther | PIT_I1 | 기타 오프셋 |
| OffsetSymbol | PIT_I1 | 심벌 오프셋 |
| OffsetUser | PIT_I1 | 사용자 오프셋 |

#### 글자 스타일 (굵기, 기울임 등)

| Item ID | Type | 설명 |
|---------|------|------|
| Bold | PIT_UI1 | 굵게: 0=off, 1=on |
| Italic | PIT_UI1 | 기울임: 0=off, 1=on |
| SmallCaps | PIT_UI1 | 작은 대문자: 0=off, 1=on |
| Emboss | PIT_UI1 | 양각: 0=off, 1=on |
| Engrave | PIT_UI1 | 음각: 0=off, 1=on |
| SuperScript | PIT_UI1 | 위첨자: 0=off, 1=on |
| SubScript | PIT_UI1 | 아래첨자: 0=off, 1=on |

#### 색상 (COLORREF 형식: 0x00BBGGRR)

| Item ID | Type | 설명 |
|---------|------|------|
| TextColor | PIT_UI4 | 글자색 |
| ShadeColor | PIT_UI4 | 음영색 |
| UnderlineColor | PIT_UI4 | 밑줄 색 |
| ShadowColor | PIT_UI4 | 그림자 색 |
| StrikeOutColor | PIT_UI4 | 취소선 색 |

#### 밑줄 (Underline)

| Item ID | Type | 설명 |
|---------|------|------|
| UnderlineType | PIT_UI1 | 밑줄 종류: 0=none, 1=bottom, 2=center, 3=top |
| UnderlineShape | PIT_UI1 | 밑줄 모양 (선 종류) |
| UnderlineColor | PIT_UI4 | 밑줄 색 |

#### 외곽선/그림자 (Outline/Shadow)

| Item ID | Type | 설명 |
|---------|------|------|
| OutlineType | PIT_UI1 | 외곽선 종류: 0=none, 1=solid, 2=dot, 3=thick, 4=dash, 5=dashdot, 6=dashdotdot |
| ShadowType | PIT_UI1 | 그림자 종류: 0=none, 1=drop, 2=continuous |
| ShadowOffsetX | PIT_I1 | 그림자 X 간격 (-100% - 100%) |
| ShadowOffsetY | PIT_I1 | 그림자 Y 간격 (-100% - 100%) |
| ShadowColor | PIT_UI4 | 그림자 색 |

#### 취소선 (StrikeOut)

| Item ID | Type | 설명 |
|---------|------|------|
| StrikeOutType | PIT_UI1 | 취소선 종류: 0=none, 1=red single, 2=red double, 3=text single, 4=text double |
| StrikeOutShape | PIT_UI1 | 취소선 모양 (선 종류) |
| StrikeOutColor | PIT_UI4 | 취소선 색 |

##### StrikeOutShape 값

| 값 | 설명 |
|----|------|
| 0 | 실선 |
| 1 | 파선 |
| 2 | 점선 |
| 3 | 일점쇄선 |
| 4 | 이점쇄선 |
| 5 | 긴 파선 |
| 6 | 원형 점선 |
| 7 | 이중 실선 |
| 8 | 얇고 굵은 이중선 |
| 9 | 굵고 얇은 이중선 |
| 10 | 얇고 굵고 얇은 삼중선 |
| 11 | 물결선 |
| 12 | 이중 물결선 |
| 13 | 3D 굵은선 |
| 14 | 3D 굵은선 (광원 반대) |
| 15 | 3D 실선 |
| 16 | 3D 실선 (광원 반대) |

#### 기타 속성

| Item ID | Type | 설명 |
|---------|------|------|
| DiacSymMark | PIT_UI1 | 강조점 종류: 0=none, 1=검정 동그라미, 2=속 빈 동그라미 |
| UseFontSpace | PIT_UI1 | 글꼴에 어울리는 빈칸: 0=off, 1=on |
| UseKerning | PIT_UI1 | 커닝: 0=off, 1=on |
| BorderFill | PIT_SET | 테두리/배경 (BorderFill ParameterSet) |

---

## 2. ParaShape - 문단 모양

### 2.1 속성 접근 방법

```python
# 현재 Selection의 문단 모양을 조회
para_shape = hwp.ParaShape

# Item() 메서드로 개별 속성 조회
align_type = para_shape.Item("AlignType")      # 정렬 방식
line_spacing = para_shape.Item("LineSpacing")  # 줄 간격
left_margin = para_shape.Item("LeftMargin")    # 왼쪽 여백
```

### 2.2 주요 Item 키 목록

#### 여백 및 들여쓰기

| Item ID | Type | 설명 |
|---------|------|------|
| LeftMargin | PIT_I4 | 왼쪽 여백 (HWPUNIT) |
| RightMargin | PIT_I4 | 오른쪽 여백 (HWPUNIT) |
| Indentation | PIT_I4 | 들여쓰기/내어쓰기 (HWPUNIT, 음수=내어쓰기) |

#### 문단 간격

| Item ID | Type | 설명 |
|---------|------|------|
| PrevSpacing | PIT_I4 | 문단 간격 위 (HWPUNIT) |
| NextSpacing | PIT_I4 | 문단 간격 아래 (HWPUNIT) |

#### 줄 간격

| Item ID | Type | 설명 |
|---------|------|------|
| LineSpacingType | PIT_UI1 | 줄 간격 종류: 0=글자에 따라, 1=고정값, 2=여백만 지정 |
| LineSpacing | PIT_I4 | 줄 간격 값 (종류에 따라 단위 다름) |

**LineSpacing 값 해석:**
- `LineSpacingType=0` (글자에 따라): 0-500% 비율
- `LineSpacingType=1` (고정값): HWPUNIT 값
- `LineSpacingType=2` (여백만 지정): HWPUNIT 값

#### 정렬

| Item ID | Type | 설명 |
|---------|------|------|
| AlignType | PIT_UI1 | 정렬 방식 |
| TextAlignment | PIT_UI1 | 세로 정렬 |

**AlignType 값:**

| 값 | 설명 |
|----|------|
| 0 | 양쪽 정렬 |
| 1 | 왼쪽 정렬 |
| 2 | 오른쪽 정렬 |
| 3 | 가운데 정렬 |
| 4 | 배분 정렬 |
| 5 | 나눔 정렬 (공백에만 배분) |

**TextAlignment 값:**

| 값 | 설명 |
|----|------|
| 0 | 글꼴 기준 |
| 1 | 위 |
| 2 | 가운데 |
| 3 | 아래 |

#### 줄 나눔

| Item ID | Type | 설명 |
|---------|------|------|
| BreakLatinWord | PIT_UI1 | 줄 나눔 단위 (라틴 문자): 0=단어, 1=하이픈, 2=글자 |
| BreakNonLatinWord | PIT_UI1 | 줄 나눔 단위 (비라틴 문자): TRUE=글자, FALSE=어절 |
| SnapToGrid | PIT_UI1 | 편집 용지의 줄 격자 사용: on/off |
| Condense | PIT_UI1 | 공백 최소값 (0-75%) |

#### 문단 보호

| Item ID | Type | 설명 |
|---------|------|------|
| WidowOrphan | PIT_UI1 | 외톨이줄 보호: on/off |
| KeepWithNext | PIT_UI1 | 다음 문단과 함께: on/off |
| KeepLinesTogether | PIT_UI1 | 문단 보호: on/off |
| PagebreakBefore | PIT_UI1 | 문단 앞에서 항상 쪽 나눔: on/off |

#### 문단 머리

| Item ID | Type | 설명 |
|---------|------|------|
| HeadingType | PIT_UI1 | 문단 머리 모양: 0=없음, 1=개요, 2=번호, 3=불릿 |
| Level | PIT_UI1 | 단계 (0-6) |
| Numbering | PIT_SET | 문단 번호 (NumberingShape) |
| Bullet | PIT_SET | 불릿 모양 (BulletShape) |
| Checked | PIT_UI1 | 체크 글머리표 체크 여부: on/off |

#### 테두리/배경

| Item ID | Type | 설명 |
|---------|------|------|
| BorderConnect | PIT_UI1 | 테두리 연결: on/off |
| BorderText | PIT_UI1 | 여백 무시: 0=단, 1=텍스트 |
| BorderOffsetLeft | PIT_I | 왼쪽 간격 (HWPUNIT) |
| BorderOffsetRight | PIT_I | 오른쪽 간격 (HWPUNIT) |
| BorderOffsetTop | PIT_I | 위 간격 (HWPUNIT) |
| BorderOffsetBottom | PIT_I | 아래 간격 (HWPUNIT) |
| BorderFill | PIT_SET | 테두리/배경 (BorderFill ParameterSet) |

#### 기타

| Item ID | Type | 설명 |
|---------|------|------|
| FontLineHeight | PIT_UI1 | 글꼴에 어울리는 줄 높이: on/off |
| TailType | PIT_UI1 | 문단 꼬리 모양: on/off |
| LineWrap | PIT_UI1 | 글꼴에 어울리는 줄 높이: on/off |
| TabDef | PIT_SET | 탭 정의 (TabDef ParameterSet) |
| AutoSpaceEAsianEng | PIT_UI1 | 한글과 영어 간격 자동 조절: on/off |
| AutoSpaceEAsianNum | PIT_UI1 | 한글과 숫자 간격 자동 조절: on/off |
| SuppressLineNum | PIT_UI1 | 줄 번호 표시 여부 |
| TextDir | PIT_UI1 | 텍스트 방향: 0=자동, 1=오른편->왼편, 2=왼편->오른편 |

---

## 3. CellShape - 셀 모양

### 3.1 속성 접근 방법

```python
# 현재 선택된 표/셀의 모양 정보 조회
# 주의: 표 내부에 캐럿이 위치해야 함
cell_shape = hwp.CellShape

# Table 속성 조회
page_break = cell_shape.Item("PageBreak")
cell_spacing = cell_shape.Item("CellSpacing")

# Cell 세부 속성 조회 (중첩된 ParameterSet)
cell = cell_shape.Item("Cell")
cell_width = cell.Item("Width")
cell_height = cell.Item("Height")
```

### 3.2 CellShape 구조

CellShape는 `ParameterSet/Table` 형식이며, 내부에 `Cell` 아이템이 `ParameterSet/Cell`로 셀 속성을 포함합니다.

### 3.3 Table 속성 (Table ParameterSet)

#### 페이지 처리

| Item ID | Type | 설명 |
|---------|------|------|
| PageBreak | PIT_UI1 | 페이지 경계 처리: 0=나누지 않음, 1=테이블만 나눔, 2=셀 내 텍스트도 나눔 |
| RepeatHeader | PIT_UI1 | 제목 행 반복: on/off |

#### 셀 여백/간격

| Item ID | Type | 설명 |
|---------|------|------|
| CellSpacing | PIT_UI4 | 셀 간격 (HWPUNIT) |
| CellMarginLeft | PIT_I4 | 기본 셀 안쪽 여백 (왼쪽) |
| CellMarginRight | PIT_I4 | 기본 셀 안쪽 여백 (오른쪽) |
| CellMarginTop | PIT_I4 | 기본 셀 안쪽 여백 (위쪽) |
| CellMarginBottom | PIT_I4 | 기본 셀 안쪽 여백 (아래쪽) |

#### 배치

| Item ID | Type | 설명 |
|---------|------|------|
| TreatAsChar | PIT_UI1 | 글자처럼 취급: on/off |
| AffectsLine | PIT_UI1 | 줄 간격에 영향 (TreatAsChar=TRUE일 때): on/off |
| VertRelTo | PIT_UI1 | 세로 위치 기준: 0=종이, 1=본문영역, 2=문단 |
| VertAlign | PIT_UI1 | 세로 정렬: 0=위, 1=가운데, 2=아래, 3=안쪽, 4=바깥쪽 |
| VertOffset | PIT_UI4 | 세로 오프셋 (HWPUNIT) |
| HorzRelTo | PIT_UI1 | 가로 위치 기준: 0=종이, 1=본문영역, 2=단, 3=문단 |
| HorzAlign | PIT_UI1 | 가로 정렬: 0=왼쪽, 1=가운데, 2=오른쪽, 3=안쪽, 4=바깥쪽 |
| HorzOffset | PIT_I4 | 가로 오프셋 (HWPUNIT) |

#### 크기

| Item ID | Type | 설명 |
|---------|------|------|
| WidthRelTo | PIT_UI1 | 폭 기준: 0=종이, 1=본문영역, 2=단, 3=문단, 4=절대값 |
| Width | PIT_I4 | 폭 값 (기준에 따라 % 또는 HWPUNIT) |
| HeightRelTo | PIT_UI1 | 높이 기준: 0=종이, 1=본문영역, 2=절대값 |
| Height | PIT_I4 | 높이 값 (기준에 따라 % 또는 HWPUNIT) |
| ProtectSize | PIT_UI1 | 크기 보호: on/off |
| LayoutWidth | PIT_I4 | 레이아웃 계산된 폭 |
| LayoutHeight | PIT_I4 | 레이아웃 계산된 높이 |

#### 텍스트 배치

| Item ID | Type | 설명 |
|---------|------|------|
| FlowWithText | PIT_UI1 | 세로 위치 본문 영역 제한: on/off |
| AllowOverlap | PIT_UI1 | 다른 오브젝트와 겹침 허용: on/off |
| TextWrap | PIT_UI1 | 텍스트 배치 방식 |
| TextFlow | PIT_UI1 | 좌/우 텍스트 배치: 0=양쪽, 1=왼쪽, 2=오른쪽, 3=큰쪽 |

**TextWrap 값:**

| 값 | 설명 |
|----|------|
| 0 | bound rect를 따라 |
| 1 | 좌, 우에는 텍스트 배치 안함 |
| 2 | 글과 겹치게 하여 글 뒤로 |
| 3 | 글과 겹치게 하여 글 앞으로 |
| 4 | 오브젝트의 outline을 따라 |
| 5 | 오브젝트 내부의 빈 공간까지 |

#### 바깥 여백

| Item ID | Type | 설명 |
|---------|------|------|
| OutsideMarginLeft | PIT_I4 | 왼쪽 바깥 여백 (HWPUNIT) |
| OutsideMarginRight | PIT_I4 | 오른쪽 바깥 여백 (HWPUNIT) |
| OutsideMarginTop | PIT_I4 | 위 바깥 여백 (HWPUNIT) |
| OutsideMarginBottom | PIT_I4 | 아래 바깥 여백 (HWPUNIT) |

#### 테두리/배경

| Item ID | Type | 설명 |
|---------|------|------|
| BorderFill | PIT_SET | 표에 적용되는 테두리/배경 |
| TableBorderFill | PIT_SET | 표에 적용되는 테두리/배경 |

#### 기타

| Item ID | Type | 설명 |
|---------|------|------|
| Cell | PIT_SET | 셀 속성 (Cell ParameterSet) |
| NumberingType | PIT_UI1 | 번호 범주: 1=그림, 2=표, 3=수식 |
| Lock | PIT_UI1 | 오브젝트 선택 가능: on/off |
| HoldAnchorObj | PIT_UI1 | 쪽 나눔 방지: on/off |
| PageNumber | PIT_UI | 개체가 존재하는 페이지 |

### 3.4 Cell 속성 (Cell ParameterSet)

Cell은 ListProperties를 상속받습니다.

#### Cell 고유 속성

| Item ID | Type | 설명 |
|---------|------|------|
| HasMargin | PIT_UI1 | 자체 셀 여백 적용: on/off |
| Protected | PIT_UI1 | 사용자 편집 막기: 0=off, 1=on |
| Header | PIT_UI1 | 제목 셀 여부: 0=off, 1=on |
| Width | PIT_I4 | 셀의 폭 (HWPUNIT) |
| Height | PIT_I4 | 셀의 높이 (HWPUNIT) |
| Editable | PIT_UI1 | 양식모드에서 편집 가능: 0=off, 1=on |
| Dirty | PIT_UI1 | 수정 상태: 0=초기화, 1=수정됨 |
| CellCtrlData | PIT_SET | 셀 데이터 (CtrlData) |

#### ListProperties 상속 속성

| Item ID | Type | 설명 |
|---------|------|------|
| TextDirection | PIT_UI1 | 글자 방향 |
| LineWrap | PIT_UI1 | 줄 나눔 방식: 0=일반, 1=줄 바꾸지 않음, 2=자간 조정 |
| VertAlign | PIT_UI1 | 세로 정렬: 0=위, 1=가운데, 2=아래 |
| MarginLeft | PIT_I4 | 왼쪽 여백 |
| MarginRight | PIT_I4 | 오른쪽 여백 |
| MarginTop | PIT_I4 | 위 여백 |
| MarginBottom | PIT_I4 | 아래 여백 |

---

## 4. BorderFill - 테두리/배경

BorderFill은 CharShape, ParaShape, Cell 등에서 공통으로 사용됩니다.

### 4.1 테두리 속성

#### 테두리 종류 (BorderType)

| Item ID | Type | 설명 |
|---------|------|------|
| BorderTypeLeft | PIT_UI2 | 왼쪽 테두리 종류 (선 종류) |
| BorderTypeRight | PIT_UI2 | 오른쪽 테두리 종류 |
| BorderTypeTop | PIT_UI2 | 위 테두리 종류 |
| BorderTypeBottom | PIT_UI2 | 아래 테두리 종류 |

#### 테두리 두께 (BorderWidth)

| Item ID | Type | 설명 |
|---------|------|------|
| BorderWidthLeft | PIT_UI1 | 왼쪽 테두리 두께 |
| BorderWidthRight | PIT_UI1 | 오른쪽 테두리 두께 |
| BorderWidthTop | PIT_UI1 | 위 테두리 두께 |
| BorderWidthBottom | PIT_UI1 | 아래 테두리 두께 |

#### 테두리 색상 (BorderColor, COLORREF 0x00BBGGRR)

| Item ID | Type | 설명 |
|---------|------|------|
| BorderColorLeft | PIT_UI4 | 왼쪽 테두리 색상 |
| BorderColorRight | PIT_UI4 | 오른쪽 테두리 색상 |
| BorderColorTop | PIT_UI4 | 위 테두리 색상 |
| BorderColorBottom | PIT_UI4 | 아래 테두리 색상 |

#### 대각선

| Item ID | Type | 설명 |
|---------|------|------|
| SlashFlag | PIT_UI2 | 슬래쉬 대각선 플래그 (bit 0=하단, 1=중앙, 2=상단) |
| BackSlashFlag | PIT_UI2 | 백슬래쉬 대각선 플래그 |
| DiagonalType | PIT_UI2 | 대각선 종류 |
| DiagonalWidth | PIT_UI1 | 대각선 두께 |
| DiagonalColor | PIT_UI4 | 대각선 색상 |
| CrookedSlashFlag | PIT_UI2 | 꺾인 대각선 플래그 |
| CounterSlashFlag | PIT_UI1 | 슬래쉬 역방향 플래그: 0=순방향, 1=역방향 |
| CounterBackSlashFlag | PIT_UI1 | 백슬래쉬 역방향 플래그 |
| CenterLineFlag | PIT_UI1 | 중심선: 0=없음, 1=있음 |

#### 효과

| Item ID | Type | 설명 |
|---------|------|------|
| BorderFill3D | PIT_UI1 | 3차원 효과: 0=off, 1=on |
| Shadow | PIT_UI1 | 그림자 효과: 0=off, 1=on |

### 4.2 배경 속성 (DrawFillAttr)

FillAttr 아이템으로 배경 채우기 속성에 접근합니다.

#### 배경 유형

| Item ID | Type | 설명 |
|---------|------|------|
| Type | PIT_UI | 배경 유형: 0=없음, 1=면색/무늬색, 2=그림, 3=그러데이션 |

#### 면색/무늬 (Type=1)

| Item ID | Type | 설명 |
|---------|------|------|
| WinBrushFaceColor | PIT_UI | 면 색 (RGB 0x00BBGGRR) |
| WinBrushHatchColor | PIT_UI | 무늬 색 |
| WinBrushFaceStyle | PIT_I1 | 무늬 스타일 |
| WinBrushAlpha | PIT_UI | 투명도 |
| WindowsBrush | PIT_UI1 | 면/무늬 브러시 여부 |

**WinBrushFaceStyle 값:**

| 값 | 설명 |
|----|------|
| 0 | horizontal lines |
| 1 | vertical lines |
| 2 | forward diagonal |
| 3 | backward diagonal |
| 4 | cross |
| 5 | diagonal cross |
| 6 | none |

#### 그림 (Type=2)

| Item ID | Type | 설명 |
|---------|------|------|
| FileName | PIT_BSTR | 그림 파일 경로 |
| Embedded | PIT_UI1 | 문서에 삽입(TRUE) / 파일로 연결(FALSE) |
| PicEffect | PIT_UI1 | 그림 효과: 0=원본, 1=그레이스케일, 2=흑백 |
| Brightness | PIT_I1 | 명도 (-100 ~ 100) |
| Contrast | PIT_I1 | 밝기 (-100 ~ 100) |
| Reverse | PIT_UI1 | 반전 유무 |
| DrawFillImageType | PIT_I | 배경 채우기 방식 |
| ImageAlpha | PIT_UI1 | 투명도 |
| ImageBrush | PIT_UI1 | 그림 브러시 여부 |

**DrawFillImageType 값:**

| 값 | 설명 |
|----|------|
| 0 | 바둑판식으로 |
| 1 | 가로/위만 바둑판식 |
| 2 | 가로/아래만 바둑판식 |
| 3 | 세로/왼쪽만 바둑판식 |
| 4 | 세로/오른쪽만 바둑판식 |
| 5 | 크기에 맞추어 |
| 6 | 가운데로 |
| 7 | 가운데 위로 |
| 8 | 가운데 아래로 |
| 9 | 왼쪽 가운데로 |
| 10 | 왼쪽 위로 |
| 11 | 왼쪽 아래로 |
| 12 | 오른쪽 가운데로 |
| 13 | 오른쪽 위로 |
| 14 | 오른쪽 아래로 |

#### 그러데이션 (Type=3)

| Item ID | Type | 설명 |
|---------|------|------|
| GradationType | PIT_I | 형태: 1=줄무늬형, 2=원형, 3=원뿔형, 4=사각형 |
| GradationAngle | PIT_I | 기울임(시작각) |
| GradationCenterX | PIT_I | 중심 X 좌표 |
| GradationCenterY | PIT_I | 중심 Y 좌표 |
| GradationStep | PIT_I | 번짐 정도 (0..100) |
| GradationColorNum | PIT_I | 색 수 |
| GradationColor | PIT_ARRAY | 색깔 배열 |
| GradationIndexPos | PIT_ARRAY | 다음 색깔과의 거리 |
| GradationStepCenter | PIT_UI1 | 번짐 중심 (0..100) |
| GradationAlpha | PIT_UI1 | 투명도 |
| GradationBrush | PIT_UI1 | 그러데이션 브러시 여부 |

---

## 5. 셀 테두리/배경 (CellBorderFill)

CellBorderFill은 BorderFillExt를 상속받으며, BorderFillExt는 BorderFill을 상속받습니다.

### 5.1 CellBorderFill 고유 속성

| Item ID | Type | 설명 |
|---------|------|------|
| ApplyTo | PIT_UI1 | 적용 대상: 0=선택된 셀, 1=전체 셀, 2=여러 셀에 걸쳐 |
| NoNeighborCell | PIT_UI1 | 주변 셀에 선 모양 미적용: 1=미적용 |
| TableBorderFill | PIT_SET | 표 테두리/배경 |
| AllCellsBorderFill | PIT_SET | 전체 셀의 테두리/배경 |
| SelCellsBorderFill | PIT_SET | 선택된 셀의 테두리/배경 |

### 5.2 BorderFillExt 상속 속성 (중앙선)

| Item ID | Type | 설명 |
|---------|------|------|
| TypeHorz | PIT_UI2 | 가로 중앙선 종류 |
| TypeVert | PIT_UI2 | 세로 중앙선 종류 |
| WidthHorz | PIT_UI1 | 가로 중앙선 두께 |
| WidthVert | PIT_UI1 | 세로 중앙선 두께 |
| ColorHorz | PIT_UI4 | 가로 중앙선 색상 (0x00BBGGRR) |
| ColorVert | PIT_UI4 | 세로 중앙선 색상 |

---

## 6. 사용 예시

### 6.1 글자 모양 조회

```python
import win32com.client as win32

hwp = win32.Dispatch('HWPFrame.HwpObject')
hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')
hwp.Open("C:\\test.hwp")

# 현재 Selection의 글자 모양 조회
char_shape = hwp.CharShape

# 폰트 정보
font_name = char_shape.Item("FaceNameHangul")
font_size_hwpunit = char_shape.Item("Height")
font_size_pt = font_size_hwpunit / 100  # pt 변환

# 스타일 정보
is_bold = char_shape.Item("Bold")
is_italic = char_shape.Item("Italic")

# 색상 정보 (COLORREF: 0x00BBGGRR)
text_color = char_shape.Item("TextColor")
red = text_color & 0xFF
green = (text_color >> 8) & 0xFF
blue = (text_color >> 16) & 0xFF

print(f"폰트: {font_name}, 크기: {font_size_pt}pt")
print(f"굵게: {is_bold}, 기울임: {is_italic}")
print(f"색상: RGB({red}, {green}, {blue})")
```

### 6.2 문단 모양 조회

```python
# 현재 Selection의 문단 모양 조회
para_shape = hwp.ParaShape

# 정렬
align_type = para_shape.Item("AlignType")
align_names = {0: "양쪽", 1: "왼쪽", 2: "오른쪽", 3: "가운데", 4: "배분", 5: "나눔"}
print(f"정렬: {align_names.get(align_type, '알 수 없음')}")

# 줄 간격
line_spacing_type = para_shape.Item("LineSpacingType")
line_spacing = para_shape.Item("LineSpacing")

if line_spacing_type == 0:  # 글자에 따라
    print(f"줄 간격: {line_spacing}%")
else:  # 고정값 또는 여백만
    print(f"줄 간격: {line_spacing / 100}pt")

# 여백
left_margin = para_shape.Item("LeftMargin")
indentation = para_shape.Item("Indentation")
print(f"왼쪽 여백: {left_margin / 100}pt, 들여쓰기: {indentation / 100}pt")
```

### 6.3 셀 모양 조회

```python
# 표 내부에 캐럿이 위치해야 함
try:
    cell_shape = hwp.CellShape

    # 표 속성
    cell_spacing = cell_shape.Item("CellSpacing")
    page_break = cell_shape.Item("PageBreak")

    # 셀 속성
    cell = cell_shape.Item("Cell")
    cell_width = cell.Item("Width")
    cell_height = cell.Item("Height")
    vert_align = cell.Item("VertAlign")

    vert_align_names = {0: "위", 1: "가운데", 2: "아래"}

    print(f"셀 크기: {cell_width / 100}pt x {cell_height / 100}pt")
    print(f"세로 정렬: {vert_align_names.get(vert_align, '알 수 없음')}")

except Exception as e:
    print("표 내부에 캐럿이 위치하지 않습니다.")
```

---

## 7. 단위 변환

### 7.1 HWPUNIT 기본 변환

| 단위 | HWPUNIT 값 | 변환식 |
|------|-----------|--------|
| 1 pt (포인트) | 100 | `hwpunit / 100 = pt` |
| 1 inch (인치) | 7,200 | `hwpunit / 7200 = inch` |
| 1 cm | 약 2,834.6 | `hwpunit / 2834.6 = cm` |
| 1 mm | 약 283.46 | `hwpunit / 283.46 = mm` |

### 7.2 색상 변환 (COLORREF)

```python
# COLORREF (0x00BBGGRR) -> RGB
def colorref_to_rgb(colorref):
    red = colorref & 0xFF
    green = (colorref >> 8) & 0xFF
    blue = (colorref >> 16) & 0xFF
    return (red, green, blue)

# RGB -> COLORREF
def rgb_to_colorref(red, green, blue):
    return red | (green << 8) | (blue << 16)

# 예시
colorref = 0x00FF8040  # BGR
r, g, b = colorref_to_rgb(colorref)
print(f"RGB: ({r}, {g}, {b})")  # RGB: (64, 128, 255)
```

---

## 8. 참고 사항

1. **속성 존재 여부**: Selection 내에서 특정 항목이 서로 다른 속성을 가지면, 해당 아이템 자체가 존재하지 않습니다.

2. **에러 처리**: CellShape는 표 내부에 캐럿이 위치하지 않으면 에러가 발생합니다.

3. **속성 설정**: CharShape, ParaShape에 property set을 수행하면 아이템이 존재하는 항목에 대해서만 속성이 설정됩니다.

4. **ParameterSet 형식**: 각 속성의 상세 형식은 다음을 참조하세요:
   - CharShape: ParameterSet/CharShape
   - ParaShape: ParameterSet/ParaShape
   - CellShape: ParameterSet/Table + ParameterSet/Cell
   - BorderFill: ParameterSet/BorderFill
