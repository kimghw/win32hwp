# ParameterSetTable_2504_part15

---

## 115) Table : 표

Table은 ShapeObject로부터 계승받았으므로 위 표에 정리된 Table의 아이템들 이외에 ShapeObject의 아이템들을 사용할 수 있다.

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| PageBreak | PIT_UI1 | | 표가 페이지 경계에 걸렸을 때의 처리 방식<br>0 = 나누지 않는다.<br>1 = 테이블은 나누지만 셀은 나누지 않는다.<br>2 = 셀 내의 텍스트도 나눈다. |
| RepeatHeader | PIT_UI1 | | 제목 행을 반복할지 여부. (on / off) |
| CellSpacing | PIT_UI4 | | 셀 간격(HTML의 셀 간격과 동일 의미. HWPUNIT) |
| CellMarginLeft | PIT_I4 | | 기본 셀 안쪽 여백(왼쪽) |
| CellMarginRight | PIT_I4 | | 기본 셀 안쪽 여백(오른쪽) |
| CellMarginTop | PIT_I4 | | 기본 셀 안쪽 여백(위쪽) |
| CellMarginBottom | PIT_I4 | | 기본 셀 안쪽 여백(아래쪽) |
| BorderFill | PIT_SET | BorderFill | 표에 적용되는 테두리/배경 |
| TableCharInfo | PIT_SET | TableChartInfo | 표와 연결된 차트 정보 - 차트 미완성 |
| TableBorderFill | PIT_SET | BorderFill | 표에 적용되는 테두리/배경 |
| Cell | PIT_SET | Cell | 셀 속성 |
| TreatAsChar | PIT_UI1 | | 글자처럼 취급 on / off |
| AffectsLine | PIT_UI1 | | 줄 간격에 영향을 줄지 여부 on / off<br>TREAT_AS_CHAR = TRUE일 때만 사용됨 |
| VertRelTo | PIT_UI1 | | 세로 위치의 기준.<br>TREAT_AS_CHAR = FALSE일 때만 사용됨<br>0 = 종이<br>1 = 본문 영역<br>2 = 문단 |
| VertAlign | PIT_UI1 | | VERT_REL_TO에 대한 상대적인 배열 방식.<br>VERT_REL_TO의 값에 따라 가능한 범위가 제한된다.<br>0 = 위<br>1 = 가운데<br>2 = 아래<br>3 = 안쪽<br>4 = 바깥쪽 |
| VertOffset | PIT_UI4 | | VERT_REL_TO와 VERT_ALIGN을 기준점으로 한 상대적인 오프셋 값. HWPUNIT 단위. |
| HorzRelTo | PIT_UI1 | | 가로 위치의 기준.<br>TREAT_AS_CHAR = FALSE일 때만 사용됨<br>0 = 종이<br>1 = 본문 영역<br>2 = 단<br>3 = 문단 |
| HorzAlign | PIT_UI1 | | HORZ_REL_TO에 대한 상대적인 배열 방식.<br>0 = 왼쪽<br>1 = 가운데<br>2 = 오른쪽<br>3 = 안쪽<br>4 = 바깥쪽 |
| HorzOffset | PIT_I4 | | HORZ_REL_TO와 HORZ_ALIGN을 기준점으로 한 상대적인 오프셋 값. HWPUNIT 단위. |
| FlowWithText | PIT_UI1 | | 오브젝트의 세로 위치를 본문 영역으로 제한할지 여부 on / off<br>VERT_REL_TO = PARA일 때만 사용됨 |
| AllowOverlap | PIT_UI1 | | 다른 오브젝트와 겹치는 것을 허용할지 여부 on / off<br>TREAT_AS_CHAR = FALSE일 때만 사용됨<br>FLOW_WITH_TEXT = TRUE이면 언제나 FALSE로 간주한다. |
| WidthRelTo | PIT_UI1 | | 오브젝트 폭의 기준<br>0 = 종이<br>1 = 본문 영역<br>2 = 단<br>3 = 문단(VertRelTo = 문단일 때만 가능)<br>4 = 절대값 |
| Width | PIT_I4 | | 오브젝트 폭의 값<br>WIDTH_REL_TO의 값에 따라 다음과 같은 다른 단위를 뜻한다.<br>0 = 종이의 몇 %<br>1 = 본문 영역의 몇 %<br>2 = 단의 몇 %<br>3 = 문단의 몇 %<br>4 = 절대값 HWPUNIT |
| HeightRelTo | PIT_UI1 | | 오브젝트 높이의 기준<br>0 = 종이<br>1 = 본문 영역<br>2 = 절대값 |
| Height | PIT_I4 | | 오브젝트의 높이 값<br>HEIGHT_REL_TO의 값에 따라 다음과 같은 다른 단위를 뜻한다.<br>0 = 종이의 몇 %<br>1 = 본문 영역의 몇 %<br>2 = 절대값 HWPUNIT |
| ProtectSize | PIT_UI1 | | 크기 보호 on / off |
| TextWrap | PIT_UI1 | | 오브젝트 주위를 텍스트가 어떻게 흘러갈지 지정하는 옵션.<br>TREAT_AS_CHAR = FALSE일 때만 사용됨<br>0 = bound rect를 따라<br>1 = 좌, 우에는 텍스트를 배치하지 않음<br>2 = 글과 겹치게 하여 글 뒤로<br>3 = 글과 겹치게 하여 글 앞으로<br>4 = 오브젝트의 outline을 따라<br>5 = 오브젝트 내부의 빈 공간까지 |
| TextFlow | PIT_UI1 | | 오브젝트의 좌/우 어느쪽에 글을 배치할지 지정하는 옵션.<br>TEXT_WRAP가 SQUARE, TIGHT, THROUGH일 때만 사용된다.<br>0 = 양쪽<br>1 = 왼쪽<br>2 = 오른쪽<br>3 = 큰쪽 |
| OutsideMarginLeft | PIT_I4 | | 오브젝트의 바깥 여백. HWPUNIT 단위 |
| OutsideMarginRight | PIT_I4 | | 오브젝트의 바깥 여백. HWPUNIT 단위 |
| OutsideMarginTop | PIT_I4 | | 오브젝트의 바깥 여백. HWPUNIT 단위 |
| OutsideMarginBottom | PIT_I4 | | 오브젝트의 바깥 여백. HWPUNIT 단위 |
| NumberingType | PIT_UI1 | | 이 개체가 속하는 번호 범주<br>1 = 그림<br>2 = 표<br>3 = 수식 |
| LayoutWidth | PIT_I4 | | 오브젝트가 페이지에 배열될 때 계산되는 폭의 값<br>글상자등이 늘어나면 늘어난 값을 계산해서 가진다. 단위는 HWPITID_SO_WIDTH와 같다. |
| LayoutHeight | PIT_I4 | | 오브젝트가 페이지에 배열될 때 계산되는 높이 값<br>글상자등이 늘어나면 늘어난 값을 계산해서 가진다.<br>단위는 HWPITID_SO_HEIGHT와 같다. |
| Lock | PIT_UI1 | | 오브젝트 선택 가능 on / off |
| HoldAnchorObj | PIT_UI1 | | 쪽 나눔방지 on/off |
| PageNumber | PIT_UI | | 개체가 존재 하는 페이지 |
| AdjustSelection | PIT_UI1 | | 개체 Selection 상태 TRUE/FASLE |
| AdjustTextbox | PIT_UI1 | | 글상자로 TRUE/FASLE |
| AdjustPrevObjAttr | PIT_UI1 | | 앞개체 속성 따라가기 TRUE/FASLE |

### Example : Table ParameterSet 설정하기

*Visual Basic*

```vb
Dim TableSet As HwpParameterSet
Set TableSet = HwpCtrl.CreateSet("Table")
TableSet.SetItem "PageBreak", 1          ' <- Table의 아이템
TableSet.SetItem "TreatAsChar", True     ' <- ShapeObject의 아이템
```

---

## 116) TableCreation : 표 생성

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Rows | PIT_UI2 | | 행 수 (생략하면 5) |
| Cols | PIT_UI2 | | 칼럼 수 (생략하면 5) |
| RowHeight | PIT_ARRAY | PIT_I4 | 행의 디폴트 높이 (PIT_I4) |
| ColWidth | PIT_ARRAY | PIT_I4 | 칼럼의 디폴트 폭 (PIT_I4) |
| CellInfo | PIT_ARRAY | PIT_I4 | 정보가 없는 셀은 디폴트값을 따라가므로 모든 셀에 대해 정보를 줄 필요는 없다. |
| WidthType | PIT_UI1 | | 너비 |
| HeightType | PIT_UI1 | | 높이 |
| WidthValue | PIT_I | | 너비 값 |
| HeightValue | PIT_I | | 높이 값 |
| TableTemplateValue | PIT_UI1 | | 표 마당 적용 여부 |
| TableProperties | PIT_SET | Table | 초기 표 속성 |
| TableTemplate | PIT_SET | TableTemplate | 표마당 적용 속성 |
| TableDrawProperties | PIT_SET | TableDrawPen | 마우스로 선을 그릴 때 속성 |
| TextSelect | PIT_I | | 텍스트 선택 여부 |

---

## 117) TableDeleteLine : 표의 줄/칸 삭제

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Type | PIT_UI1 | | 0 = 줄, 1 = 칸 |

---

## 118) TableDrawPen : 마우스로 테이블을 그릴 때 쓰이는 펜

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Style | PIT_UI2 | | Table을 그리는 연필(펜)의 선 모양 |
| Width | PIT_UI1 | | Table을 그리는 연필(펜)의 선 굵기 |
| Color | PIT_UI4 | | Table을 그리는 연필(펜)의 선 색깔<br>RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR) |

---

## 119) TableInsertLine : 표의 줄/칸 삽입

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Side | PIT_UI1 | | 방향 |
| Count | PIT_UI1 | | 개수 |

---

## 120) TableSplitCell : 셀 나누기

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Cols | PIT_UI2 | | 칸 수 |
| Rows | PIT_UI2 | | 줄 수 |
| DistributeHeight | PIT_UI1 | | 줄 높이를 같게 |
| Merge | PIT_UI1 | | 나누기 전에 합치기 |
| Mode2 | PIT_UI1 | | 셀 나누기 모드 2, 셀 나누기를 할 때, adjust를 생략하고 셀이 어긋나는 것을 방지한다. |

---

## 121) TableStrToTbl : 문자열을 표로

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| DelimiterType | PIT_UI1 | | 분리 문자(탭, 쉼표, 공백) |
| UserDefine | PIT_BSTR | | 사용자 정의 필드 구분 기호 |
| AutoOrDefine | PIT_UI1 | | 자동으로 할 것인지 분리 문자를 지정 할 것인지를 결정 |
| KeepSeperator | PIT_UI1 | | 선택 사항 (구분자 유지) |
| DelimiterEtc | PIT_BSTR | | 기타 문자 필드 구분 기호 |
| TableCreation | PIT_SET | TableCreation | 만들 표 속성 |

---

*Page 141-150*
