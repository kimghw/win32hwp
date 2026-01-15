# ParameterSetTable_2504_part13

## 121

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| EquationNumber | PIT_UI2 | | 수식 시작 번호<br>0 = 앞 구역에 이어, n = 새 번호로 시작 |
| WongojiFormat | PIT_UI1 | | 원고지 방식의 포맷팅. CHAR_GRID가 지정되어야 함. |
| MemoShape | PIT_SET | MemoShape | 메모 모양 |
| TextVerticalWidthHead | PIT_I | | 머리말/꼬리말 세로쓰기 여부 |
| ApplyTo | PIT_UI1 | | 적용범위<br>0 = 선택된 구역<br>1 = 선택된 문자열<br>2 = 현재 구역<br>3 = 문서전체<br>4 = 새 구역 : 현재 위치부터 새로 |
| ApplyClass | PIT_UI1 | | 적용범위의 분류 (대화상자를 호출할 경우 사용)<br>0x01 = 선택된 구역<br>0x02 = 선택된 문자열<br>0x04 = 현재 구역<br>0x08 = 문서전체<br>0x10 = 새 구역 : 현재 위치부터 새로 |
| ApplyToPageBorderFill | PIT_UI1 | | 채울 영역 분류 (PageBorder 액션에서 사용)<br>0 = 종이, 1 = 쪽, 2 = 테두리 |
| LineNumberRestart | PIT_UI1 | | 줄 번호 형식<br>0 = 기본, 1 = 쪽마다, 2 = 구역마다, 3 = 이어서 |
| LineNumberCountBy | PIT_UI2 | | 줄 번호 표시 간격 |
| LineNumberDistance | PIT_UI | | 본문과의 줄 번호 간격 (HWPUNIT) |
| LineNumberStart | PIT_UI2 | | 줄 번호 시작 번호 |
| ShowLineNumbers | PIT_UI1 | | 줄 번호 표시 여부 (On/Off) |
| SectionApplyString | PIT_SET | PageBorderFill | 적용 범위 서브셋, (한/글 2022부터 지원) |

---

## 122

### 103) SectionApply : 적용할 구역 설정

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| ApplyTo | PIT_UI1 | | 적용범위 분류(ApplyClass)에서 하나의 값 |
| String | PIT_ARRAY | PIT_BSTR | ApplyTo를 문자열로 변환한 값의 배열 |
| Index | PIT_I | | ApplyTo를 변환한 ComboBox의 Index |
| ConvAplly2Index | PIT_I | | ApplyTo값을 ComboBox의 Index로 변환할지 여부<br>FALSE이면 IndexToApply로 변환이 이루어진다. (반대변환) |
| Category | PIT_I | | 적용범위 분류(ApplyClass)를 사용자가 직접 설정할 때 사용. 아이템이 없으면 글이 현재 상태에 맞춰 적용범위 분류(ApplyClass)를 설정한다. (일반적으로 설정하지 않고 사용) |

상기 파라메터셋은 실제로 글에서 사용되는 파라메터셋이 아닌 일종의 도구용 파라메터셋이다. 즉, 글의 기능을 위해 존재하지 않는다. 위 파라메터셋은 GetSectionApplyString과 GetSectionApplyTo 액션에서만 사용된다. 이 두 액션은 글의 상태에 따라 바뀌는 ApplyClass와 ApplyTo의 값을 쉽게 얻기 위해 고안되었다.

GetSectionApplyString 액션은 현재 글의 상태에 따라 구역의 적용범위를 얻기 위해 고안된 유틸리티 액션이다. 사용방법은 다음과 같다. 액션을 생성한 후 ApplyTo에 적용할 범위 flag를 넣어준다. 만약 적당한 ApplyTo flag값이 들어왔다면, "쪽/테두리 배경" 대화상자에서 볼 수 있는 "적용범위" ComboBox의 선택된 값을 Index 아이템을 통해서 얻을 수 있다.

그림. 적용범위 ComboBox의 모습. 위 그림에서 ApplyTo에 해당하는 것은 "현재 구역", "문서 전체", "새 구역으로" 각각의 flag값이며, 이 ApplyTo값이 모두 조합된 값이 ApplyClass가 된다. ApplyClass에 해당되는 값은 문자열로도 얻어올 수 있으며, 얻어오는 장소가 바로 String 아이템이다. 위 그림의 예를 들어 각 아이템에 할당되는 값을 나타내면 다음과 같다.

---

## 123

```
ApplyTo : 0x08
ApplyTo : 0x1C (0x04|0x08|0x10)
Index : 3
String : "현재 구역" "문서 전체" "새 구역으로"
```

※ 위 값 중 Index의 값이 1이 아닌 3인 이유는 Index의 값으로 넘어오는 값은 콤보박스를 기준으로 하지 않고 모든 ApplyTo값의 순서에 따른 Index값을 기준으로 하기 때문이다.

### Example : GetSectionApplyString 액션 사용하기

**C++**

```cpp
CDHwpAction myAction = m_HwpCtrl.CreateAction("GetSectionApplyString");
CDHwpParameterSet mySet = myAction.CreateSet();
// myAction.GetDefault();                      // 해당 Action은 GetDefault();를 생략해도 상관없다.

mySet.SetItem("ApplyTo", CComVariant(0x08));
myAction.Execute(mySet);

int nIdx = (int)mySet.Item("Index");           // 얻어온 Index값을 저장. 이후 String값도 저장한다.
...
```

GetSectionApplyTo 액션은 ApplyTo와 Index 아이템 값의 상호변환을 위한 Action이다.
ConvAplly2Index를 TRUE로 설정하면 ApplyTo -> Index 변환이, ConvAplly2Index를 FALSE로 설정하면 Index->ApplyTo 변환이 이루어진다.

### Example : GetSectionApplyTo 액션 사용하기

**C++**

```cpp
CDHwpAction myAction = m_HwpCtrl.CreateAction("GetSectionApplyTo");
CDHwpParameterSet mySet = myAction.CreateSet();
// myAction.GetDefault();                      // 해당 Action은 GetDefault();를 생략해도 상관없다.

mySet.SetItem("ApplyTo", 0x08);
mySet.SetItem("ConvAplly2Index", CComVariant(TRUE));
myAction.Execute(mySet);

int nIdx = (int)mySet.Item("Index");           // 얻어온 Index값을 저장
```

---

## 124

```
...
```

※ 현재 위 액션은 SecDef 파라메터셋(구역정보)에 대한 ApplyTo, ApplyClass만을 구해온다. 다른 정보영역의 ApplyTo, ApplyClass의 값을 얻어오는 유틸리티 액션의 구현에 대해서는 아직 정해진 바가 없다.

---

## 125

### 104) ShapeCopyPaste : 모양 복사

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Type | PIT_I | | 모양 복사 종류<br>0 = 글자 모양 복사<br>1 = 문단 모양 복사<br>2 = 글자와 문단 모양 두개 복사<br>3 = 글자 스타일 복사<br>4 = 문단 스타일 복사 |
| CellAttr | PIT_UI1 | | 셀 모양 복사 |
| CellBorder | PIT_UI1 | | 선 모양 복사 |
| CellFill | PIT_UI1 | | 셀 배경 복사 |
| TypeBodyAndCellOnly | PIT_I | | 본문과 셀 모양 둘 다 복사 or 셀 모양만 복사 |

---

## 126

### 105) ShapeObject : 그리기 개체의 공통 속성 (도형, 글상자, 표, 그림 등)

ShapeObject는 글의 컨트롤 중 편집영역을 자연스럽게 이동할 수 있는 개체를 말한다. 일반적으로 이런 개체들은 선택했을 때 8개의 사각형 포인터를 가지는 특징을 가진다.

글에서 ShapeObject에 해당하는 컨트롤에는 표, 수식, 그림, OLE, 그리기 개체 등이 있다. 이 중 표와 수식은 ShapeObject를 확장해 자신만의 ParametetSet을 사용한다. ("Table", "EqEdit")

이 외 다른 개체들(그림, OLE, 그리기 개체)은 따로 ParameterSet을 사용하지 않고, ShapeObject를 공유해서 사용한다. 아래의 Table은 ShapeObject를 사용하는 컨트롤이 공통적으로 가지는 ITEM을 나타낸 표이다. 그러므로 표("Table")와 수식("EqEdit")은 자신의 ITEM 외에 다음의 ITEM을 추가로 가진다.

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| TreatAsChar | PIT_UI1 | | 글자처럼 취급 on / off |
| AffectsLine | PIT_UI1 | | 줄 간격에 영향을 줄지 여부 on / off (TreatAsChar가 TRUE일 경우에만 사용된다) |
| VertRelTo | PIT_UI1 | | 세로 위치의 기준.<br>0 = 종이 영역(Paper)<br>1 = 쪽 영역(Page)<br>2 = 문단 영역(Paragraph)<br>(TreatAsChar가 FALSE일 경우에만 사용된다) |
| VertAlign | PIT_UI1 | | VertRelTo값에 따른 상대적인 정렬 기준.<br>VertRelTo값이 2(문단영역)일 경우 0 값만 사용할 수 있다.<br>0 = 위(Top)<br>1 = 가운데(Center)<br>2 = 아래(Bottom) |
| VertOffset | PIT_I4 | | VertRelTo와 VertAlign을 기준으로 한 Y축 위치 오프셋 값. HWPUNIT 단위. |
| HorzRelTo | PIT_UI1 | | 가로 위치의 기준.<br>0 = 종이 영역(Paper)<br>1 = 쪽 영역(Page)<br>2 = 다단 영역(Column)<br>3 = 문단 영역(Paragraph)<br>(TreatAsChar가 FALSE일 경우에만 사용된다) |
| HorzAlign | PIT_UI1 | | HorzRelTo값에 따른 상대적인 정렬 기준.<br>HorzRelTo값이 3(문단영역)일 경우 0~2 사이의 값만 사용할 수 있다. |

---

## 127

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| | | | 0 = 왼쪽(Left)<br>1 = 가운데(Center)<br>2 = 오른쪽(Right)<br>3 = 안쪽(Inside)<br>4 = 바깥쪽(Outside) |
| HorzOffset | PIT_I4 | | HorzRelTo와 HorzAlign을 기준으로 한 X축 위치 오프셋 값. HWPUNIT 단위. |
| FlowWithText | PIT_UI1 | | 그리기 개체의 세로 위치를 쪽 영역 안쪽으로 제한할지 여부 on / off<br>VertRelTo값이 2(문단영역)일 경우에만 의미가 있다. |
| AllowOverlap | PIT_UI1 | | 다른 개체와 겹치는 것을 허용할지 여부 on / off<br>TreatAsChar가 FALSE일 때만 의미가 있으며, FlowWithText가 TRUE이면 AllowOverlap은 항상 FALSE로 간주한다. |
| WidthRelTo | PIT_UI1 | | 개체 너비 기준.<br>0 = 종이(Paper)<br>1 = 쪽(Page)<br>2 = 다단(Column)<br>3 = 문단(Paragraph)<br>4 = 고정 값(Absolute) |
| Width | PIT_I4 | | 개체 너비 값.<br>WidthRelTo에 따라 값의 의미 및 단위가 달라진다.<br><table><tr><th>WidthRelTo</th><th>의미 및 단위</th></tr><tr><td>0</td><td>종이 너비의 몇 %</td></tr><tr><td>1</td><td>쪽 너비의 몇 %</td></tr><tr><td>2</td><td>단 너비의 몇 %</td></tr><tr><td>3</td><td>문단 너비의 몇 %</td></tr><tr><td>4</td><td>고정 값(단위 HWPUNIT)</td></tr></table> |
| HeightRelTo | PIT_UI1 | | 개체 높이 기준.<br>0 = 종이(Paper)<br>1 = 쪽(Page)<br>2 = 고정 값(Absolute) |
| Height | PIT_I4 | | 개체 높이 값.<br>HeightRelTo에 다라 값의 의미 및 단위가 달라진다. |

---

## 128

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| | | | <table><tr><th>HeightRelTo</th><th>의미 및 단위</th></tr><tr><td>0</td><td>종이 높이의 몇 %</td></tr><tr><td>1</td><td>쪽 높이의 몇 %</td></tr><tr><td>2</td><td>고정 값(단위 HWPUNIT)</td></tr></table> |
| ProtectSize | PIT_UI1 | | 크기 보호 on / off |
| TextWrap | PIT_UI1 | | 그리기 개체와 본문 사이의 배치 방법.<br>0 = 어울림(Square)<br>1 = 자리 차지(Top & Bottom)<br>2 = 글 뒤로(Behind Text)<br>3 = 글 앞으로(In front of Text)<br>4 = 경계를 명확히 지킴(Tight) - 현재 사용안함<br>5 = 경계를 통과함(Through) - 현재 사용안함<br>(TreatAsChar가 FALSE일 경우에만 사용된다) |
| TextFlow | PIT_UI1 | | 그리기 개체의 좌/우 어느 쪽에 글을 배치할지 지정하는 옵션. TextWrap의 값이 0일 때만 유효하다.<br>0 = 양쪽 모두(Both)<br>1 = 왼쪽만(Left Only)<br>2 = 오른쪽만(Right Only)<br>3 = 왼쪽과 오른쪽 중 넓은 쪽(Largest Only) |
| OutsideMarginLeft | PIT_I4 | | 개체의 바깥 여백. (왼쪽) HWPUNIT 단위 |
| OutsideMarginRight | PIT_I4 | | 개체의 바깥 여백. (오른쪽) HWPUNIT 단위 |
| OutsideMarginTop | PIT_I4 | | 개체의 바깥 여백. (위) HWPUNIT 단위 |
| OutsideMarginBottom | PIT_I4 | | 개체의 바깥 여백. (아래) HWPUNIT 단위 |
| NumberingType | PIT_UI1 | | 이 개체가 속하는 번호 범주.<br>0 = 없음, 1 = 그림, 2 = 표, 3 = 수식 |
| LayoutWidth | PIT_I4 | | 개체가 페이지에 배열될 때 계산되는 폭의 값 |
| LayoutHeight | PIT_I4 | | 개체가 페이지에 배열될 때 계산되는 높이 값 |
| Lock | PIT_UI1 | | 개체 보호하기 on / off |
| HoldAnchorObj | PIT_UI1 | | 쪽 나눔 방지 on / off |
| PageNumber | PIT_UI | | 개체가 존재 하는 페이지 |
| AdjustSelection | PIT_UI1 | | 개체 Selection 상태 TRUE/FASLE |
| AdjustTextBox | PIT_UI1 | | 글상자로 TRUE/FASLE |
| AdjustPrevObjAttr | PIT_UI1 | | 앞개체 속성 따라가기 TRUE/FASLE |

---

## 129

ShapeObject는 위의 공통 ITEM 외에도 다음의 ITEM을 선택적으로 가질 수 있다.

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| ShapeDrawLayOut | PIT_SET | DrawLayOut | 그리기 개체의 Layout |
| ShapeDrawLineAttr | PIT_SET | DrawLineAttr | 그리기 개체의 Line 속성 |
| ShapeDrawFillAttr | PIT_SET | DrawFillAttr | 그리기 개체의 Fill 속성 |
| ShapeDrawImageAttr | PIT_SET | DrawImageAttr | 그림 개체 속성 |
| ShapeDrawRectType | PIT_SET | DrawRectType | 사각형 그리기 개체 유형 |
| ShapeDrawArcType | PIT_SET | DrawArcType | 호 그리기 개체 유형 |
| ShapeDrawResize | PIT_SET | DrawResize | 그리기 개체 리사이징 |
| ShapeDrawRotate | PIT_SET | DrawRotate | 그리기 개체 회전 |
| ShapeDrawEditDetail | PIT_SET | DrawEditDetail | 그리기 개체 EditDetail |
| ShapeDrawImageScissoring | PIT_SET | DrawImageScissoring | 그림 개체 자르기 |
| ShapeDrawScAction | PIT_SET | DrawScAction | 그리기 개체 회전 |
| ShapeDrawCtrlHyperlink | PIT_SET | DrawCtrlHyperlink | 그리기 개체 하이퍼링크 |
| ShapeDrawCoordInfo | PIT_SET | DrawCoordInfo | 그리기 개체 좌표정보 |
| ShapeDrawShear | PIT_SET | DrawShear | 그리기 개체 기울이기 |
| ShapeDrawTextart | PIT_SET | DrawTextart | 글맵시 |
| ShapeDrawShadow | PIT_SET | DrawShadow | 그림자 |
| ShapeTableCell | PIT_SET | Cell | 셀 정보 |
| ShapeListProperties | PIT_SET | ListProperties | 서브 list 속성 |
| ShapeCaption | PIT_SET | Caption | 캡션 |
| ShapeFormGeneral | PIT_SET | FormGeneral | 양식개체 일반 |
| ShapeFormCommonAttr | PIT_SET | FormCommonAttr | 양식개체 공통속성 |
| ShapeFormCharshapeattr | PIT_SET | FormCharshapeattr | 양식개체 글자모양 속성 |
| ShapeFormButtonAttr | PIT_SET | FormButtonAttr | 양식개체 버튼 속성 |
| ShapeFormComboboxAttr | PIT_SET | FormComboboxAttr | 양식개체 콤보박스 속성 |
| ShapeFormEditAttr | PIT_SET | FormEditAttr | 양식개체 에디트박스 속성 |
| ShapeFormScrollbarAttr | PIT_SET | FormScrollbarAttr | 양식개체 스크롤바 속성 |
| ShapeFormListBoxAttr | PIT_SET | FormListBoxAttr | 양식개체 리스트박스 속성 |
| ShapeType | PIT_UI1 | | TablePropertyDialog 의 종류 |
| ShapeCellSize | PIT_UI1 | | 셀 크기 적용 여부 ( on / off ) |
| ShapeCreationType | PIT_UI1 | | 그리기 개체 형태 (DrawObjCreatorObject에서 사용) |

---

## 130

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| ShapeCreationMode | PIT_UI1 | | 마우스로 그리기 여부 ( on / off )<br>(DrawObjCreatorObject에서 사용) |
| ShapeComment | PIT_SET | ShapeObjComment | 개체 설명문 |
