# ParameterSet Table (Part 04)

---

## 25) DocFilters : Document 필터 리스트

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| DocFilters | PIT_BSTR | | '\|'문자로 분리된 필터 리스트 스트링(MFC 와 같은 방법)을 가져옴 |
| Format | PIT_ARRAY | PIT_BSTR | 필터 리스트를 string array형태로 가져옴 (Export시에만) |
| Type | PIT_UI1 | | Import 리스트를 가져올 것인지 Export 리스트를 가져올 것인지의 관한 타입.<br>Import = 1<br>Export = 0 (on input) |
| Hide | PIT_BSTR | | 필터 목록에서 제거를 위한 필터 아이템 Format 문자열. 구분자 '\|' 사용 |

---

## 26) DocFindInfo : 문서 찾기

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| ListID | PIT_UI | | 현재 위치한 리스트 |
| ParaID | PIT_UI | | 현재 위치한 문단 |
| Pos | PIT_UI | | 현재 위치한 글자 |

---

## 27) DocumentInfo : 문서에 대한 정보

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| CurPara | PIT_UI | | 현재 위치한 문단 |
| CurPos | PIT_UI | | 현재 위치한 오프셋 |
| CurParaLen | PIT_UI | | 현재 위치한 문단의 길이 |
| CurCtrl | PIT_UI | | 현재 리스트의 종류<br>0 = 일반 텍스트<br>1 = 글상자<br>기타 = 컨트롤 ID |
| CurParaCount | PIT_UI | | 현재 리스트의 문단 수 |
| RootPara | PIT_UI | | 루트 리스트의 현재 문단 |
| RootPos | PIT_UI | | 루트 리스트의 현재 오프셋 |
| RootParaCout | PIT_UI | | 루트 리스트의 문단 수 |
| DetailInfo | PIT_I1 | | 자세한 정보를 구할지 여부<br>Detail~ 로 시작하는 아이템의 정보를 얻기 위해서는 이 값을 1로 넣어준 후에 액션을 실행해준다. |
| DetailCharCount | PIT_UI | | 문서에 포함된 글자 수 |
| DetailWordCount | PIT_UI | | 문서에 포함된 어절 수 |
| DetailLineCount | PIT_UI | | 문서에 포함된 줄 수 |
| DetailPageCount | PIT_UI | | 문서에 포함된 쪽 수 |
| DetailCurPage | PIT_UI | | 현재 쪽 번호 |
| DetailCurPrtPage | PIT_UI | | 현재 쪽 번호 (인쇄 번호) |
| SectionInfo | PIT_UI | | 구역의 속성까지 구할지 여부<br>SecDef 아이템은 이 값을 1로 넣어준 후 액션을 실행한 후에 얻을 수 있다. |
| SecDef | PIT_SET | SecDef | 구역의 속성 |

---

## 28) DrawArcType : 그리기 개체의 부채꼴 테두리 모양

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Type | PIT_UI | | 호의 유형 :<br>0 = 호, 1 = 부채꼴, 2 = 화살모양 |
| Interval | PIT_ARRAY | PIT_I | 확장타원(호)에서 호의 시작점과 끝점을 나타내게 한다. |

---

## 29) DrawCoordInfo : 그리기 개체의 좌표 정보

정보를 얻을 때만 사용하도록 한다.

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Count | PIT_UI4 | | 점의 개수 |
| Point | PIT_ARRAY | PIT_I | 좌표 Array (X1,Y1), (X2,Y2), ..., (Xn,Yn) |
| Line | PIT_ARRAY | PIT_UI1 | 선 속성 Array(곡선에서만 쓰임) |

---

## 30) DrawCtrlHyperlink : 그리기 개체의 Hyperlink 정보

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Command | PIT_BSTR | | Command String 참조 |

---

## 31) DrawEditDetail : 그리기 개체의 다각형 편집

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Command | PIT_UI | | Reserved |
| Index | PIT_UI | | 고칠 점의 인덱스 |
| PointX | PIT_I | | 점의 X 좌표 |
| PointY | PIT_I | | 점의 Y 좌표 |

---

## 32) DrawFillAttr : 그리기 개체의 채우기 속성

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Type | PIT_UI | | 배경 유형<br>0 = 채우기 없음<br>1 = 면색 또는 무늬색<br>2 = 그림<br>3 = 그러데이션 |
| GradationType | PIT_I | | 배경 유형이 그러데이션일 때 그러데이션 형태<br>1 = 줄무늬형<br>2 = 원형<br>3 = 원뿔형<br>4 = 사각형 |
| GradationAngle | PIT_I | | 그러데이션의 기울임(시작각) |
| GradationCenterX | PIT_I | | 그러데이션의 가로중심(중심 X 좌표) |
| GradationCenterY | PIT_I | | 그러데이션의 가로중심(중심 Y 좌표) |
| GradationStep | PIT_I | | 그러데이션 번짐 정도 (0..100) |
| GradationColorNum | PIT_I | | 그러데이션의 색 수 |
| GradationColor | PIT_ARRAY | PIT_I | 그러데이션의 색깔<br>(시작색, 중간색1,..중간색 n-2, 끝색) 2<= n <=10 |
| GradationIndexPos | PIT_ARRAY | PIT_I | 그러데이션 다음 색깔과의 거리(얼마나 번지고 나서 다음색깔로 가는가) |
| GradationStepCenter | PIT_UI1 | | 그러데이션 번짐 정도의 중심 (0..100) |
| GradationAlpha | PIT_UI1 | | 그러데이션 투명도 |
| WinBrushFaceColor | PIT_UI | | 면 색 (RGB 0x00BBGGRR) |
| WinBrushHatchColor | PIT_UI | | 무늬 색 (RGB 0x00BBGGRR) |
| WinBrushFaceStyle | PIT_I1 | | 무늬 스타일<br>0 = (horizontal lines)<br>1 = (vertical lines)<br>2 = (forward diagonal)<br>3 = (backward diagonal)<br>4 = (cross)<br>5 = (diagonal cross)<br>6 = (none) |
| WinBrushAlpha | PIT_UI | | 면/무늬 색 투명도 |
| FileName | PIT_BSTR | | ShapeObject의 배경을 그림으로 선택했을 경우. 또는 그림개체일 경우의 그림파일 경로 |
| Embedded | PIT_UI1 | | 그림이 문서에 직접 삽입(TRUE) / 파일로 연결(FALSE) |
| PicEffect | PIT_UI1 | | 그림 효과<br>0 = 실제 이미지 그대로<br>1 = 그레이스케일<br>2 = 흑백 효과 |
| Brightness | PIT_I1 | | 명도 (-100 ~ 100) |
| Contrast | PIT_I1 | | 밝기 (-100 ~ 100) |
| Reverse | PIT_UI1 | | 반전 유무 |
| DrawFillImageType | PIT_I | | ShapeObject의 배경일 경우에만 의미 있는 아이템, 배경을 채우는 방식을 결정한다. (그림개체에는 해당사항 없음)<br>0 = 바둑판식으로<br>1 = 가로/위만 바둑판식으로 배열<br>2 = 가로/아래만 바둑판식으로 배열<br>3 = 세로/왼쪽만 바둑판식으로 배열<br>4 = 세로/오른쪽만 바둑판식으로 배열<br>5 = 크기에 맞추어<br>6 = 가운데로<br>7 = 가운데 위로<br>8 = 가운데 아래로<br>9 = 왼쪽 가운데로<br>10 = 왼쪽 위로<br>11 = 왼쪽 아래로<br>12 = 오른쪽 가운데로<br>13 = 오른쪽 위로<br>14 = 오른쪽 아래로 |
| SkipLeft | PIT_UI | | 그림 개체일 경우에만 의미 있는 아이템, 왼쪽 자르기 |
| SkipRight | PIT_UI | | 그림 개체일 경우에만 의미 있는 아이템, 오른쪽 자르기 |
| SkipTop | PIT_UI | | 그림 개체일 경우에만 의미 있는 아이템, 위 자르기 |
| SkipBottom | PIT_UI | | 그림 개체일 경우에만 의미 있는 아이템, 아래 자르기 |
| OriginalSizeX | PIT_UI | | 그림 개체일 경우에만 의미 있는 아이템, 이미지 원본 크기 X size |
| OriginalSizeY | PIT_UI | | 그림 개체일 경우에만 의미 있는 아이템, 이미지 원본 크기 Y size |
| InsideMarginLeft | PIT_I4 | | 그림 개체일 경우에만 의미 있는 아이템, 이미지 안쪽 여백. (왼쪽) |
| InsideMarginRight | PIT_I4 | | 그림 개체일 경우에만 의미 있는 아이템, 이미지 안쪽 여백. (오른쪽) |
| InsideMarginTop | PIT_I4 | | 그림 개체일 경우에만 의미 있는 아이템, 이미지 안쪽 여백. (위) |
| InsideMarginBottom | PIT_I4 | | 그림 개체일 경우에만 의미 있는 아이템, 이미지 안쪽 여백. (아래) |
| WindowsBrush | PIT_UI1 | | 현재 선택된 brush의 type이 면/무늬 브러시인가를 나타냄 |
| GradationBrush | PIT_UI1 | | 현재 선택된 brush의 type이 그러데이션 브러시인가를 나타냄 |
| ImageBrush | PIT_UI1 | | 현재 선택된 brush의 type이 그림 브러시인가를 나타냄 |
| ImageCreateOnDrag | PIT_UI1 | | 그림 개체 생성 시 마우스로 끌어 생성할지 여부 |
| ImageAlpha | PIT_UI1 | | 그림 개체/배경 투명도 |
| FileNameStr | PIT_BSTR | | 브러쉬 설정을 위한 itemid |

---

*Pages 31-40*
