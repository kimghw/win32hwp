# ParameterSetTable_2504_part05

---

## 33) DrawImageAttr : 그림 개체 속성

그림 개체의 속성을 지정하기 위한 파라메터셋.
DrawFillAttr에서 그림과 관련된 속성만 빼서 파라메터셋으로 지정되었다. 현재 DrawFillAttr와 상속관계가 지정되지 않았다. (차후 상속관계로 묶일 예정)

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| FileName | PIT_BSTR | | ShapeObject의 배경을 그림으로 선택했을 경우 또는 그림개체일 경우의 그림파일 경로 |
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
| ImageCreateTreatAsChar | PIT_UI1 | | 그림 넣을때 글자처럼 취급 여부 |
| ImageAdjustPrevObject | PIT_UI1 | | 앞 개체 속성 적용 |
| ImageAdjustTableCell | PIT_UI1 | | 테이블 셀 크기에 맞춰 이미지 조정 |
| ImageInsertFileNameInCaption | PIT_UI1 | | 캡션에 파일이름 넣기 |
| ImageAutoRotate | PIT_UI1 | | 이미지 자동 회전 |
| FileNameStr | PIT_BSTR | | 브러쉬 설정을 위한 itemid |
| ImageAlphaEffect | PIT_I1 | | 이미지 투명도 |
| ImageUseTextInPicture | PIT_UI1 | | 이미지에서 텍스트 추출, (한/글 2024부터 지원) |

---

## 34) DrawImageScissoring : 그림 개체의 자르기 정보

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Xoffset | PIT_I | | 자를 x축 오프셋 |
| Yoffset | PIT_I | | 자를 y축 오프셋 |
| HandleIndex | PIT_UI | | Reserved |

---

## 35) DrawLayOut : 그리기 개체의 Layout

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| CreateNumPt | PIT_UI | | 생성할 점의 수 |
| CreatePt | PIT_ARRAY | PIT_I | 생성할 점의 위치정보<br>POINT(x,y)로 계산되므로 CreateNumPt*2의 개수로 구성 됨. |
| CurveSegmentInfo | PIT_ARRAY | PIT_UI1 | 곡선 세그먼트 정보 |

---

## 36) DrawLineAttr : 그리기 개체의 선 속성

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Color | PMT_UINT32 | | 선 색깔<br>RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR) |
| Style | PMT_INT | | 선의 style |
| Width | PMT_INT | | 선의 width |
| EndCap | PMT_INT | | 선의 endcap |
| HeadStyle | PMT_INT | | 선의 시작 화살표 모양 |
| TailStyle | PMT_INT | | 선의 끝 화살표 모양 |
| HeadSize | PMT_INT | | 선의 시작 화살표 크기 |
| TailSize | PMT_INT | | 선의 끝 화살표 크기 |
| HeadFill | PMT_BOOL | | 선의 시작 화살표 채움 유무 |
| TailFill | PMT_BOOL | | 선의 끝 화살표 채움 유무 |
| OutLineStyle | PMT_UINT | | 선두께 외부/내부/일반 적용 |
| Alpha | PIT_UI1 | | 투명도 |

---

## 37) DrawRectType : 사각형 모서리 모양

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Type | PIT_UI | | 모서리의 곡률 지정 (0 ~ 50까지) |

---

## 38) DrawResize : 그리기 개체 Resizing 정보

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Xoffset | PIT_I | | 그리기 개체 X좌표 위치 |
| Yoffset | PIT_I | | 그리기 개체 Y좌표 위치 |
| HandleIndex | PIT_UI | | Reserved |
| Mode | PIT_UI | | Reserved |

---

## 39) DrawRotate : 그리기 개체 회전 정보

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Command | PIT_UI | | 회전 중심의 기준 설정<br>0 = 일반 회전<br>1 = 회전 영역의 중심을 기준으로 회전<br>2 = 그리기 개체 중심을 기준으로 회전<br>3 = 회전 영역의 중심을 기준으로 회전 & 중심이 변경됨 |
| CenterX | PIT_I | | 회전 중심의 X 좌표 |
| CenterY | PIT_I | | 회전 중심의 Y 좌표 |
| ObjectCenterX | PIT_I | | 그리기 개체의 중심 X 좌표 |
| ObjectCenterY | PIT_I | | 그리기 개체의 중심 Y 좌표 |
| Angle | PIT_I | | 회전 각 |
| RotateImage | PIT_UI1 | | 그림 회전 여부<br>0 = 회전 안 함<br>1 = 회전함 |

---

## 40) DrawScAction : 그리기 개체 90도 회전 및 좌우/상하 뒤집기

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| RotateCenterX | PIT_I4 | | 회전 중심 X 좌표 |
| RotateCenterY | PIT_I4 | | 회전 중심 Y 좌표 |
| RotateAngel | PIT_I | | 회전각 |
| HorzFlip | PIT_UI | | 수평 flip (좌우 뒤집기) |
| VertFlip | PIT_UI | | 수직 flip (상하 뒤집기) |

---

## 41) DrawShadow : 그리기 개체 그림자 정보

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| ShadowType | PIT_I4 | | 그림자 종류.<br>0 = none, 1 = drop, 2 = continuous |
| ShadowColor | PIT_UI4 | | 그림자 색 (COLORREF) |
| ShadowOffsetX | PIT_I4 | | 그림자 X축 간격 (-48% ~ 48%) |
| ShadowOffsetY | PIT_I4 | | 그림자 Y축 간격 (-48% ~ 48%) |
| ShadowAlpha | PIT_UI1 | | 그림자 투명도 |
