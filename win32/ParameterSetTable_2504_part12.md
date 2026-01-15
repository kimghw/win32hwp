# ParameterSetTable_2504_part12

## Page 111

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| PrintAutoFootnoteLtext | PIT_BSTR | | 자동 꼬리말의 왼쪽 String |
| PrintAutoFootnoteCtext | PIT_BSTR | | 자동 꼬리말의 가운데 String |
| PrintAutoFootnoteRtext | PIT_BSTR | | 자동 꼬리말의 오른쪽 String |
| PrinterName | PIT_BSTR | | 프린터 |
| PrintToFile | PIT_UI1 | | 인쇄 결과를 파일로 저장 |
| FileName | PIT_BSTR | | 인쇄 결과를 저장할 파일 이름 |
| ReverseOrder | PIT_UI1 | | 역순 인쇄 |
| Pause | PIT_UI2 | | 끊어 찍기 매수 |
| PrintImage | PIT_UI1 | | 그림 개체 |
| PrintDrawObj | PIT_UI1 | | 그리기 개체 |
| PrintClickHere | PIT_UI1 | | 누름틀 |
| PrintCropMark | PIT_UI1 | | 편집 용지 표시 |
| IdcPrintWallPaper | PIT_UI1 | | 배경 그림 |
| LastBlankPage | PIT_UI1 | | 빈 마지막 쪽 |
| BinderHoleType | PIT_UI1 | | 바인더 구멍 |
| EvenOddPageType | PIT_UI1 | | 홀짝 인쇄 |
| ZoomX | PIT_UI2 | | 가로 확대 |
| ZoomY | PIT_UI2 | | 세로 확대 |
| Flags | PIT_UI | | 문제 해결을 위한 고급 선택 사항 |
| Device | PIT_UI1 | | 인쇄 방향(장치)<br>0 : 프린터<br>1: 팩스<br>2: 그림으로 저장<br>3: PDF 파일로 저장<br>4: 미리보기 |
| PrintFormObj | PIT_UI1 | | 양식 개체 출력여부 |
| PrintMarkPen | PIT_UI1 | | 형광펜 출력여부 |
| PrintMemo | PIT_UI1 | | 메모 출력여부 |
| PrintMemoContents | PIT_UI1 | | 메모 내용 출력여부 |
| PrintRevision | PIT_UI1 | | 교정부호 출력여부 |
| PrintWatermark | PIT_SET | PrintWatermark | 인쇄워터마크 |
| UserOrder | PIT_UI1 | | 사용자가 입력한 순서대로 인쇄 |

---

## Page 112

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| OverlapSize | PIT_UI2 | | 나눠찍기시 내용이 겹치는 크기 |
| UsingPagenum | PIT_UI1 | | 문서의 쪽 번호로 입력 |
| PrintBarcode | PIT_UI1 | | 바코드 |
| PrintPronounce | PIT_UI1 | | 한자/일어 발음 표시 |
| PrintColorSet | PIT_UI | | 색 변경 인쇄, 0 = normal, 1 = gary, 2 = light gray |
| PrintWithoutBlank | PIT_I | | 빈 쪽 없이 이어서 인쇄 |

---

## Page 113

### 96) PrintToImage : 그림으로 저장

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Format | PIT_UI1 | | 그림 형식<br>0 : none<br>1 : BMP<br>2 : GIF<br>3 : PNG<br>4 : JPG<br>5 : WMF |
| FileName | PIT_BSTR | | 그림 경로 |
| ColorDepth | PIT_UI1 | | 색상수 (bits: 8, 16...) |
| Resolution | PIT_UI2 | | 해상도 |
| Range | PIT_UI1 | | 인쇄 범위<br>0 : 문서 전체 (연결된 문서 포함)<br>1 : 현재 페이지만<br>2 : 현재 페이지부터<br>3 : 현재 페이지까지<br>4 : 사용자 정의<br>5 : 선택한 쪽만<br>6 : 현재 문서 (연결된 문서 미포함)<br>7 : 현재 구역 |
| Width | PIT_UI | | 그림 너비(pixel), (한/글 2022부터 지원) |
| Height | PIT_UI | | 그림 높이(pixel), (한/글 2022부터 지원) |

---

## Page 114

### 97) PrintWatermark : 워터마크 속성

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| WatermarkType | PIT_UI | | 현재 선택된 워터마크의 유형을 나타냄.<br>0 = 워터마크 없음<br>1 = 그림 워터마크<br>2 = 글자 워터마크 |
| PosPage | PIT_UI1 | | 워터마크의 위치 기준 : 0 = 종이 기준, 1 = 쪽 기준 |
| TextWrap | PIT_UI1 | | 워터마크의 배치 : 0 = 글 뒤로, 1 = 글 앞으로 |
| AlphaText | PIT_UI1 | | 글자 투명도 (0 ~ 255) |
| AlphaImage | PIT_UI1 | | 그림 투명도 (0 ~ 255) |
| FileName | PIT_BSTR | | 그림 파일의 경로 or 그림파일 삽입일 경우에는 binary data |
| PicEffect | PIT_UI1 | | 그림 효과 :<br>0 = 실제 이미지 그대로, 1 = 그레이스케일, 2 = 흑백효과 |
| Brightness | PIT_I1 | | 명도 (-100 ~ 100) |
| Contrast | PIT_I1 | | 밝기 (-100 ~ 100) |
| DrawFillImageType | PIT_I | | 채우기 유형<br>0 = 바둑판식으로 - 모두<br>1 = 바둑판식으로 - 가로/위<br>2 = 바둑판식으로 - 가로/아래<br>3 = 바둑판식으로 - 세로/왼쪽<br>4 = 바둑판식으로 - 세로/오른쪽<br>5 = 크기에 맞추어<br>6 = 가운데로<br>7 = 가운데 위로<br>8 = 가운데 아래로<br>9 = 왼쪽 가운데로<br>10 = 왼쪽 위로<br>11 = 왼쪽 아래로<br>12 = 오른쪽 가운데로<br>13 = 오른쪽 위로<br>14 = 오른쪽 아래로<br>15 = 원래크기에 비례하여 |
| String | PIT_BSTR | | 글맵시에 넣을 문자열 내용 : 내용 |
| FontName | PIT_BSTR | | 글꼴 |
| FontType | PIT_UI1 | | 글꼴 속성 : 0 = don't care, 1 = TTF, 2 = HFT |
| FontSize | PIT_I | | 글꼴 크기 (HWPUNIT : 2500(25pt) ~ 25400(254pt) |
| ShadowType | PIT_UI1 | | 그림자 종류 :<br>0 = none, 1 = drop, 2 = continuous |
| ShadowOffsetX | PIT_I1 | | X축 그림자 간격 (-48% ~ 48% ) |
| ShadowOffsetY | PIT_I1 | | Y축 그림자 간격 (-48% ~ 48% ) |
| ShadowColor | PIT_UI4 | | 그림자 색 (COLORREF) |

---

## Page 115

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| FontColor | PIT_UI4 | | 글자색 (COLORREF) |
| RotateAngle | PIT_I | | 회전각도 (-360 ~ 360) |
| WaterMarkEff | PIT_UI1 | | 워터마크 효과 : 0 = off, 1 = on |

---

## Page 116

### 98) QCorrect : 빠른 교정

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| LauncherKey | PIT_UI1 | | 빠른 교정을 실행한 키 정보 |
| HyperLinkRunKey | PIT_UI1 | | URL 또는 email 하이퍼링크 작성 키 정보 |

---

## Page 117

### 99) RevisionDef : 교정부호 데이터

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| SignType | PIT_UI | | 교정부호 종류 :<br>0 = 교정부호 없음<br>1 = 띄움표<br>2 = 줄 바꿈표<br>3 = 줄 비움표<br>4 = 메모형 고침표<br>5 = 지움표<br>6 = 붙임표<br>7 = 뺌표<br>8 = 줄 이음표<br>9 = 줄 붙임표<br>10 = 톱니표<br>11 = 생각표<br>12 = 칭찬표<br>13 = 줄표<br>14 = 부호 넣음표<br>15 = 넣음표<br>16 = 고침표<br>17 = 자리 바꿈표<br>18 = 오른자리 옮김표<br>19 = 자료연결<br>20 = 왼자리 옮김표<br>21 = 부분자리 옮김표<br>22 = 줄 서로 바꿈표<br>23 = 자리바꿈 나눔표(내부용)<br>24 = 줄 서로 바꿈 나눔표(내부용) |
| SubText | PIT_BSTR | | 교정 문자열<br>교정 문자열을 가질 수 있는 교정부호만 적용. 나머지는 무시 |
| Margin | PIT_I4 | | 여백(HWPUNIT). 오른자리 옮김표와 왼자리 옮김표일 경우에만 적용. |
| BeginPos | PIT_I4 | | 시작위치(HWPUNIT). 오른자리 옮김표와 왼자리 옮김표일 경우에만 적용. |

---

## Page 118

### 100) SaveFootnote : 주석 저장

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| FileName | PIT_BSTR | | 파일 이름 |
| Flag | PIT_UI1 | | 옵션<br>1 : 각주 저장<br>2 : 미주 저장<br>3 : 각주/미주 저장 |

---

## Page 119

### 101) ScriptMacro : 스크립트 매크로

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Index | PIT_I | | 정의(or 실행)할 매크로의 인덱스 |
| RepeatCount | PIT_I | | 실행 반복 횟수 |
| Name | PIT_BSTR | | 매크로 이름 |
| Detail | PIT_BSTR | | 매크로 설명 |

---

## Page 120

### 102) SecDef : 구역의 속성

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| TextDirection | PIT_UI1 | | 글자 방향 |
| StartsOn | PIT_UI1 | | 구역 나눔으로 새 페이지가 생길 때의 페이지 번호 적용 옵션<br>0 = 이어서, 1 = 홀수, 2 = 짝수, 3 = 임의 값 |
| LineGrid | PIT_I4 | | 세로로 줄맞춤을 할지 여부.<br>0 = off, 1 - n = 간격을 HWPUNIT 단위로 지정 |
| CharGrid | PIT_I4 | | 가로로 줄맞춤을 할지 여부.<br>0 = off, 1 - n = 간격을 HWPUNIT 단위로 지정 |
| PageDef | PIT_SET | PageDef | 용지 설정 정보 |
| HideEmptyLine | PIT_UI1 | | 빈 줄 감춤 on / off |
| SpaceBetweenColumns | PIT_I4 | | 동일한 페이지에서 서로 다른 단 사이의 간격 |
| TabStop | PIT_I4 | | 기본 탭 간격 |
| FootnoteShape | PIT_SET | FootnoteShape | 각주 모양 |
| EndnoteShape | PIT_SET | FootnoteShape | 미주 모양 |
| HideHeader | PIT_UI1 | | 구역의 첫 쪽에만 머리말 감추기 옵션 on / off |
| HideFooter | PIT_UI1 | | 구역의 첫 쪽에만 꼬리말 감추기 옵션 on / off |
| HideMasterPage | PIT_UI1 | | 구역의 첫 쪽에만 바탕쪽 감추기 옵션 on / off |
| HideBorder | PIT_UI1 | | 구역의 첫 쪽에만 테두리 감추기 옵션 on / off |
| HideFill | PIT_UI1 | | 구역의 첫 쪽에만 배경 감추기 옵션 on / off |
| HidePageNumPos | PIT_UI1 | | 구역의 첫 쪽에만 쪽번호 감추기 옵션 on / off |
| FirstBorder | PIT_UI1 | | 구역의 첫 쪽에만 테두리 표시 옵션 on / off |
| FirstFill | PIT_UI1 | | 구역의 첫 쪽에만 배경 표시 옵션 on / off |
| OutlineShape | PIT_SET | NumberingShape | 개요 번호 |
| PageBorderFillBoth | PIT_SET | PageBorderFill | 쪽 테두리/배경 (양쪽) |
| PageBorderFillEven | PIT_SET | PageBorderFill | 쪽 테두리/배경 (짝수) |
| PageBorderFillOdd | PIT_SET | PageBorderFill | 쪽 테두리/배경 (홀수) |
| PageNumber | PIT_UI2 | | 쪽 시작 번호<br>0 = 앞 구역에 이어, n = 새 번호로 시작 |
| FigureNumber | PIT_UI2 | | 그림 시작 번호<br>0 = 앞 구역에 이어, n = 새 번호로 시작 |
| TableNumber | PIT_UI2 | | 표 시작 번호<br>0 = 앞 구역에 이어, n = 새 번호로 시작 |
