# HWP 표(Table) API 정리

> 최종 수정일: 2026년 1월 (프로젝트 테이블 모듈 문서 추가)

---

## 1. 표 식별하기

### 컨트롤 ID로 식별
표의 CtrlID는 `"tbl"`입니다.

```python
# 현재 커서 위치에서 컨트롤 찾기
ctrl_id = hwp.FindCtrl()
if ctrl_id == "tbl":
    print("현재 커서가 표 위에 있습니다")
```

### 문서 내 모든 표 찾기

```python
# HeadCtrl부터 순회
ctrl = hwp.HeadCtrl
table_count = 0
while ctrl:
    if ctrl.CtrlID == "tbl":
        table_count += 1
        print(f"표 {table_count} 발견")
        # ctrl.Properties로 속성 접근 가능
    ctrl = ctrl.Next

print(f"총 {table_count}개의 표")
```

### 컨트롤 ID 목록

| CtrlID | 설명 |
|--------|------|
| `tbl` | 표 |
| `fn` | 각주 |
| `en` | 미주 |
| `eqed` | 수식 |
| `gso` | 그리기 개체 |
| `head` | 머리말 |
| `foot` | 꼬리말 |

---

## 2. 표 생성

### 기본 표 만들기

```python
ctrl = hwp.HParameterSet.HTableCreation
hwp.HAction.GetDefault("TableCreate", ctrl.HSet)
ctrl.Rows = 5      # 행 수
ctrl.Cols = 5      # 열 수
ctrl.WidthType = 2  # 너비 타입 (2=단에 맞춤)
ctrl.HeightType = 0 # 높이 타입
hwp.HAction.Execute("TableCreate", ctrl.HSet)
```

### pyhwpx를 사용한 간편 생성

```python
# pyhwpx 사용 시
hwp.create_table(5, 5, treat_as_char=True)  # 5행 5열, 글자처럼 취급
hwp.create_table(3, 4)  # 3행 4열 기본 생성
```

### TableCreation 파라미터셋 (표 생성)

| Item ID | Type | Description |
|---------|------|-------------|
| Rows | PIT_UI2 | 행 수 (기본값 5) |
| Cols | PIT_UI2 | 열 수 (기본값 5) |
| RowHeight | PIT_ARRAY | 행의 기본 높이 배열 (HWPUNIT) |
| ColWidth | PIT_ARRAY | 열의 기본 너비 배열 (HWPUNIT) |
| CellInfo | PIT_ARRAY | 셀 정보 배열 (정보 없는 셀은 기본값 사용) |
| WidthType | PIT_UI1 | 너비 타입 (0=지정값, 1=종이에 맞춤, 2=단에 맞춤) |
| HeightType | PIT_UI1 | 높이 타입 |
| WidthValue | PIT_I | 너비 값 (HWPUNIT) |
| HeightValue | PIT_I | 높이 값 (HWPUNIT) |
| TableProperties | PIT_SET | 초기 표 속성 (Table 파라미터셋) |
| TableTemplate | PIT_SET | 표 마당 적용 속성 |
| TableDrawProperties | PIT_SET | 마우스로 선 그릴 때 속성 |
| TableTemplateValue | PIT_UI1 | 표 마당 적용 여부 |
| TextSelect | PIT_I | 텍스트 선택 여부 |

---

## 3. 표 속성 (Table ParameterSet)

Table은 ShapeObject로부터 상속받으므로 ShapeObject의 속성도 사용 가능합니다.

### 주요 속성

| Item ID | Type | Description |
|---------|------|-------------|
| PageBreak | PIT_UI1 | 페이지 경계 처리<br>0=나누지 않음<br>1=셀 단위로 나눔<br>2=텍스트도 나눔 |
| RepeatHeader | PIT_UI1 | 제목 행 반복 여부 (on/off) |
| CellSpacing | PIT_UI4 | 셀 간격 (HWPUNIT) |
| CellMarginLeft | PIT_I4 | 기본 셀 안쪽 여백 (왼쪽) |
| CellMarginRight | PIT_I4 | 기본 셀 안쪽 여백 (오른쪽) |
| CellMarginTop | PIT_I4 | 기본 셀 안쪽 여백 (위쪽) |
| CellMarginBottom | PIT_I4 | 기본 셀 안쪽 여백 (아래쪽) |
| BorderFill | PIT_SET | 표에 적용되는 테두리/배경 |
| Cell | PIT_SET | 셀 속성 |
| TreatAsChar | PIT_UI1 | 글자처럼 취급 (on/off) |

### 위치 관련 속성 (TreatAsChar=FALSE일 때)

| Item ID | Type | Description |
|---------|------|-------------|
| VertRelTo | PIT_UI1 | 세로 위치 기준<br>0=종이, 1=본문영역, 2=문단 |
| VertAlign | PIT_UI1 | 세로 정렬<br>0=위, 1=가운데, 2=아래, 3=안쪽, 4=바깥쪽 |
| VertOffset | PIT_UI4 | 세로 오프셋 (HWPUNIT) |
| HorzRelTo | PIT_UI1 | 가로 위치 기준<br>0=종이, 1=본문영역, 2=단, 3=문단 |
| HorzAlign | PIT_UI1 | 가로 정렬<br>0=왼쪽, 1=가운데, 2=오른쪽, 3=안쪽, 4=바깥쪽 |
| HorzOffset | PIT_I4 | 가로 오프셋 (HWPUNIT) |

### 크기 관련 속성

| Item ID | Type | Description |
|---------|------|-------------|
| WidthRelTo | PIT_UI1 | 너비 기준<br>0=종이, 1=본문영역, 2=단, 3=문단, 4=절대값 |
| Width | PIT_I4 | 너비 값 (%/HWPUNIT) |
| HeightRelTo | PIT_UI1 | 높이 기준<br>0=종이, 1=본문영역, 2=절대값 |
| Height | PIT_I4 | 높이 값 (%/HWPUNIT) |
| ProtectSize | PIT_UI1 | 크기 보호 (on/off) |

### 텍스트 배치 속성

| Item ID | Type | Description |
|---------|------|-------------|
| TextWrap | PIT_UI1 | 텍스트 배치<br>0=사각형에 맞춤, 1=좌우 배치 안함, 2=글 뒤로, 3=글 앞으로, 4=외곽선 따라, 5=빈 공간까지 |
| TextFlow | PIT_UI1 | 배치 방향<br>0=양쪽, 1=왼쪽, 2=오른쪽, 3=큰쪽 |
| FlowWithText | PIT_UI1 | 본문 영역 제한 (on/off) |
| AllowOverlap | PIT_UI1 | 겹침 허용 (on/off) |

### 예제: Table 속성 설정

```python
# 표 속성 설정 (Visual Basic 스타일)
table_set = hwp.CreateSet("Table")
table_set.SetItem("PageBreak", 1)       # 셀 단위로 나눔
table_set.SetItem("TreatAsChar", True)  # 글자처럼 취급
table_set.SetItem("RepeatHeader", True) # 제목 행 반복
```

---

## 4. 셀 속성 (Cell ParameterSet)

Cell은 ListProperties로부터 상속받습니다.

| Item ID | Type | Description |
|---------|------|-------------|
| HasMargin | PIT_UI1 | 자체 셀 여백 사용 여부 (on/off) |
| Protected | PIT_UI1 | 편집 보호 (0=off, 1=on) |
| Header | PIT_UI1 | 제목 셀 여부 (0=off, 1=on) |
| Width | PIT_I4 | 셀 너비 (HWPUNIT) |
| Height | PIT_I4 | 셀 높이 (HWPUNIT) |
| Editable | PIT_UI1 | 양식모드에서 편집 가능 여부 |
| Dirty | PIT_UI1 | 수정 상태 여부 |

---

## 5. 셀 테두리/배경 (CellBorderFill ParameterSet)

CellBorderFill은 BorderFillExt로부터 상속받습니다.

### CellBorderFill 고유 속성

| Item ID | Type | Description |
|---------|------|-------------|
| ApplyTo | PIT_UI1 | 적용 대상<br>0=선택된 셀, 1=전체 셀, 2=여러 셀에 걸쳐 |
| NoNeighborCell | PIT_UI1 | 주변 셀에 적용 안함 (1=적용 안함) |
| TableBorderFill | PIT_SET | 표 테두리/배경 |
| AllCellsBorderFill | PIT_SET | 전체 셀 테두리/배경 |
| SelCellsBorderFill | PIT_SET | 선택된 셀 테두리/배경 |

### BorderFill 공통 속성 (상속됨)

| Item ID | Type | Description |
|---------|------|-------------|
| BorderTypeLeft/Right/Top/Bottom | PIT_UI2 | 4방향 테두리 종류 |
| BorderWidthLeft/Right/Top/Bottom | PIT_UI1 | 4방향 테두리 두께 |
| BorderColorLeft/Right/Top/Bottom | PIT_UI4 | 4방향 테두리 색 (0x00BBGGRR) |
| TypeHorz/TypeVert | PIT_UI2 | 중앙선 종류 (가로/세로) |
| WidthHorz/WidthVert | PIT_UI1 | 중앙선 두께 |
| ColorHorz/ColorVert | PIT_UI4 | 중앙선 색 |
| DiagonalType | PIT_UI2 | 대각선 종류 |
| DiagonalWidth | PIT_UI1 | 대각선 두께 |
| DiagonalColor | PIT_UI4 | 대각선 색 |
| SlashFlag | PIT_UI2 | 슬래시 대각선 플래그 |
| BackSlashFlag | PIT_UI2 | 백슬래시 대각선 플래그 |
| BorderFill3D | PIT_UI1 | 3D 효과 (on/off) |
| Shadow | PIT_UI1 | 그림자 효과 (on/off) |
| FillAttr | PIT_SET | 배경 채우기 속성 |

### 예제: 셀 테두리 점선으로 변경

```python
# 셀 테두리를 점선으로 설정
hwp.HAction.GetDefault("CellBorder", hwp.HParameterSet.HCellBorderFill.HSet)
hwp.HParameterSet.HCellBorderFill.BorderTypeBottom = 3  # 점선
hwp.HParameterSet.HCellBorderFill.BorderTypeTop = 3
hwp.HParameterSet.HCellBorderFill.BorderTypeLeft = 3
hwp.HParameterSet.HCellBorderFill.BorderTypeRight = 3
hwp.HAction.Execute("CellBorder", hwp.HParameterSet.HCellBorderFill.HSet)
```

### 예제: 셀에 꺾인 대각선 넣기

```javascript
// JavaScript 예제
var vAct = vHwpCtrl.CreateAction("CellBorder");
var vSet = vAct.CreateSet();
vAct.GetDefault(vSet);

vSet.SetItem("DiagonalType", 1);      // 실선
vSet.SetItem("BackSlashFlag", 0x02);  // 중앙 대각선
vSet.SetItem("CrookedSlashFlag2", 1); // 꺾인 대각선

vAct.Execute(vSet);
```

---

## 6. 표 관련 액션 (Action)

### 표 생성/삭제

| Action ID | ParameterSet | Description |
|-----------|--------------|-------------|
| TableCreate | TableCreation | 표 만들기 |
| TableStringToTable | TableStrToTbl | 문자열을 표로 변환 |
| TableTableToString | TableTblToStr | 표를 문자열로 변환 |
| TableSplitTable | - | 표 나누기 |
| TableMergeTable | - | 표 붙이기 |

### 줄/칸 삽입

| Action ID | ParameterSet | Description |
|-----------|--------------|-------------|
| TableInsertUpperRow | TableInsertLine | 위쪽 줄 삽입 |
| TableInsertLowerRow | TableInsertLine | 아래쪽 줄 삽입 |
| TableInsertLeftColumn | TableInsertLine | 왼쪽 칸 삽입 |
| TableInsertRightColumn | TableInsertLine | 오른쪽 칸 삽입 |
| TableInsertRowColumn | TableInsertLine | 줄-칸 삽입 |
| TableAppendRow | - | 줄 추가 |

### TableInsertLine 파라미터

| Item ID | Type | Description |
|---------|------|-------------|
| Side | PIT_UI1 | 방향 |
| Count | PIT_UI1 | 삽입할 개수 |

### 줄/칸 삭제

| Action ID | ParameterSet | Description |
|-----------|--------------|-------------|
| TableDeleteRow | TableDeleteLine | 줄 삭제 |
| TableDeleteColumn | TableDeleteLine | 칸 삭제 |
| TableDeleteRowColumn | TableDeleteLine | 줄-칸 삭제 |
| TableDeleteCell | - | 셀 삭제 |
| TableSubtractRow | TableDeleteLine | 줄 빼기 |

### TableDeleteLine 파라미터

| Item ID | Type | Description |
|---------|------|-------------|
| Type | PIT_UI1 | 0=줄, 1=칸 |

### 셀 합치기/나누기

| Action ID | ParameterSet | Description |
|-----------|--------------|-------------|
| TableMergeCell | - | 셀 합치기 |
| TableSplitCell | TableSplitCell | 셀 나누기 |
| TableSplitCellRow2 | TableSplitCell | 셀 줄 나누기 |
| TableSplitCellCol2 | TableSplitCell | 셀 칸 나누기 |

### TableSplitCell 파라미터

| Item ID | Type | Description |
|---------|------|-------------|
| Cols | PIT_UI2 | 나눌 칸 수 |
| Rows | PIT_UI2 | 나눌 줄 수 |
| DistributeHeight | PIT_UI1 | 줄 높이를 같게 |
| Merge | PIT_UI1 | 나누기 전에 합치기 |
| Mode2 | PIT_UI1 | 셀 어긋남 방지 모드 |

### 셀 크기 조정

| Action ID | Description |
|-----------|-------------|
| TableDistributeCellHeight | 셀 높이를 같게 |
| TableDistributeCellWidth | 셀 너비를 같게 |
| TableResizeCellUp | 셀 크기 변경 (위) |
| TableResizeCellDown | 셀 크기 변경 (아래) |
| TableResizeCellLeft | 셀 크기 변경 (왼쪽) |
| TableResizeCellRight | 셀 크기 변경 (오른쪽) |
| TableResizeExUp/Down/Left/Right | 셀 블록 없이도 동작 |
| TableResizeLineUp/Down/Left/Right | 선 기준 크기 변경 |

### 셀 이동

| Action ID | Description |
|-----------|-------------|
| TableLeftCell | 왼쪽 셀로 이동 |
| TableRightCell | 오른쪽 셀로 이동 |
| TableRightCellAppend | 오른쪽 셀에 이어서 |
| TableUpperCell | 위쪽 셀로 이동 |
| TableLowerCell | 아래쪽 셀로 이동 |
| TableColBegin | 열 시작으로 이동 |
| TableColEnd | 열 끝으로 이동 |
| TableColPageUp | 페이지 업 |
| TableColPageDown | 페이지 다운 |

### 셀 블록 선택

| Action ID | Description |
|-----------|-------------|
| TableCellBlock | 셀 블록 선택 |
| TableCellBlockRow | 줄 단위 블록 |
| TableCellBlockCol | 칸 단위 블록 |
| TableCellBlockExtend | 셀 블록 연장 (F5+F5) |
| TableCellBlockExtendAbs | 셀 블록 연장 (Shift+F5) |

### 셀 정렬

| Action ID | Description |
|-----------|-------------|
| TableCellAlignLeftTop | 왼쪽 위 정렬 |
| TableCellAlignLeftCenter | 왼쪽 가운데 정렬 |
| TableCellAlignLeftBottom | 왼쪽 아래 정렬 |
| TableCellAlignCenterTop | 가운데 위 정렬 |
| TableCellAlignCenterCenter | 가운데 정렬 |
| TableCellAlignCenterBottom | 가운데 아래 정렬 |
| TableCellAlignRightTop | 오른쪽 위 정렬 |
| TableCellAlignRightCenter | 오른쪽 가운데 정렬 |
| TableCellAlignRightBottom | 오른쪽 아래 정렬 |
| TableVAlignTop | 세로 위 정렬 |
| TableVAlignCenter | 세로 가운데 정렬 |
| TableVAlignBottom | 세로 아래 정렬 |

### 셀 테두리

| Action ID | Description |
|-----------|-------------|
| TableCellBorderAll | 모든 테두리 토글 |
| TableCellBorderOutside | 바깥 테두리 토글 |
| TableCellBorderInside | 안쪽 테두리 토글 |
| TableCellBorderInsideHorz | 안쪽 가로 테두리 토글 |
| TableCellBorderInsideVert | 안쪽 세로 테두리 토글 |
| TableCellBorderTop | 위 테두리 토글 |
| TableCellBorderBottom | 아래 테두리 토글 |
| TableCellBorderLeft | 왼쪽 테두리 토글 |
| TableCellBorderRight | 오른쪽 테두리 토글 |
| TableCellBorderDiagonalUp | 대각선(⍁) 토글 |
| TableCellBorderDiagonalDown | 대각선(⍂) 토글 |
| TableCellBorderNo | 모든 테두리 제거 |

### 셀 배경/음영

| Action ID | ParameterSet | Description |
|-----------|--------------|-------------|
| CellBorder | CellBorderFill | 셀 테두리 설정 |
| CellBorderFill | CellBorderFill | 셀 테두리/배경 설정 |
| CellFill | CellBorderFill | 셀 배경 설정 |
| CellZoneBorder | CellBorderFill | 여러 셀 테두리 |
| CellZoneBorderFill | CellBorderFill | 여러 셀 테두리/배경 |
| CellZoneFill | CellBorderFill | 여러 셀 배경 |
| TableCellShadeInc | CellBorderFill | 음영 증가 (어둡게) |
| TableCellShadeDec | CellBorderFill | 음영 감소 (밝게) |

### 셀 문자 방향

| Action ID | ParameterSet | Description |
|-----------|--------------|-------------|
| TableCellTextHorz | CellBorderFill | 가로 쓰기 |
| TableCellTextVert | CellBorderFill | 세로 쓰기 |
| TableCellTextVertAll | CellBorderFill | 세로 쓰기 (영문 세움) |
| TableCellToggleDirection | CellBorderFill | 방향 토글 |

### 표 계산 (수식)

| Action ID | Description |
|-----------|-------------|
| TableFormulaSumHor | 가로 합계 |
| TableFormulaSumVer | 세로 합계 |
| TableFormulaSumAuto | 블록 합계 |
| TableFormulaAvgHor | 가로 평균 |
| TableFormulaAvgVer | 세로 평균 |
| TableFormulaAvgAuto | 블록 평균 |
| TableFormulaProHor | 가로 곱 |
| TableFormulaProVer | 세로 곱 |
| TableFormulaProAuto | 블록 곱 |
| TableFormula | 계산식 (FieldCtrl 사용) |

### 기타 표 작업

| Action ID | ParameterSet | Description |
|-----------|--------------|-------------|
| TablePropertyDialog | ShapeObject | 표 속성 대화상자 |
| TableTreatAsChar | ShapeObject | 글자처럼 취급 토글 |
| TableBreak | Table | 쪽 경계에서 나누지 않음 |
| TableBreakCell | Table | 셀 단위로 나눔 |
| TableBreakNone | Table | 텍스트도 나눔 |
| TableSwap | TableSwap | 표 뒤집기 |
| TableTemplate | TableTemplate | 표 마당 적용 |
| TableAutoFill | AutoFill | 자동 채우기 |
| TableDrawPen | - | 표 그리기 |
| TableEraser | - | 표 지우개 |
| TableInsertComma | - | 세 자리마다 쉼표 넣기 |
| TableDeleteComma | - | 세 자리 쉼표 빼기 |
| ShapeObjTableSelCell | - | 표 선택 상태에서 첫 셀 선택 |

### 캡션 위치

| Action ID | Description |
|-----------|-------------|
| TableCaptionPosTop | 캡션 위치 - 위 |
| TableCaptionPosBottom | 캡션 위치 - 아래 |
| TableCaptionPosLeftTop | 캡션 위치 - 왼쪽 위 |
| TableCaptionPosLeftCenter | 캡션 위치 - 왼쪽 가운데 |
| TableCaptionPosLeftBottom | 캡션 위치 - 왼쪽 아래 |
| TableCaptionPosRightTop | 캡션 위치 - 오른쪽 위 |
| TableCaptionPosRightCenter | 캡션 위치 - 오른쪽 가운데 |
| TableCaptionPosRightBottom | 캡션 위치 - 오른쪽 아래 |

---

## 7. 추가 파라미터셋

### TableSwap (표 뒤집기)

| Item ID | Type | Description |
|---------|------|-------------|
| Type | PIT_UI1 | 뒤집기 형식<br>0=상하, 1=좌우, 2=X/Y 교환, 3=반시계 90도, 4=180도, 5=시계 90도 |
| SwapMargin | PIT_UI1 | 여백 뒤집기 지원 |

### TableStrToTbl (문자열을 표로)

| Item ID | Type | Description |
|---------|------|-------------|
| DelimiterType | PIT_UI1 | 분리 문자 (탭, 쉼표, 공백) |
| UserDefine | PIT_BSTR | 사용자 정의 구분 기호 |
| AutoOrDefine | PIT_UI1 | 자동/지정 선택 |
| KeepSeperator | PIT_UI1 | 구분자 유지 |
| DelimiterEtc | PIT_BSTR | 기타 구분 기호 |
| TableCreation | PIT_SET | 표 속성 (TableCreation) |

### TableTblToStr (표를 문자열로)

| Item ID | Type | Description |
|---------|------|-------------|
| DelimiterType | PIT_UI1 | 분리 문자 (탭, 쉼표, 공백) |
| UserDefine | PIT_BSTR | 사용자 정의 구분 기호 |

### TableTemplate (표 마당)

| Item ID | Type | Description |
|---------|------|-------------|
| Format | PIT_UI | 적용할 서식 (비트 조합)<br>0x0001=테두리, 0x0002=글자/문단 모양, 0x0004=셀 배경, 0x0008=그레이스케일 |
| ApplyTarget | PIT_UI | 적용 대상 (비트 조합)<br>0x0001=제목줄, 0x0002=마지막줄, 0x0004=첫째칸, 0x0008=마지막칸 |

### TableDrawPen (표 그리기 펜)

| Item ID | Type | Description |
|---------|------|-------------|
| Style | PIT_UI2 | 선 모양 |
| Width | PIT_UI1 | 선 굵기 |
| Color | PIT_UI4 | 선 색 (0x00BBGGRR) |

---

## 8. 표 가운데 정렬

표를 페이지 가운데에 배치하려면 **표가 포함된 문단**을 가운데 정렬합니다.

```python
# 표 위치로 이동 후 문단 정렬
hwp.HAction.Run('MoveDocEnd')  # 문서 끝으로 (표가 있는 곳)
hwp.HAction.Run('MoveParaBegin')
hwp.HAction.Run('MoveSelParaEnd')
hwp.HAction.Run('ParagraphShapeAlignCenter')
hwp.HAction.Run('Cancel')
```

---

## 9. 예제 코드

### 예제 1: 3x3 표 생성 후 데이터 입력

```python
# 표 생성
ctrl = hwp.HParameterSet.HTableCreation
hwp.HAction.GetDefault("TableCreate", ctrl.HSet)
ctrl.Rows = 3
ctrl.Cols = 3
hwp.HAction.Execute("TableCreate", ctrl.HSet)

# 각 셀에 데이터 입력
data = [["A1", "B1", "C1"], ["A2", "B2", "C2"], ["A3", "B3", "C3"]]

for row in range(3):
    for col in range(3):
        # 텍스트 입력
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = data[row][col]
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

        # 다음 셀로 이동 (마지막 셀 제외)
        if not (row == 2 and col == 2):
            hwp.HAction.Run('TableRightCell')
```

### 예제 2: 표 찾아서 삭제

```python
ctrl = hwp.HeadCtrl
while ctrl:
    if ctrl.CtrlID == "tbl":
        hwp.DeleteCtrl(ctrl)
        break  # 첫 번째 표만 삭제
    ctrl = ctrl.Next
```

### 예제 3: 셀 합치기

```python
# 먼저 합칠 셀을 블록 선택
hwp.HAction.Run('TableCellBlock')  # 셀 블록 시작
# 화살표로 범위 확장 또는
hwp.HAction.Run('TableCellBlockExtend')  # F5+F5 효과

# 셀 합치기
hwp.HAction.Run('TableMergeCell')
```

### 예제 4: 줄/칸 삽입

```python
# 아래에 3줄 삽입
hwp.HAction.GetDefault("TableInsertLowerRow", hwp.HParameterSet.HTableInsertLine.HSet)
hwp.HParameterSet.HTableInsertLine.Count = 3
hwp.HAction.Execute("TableInsertLowerRow", hwp.HParameterSet.HTableInsertLine.HSet)

# 오른쪽에 2칸 삽입
hwp.HAction.GetDefault("TableInsertRightColumn", hwp.HParameterSet.HTableInsertLine.HSet)
hwp.HParameterSet.HTableInsertLine.Count = 2
hwp.HAction.Execute("TableInsertRightColumn", hwp.HParameterSet.HTableInsertLine.HSet)
```

### 예제 5: 셀 나누기

```python
# 현재 셀을 2행 3열로 나누기
hwp.HAction.GetDefault("TableSplitCell", hwp.HParameterSet.HTableSplitCell.HSet)
hwp.HParameterSet.HTableSplitCell.Rows = 2
hwp.HParameterSet.HTableSplitCell.Cols = 3
hwp.HParameterSet.HTableSplitCell.DistributeHeight = 1  # 높이 같게
hwp.HAction.Execute("TableSplitCell", hwp.HParameterSet.HTableSplitCell.HSet)
```

### 예제 6: 셀 배경색 설정

```python
# 셀 블록 선택 후
hwp.HAction.GetDefault("CellFill", hwp.HParameterSet.HCellBorderFill.HSet)

# 배경 채우기 설정
fill_attr = hwp.HParameterSet.HCellBorderFill.FillAttr
fill_attr.SetItem("FillType", 1)  # 단색 채우기
fill_attr.SetItem("FaceColor", 0x00FFFF00)  # 노란색 (BGR)

hwp.HAction.Execute("CellFill", hwp.HParameterSet.HCellBorderFill.HSet)
```

### 예제 7: 표 뒤집기 (좌우)

```python
# 표를 선택한 상태에서
hwp.HAction.GetDefault("TableSwap", hwp.HParameterSet.HTableSwap.HSet)
hwp.HParameterSet.HTableSwap.Type = 1  # 좌우 뒤집기
hwp.HAction.Execute("TableSwap", hwp.HParameterSet.HTableSwap.HSet)
```

### 예제 8: 문자열을 표로 변환

```python
# 탭으로 구분된 텍스트를 표로 변환
hwp.HAction.GetDefault("TableStringToTable", hwp.HParameterSet.HTableStrToTbl.HSet)
hwp.HParameterSet.HTableStrToTbl.DelimiterType = 0  # 탭
hwp.HParameterSet.HTableStrToTbl.AutoOrDefine = 0   # 자동

# TableCreation 속성 설정
creation = hwp.HParameterSet.HTableStrToTbl.TableCreation
creation.SetItem("WidthType", 2)  # 단에 맞춤

hwp.HAction.Execute("TableStringToTable", hwp.HParameterSet.HTableStrToTbl.HSet)
```

### 예제 9: 셀 여백 설정 (pyhwpx)

```python
# pyhwpx 사용 시
hwp.set_table_inside_margin(1, 1, 1, 1)  # 상하좌우 1mm
hwp.set_cell_margin(0.5, 0.5, 0.5, 0.5)  # 현재 셀만
```

### 예제 10: 표를 DataFrame으로 변환 (pyhwpx)

```python
# pyhwpx 사용 시
import pandas as pd

# 표를 DataFrame으로 추출
df = hwp.table_to_df()
print(df)

# DataFrame을 표로 삽입
data = {"이름": ["홍길동", "김철수"], "나이": [30, 25]}
df = pd.DataFrame(data)
hwp.table_from_data(df)
```

---

## 10. 참고 자료

### 공식 문서
- [한컴 오피스 도움말 - 셀 테두리/배경](https://help.hancom.com/hoffice/multi/ko_kr/hwp/table/cellborder/cellborder.htm)
- [한컴 오피스 도움말 - 표 테두리/배경](https://help.hancom.com/hoffice/multi/ko_kr/hwp/table/tableborder/tableborder(background).htm)

### 커뮤니티/라이브러리
- [pyhwpx PyPI](https://pypi.org/project/pyhwpx/) - 한글 자동화 래퍼 라이브러리
- [hwpapi PyPI](https://pypi.org/project/hwpapi/) - HWP API 래퍼
- [pyhwpx Cookbook (WikiDocs)](https://wikidocs.net/book/8956) - 행정업무 자동화 가이드
- [한컴디벨로퍼 포럼](https://forum.developer.hancom.com/) - 공식 개발자 포럼

### 선 종류 참조값

| 값 | 선 종류 |
|----|--------|
| 0 | 없음 |
| 1 | 실선 (Solid) |
| 2 | 점선 (Dot) |
| 3 | 파선 (Dash) |
| 4 | 일점쇄선 (DashDot) |
| 5 | 이점쇄선 (DashDotDot) |
| 6 | 긴 파선 |
| 7 | 이중 실선 |

### HWPUNIT 단위 변환

```python
# 1 HWPUNIT = 1/7200 인치
# 1 pt = 100 HWPUNIT
# 1 mm = 약 283 HWPUNIT

def pt_to_hwpunit(pt):
    return pt * 100

def mm_to_hwpunit(mm):
    return int(mm * 283.46)

def hwpunit_to_pt(hwpunit):
    return hwpunit / 100

def hwpunit_to_mm(hwpunit):
    return hwpunit / 283.46
```

---

## 11. 프로젝트 테이블 모듈 (table/)

프로젝트에서 제공하는 테이블 관련 모듈들입니다.

### 모듈 구조

| 모듈 | 파일 | 설명 |
|------|------|------|
| `table_info` | `table/table_info.py` | 테이블 구조 탐색, 셀 수집, 좌표 매핑 |
| `table_boundary` | `table/table_boundary.py` | 테이블 경계 판별, 서브테이블 분할, 인접 관계 |
| `table_cell_info` | `table/table_cell_info.py` | 셀 순회, 컨트롤 탐색, 서식 조회 |
| `table_field` | `table/table_field.py` | 필드 CRUD, 좌표 연동 |

---

## 12. MovePos 셀 이동 상수 (100~107)

셀 간 이동에 사용하는 MovePos 상수입니다.

```python
from table.table_info import (
    MOVE_LEFT_OF_CELL,    # 100: 왼쪽 셀로 이동
    MOVE_RIGHT_OF_CELL,   # 101: 오른쪽 셀로 이동
    MOVE_UP_OF_CELL,      # 102: 위쪽 셀로 이동
    MOVE_DOWN_OF_CELL,    # 103: 아래쪽 셀로 이동
    MOVE_START_OF_CELL,   # 104: 행의 시작 셀 (맨 왼쪽)
    MOVE_END_OF_CELL,     # 105: 행의 끝 셀 (맨 오른쪽)
    MOVE_TOP_OF_CELL,     # 106: 열의 시작 셀 (맨 위)
    MOVE_BOTTOM_OF_CELL,  # 107: 열의 끝 셀 (맨 아래)
)

# 사용 예시
hwp.MovePos(MOVE_RIGHT_OF_CELL, 0, 0)  # 오른쪽 셀로 이동
hwp.MovePos(MOVE_DOWN_OF_CELL, 0, 0)   # 아래쪽 셀로 이동
hwp.MovePos(MOVE_START_OF_CELL, 0, 0)  # 행의 첫 번째 셀로 이동
hwp.MovePos(MOVE_TOP_OF_CELL, 0, 0)    # 열의 첫 번째 셀로 이동
```

**참고:** 기본 MovePos 상수(0~13)와 달리 100~107은 테이블 셀 간 이동 전용입니다.

---

## 13. 셀 정보 수집: collect_cells_bfs()

`TableInfo.collect_cells_bfs()`는 행 우선 순회로 테이블의 모든 셀을 탐색하고 이웃 정보를 수집합니다.

### 사용법

```python
from table.table_info import TableInfo
from cursor import get_hwp_instance

hwp = get_hwp_instance()
table = TableInfo(hwp, debug=True)

# 셀 수집 (커서가 테이블 내부에 있어야 함)
cells = table.collect_cells_bfs()

# 결과: Dict[int, CellInfo]
# - key: list_id
# - value: CellInfo(list_id, left, right, up, down, width, height)
for list_id, cell in cells.items():
    print(f"셀 {list_id}: 좌={cell.left}, 우={cell.right}, 상={cell.up}, 하={cell.down}")
    print(f"  크기: {cell.width}x{cell.height} HWPUNIT")
```

### CellInfo 데이터클래스

```python
@dataclass
class CellInfo:
    list_id: int      # 셀의 고유 ID
    left: int = 0     # 왼쪽 셀의 list_id (없으면 0)
    right: int = 0    # 오른쪽 셀의 list_id
    up: int = 0       # 위쪽 셀의 list_id
    down: int = 0     # 아래쪽 셀의 list_id
    width: int = 0    # 셀 너비 (HWPUNIT)
    height: int = 0   # 셀 높이 (HWPUNIT)
```

### 알고리즘

1. 첫 셀에서 시작하여 오른쪽으로 이동하며 행의 모든 셀 수집
2. 행 끝에서 첫 셀로 돌아가 아래로 이동
3. 다음 행 반복
4. 각 셀에서 4방향(좌/우/상/하) 이웃 정보 저장

---

## 14. 좌표 매핑: map_cell_coordinates_v4()

`TableBoundary.map_cell_coordinates_v4()`는 병합 셀을 포함한 복잡한 테이블의 정확한 (row, col) 좌표를 계산합니다.

### 사용법

```python
from table.table_boundary import TableBoundary
from cursor import get_hwp_instance

hwp = get_hwp_instance()
boundary = TableBoundary(hwp, debug=True)

# 경계 분석
result = boundary.check_boundary_table()

# 좌표 매핑 (v4: 서브셀 + 위쪽 셀 기준 열 조정)
coord_result = boundary.map_cell_coordinates_v4(result)

# 결과: CellCoordinateResult
# - cells: Dict[int, CellCoordinate]  # list_id -> 좌표 정보
# - max_row, max_col: 최대 행/열 번호
# - traversal_order: 순회 순서

for list_id, cell in coord_result.cells.items():
    print(f"셀 {list_id}: ({cell.row}, {cell.col}), 방문횟수={cell.visit_count}")
```

### CellCoordinateResult 데이터클래스

```python
@dataclass
class CellCoordinate:
    list_id: int
    row: int = 0          # 행 번호 (0부터 시작)
    col: int = 0          # 열 번호 (0부터 시작)
    visit_count: int = 0  # 해당 셀 방문 횟수

@dataclass
class CellCoordinateResult:
    cells: Dict[int, CellCoordinate]  # list_id -> CellCoordinate
    max_row: int = 0
    max_col: int = 0
    traversal_order: List[int]        # 순회 순서
```

### 알고리즘 (v4)

**1단계: v3 기본 매핑**
- first_col[0]에서 시작하여 오른쪽으로 순회
- 새 셀(처음 방문)을 만나면 → 현재 행, 열+1
- 이미 방문한 셀(2회 이상)을 만나면 → 행+1 (한 번만), 열은 왼쪽 열+1 유지
- first_col을 만나면 → 열=0

**2단계: 위쪽 셀 기준 열 조정**
- 2행부터 각 셀에 대해
- 커서를 0행까지 위로 올리면서 모든 위쪽 셀의 열 중 가장 큰 값 찾음
- 열 번호 = max(위쪽 셀들의 최대 열, 왼쪽 셀의 열 + 1)

### 버전 비교

| 버전 | 설명 | 용도 |
|------|------|------|
| `map_cell_coordinates()` | 방문 횟수로 행 결정 | 간단한 테이블 |
| `map_cell_coordinates_v2()` | first_col 기반 개선 | 중간 복잡도 |
| `map_cell_coordinates_v3()` | 서브셀 기반 | 병합 셀 포함 |
| `map_cell_coordinates_v4()` | v3 + 위쪽 셀 기준 열 조정 | **권장** |

---

## 15. 인접 관계: build_cell_adjacency()

`TableBoundary.build_cell_adjacency()`는 모든 셀의 인접 노드(상/하/좌/우) 관계를 계산합니다.

### 사용법

```python
from table.table_boundary import TableBoundary
from cursor import get_hwp_instance

hwp = get_hwp_instance()
boundary = TableBoundary(hwp, debug=True)

# 좌표 매핑 먼저 수행
coord_result = boundary.map_cell_coordinates_v4()

# 인접 관계 계산
adjacency = boundary.build_cell_adjacency(coord_result)

# 결과: CellAdjacencyResult
# - nodes: Dict[int, CellAdjacency]  # list_id -> 인접 정보
# - max_row, max_col

for list_id, node in adjacency.nodes.items():
    print(f"노드 {list_id} ({node.row},{node.col}):")
    print(f"  왼쪽: {node.left}, 오른쪽: {node.right}")
    print(f"  위쪽: {node.up}, 아래쪽: {node.down}")
```

### CellAdjacency 데이터클래스

```python
@dataclass
class CellAdjacency:
    list_id: int
    row: int = 0
    col: int = 0
    left: List[int]   # 왼쪽 인접 노드들 (병합 셀이면 여러 개)
    right: List[int]  # 오른쪽 인접 노드들
    up: List[int]     # 위쪽 인접 노드들
    down: List[int]   # 아래쪽 인접 노드들
```

### 알고리즘

1. 각 셀에서 상/하/좌/우로 커서 이동하여 인접 셀 찾음
2. **양방향 관계 활용**: A→B 이면 B→A도 성립
   - 예: 13의 up이 3이면, 3의 down에 13 추가
3. 병합 셀(예: 13)이 여러 셀(3~8)과 인접한 경우도 처리

### JSON 저장

```python
# 인접 관계를 JSON 파일로 저장
filepath = boundary.save_adjacency_to_json(adjacency, "cell_adjacency.json")
```

---

## 16. 필드 CRUD 함수 (TableField)

`TableField` 클래스는 테이블 내 필드의 생성/조회/수정/삭제와 셀 좌표 연동을 제공합니다.

### 초기화

```python
from table.table_field import TableField
from cursor import get_hwp_instance

hwp = get_hwp_instance()
tf = TableField(hwp, debug=True)

# 첫 번째 테이블 진입 (좌표 맵 자동 초기화)
tf.enter_table(0)
```

### 필드 조회 (Read)

```python
# 모든 필드 조회
fields = tf.get_all_fields()

# 셀 필드만 조회
cell_fields = tf.get_cell_fields()

# 누름틀 필드만 조회
clickhere_fields = tf.get_clickhere_fields()

# 이름으로 필드 조회
field = tf.get_field_by_name("필드이름")

# 특정 좌표의 필드들 조회
fields_at = tf.get_fields_at_coord(row=1, col=2)

# 필드 텍스트 조회
text = tf.get_field_text("필드이름")

# 필드 존재 여부
exists = tf.field_exists("필드이름")
```

### 필드 생성 (Create)

```python
# 현재 커서 위치에 누름틀 필드 생성
tf.create_field_at_cursor("필드이름", direction="안내문", memo="설명")

# 특정 좌표에 필드 생성
tf.create_field_at_coord(row=1, col=2, name="필드이름")

# 셀 자체를 필드로 설정
tf.set_cell_field_name(row=1, col=2, name="셀필드")

# 구조화된 필드 이름 일괄 설정
# 형식: {캡션}_{영역}_{row}_{col} (예: 표1_head_0_0)
result = tf.set_structured_field_names(
    caption="표1",
    header_rows=1,
    footer_rows=0
)
# 결과: {'head': [...], 'body': [...], 'foot': [...]}

# JSON 형식 필드 이름 설정
# 형식: {"table":"표이름","coord":[row,col],"desc":"","target":"","author":""}
result = tf.set_table_fields_json(table_name="표1")
```

### 필드 수정 (Update)

```python
# 필드에 텍스트 입력
tf.put_field_text("필드이름", "입력할 텍스트")

# 특정 인덱스의 필드에 텍스트 입력 (동일 이름 필드가 여러 개일 때)
tf.put_field_text_by_index("필드이름", index=0, text="텍스트")

# 여러 필드에 일괄 입력
tf.put_fields_text({
    "필드1": "값1",
    "필드2": "값2"
})

# 필드 이름 변경
tf.rename_field("기존이름", "새이름")

# 여러 필드 이름 일괄 변경
tf.rename_fields({
    "old1": "new1",
    "old2": "new2"
})

# 필드 편집 가능 속성 변경
tf.modify_field_editable("필드이름", editable=True)
```

### 필드 삭제 (Delete)

```python
# 현재 커서 위치의 필드 삭제 (내용은 유지)
tf.delete_field_at_cursor()

# 이름으로 필드 삭제
tf.delete_field_by_name("필드이름")

# 특정 좌표의 필드 삭제
tf.delete_field_at_coord(row=1, col=2)

# 필드 텍스트 비우기
tf.clear_field_text("필드이름")
```

### 필드 이동 (Navigation)

```python
# 필드로 이동
tf.move_to_field("필드이름", select=False)

# 좌표의 첫 번째 필드로 이동
tf.move_to_field_by_coord(row=1, col=2)
```

### JSON 필드 관련

```python
# JSON 필드명 생성
field_name = tf.make_field_name(
    table="표1", row=1, col=2,
    desc="설명", target="대상", author="작성자"
)

# table과 coord로 필드 찾기
full_name = tf.find_field_by_coord("표1", row=1, col=2)

# table과 coord로 필드 값 조회
value = tf.get_field_value_by_coord("표1", row=1, col=2)

# table과 coord로 필드 값 설정
tf.set_field_value_by_coord("표1", row=1, col=2, text="값")
```

### 유틸리티

```python
# 테이블 크기
size = tf.get_table_size()  # {'rows': 5, 'cols': 3}

# 필드 요약 정보
summary = tf.get_field_summary()
# {'total': 10, 'in_table': 8, 'outside_table': 2, 'fields': [...]}

# 필드 맵 출력 (디버그용)
tf.print_field_map()

# 병합 셀 정보
merge_info = tf.get_merge_info(list_id)
# {'colspan': 2, 'rowspan': 1, 'representative': (0, 1), 'coords': [(0,1), (0,2)]}
```

---

## 17. 테이블 경계 분석

`TableBoundary.check_boundary_table()`은 테이블의 경계 정보를 분석합니다.

### 사용법

```python
from table.table_boundary import TableBoundary
from cursor import get_hwp_instance

hwp = get_hwp_instance()
boundary = TableBoundary(hwp, debug=True)

# 경계 분석 (커서가 테이블 내부에 있어야 함)
result = boundary.check_boundary_table()

print(f"table_origin: {result.table_origin}")     # 첫 번째 셀
print(f"table_end: {result.table_end}")           # 마지막 셀
print(f"table_cell_counts: {result.table_cell_counts}")  # 총 셀 수
print(f"first_rows: {result.first_rows}")         # 첫 번째 행 셀들
print(f"bottom_rows: {result.bottom_rows}")       # 마지막 행 셀들
print(f"first_cols: {result.first_cols}")         # 첫 번째 열 셀들
print(f"last_cols: {result.last_cols}")           # 마지막 열 셀들
```

### TableBoundaryResult 데이터클래스

```python
@dataclass
class TableBoundaryResult:
    table_origin: int = 0       # 테이블 첫 번째 셀의 list_id
    table_end: int = 0          # 마지막 셀의 list_id
    table_cell_counts: int = 0  # 총 셀 개수
    first_rows: List[int]       # 첫 번째 행 셀들의 list_id
    bottom_rows: List[int]      # 마지막 행 셀들의 list_id
    first_cols: List[int]       # 첫 번째 열 셀들의 list_id
    last_cols: List[int]        # 마지막 열 셀들의 list_id
```

---

## 18. 셀 유틸리티 함수 (table_cell_info)

### 셀 순회

```python
from table.table_cell_info import iterate_table_cells

def my_callback(hwp, row, col, list_id):
    print(f"셀 ({row}, {col}): list_id={list_id}")
    return True  # False 반환 시 순회 중단

# 테이블 셀 순회
visited = iterate_table_cells(hwp, my_callback)
print(f"방문한 셀: {len(visited)}개")
```

### 셀 내 컨트롤 찾기

```python
from table.table_cell_info import get_ctrls_in_cell

# 특정 셀 내의 컨트롤 찾기
ctrls = get_ctrls_in_cell(hwp, target_list_id=5)
for ctrl in ctrls:
    print(f"컨트롤: {ctrl['id']} ({ctrl['desc']}), 문단={ctrl['para']}")
```

### 서식 정보 조회

```python
from table.table_cell_info import get_char_shape_info, get_para_shape_info

# 글자 모양 정보
font_name, font_size, char_spacing = get_char_shape_info(hwp)

# 문단 모양 정보
align, line_spacing = get_para_shape_info(hwp)
```

### 색상 텍스트 삽입

```python
from table.table_cell_info import insert_colored_text

# BGR 형식 색상으로 텍스트 삽입
insert_colored_text(hwp, "빨간색 텍스트", 0x0000FF)  # 빨강
insert_colored_text(hwp, "파란색 텍스트", 0xFF0000)  # 파랑
```
