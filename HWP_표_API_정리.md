# HWP 표(Table) API 정리

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

### TableCreation 파라미터

| Item ID | 설명 |
|---------|------|
| Rows | 행 수 (기본값 5) |
| Cols | 열 수 (기본값 5) |
| RowHeight | 행의 높이 배열 |
| ColWidth | 열의 너비 배열 |
| WidthType | 너비 타입 |
| HeightType | 높이 타입 |

---

## 3. 표 관련 액션

### 줄/칸 삽입

| Action ID | 설명 |
|-----------|------|
| `TableInsertUpperRow` | 위쪽 줄 삽입 |
| `TableInsertLowerRow` | 아래쪽 줄 삽입 |
| `TableInsertLeftColumn` | 왼쪽 칸 삽입 |
| `TableInsertRightColumn` | 오른쪽 칸 삽입 |
| `TableAppendRow` | 줄 추가 |

### 줄/칸 삭제

| Action ID | 설명 |
|-----------|------|
| `TableDeleteRow` | 줄 삭제 |
| `TableDeleteColumn` | 칸 삭제 |
| `TableDeleteCell` | 셀 삭제 |

### 셀 합치기/나누기

| Action ID | 설명 |
|-----------|------|
| `TableMergeCell` | 셀 합치기 |
| `TableSplitCell` | 셀 나누기 |
| `TableSplitCellRow2` | 셀 줄 나누기 |
| `TableSplitCellCol2` | 셀 칸 나누기 |

### 셀 크기 조정

| Action ID | 설명 |
|-----------|------|
| `TableDistributeCellHeight` | 셀 높이를 같게 |
| `TableDistributeCellWidth` | 셀 너비를 같게 |
| `TableResizeCellUp/Down/Left/Right` | 셀 크기 변경 |

### 셀 이동

| Action ID | 설명 |
|-----------|------|
| `TableLeftCell` | 왼쪽 셀로 이동 |
| `TableRightCell` | 오른쪽 셀로 이동 |
| `TableUpperCell` | 위쪽 셀로 이동 |
| `TableLowerCell` | 아래쪽 셀로 이동 |
| `TableColBegin` | 열 시작으로 이동 |
| `TableColEnd` | 열 끝으로 이동 |

### 셀 블록 선택

| Action ID | 설명 |
|-----------|------|
| `TableCellBlock` | 셀 블록 선택 |
| `TableCellBlockRow` | 줄 단위 블록 |
| `TableCellBlockCol` | 칸 단위 블록 |
| `TableCellBlockExtend` | 셀 블록 확장 (F5+F5) |

### 셀 정렬

| Action ID | 설명 |
|-----------|------|
| `TableCellAlignLeftTop` | 왼쪽 위 정렬 |
| `TableCellAlignCenterCenter` | 가운데 정렬 |
| `TableCellAlignRightBottom` | 오른쪽 아래 정렬 |
| `TableVAlignTop` | 세로 정렬 - 위 |
| `TableVAlignCenter` | 세로 정렬 - 가운데 |
| `TableVAlignBottom` | 세로 정렬 - 아래 |

### 셀 테두리

| Action ID | 설명 |
|-----------|------|
| `TableCellBorderAll` | 모든 테두리 토글 |
| `TableCellBorderOutside` | 바깥 테두리 토글 |
| `TableCellBorderInside` | 안쪽 테두리 토글 |
| `TableCellBorderTop/Bottom/Left/Right` | 각 방향 테두리 |
| `TableCellBorderNo` | 모든 테두리 제거 |

### 표 계산

| Action ID | 설명 |
|-----------|------|
| `TableFormulaSumHor` | 가로 합계 |
| `TableFormulaSumVer` | 세로 합계 |
| `TableFormulaSumAuto` | 블록 합계 |
| `TableFormulaAvgHor` | 가로 평균 |
| `TableFormulaAvgVer` | 세로 평균 |
| `TableFormulaProHor` | 가로 곱 |
| `TableFormulaProVer` | 세로 곱 |

### 기타 표 작업

| Action ID | 설명 |
|-----------|------|
| `TablePropertyDialog` | 표 속성 대화상자 |
| `TableMergeTable` | 표 붙이기 |
| `TableSplitTable` | 표 나누기 |
| `TableStringToTable` | 문자열을 표로 변환 |
| `TableTableToString` | 표를 문자열로 변환 |
| `TableTreatAsChar` | 글자처럼 취급 |

---

## 4. Table 속성 (ParameterSet)

### 주요 속성

| Item ID | 설명 |
|---------|------|
| PageBreak | 페이지 경계 처리 (0=나누지 않음, 1=셀 유지, 2=텍스트도 나눔) |
| RepeatHeader | 제목 행 반복 여부 |
| CellSpacing | 셀 간격 |
| CellMarginLeft/Right/Top/Bottom | 셀 안쪽 여백 |
| TreatAsChar | 글자처럼 취급 |
| HorzAlign | 가로 정렬 (0=왼쪽, 1=가운데, 2=오른쪽) |
| VertAlign | 세로 정렬 (0=위, 1=가운데, 2=아래) |
| Width / Height | 크기 |

---

## 5. 표 가운데 정렬

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

## 6. 예제 코드

### 3x3 표 생성 후 데이터 입력

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

### 표 찾아서 삭제

```python
ctrl = hwp.HeadCtrl
while ctrl:
    if ctrl.CtrlID == "tbl":
        hwp.DeleteCtrl(ctrl)
        break  # 첫 번째 표만 삭제
    ctrl = ctrl.Next
```
