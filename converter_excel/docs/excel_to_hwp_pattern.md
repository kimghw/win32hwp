# 엑셀 데이터를 HWP 테이블에 작성하는 방법

## 개요

이 문서는 Python win32com을 사용하여 엑셀 데이터를 HWP 테이블에 작성하는 다양한 방법을 정리합니다.

---

## 1. HWP 테이블 생성 API

### 1.1 TableCreate 액션 (HAction 방식)

새로운 테이블을 생성하는 기본 방법입니다.

```python
import win32com.client as win32

hwp = win32.Dispatch('HWPFrame.HwpObject')
hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')

# 테이블 생성
hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
hwp.HParameterSet.HTableCreation.Rows = 5       # 행 수
hwp.HParameterSet.HTableCreation.Cols = 3       # 열 수
hwp.HParameterSet.HTableCreation.WidthType = 2  # 너비 타입
hwp.HParameterSet.HTableCreation.HeightType = 0 # 높이 타입
hwp.HAction.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
```

### 1.2 TableCreation ParameterSet 항목

| Item ID | Type | Description |
|---------|------|-------------|
| Rows | PIT_UI2 | 행 수 (생략하면 5) |
| Cols | PIT_UI2 | 칼럼 수 (생략하면 5) |
| RowHeight | PIT_ARRAY | 행의 디폴트 높이 (HWPUNIT) |
| ColWidth | PIT_ARRAY | 칼럼의 디폴트 폭 (HWPUNIT) |
| WidthType | PIT_UI1 | 너비 타입 |
| HeightType | PIT_UI1 | 높이 타입 |
| WidthValue | PIT_I | 너비 값 |
| HeightValue | PIT_I | 높이 값 |
| TableProperties | PIT_SET | 초기 표 속성 (Table) |
| TableTemplateValue | PIT_UI1 | 표 마당 적용 여부 |

#### WidthType 값

| 값 | 설명 |
|---|------|
| 0 | 단에 맞춤 |
| 1 | 문단에 맞춤 |
| 2 | 임의 값 |

### 1.3 InsertCtrl 방식 (대안)

```python
# TableCreation ParameterSet으로 표 삽입
table_set = hwp.CreateSet("TableCreation")
table_set.SetItem("Rows", 5)
table_set.SetItem("Cols", 3)

# 표 컨트롤 삽입 (tbl = 표 CtrlID)
table_ctrl = hwp.InsertCtrl("tbl", table_set)
```

**참고:** InsertCtrl의 경우:
- ctrlid: "tbl" (표의 컨트롤 ID)
- initparam: "TableCreation" ParameterSet 사용

---

## 2. 셀 이동 및 데이터 삽입

### 2.1 HAction.Run 방식 (셀 이동)

표 내에서 셀을 이동하는 액션들입니다.

| Action ID | Description |
|-----------|-------------|
| TableLeftCell | 셀 이동: 셀 왼쪽 |
| TableRightCell | 셀 이동: 셀 오른쪽 |
| TableUpperCell | 셀 이동: 셀 위 |
| TableLowerCell | 셀 이동: 셀 아래 |
| TableColBegin | 셀 이동: 열 시작 |
| TableColEnd | 셀 이동: 열 끝 |
| TableRightCellAppend | 셀 이동: 셀 오른쪽에 이어서 (마지막 셀에서 줄 추가) |

```python
# 오른쪽 셀로 이동
hwp.HAction.Run("TableRightCell")

# 아래 셀로 이동
hwp.HAction.Run("TableLowerCell")

# 열 시작으로 이동
hwp.HAction.Run("TableColBegin")
```

### 2.2 MovePos 방식 (셀 이동)

```python
# 셀 이동 관련 MovePos ID
hwp.MovePos(100)  # moveLeftOfCell - 왼쪽 셀
hwp.MovePos(101)  # moveRightOfCell - 오른쪽 셀
hwp.MovePos(102)  # moveUpOfCell - 위쪽 셀
hwp.MovePos(103)  # moveDownOfCell - 아래쪽 셀
hwp.MovePos(104)  # moveStartOfCell - 행(row)의 시작
hwp.MovePos(105)  # moveEndOfCell - 행(row)의 끝
hwp.MovePos(106)  # moveTopOfCell - 열(column)의 시작
hwp.MovePos(107)  # moveBottomOfCell - 열(column)의 끝
```

### 2.3 셀에 텍스트 삽입

```python
def insert_text(hwp, text):
    """현재 캐럿 위치에 텍스트 삽입"""
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = text
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
```

### 2.4 표 순회하며 데이터 입력 (실전 패턴)

```python
def fill_table_data(hwp, data):
    """
    2차원 리스트 데이터를 표에 채우기
    data: [[row1_col1, row1_col2, ...], [row2_col1, row2_col2, ...], ...]

    주의: 캐럿이 표의 첫 번째 셀에 있어야 함
    """
    for row_idx, row in enumerate(data):
        for col_idx, cell_value in enumerate(row):
            # 텍스트 삽입
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
            hwp.HParameterSet.HInsertText.Text = str(cell_value) if cell_value else ""
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

            # 다음 셀로 이동 (마지막 열이 아닌 경우)
            if col_idx < len(row) - 1:
                hwp.HAction.Run("TableRightCell")

        # 다음 행으로 이동 (마지막 행이 아닌 경우)
        if row_idx < len(data) - 1:
            hwp.HAction.Run("TableColBegin")  # 열 시작으로
            hwp.HAction.Run("TableLowerCell")  # 아래 행으로
```

### 2.5 VB.NET 스타일 예제 (참고)

```python
# VB.NET에서의 패턴을 Python으로 변환
def fill_table_vbnet_style(hwp, data, rows, cols):
    """VB.NET 스타일의 표 채우기"""
    for i in range(rows):
        for j in range(cols):
            # 텍스트 입력
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
            hwp.HParameterSet.HInsertText.Text = str(data[i][j])
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

            if j < cols - 1:
                # 오른쪽으로 이동
                hwp.HAction.Run("MoveRight")
            else:
                # 행의 끝: 다음 행 첫 번째 셀로 이동
                hwp.HAction.Run("TableCellBlock")   # 셀 블록 설정
                hwp.HAction.Run("TableColBegin")    # 열 시작으로
                hwp.HAction.Run("Cancel")           # 블록 해제
                hwp.HAction.Run("MoveDown")         # 아래로
```

---

## 3. 셀 병합 방법

### 3.1 TableMergeCell 액션

셀을 병합하려면 먼저 셀 블록을 선택해야 합니다.

```python
def merge_cells(hwp):
    """
    현재 선택된 셀 블록 병합
    주의: 미리 셀 블록이 선택되어 있어야 함
    """
    hwp.HAction.Run("TableMergeCell")
```

### 3.2 셀 블록 선택 방법

| Action ID | Description |
|-----------|-------------|
| TableCellBlock | 셀 블록 (현재 셀 선택) |
| TableCellBlockCol | 셀 블록 (칸 전체) |
| TableCellBlockRow | 셀 블록 (줄 전체) |
| TableCellBlockExtend | 셀 블록 연장 (F5 + F5) |
| TableCellBlockExtendAbs | 셀 블록 연장 (SHIFT + F5) |

### 3.3 셀 블록 선택 후 병합 예제

```python
def select_and_merge_cells(hwp, start_row, start_col, end_row, end_col):
    """
    특정 범위의 셀을 선택하고 병합

    참고: 표의 첫 번째 셀에 캐럿이 있다고 가정
    """
    # 시작 셀로 이동
    for _ in range(start_row):
        hwp.HAction.Run("TableLowerCell")
    for _ in range(start_col):
        hwp.HAction.Run("TableRightCell")

    # 셀 블록 시작
    hwp.HAction.Run("TableCellBlock")

    # 끝 셀까지 선택 확장
    for _ in range(end_row - start_row):
        hwp.HAction.Run("TableCellBlockExtend")
        hwp.HAction.Run("TableLowerCell")
    for _ in range(end_col - start_col):
        hwp.HAction.Run("TableCellBlockExtend")
        hwp.HAction.Run("TableRightCell")

    # 병합 실행
    hwp.HAction.Run("TableMergeCell")
```

### 3.4 셀 나누기 (TableSplitCell)

```python
# 셀 나누기 ParameterSet
hwp.HAction.GetDefault("TableSplitCell", hwp.HParameterSet.HTableSplitCell.HSet)
hwp.HParameterSet.HTableSplitCell.Cols = 2  # 칸 수
hwp.HParameterSet.HTableSplitCell.Rows = 2  # 줄 수
hwp.HParameterSet.HTableSplitCell.DistributeHeight = 1  # 줄 높이를 같게
hwp.HAction.Execute("TableSplitCell", hwp.HParameterSet.HTableSplitCell.HSet)
```

#### TableSplitCell ParameterSet

| Item ID | Type | Description |
|---------|------|-------------|
| Cols | PIT_UI2 | 칸 수 |
| Rows | PIT_UI2 | 줄 수 |
| DistributeHeight | PIT_UI1 | 줄 높이를 같게 |
| Merge | PIT_UI1 | 나누기 전에 합치기 |

---

## 4. 기존 테이블에 데이터 채우기 vs 새 테이블 생성

### 4.1 방법 비교

| 방식 | 장점 | 단점 | 사용 상황 |
|------|------|------|----------|
| 새 테이블 생성 | 완전한 제어 가능, 셀 크기/병합 자유롭게 설정 | 템플릿 활용 불가, 복잡한 서식 재현 어려움 | 단순한 데이터 표, 정형화된 출력 |
| 기존 테이블 채우기 | 템플릿 활용 가능, 복잡한 서식 유지 | 템플릿 필요, 셀 구조 변경 어려움 | 양식 문서, 복잡한 레이아웃 |

### 4.2 새 테이블 생성 패턴

```python
def create_table_from_excel(hwp, excel_data):
    """
    엑셀 데이터로 새 테이블 생성
    excel_data: 2차원 리스트
    """
    rows = len(excel_data)
    cols = len(excel_data[0]) if rows > 0 else 0

    if rows == 0 or cols == 0:
        return

    # 1. 테이블 생성
    hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
    hwp.HParameterSet.HTableCreation.Rows = rows
    hwp.HParameterSet.HTableCreation.Cols = cols
    hwp.HParameterSet.HTableCreation.WidthType = 2
    hwp.HAction.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet)

    # 2. 표 내부로 진입 (첫 번째 셀)
    hwp.HAction.Run("TableCellBlock")
    hwp.HAction.Run("Cancel")

    # 3. 데이터 채우기
    fill_table_data(hwp, excel_data)
```

### 4.3 기존 테이블에 필드로 데이터 채우기

**필드(Field) 방식의 장점:**
- 템플릿 문서의 서식 유지
- 특정 셀에 직접 접근 가능
- 데이터프레임과 연동 용이

```python
def fill_table_by_field(hwp, field_data):
    """
    필드명으로 데이터 입력
    field_data: {필드명: 값, ...}
    """
    for field_name, value in field_data.items():
        hwp.PutFieldText(field_name, str(value))
```

#### 다중 필드 입력 (구분자 사용)

```python
def fill_multiple_fields(hwp, fields, values):
    """
    여러 필드에 동시에 값 입력
    fields: ["필드1", "필드2", ...]
    values: ["값1", "값2", ...]
    """
    field_str = "\x02".join(fields)
    value_str = "\x02".join(str(v) for v in values)
    hwp.PutFieldText(field_str, value_str)
```

### 4.4 기존 테이블 탐색 및 채우기

```python
def fill_existing_table(hwp, data):
    """
    문서 내 첫 번째 표를 찾아 데이터 채우기
    """
    # 문서 시작으로 이동
    hwp.MovePos(2)  # moveTopOfFile

    # 표 찾기
    ctrl = hwp.HeadCtrl
    while ctrl:
        if ctrl.CtrlID == "tbl":
            # 표 발견: 첫 번째 셀로 이동
            hwp.SetPosBySet(ctrl.GetAnchorPos(0))
            hwp.HAction.Run("TableCellBlock")
            hwp.HAction.Run("TableColBegin")
            hwp.HAction.Run("Cancel")

            # 데이터 채우기
            fill_table_data(hwp, data)
            return True
        ctrl = ctrl.Next

    return False  # 표를 찾지 못함
```

---

## 5. 표 속성 설정 (Table ParameterSet)

### 5.1 Table ParameterSet 주요 항목

| Item ID | Type | Description |
|---------|------|-------------|
| PageBreak | PIT_UI1 | 페이지 경계 처리: 0=나누지 않음, 1=표만 나눔, 2=셀 내 텍스트도 나눔 |
| RepeatHeader | PIT_UI1 | 제목 행 반복 (on/off) |
| CellSpacing | PIT_UI4 | 셀 간격 (HWPUNIT) |
| CellMarginLeft | PIT_I4 | 기본 셀 안쪽 여백 (왼쪽) |
| CellMarginRight | PIT_I4 | 기본 셀 안쪽 여백 (오른쪽) |
| CellMarginTop | PIT_I4 | 기본 셀 안쪽 여백 (위쪽) |
| CellMarginBottom | PIT_I4 | 기본 셀 안쪽 여백 (아래쪽) |
| TreatAsChar | PIT_UI1 | 글자처럼 취급 (on/off) |
| Width | PIT_I4 | 표 너비 (HWPUNIT) |
| Height | PIT_I4 | 표 높이 (HWPUNIT) |

### 5.2 표 속성 설정 예제

```python
# 표 속성 설정 (표가 선택된 상태에서)
table_set = hwp.CreateSet("Table")
table_set.SetItem("PageBreak", 1)        # 표만 나눔
table_set.SetItem("TreatAsChar", True)   # 글자처럼 취급
```

---

## 6. 셀 정렬 및 서식

### 6.1 셀 정렬 액션

| Action ID | Description |
|-----------|-------------|
| TableVAlignTop | 셀 세로정렬 위 |
| TableVAlignCenter | 셀 세로정렬 가운데 |
| TableVAlignBottom | 셀 세로정렬 아래 |
| TableCellAlignCenterCenter | 가운데/가운데 정렬 |
| TableCellAlignLeftTop | 왼쪽/위 정렬 |
| TableCellAlignRightBottom | 오른쪽/아래 정렬 |

```python
# 셀 정렬 예제 (셀 내부에서 실행)
hwp.HAction.Run("TableVAlignCenter")  # 세로 가운데
```

### 6.2 셀 배경/테두리

| Action ID | ParameterSet | Description |
|-----------|--------------|-------------|
| TableCellShadeDec | CellBorderFill | 셀 배경 음영 낮추기 |
| TableCellShadeInc | CellBorderFill | 셀 배경 음영 높이기 |
| TableCellBorderAll | - | 모든 셀 테두리 토글 |
| TableCellBorderNo | - | 모든 셀 테두리 지움 |

---

## 7. 엑셀에서 HWP로 변환 실전 예제

### 7.1 pandas DataFrame을 HWP 테이블로

```python
import pandas as pd

def dataframe_to_hwp_table(hwp, df, include_header=True):
    """
    pandas DataFrame을 HWP 테이블로 변환
    """
    # 데이터 준비
    data = []
    if include_header:
        data.append(df.columns.tolist())
    data.extend(df.values.tolist())

    rows = len(data)
    cols = len(data[0]) if rows > 0 else 0

    # 테이블 생성
    hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
    hwp.HParameterSet.HTableCreation.Rows = rows
    hwp.HParameterSet.HTableCreation.Cols = cols
    hwp.HAction.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet)

    # 표 내부로 이동
    hwp.HAction.Run("TableCellBlock")
    hwp.HAction.Run("Cancel")

    # 데이터 채우기
    for row_idx, row in enumerate(data):
        for col_idx, value in enumerate(row):
            # None/NaN 처리
            text = "" if pd.isna(value) else str(value)

            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
            hwp.HParameterSet.HInsertText.Text = text
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

            if col_idx < cols - 1:
                hwp.HAction.Run("TableRightCell")

        if row_idx < rows - 1:
            hwp.HAction.Run("TableColBegin")
            hwp.HAction.Run("TableLowerCell")
```

### 7.2 openpyxl에서 읽어 HWP로

```python
from openpyxl import load_workbook

def excel_to_hwp(hwp, excel_path, sheet_name=None):
    """
    엑셀 파일을 읽어 HWP 테이블로 변환
    """
    wb = load_workbook(excel_path, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active

    # 데이터 추출
    data = []
    for row in ws.iter_rows(values_only=True):
        data.append(list(row))

    # HWP 테이블로 변환
    if data:
        create_table_from_excel(hwp, data)

    wb.close()
```

### 7.3 필드 기반 템플릿 채우기

```python
def fill_template_from_excel(hwp, template_path, excel_data):
    """
    템플릿 HWP 파일의 필드에 엑셀 데이터 채우기

    excel_data: {필드명: 값, ...}
    """
    # 템플릿 열기
    hwp.Open(template_path, "HWP", "lock:false")

    # 필드에 데이터 입력
    for field_name, value in excel_data.items():
        try:
            hwp.PutFieldText(field_name, str(value) if value else "")
        except:
            print(f"필드 '{field_name}' 입력 실패")

    # 저장
    hwp.Save()
```

---

## 8. 주요 API 요약

### 테이블 관련 Action ID

| Action ID | ParameterSet ID | Description |
|-----------|-----------------|-------------|
| TableCreate | TableCreation | 표 만들기 |
| TableMergeCell | - | 셀 합치기 |
| TableSplitCell | TableSplitCell | 셀 나누기 |
| TableAppendRow | - | 줄 추가 |
| TableInsertUpperRow | TableInsertLine | 위쪽 줄 삽입 |
| TableInsertLowerRow | TableInsertLine | 아래쪽 줄 삽입 |
| TableInsertLeftColumn | TableInsertLine | 왼쪽 칸 삽입 |
| TableInsertRightColumn | TableInsertLine | 오른쪽 칸 삽입 |
| TableDeleteRow | TableDeleteLine | 줄 지우기 |
| TableDeleteColumn | TableDeleteLine | 칸 지우기 |
| TablePropertyDialog | ShapeObject | 표 고치기 (대화상자) |

### 셀 이동 관련

| Action ID | Description |
|-----------|-------------|
| TableLeftCell | 왼쪽 셀 |
| TableRightCell | 오른쪽 셀 |
| TableUpperCell | 위 셀 |
| TableLowerCell | 아래 셀 |
| TableColBegin | 열 시작 |
| TableColEnd | 열 끝 |

### 셀 선택 관련

| Action ID | Description |
|-----------|-------------|
| TableCellBlock | 셀 블록 시작 |
| TableCellBlockExtend | 셀 블록 확장 |
| TableCellBlockRow | 줄 전체 선택 |
| TableCellBlockCol | 칸 전체 선택 |

---

## 9. 참고 자료

### 로컬 문서
- `/mnt/d/hwp_docs/win32/ActionTable_2504_part05.md` - 테이블 관련 액션
- `/mnt/d/hwp_docs/win32/ParameterSetTable_2504_part15.md` - Table, TableCreation ParameterSet
- `/mnt/d/hwp_docs/win32/HwpAutomation_2504_part03.md` - MovePos, InsertCtrl 등

### 웹 참고 자료
- [pyhwpx Cookbook](https://wikidocs.net/book/8956) - 한글 자동화 실습 가이드
- [한컴디벨로퍼 포럼](https://forum.developer.hancom.com/c/hwp-automation/52) - 한글 오토메이션 Q&A
- [hwp-mcp GitHub](https://github.com/jkf87/hwp-mcp) - HWP MCP 서버 구현 참고

### HWPUNIT 단위 변환

| 단위 | HWPUNIT 값 | 변환 공식 |
|------|-----------|----------|
| 1 pt | 100 | `hwpunit / 100 = pt` |
| 1 inch | 7,200 | `hwpunit / 7200 = inch` |
| 1 cm | 2,834.6 | `hwpunit / 2834.6 = cm` |
| 1 mm | 283.46 | `hwpunit / 283.46 = mm` |

---

## 10. 주의사항

1. **보안 모듈 등록 필수**
   ```python
   hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')
   ```

2. **표 내부에서만 동작하는 액션**
   - `TableRightCell`, `TableLowerCell` 등은 캐럿이 표 안에 있어야 동작

3. **셀 블록 필요한 액션**
   - `TableMergeCell`은 먼저 `TableCellBlock`으로 셀 선택 필요

4. **WSL 환경에서 실행**
   ```bash
   cmd.exe /c "python C:\\win32hwp\\script.py"
   ```

5. **필드 방식 권장 상황**
   - 복잡한 서식의 템플릿 문서
   - 반복적인 데이터 입력 작업
   - 데이터프레임 연동
