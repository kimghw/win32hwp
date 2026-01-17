# ParameterSet Table (Part 03)

### 16) CodeTable : 문자표

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Text | PIT_BSTR | | 삽입될 스트링 |

---

### 17) ColDef : 단 정의 속성

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Type | PIT_UI1 | | 단 종류 :<br>0 = 보통 다단, 1 = 배분 다단, 2 = 평행 다단 |
| Count | PIT_UI1 | | 단 개수. 1-255까지. |
| SameSize | PIT_UI1 | | 단의 너비를 같도록 할지 여부 :<br>0 = 단 너비 각자 지정, 1 = 단 너비 동일 |
| SameGap | PIT_I4 | | 단 사이 간격(HWPUNIT)<br>SAME_SIZE가 1일 때만 사용된다. |
| WidthGap | PIT_ARRAY | PIT_UI2 | 각 단의 너비와 간격(HWPUNIT)<br>col*2 = 단의 폭, col*2 + 1 = 단 사이 간격.<br>영역 전체의 폭을 Column ratio base (=32768)로 보았을 때의 비율로 환산한다.<br>SameSize가 0일 때만 사용된다. 배열의 아이템의 개수는 Count*2-1과 같아야 한다. |
| Layout | PIT_UI1 | | 단 방향 지정 :<br>0 = 왼쪽부터, 1 = 오른쪽부터, 2 = 맞쪽 |
| LineType | PIT_UI1 | | 선 종류 |
| LineWidth | PIT_UI1 | | 선 종류 |
| LineColor | PIT_UI4 | | 선 색깔. (COLORREF)<br>RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR) |
| ApplyTo | PIT_UI1 | | 적용범위 (아래 표 참조) |
| ApplyClass | PIT_UI1 | | 적용범위의 분류. 아래 값의 조합이다. |

#### ApplyTo 값

| 값 | 설명 |
|---|------|
| 0 | 선택된 다단 |
| 1 | 선택된 문자열 |
| 2 | 현재 다단 |
| 3 | 개체 전체 |
| 4 | 선택된 셀 |
| 5 | 현재 구역 |
| 6 | 문서 전체 |
| 7 | 현재 셀 |
| 8 | 새 쪽으로 |
| 9 | 새 다단으로 |
| 10 | 모든 셀 |

#### ApplyClass 값 (비트 조합)

| 값 | 설명 |
|---|------|
| 0x0001 | 선택된 다단 |
| 0x0002 | 선택된 문자열 |
| 0x0004 | 현재 다단 |
| 0x0008 | 개체 전체 |
| 0x0010 | 선택된 셀 |
| 0x0020 | 현재 구역 |
| 0x0040 | 문서전체 |
| 0x0080 | 현재 셀 |
| 0x0100 | 새 쪽으로 |
| 0x0200 | 새 다단으로 |
| 0x0400 | 모든 셀 |

---

### 18) ConvertCase : 대/소문자 변환

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Type | PMT_UINT | | 공통사용. |

---

### 19) ConvertFullHalf : 전/반각 변환

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Type | PMT_UINT | | 전각으로 변경할지, 반각으로 변경할지 여부 : 0 = 반각, 1 = 전각 |
| Number | PMT_UINT | | 변경 대상에 숫자를 추가할지 여부 : 0 = off, 1 = on |
| Alpha | PMT_UINT | | 변경 대상에 영문자를 추가할지 여부 : 0 = off, 1 = on |
| Symbol | PMT_UINT | | 변경 대상에 기호를 추가할지 여부 : 0 = off, 1 = on |
| Gata | PMT_UINT | | 변경 대상에 가타가나를 추가할지 여부 : 0 = off, 1 = on |
| HGJamo | PMT_UINT | | 변경 대상에 한글자모를 추가할지 여부 : 0 = off, 1 = on |

---

### 20) ConvertHiraToGata : 히라가나/가타가나 변환

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Type | PMT_UINT | | 히라가나로 변경할지, 가타가나로 변경할지 여부<br>0 = 가타가나로 변경, 1 = 히라가나로 변경 |

---

### 21) ConvertJianFan : 간/번체 변환

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Type | PMT_UINT | | 간체로 변경할지, 번체로 변경할지 여부<br>0 = 번체로 변경, 1 = 간체로 변경 |

---

### 22) ConvertToHangul : 한자, 일어, 구결을 한글로

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Type | PMT_UINT | | 변경할 유형. 확장 또는 변환 액션 간 파라메터셋의 호환을 위해 선언됨. |
| Hanja | PMT_UINT | | 한자를 한글로 변경할지 여부 : 0 = off, 1 = on |
| Hira | PMT_UINT | | 히라가나를 한글로 변경할지 여부 : 0 = off, 1 = on<br>확장 또는 변환 액션 간 파라메터셋의 호환을 위해 선언됨. |
| Gata | PMT_UINT | | 가타가나를 한글로 변경할지 여부 : 0 = off, 1 = on<br>확장 또는 변환 액션 간 파라메터셋의 호환을 위해 선언됨. |
| Gu | PMT_UINT | | 구결을 한글로 변경할지 여부 : 0 = off, 1 = on |
| HanjaHangul | PMT_UINT | | 한자를 漢字(한글)로 변경할지 여부 : 0 = off, 1 = on |

---

### 23) CtrlData : 컨트롤 데이터

컨트롤 데이터. 컨트롤에 임의로 설정할 수 있는 데이터 셋. 기본적으로 서브셋을 사용하는 것을 원칙으로 한다.

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Name | PIT_BSTR | | 사용자가 지정한 컨트롤의 이름. |

---

### 24) DeleteCtrls : 조판 부호 컨트롤 지우기

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| DeleteCtrlType | PIT_ARRAY | PIT_UI | 지울 개체 목록 |
