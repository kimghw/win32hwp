# 한글 OLE Automation (Part 07)

## HAction (계속)

| Item Name | Description |
|-----------|-------------|
| GetDefault(Method) | (이전 파트에서 계속)<br>object : HSet Object |
| Execute(Method) | **Description**<br>지정된 액션이름과 LPDISPATCH로 HSet을 받아 액션을 수행 한다.<br>**Declaration**<br>`BOOL Execute(BSTR actname, LPDISPATCH object)`<br>**Parameters**<br>actname : 액션 이름<br>object : HSet Object |
| PopupDialog(Method) | **Description**<br>지정된 액션이름과 LPDISPATCH로 HSet을 받아 액션의 다이얼로그를 생성한다.<br>**Declaration**<br>`BOOL PopupDialog(BSTR actname, LPDISPATCH object)`<br>**Parameters**<br>actname : 액션 이름<br>object : HSet Object |
| Run(Method) | **Description**<br>GetDefault, PopupDialog, Execute를 동시에 수행하도록 한다.<br>**Declaration**<br>`BOOL Run(BSTR actname)`<br>**Parameters**<br>actname : 액션 이름 |

---

## HArray

**ParameterArray의 데이터 집합**

| Item Name | Description |
|-----------|-------------|
| Count(Property) | **Description**<br>Array에 크기를 지정하거나 얻을 수 있다. |
| Item(Property) | **Description**<br>Array에 지정된 index에 해당하는 값을 VARIANT로 지정하거나 얻을 수 있다. |

---

## IXHwpMessageBox

**OLE Automation standard object**
(메세지박스 MessageBox 개체 - 사용 편이를 위해 제공됨)

| Item Name | Description |
|-----------|-------------|
| Application(Property) | **Description**<br>최상위 오브젝트를 얻어옴(IHwpObject - 읽기 전용) |
| String(Property) | **Description**<br>메시지 박스에 넣을 문자열 |
| Flag(Property) | **Description**<br>메시지 박스에 사용할 Flag |
| Result(Property) | **Description**<br>메시지 박스의 리턴값 |
| DoModal(Method) | **Description**<br>메시지 박스 보이기<br>**Declaration**<br>`void DoModal(void)` |

---

## IDHwpAction

**한/글에서 특정 기능을 수행하기 위한 액션 오브젝트**
(호환 유지를 위해 제공됨 - HwpCtrl에서 사용되는 DHwpAction과 동일)

| Item Name | Description |
|-----------|-------------|
| ActID(Property) | **Description**<br>액션 ID를 나타낸다. 읽기 전용. |
| SetID(Property) | **Description**<br>액션이 사용하는 parameter set ID를 나타낸다. 읽기 전용. |
| GetDefault(Method) | **Description**<br>현재 상태에 따라 액션 실행에 필요한 인수를 구한다.<br>**Declaration**<br>`void GetDefault(LPDISPATCH param)`<br>**Parameters**<br>param : 인수를 저장할 Parameter Set |
| CreateSet(Method) | **Description**<br>액션과 대응하는 Parameter Set을 생성한다.<br>**Declaration**<br>`LPDISPATCH CreateSet(void)`<br>**Parameters**<br>return : ParameterSet Object (IDHwpParameterSet) |
| Execute(Method) | **Description**<br>지정한 인수로 액션을 실행한다.<br>**Declaration**<br>`void Execute(LPDISPATCH param)`<br>**Parameters**<br>param : 액션의 실행을 제어할 인수, ParameterSet의 종류와 아이템의 의미는 액션이 정의한 바에 따라 다르다.(IDHwpParameterSet) |
| PopupDialog(Method) | **Description**<br>액션의 대화상자를 띄운다.<br>**Declaration**<br>`long PopupDialog(LPDISPATCH param)`<br>**Parameters**<br>param : 여기에 지정된 아이템의 값에 따라 대화상자의 컨트롤의 초기값이 결정되고, 대화상자가 닫힌 후에는 사용자가 지정한 값들이 담겨 돌아온다.<br>return : 액션이 정의하기에 따라 다르지만, 일반적으로 modal dialog result를 리턴한다. |
| Run(Method) | **Description**<br>액션을 실행한다.<br>**Declaration**<br>`void Run(void)`<br>**Remarks**<br>CreateSet, GetDefault, PopupDialog, Execute를 차례대로 부른 것과 같다. |

### PopupDialog Return 값

| ID | 값 | 설명 |
|----|---|------|
| hwpOK | IDOK | 다이얼로그 박스의 확인버튼을 눌렀을 경우 리턴 되는 값 |
| hwpCancel | IDCANCEL | 다이얼로그 박스의 취소버튼을 눌렀을 경우 리턴 되는 값 |
| hwpError | -1 | 실행시 에러가 발생 하였을 경우 리턴 되는 값 |

---

## IDHwpParameterSet

**오브젝트 또는 액션의 실행에 필요한 정보를 주고 받을 수 있도록 하기 위한 오브젝트**
(호환 유지를 위해 제공됨 - HwpCtrl에서 사용되는 DHwpParameterSet과 동일)

| Item Name | Description |
|-----------|-------------|
| Count(Property) | **Description**<br>현재 존재하는 아이템의 개수를 나타낸다. (읽기 전용) |
| IsSet(Property) | **Description**<br>Parameter Set인지 여부를 나타낸다. (읽기 전용)<br>**Remarks**<br>임의의 IDispatch 포인터로부터 Parameter set / parameter array를 구분하기 위해 동일한 이름의 Property를 가지고 종류에 따라 TRUE/FALSE를 돌려준다. (Parameter set은 TRUE를 리턴한다.) |
| GetSetID(Property) | **Description**<br>Parameter Set의 ID를 나타낸다. (읽기 전용) |
| Clone(Method) | **Description**<br>동일한 데이터를 가진 Parameter Set을 복사하여 리턴한다.<br>**Declaration**<br>`LPDISPATCH Clone(void)`<br>**Parameters**<br>return : ParameterSet을 리턴한다. (IDHwpParameterSet) |
| CreateItemArray(Method) | **Description**<br>아이템으로 Parameter Array 타입의 배열을 생성한다.<br>**Declaration**<br>`LPDISPATCH CreateItemArray(BSTR itemid, long count)`<br>**Parameters**<br>itemid : 아이템 ID<br>count : 생성할 배열의 초기 크기<br>return : 생성된 parameter array 오브젝트 (IDHwpParameterArray)<br>**Remarks**<br>동일한 ID를 가진 기존의 아이템은 삭제된다. |
| CreateItemSet(Method) | **Description**<br>아이템으로 Parameter Set을 생성한다.<br>**Declaration**<br>`LPDISPATCH CreateItemSet(BSTR itemid, BSTR setid)`<br>**Parameters**<br>itemid : 아이템 ID<br>setid : 생성할 Parameter Set ID<br>return : 생성된 서브 parameter Set 오브젝트 (IDHwpParameterSet)<br>**Remarks**<br>ParameterSet 내부에 아이템으로 또 다른 Parameter Set을 가지는 서브셋의 개념이다. |
| GetIntersection(Method) | **Description**<br>두 Parameter Set에 공통적으로 존재하고, 값도 동일한 아이템만으로 구성된 intersection Set을 구한다.<br>**Declaration**<br>`void GetIntersection(LPDISPATCH srcset)`<br>**Parameters**<br>srcset : this와 srcset의 intersection이 this에 저장된다. |
| IsEquivalent(Method) | **Description**<br>두 Parameter Set의 내용이 동일한 값을 가지고 있는지 검사한다.<br>**Declaration**<br>`BOOL IsEquivalent(LPDISPATCH srcset)`<br>**Parameters**<br>srcset : this와 srcset의 비교한 결과를 리턴한다.<br>return : 동일하면 TRUE, 다르면 FALSE |
| Item(Method) | **Description**<br>지정한 아이템의 값을 리턴한다.<br>**Declaration**<br>`Variant Item(BSTR itemid)`<br>**Parameters**<br>itemid : 아이템 ID<br>return : 아이템의 값<br>**Remarks**<br>만약 지정한 아이템이 존재하지 않으면 아이템의 포맷에 따라 0 또는 빈 문자열을 리턴한다. |
| ItemExist(Method) | **Description**<br>지정한 아이템이 존재하는지 검사한다.<br>**Declaration**<br>`BOOL ItemExist(BSTR itemid)`<br>**Parameters**<br>itemid : 아이템 ID<br>return : 존재하면 TRUE, 존재하지 않으면 FALSE |
| Merge(Method) | **Description**<br>두 Parameter Set의 내용을 병합한다.<br>**Declaration**<br>`void Merge(LPDISPATCH srcset)`<br>**Parameters**<br>srcset : this와 srcset이 병합되어 this에 저장된다.<br>**Remarks**<br>결과는 "this의 모든 아이템 + srcset에만 존재하는 아이템"이다. |
| RemoveAll(Method) | **Description**<br>Parameter Set을 초기화 한다.<br>**Declaration**<br>`void RemoveAll(BSTR setid)`<br>**Parameters**<br>setid : 새로 적용할 Set ID<br>**Remarks**<br>이미 존재하는 Parameter Set 오브젝트를 이용해 새로운 타입의 Parameter Set으로 초기화하여 재사용하는 목적에 사용된다. |
| RemoveItem(Method) | **Description**<br>지정한 아이템을 삭제한다.<br>**Declaration**<br>`void RemoveItem(BSTR itemid)`<br>**Parameters**<br>itemid : 아이템 ID |
| SetItem(Method) | **Description**<br>지정한 아이템의 값을 설정한다.<br>**Declaration**<br>`void SetItem(BSTR itemid, Variant value)`<br>**Parameters**<br>itemid : 아이템 ID<br>value : 설정할 값<br>**Remarks**<br>이미 동일한 ID의 아이템이 존재하면 지정한 값으로 바뀌고, 존재하지 않으면 아이템이 생성된다. |

---

## IDHwpParameterArray

**Parameter Set의 아이템으로 배열을 표현하는데 사용된다.**
일반적인 Method의 독립적인 인수로 사용되는 일은 없고, Parameter Set의 아이템으로만 사용된다.
(호환 유지를 위해 제공됨 - HwpCtrl에서 사용되는 DHwpParameterArray와 동일)

| Item Name | Description |
|-----------|-------------|
| Count(Property) | **Description**<br>배열의 크기를 나타낸다.<br>**Remarks**<br>배열의 크기는 runtime에 dynamic하게 조절 할 수 있다. |
| IsSet(Property) | **Description**<br>Parameter Set인지 여부를 나타낸다.(읽기 전용)<br>**Remarks**<br>임의의 IDispatch 포인터로부터 Parameter Set / Parameter Array를 구분하기 위해 동일한 이름의 Property를 가지고 종류에 따라 TRUE/FALSE를 돌려준다.(ParameterArray는 FALSE를 리턴한다.) |
| Clone(Method) | **Description**<br>동일한 크기와 데이터를 갖는 ParameterArray 개체를 복사하여 돌려준다<br>**Declaration**<br>`LPDISPATCH Clone(void)` |
| Copy(Method) | **Description**<br>배열을 복사한다.<br>**Declaration**<br>`void Copy(LPDISPATCH srcarray)`<br>**Parameters**<br>srcarray : srcarray의 내용이 그대로 this로 복사된다. |
| Item(Method) | **Description**<br>지정한 원소의 값을 리턴한다.<br>**Declaration**<br>`Variant Item(long index)`<br>**Parameters**<br>index : 원소의 인덱스. 1부터 시작한다. |
| SetItem(Method) | **Description**<br>지정한 원소의 값을 설정한다.<br>**Declaration**<br>`void SetItem(long index, Variant value)`<br>**Parameters**<br>index : 원소의 인덱스, 1부터 시작<br>value : 원소의 값 |

---

## IDHwpCtrlCode

**문서 내부의 표, 각주 등의 컨트롤(특수 문자 포함)를 나타내는 오브젝트이다.**
(호환 유지를 위해 제공됨 - HwpCtrl에서 사용되는 DHwpCtrlCode와 동일)

| Item Name | Description |
|-----------|-------------|
| CtrlCh(Property) | **Description**<br>컨트롤 문자. (읽기전용)<br>**Remarks**<br>일반적으로 컨트롤 ID를 사용해 컨트롤의 종류를 판별하지만, 이보다 더 포괄적인 범주를 나타내는 컨트롤 문자로 판별할 수도 있다. 예를 들어 각주와 미주는 ID는 다르지만, 컨트롤 문자는 17로 동일하다. 컨트롤 문자는 1-31 사이의 값을 사용한다. |
| CtrlID(Property) | **Description**<br>컨트롤 ID. (읽기 전용)<br>**Remarks**<br>컨트롤 ID는 컨트롤의 종류를 나타내기 위해 할당된 ID로서, 최대 4개의 문자로 구성된 문자열이다. 예를 들어 표는 "tbl", 각주는 "fn"이다. |
| HasList(Property) | **Description**<br>글상자를 지원하는지의 여부(읽기 전용) |
| Next(Property) | **Description**<br>다음 컨트롤.(읽기 전용)<br>**Remarks**<br>문서 중의 모든 컨트롤(표, 그림등의 특수 문자들)은 linked list로 서로 연결되어 있는데, list중 다음 컨트롤을 나타낸다. |
| Prev(Property) | **Description**<br>앞 컨트롤.(읽기 전용)<br>**Remarks**<br>문서 중의 모든 컨트롤(표, 그림등의 특수 문자들)은 linked list로 서로 연결되어 있는데, list중 다음 컨트롤을 나타낸다. |
| Properties(Property) | **Description**<br>컨트롤의 속성을 나타낸다.<br>**Parameters**<br>모든 컨트롤은 대응하는 parameter set으로 속성을 읽고 쓸 수 있다. |
| UserDesc(Property) | **Description**<br>컨트롤의 종류를 사용자에게 보여줄 수 있는 localize된 문자열로 나타낸다. (읽기 전용) |
| GetAnchorPos(Method) | **Description**<br>컨트롤의 anchor의 위치를 리턴한다.<br>**Declaration**<br>`LPDISPATCH GetAnchorPos(long type)`<br>**Parameters**<br>type : 기준 위치<br>value : 성공했을 경우 LispParaPos Parameter Set이 리턴 된다. 실패했을 경우 NULL이 리턴된다. |

### 컨트롤 문자 (CtrlCh) 표

| Ch | 설명 |
|----|------|
| 1 | 예약 |
| 2 | 구역/단 정의 |
| 3 | 필드 시작 |
| 4 | 필드 끝 |
| 5 | 예약 |
| 6 | 예약 |
| 7 | 예약 |
| 8 | 예약 |
| 9 | 탭 |
| 10 | 강제 줄 나눔 |
| 11 | 그리기 개체 / 표 |
| 12 | 예약 |
| 13 | 문단 나누기 |
| 14 | 예약 |
| 15 | 주석 |
| 16 | 머리말 / 꼬리말 |
| 17 | 각주 / 미주 |
| 18 | 자동 번호 |
| 19 | 예약 |
| 20 | 예약 |
| 21 | 쪽바뀜 |
| 22 | 책갈피 / 찾아보기 표시 |
| 23 | 덧말 / 글짜 겹침... |
| 24 | 하이픈 |
| 25 | 예약 |
| 26 | 예약 |
| 27 | 예약 |
| 28 | 예약 |
| 29 | 예약 |
| 30 | 묶음 빈칸 |
| 31 | 고정 폭 빈칸 |

### 컨트롤 ID (CtrlID) 표

| ID | Property Set | Initialization Set | 설명 |
|----|--------------|-------------------|------|
| cold | ColDef | ColDef | 단 |
| secd | SecDef | SecDef | 구역 |
| fn | FootnoteShape | FootnoteShape | 각주 |
| en | FootnoteShape | FootnoteShape | 미주 |
| tbl | Table | TableCreation | 표 |
| eqed | EqEdit | EqEdit | 수식 |
| atno | AutoNum | AutoNum | 번호넣기 |
| nwno | AutoNum | AutoNum | 새번호로 |
| pgct | PageNumCtrl | PageNumCtrl | 페이지 번호 제어 (97의 홀수쪽에서 시작) |
| pghd | PageHiding | PageHiding | 감추기 |
| pgnp | PageNumPos | PageNumPos | 쪽번호 위치 |
| head | HeaderFooter | HeaderFooter | 머리말 |
| foot | HeaderFooter | HeaderFooter | 꼬리말 |
| %dte | FieldCtrl | FieldCtrl | 현재의 날짜/시간 필드 |
| %ddt | FieldCtrl | FieldCtrl | 파일 작성 날짜/시간 필드 |
| %pat | FieldCtrl | FieldCtrl | 문서 경로 필드 |
| %bmk | FieldCtrl | FieldCtrl | 블럭 책갈피 |
| %mmg | FieldCtrl | FieldCtrl | 메일 머지 |
| %xrf | FieldCtrl | FieldCtrl | 상호 참조 |
| %fmu | FieldCtrl | FieldCtrl | 계산식 |
| %clk | FieldCtrl | FieldCtrl | 누름틀 |
| %smr | FieldCtrl | FieldCtrl | 문서 요약 정보 필드 |
| %usr | FieldCtrl | FieldCtrl | 사용자 정보 필드 |
| %hlk | FieldCtrl | FieldCtrl | 하이퍼링크 |
| %sig | RevisionDef | RevisionDef | 교정부호(띄움표) |
| %%*d | RevisionDef | RevisionDef | 교정부호(지움표) |
| %%*a | RevisionDef | RevisionDef | 교정부호(붙임표) |
| %%*C | RevisionDef | RevisionDef | 교정부호(뺌표) |
| %%*S | RevisionDef | RevisionDef | 교정부호(톱니표) |
| %%*T | RevisionDef | RevisionDef | 교정부호(생각표) |
| %%*P | RevisionDef | RevisionDef | 교정부호(칭찬표) |
| %%*L | RevisionDef | RevisionDef | 교정부호(줄표) |
| %%*c | RevisionDef | RevisionDef | 교정부호(고침표) |
| %%*h | HyperLink | HyperLink | 교정부호(자료연결) |
| %%*A | RevisionDef | RevisionDef | 교정부호(줄붙임표) |
| %%*i | RevisionDef | RevisionDef | 교정부호(줄이음표) |
| %%*t | RevisionDef | RevisionDef | 교정부호(줄서로바꿈표) |
| %%*r | RevisionDef | RevisionDef | 교정부호(오른자리옮김표) |
| %%*l | RevisionDef | RevisionDef | 교정부호(왼자리옮김표) |
| %%*n | RevisionDef | RevisionDef | 교정부호(자리바꿈표) |
| %spl | RevisionDef | RevisionDef | 교정부호(나눔표) |
| %%mr | RevisionDef | RevisionDef | 교정부호(메모고침표) |
| %%me | RevisionDef | RevisionDef | 메모 |
| bokm | TextCtrl | TextCtrl | 책갈피 |
| idxm | IndexMark | IndexMark | 찾아보기 |
| tdut | Dutmal | Dutmal | 덧말 |
| tcmt | 없음 | 없음 | 주석 |
| $con | ShapeObject | ShapeObject | 여러 개체를 묶은 개체 |
| $lin | ShapeObject | ShapeObject | 직선 |
| $rec | ShapeObject | ShapeObject | 사각형 |
| $ell | ShapeObject | ShapeObject | 원형 |
| $arc | ShapeObject | ShapeObject | 호 |
| $pol | ShapeObject | ShapeObject | 다각형 |
| $cur | ShapeObject | ShapeObject | 곡선 |
| $pic | ShapeObject | ShapeObject | 그림 |
| form | ShapeObject | ShapeObject | 양식 개체 |
| +pbt | ShapeObject | ShapeObject | 명령 단추 |
| +rbt | ShapeObject | ShapeObject | 라디오 단추 |
| +cbt | ShapeObject | ShapeObject | 선택 상자 |
| +cob | ShapeObject | ShapeObject | 콤보 상자 |
| +edt | ShapeObject | ShapeObject | 입력 상자 |
| $ole | ShapeObject | OleCreation | OLE개체 |

**참고:**
- **Property Set** : Ctrl.Properties를 통해 액세스할 수 있는 속성 parameter set ID
- **Initialization Set** : HwpCtrl.InsertCtrl에 지정할 수 있는 initparam의 parameter set ID

### GetAnchorPos type 값

| 값 | 설명 | 비고 |
|---|------|-----|
| 0 | 바로 상위 리스트에서의 anchor position | default |
| 1 | 탑레벨 리스트에서의 anchor position | |
| 2 | 루트 리스트에서의 anchor position | |

---

## 관련 문서 참조

- **Action Object** : Action Object 매뉴얼 참고(ActionObject.hwp)
- **ParameterSet Object** : ParameterSet Object 매뉴얼 참고(ParameterSetObject.hwp)
- **Add-On Object** : 메뉴, 툴바 제어 기능, 사용자 정의 액션 기능
