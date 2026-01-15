# 한글 OLE Automation (Part 06)

## codepage 종류

| 코드 | 설명 |
|------|------|
| ks | 한글 KS 완성형 |
| kssm | 한글 조합형 |
| sjis | 일본 |
| gb | 중국 간체 |
| big5 | 중국 번체 |
| acp | Active Codepage 현재 시스템의 코드 페이지 |

---

## IXHwpDocument 메서드 (계속)

### SendBrowser (Method)

| 항목 | 내용 |
|------|------|
| **Description** | 문서를 브라우저로 내보내기 기능 |
| **Declaration** | `BOOL SendBrowser(void)` |

### SetActive_XHwpDocument (Method)

| 항목 | 내용 |
|------|------|
| **Description** | 문서를 활성화 상태로 하기 |
| **Declaration** | `void SetActive_XHwpDocument(void)` |

### Clear (Method)

| 항목 | 내용 |
|------|------|
| **Description** | 문서의 내용을 닫고 빈문서 편집 상태로 돌아간다. |
| **Declaration** | `void Clear(Variant option)` |
| **Parameters** | option : 편집중인 문서의 내용에 대한 처리 방법. 생략하면 hwpAskSave가 선택된다. |
| **Remark** | format, arg에 대해서는 Open 참조. hwpSaveIfDirty, hwpSave가 지정된 경우 현재 문서 경로가 지정되어 있지 않으면 "새이름으로 저장" 대화상자를 띄워 사용자에게 경로를 묻는다. |

#### option 값

| ID | 값 | 설명 |
|----|---|------|
| hwpAskSave | 0 | 문서의 내용이 변경되었을 때 사용자에게 저장할지 묻는 대화상자를 띄운다. |
| hwpDiscard | 1 | 문서 내용을 버린다. |
| hwpSaveIfDirty | 2 | 문서가 변경된 경우 저장 한다. |
| hwpSave | 3 | 무조건 저장한다. |

---

## IXHwpFormPushButtons

**IXHwpFormPushButton 오브젝트를 관리하는 오브젝트**
(양식개체 PushButton을 관리하는 Collection 개체 - 사용 편이를 위해 제공됨)

| Item Name | Description |
|-----------|-------------|
| Application(Property) | **Description**<br>최상위 오브젝트를 얻어옴(IHwpObject) |
| Item(Property) | **Description**<br>지정한 원소의 양식개체 PushButton 오브젝트를 얻어온다.<br>**Declaration**<br>`LPDISPATCH Item(long index)`<br>**Parameters**<br>index : 원소의 인덱스<br>return : 원소의 값(IXHwpFormPushButton) |
| Count(Property) | **Description**<br>원소의 총개수 |
| ItemFromName(Property) | **Description**<br>양식 개체 PushButton의 이름으로 원소를 찾는다.<br>**Declaration**<br>`LPDISPATCH ItemFormName(BSTR name)`<br>**Parameters**<br>name : 양식 개체 PushButton의 이름<br>return : 양식 개체 PushButton Object(IXHwpFormPushButton) |

---

## IXHwpFormPushButton

**양식 개체 푸쉬 버튼 오브젝트**
(PushButton 개체 - 사용 편이를 위해 제공됨)

| Item Name | Description |
|-----------|-------------|
| Application(Property) | **Description**<br>최상위 오브젝트를 얻어옴(IHwpObject) |
| Caption(Property) | 캡션 |
| Name(Property) | 이름 |
| ForeColor(Property) | 전경색 |
| BackColor(Property) | 배경색 |
| GroupName(Property) | 그룹 이름 |
| TabStop(Property) | Tab Stop |
| TabOrder(Property) | 탭 순서 |
| Width(Property) | 너비 |
| Height(Property) | 높이 |
| Left(Property) | 왼쪽 좌표 |
| Top(Property) | 위쪽 좌표 |
| CharShapeID(Property) | 그리기 개체 Control ID |
| FollowContext(Property) | 주위의 글자 속성을 따를지의 여부 |
| AutoSize(Property) | 글자 크기에 맞게 개체 크기가 바뀜 |
| BorderType(Property) | 테두리 타입 |
| DrawFrame(Property) | 틀을 그릴지의 여부 |
| Enabled(Property) | 활성, 비활성의 여부 |
| WordWrap(Property) | 자동 줄 바꿈 |

---

## IXHwpFormCheckButtons

**IXHwpFormCheckButton 오브젝트를 관리하는 오브젝트**
(양식개체 CheckButton을 관리하는 Collection 개체 - 사용 편이를 위해 제공됨)

| Item Name | Description |
|-----------|-------------|
| Application(Property) | **Description**<br>최상위 오브젝트를 얻어옴(IHwpObject) |
| Item(Property) | **Description**<br>지정한 원소의 양식개체 CheckButton 오브젝트를 얻어온다.<br>**Declaration**<br>`LPDISPATCH Item(long index)`<br>**Parameters**<br>index : 원소의 인덱스<br>return : 원소의 값(IXHwpFormCheckButton) |
| Count(Property) | **Description**<br>원소의 총개수 |
| ItemFromName(Property) | **Description**<br>양식 개체 CheckButton의 이름으로 원소를 찾는다.<br>**Declaration**<br>`LPDISPATCH ItemFormName(BSTR name)`<br>**Parameters**<br>name : 양식 개체 CheckButton의 이름<br>return : 양식 개체 CheckButton Object(IXHwpFormCheckButton) |

---

## IXHwpFormCheckButton

**양식 개체 체크 버튼 오브젝트**
(CheckButton 개체 - 사용 편이를 위해 제공됨)

| Item Name | Description |
|-----------|-------------|
| Application(Property) | **Description**<br>최상위 오브젝트를 얻어옴(IHwpObject) |
| Caption(Property) | 캡션 |
| Name(Property) | 이름 |
| ForeColor(Property) | 전경색 |
| BackColor(Property) | 배경색 |
| GroupName(Property) | 그룹 이름 |
| TabStop(Property) | Tap Stop |
| TabOrder(Property) | 탭 순서 |
| Width(Property) | 너비 |
| Height(Property) | 높이 |
| Left(Property) | 왼쪽 좌표 |
| Top(Property) | 위쪽 좌표 |
| CharShapeID(Property) | 그리기 개체 Control ID |
| FollowContext(Property) | 주위의 글자 속성을 따를지의 여부 |
| AutoSize(Property) | 글자 크기에 맞게 개체 크기가 바뀜 |
| BackStyle(Property) | 배경 투명도 |
| BorderType(Property) | 테두리 타입 |
| DrawFrame(Property) | 틀을 그릴지의 여부 |
| Enabled(Property) | 활성, 비활성의 여부 |
| TriState(Property) | 체크 상태 옵션 |
| Value(Property) | 값 |
| WordWrap(Property) | 자동 줄 바꿈 |

---

## IXHwpFormRadioButtons

**IXHwpFormRadioButton 오브젝트를 관리하는 오브젝트**
(양식개체 RadioButton을 관리하는 Collection 개체 - 사용 편이를 위해 제공됨)

| Item Name | Description |
|-----------|-------------|
| Application(Property) | **Description**<br>최상위 오브젝트를 얻어옴(IHwpObject) |
| Item(Property) | **Description**<br>지정한 원소의 양식개체 RadioButton 오브젝트를 얻어온다.<br>**Declaration**<br>`LPDISPATCH Item(long index)`<br>**Parameters**<br>index : 원소의 인덱스<br>return : 원소의 값(IXHwpFormRadioButton) |
| Count(Property) | **Description**<br>원소의 총개수 |
| ItemFromName(Property) | **Description**<br>양식 개체 RadioButton의 이름으로 원소를 찾는다.<br>**Declaration**<br>`LPDISPATCH ItemFormName(BSTR name)`<br>**Parameters**<br>name : 양식 개체 RadioButton의 이름<br>return : 양식 개체 RadioButton Object(IXHwpFormCheckButton) |

---

## HwpFormRadioButton

**양식 개체 라디오 버튼 오브젝트**
(RadioButton 개체 - 사용 편이를 위해 제공됨)

| Item Name | Description |
|-----------|-------------|
| Application(Property) | **Description**<br>최상위 오브젝트를 얻어옴(IHwpObject) |
| Name(Property) | 이름 |
| ForeColor(Property) | 전경색 |
| BackColor(Property) | 배경색 |
| GroupName(Property) | 그룹 이름 |
| TabStop(Property) | Tap Stop |
| TabOrder(Property) | 탭 순서 |
| Width(Property) | 너비 |
| Height(Property) | 높이 |
| Left(Property) | 왼쪽 좌표 |
| Top(Property) | 위쪽 좌표 |
| CharShapeID(Property) | 그리기 개체 Control ID |
| FollowContext(Property) | 주위의 글자 속성을 따를지의 여부 |
| AutoSize(Property) | 글자 크기에 맞게 개체 크기가 바뀜 |
| BackStyle(Property) | 배경 투명도 |
| BorderType(Property) | 테두리 타입 |
| Caption(Property) | 캡션 |
| DrawFrame(Property) | 틀을 그릴지의 여부 |
| Enabled(Property) | 활성, 비활성의 여부 |
| RadioGroupName(Property) | 라디오 그룹 이름 |
| TriState(Property) | 체크 상태 옵션 |
| Value(Property) | 값 |
| WordWrap(Property) | 자동 줄 바꿈 |

---

## IXHwpFormComboBoxs

**IXHwpFormComboBox 오브젝트를 관리하는 오브젝트**
(양식개체 ComboBox을 관리하는 Collection 개체 - 사용 편이를 위해 제공됨)

| Item Name | Description |
|-----------|-------------|
| Application(Property) | **Description**<br>최상위 오브젝트를 얻어옴(IHwpObject) |
| Item(Property) | **Description**<br>지정한 원소의 양식개체 ComboBox 오브젝트를 얻어온다.<br>**Declaration**<br>`LPDISPATCH Item(long index)`<br>**Parameters**<br>index : 원소의 인덱스<br>return : 원소의 값(IXHwpFormComboBox) |
| Count(Property) | **Description**<br>원소의 총개수 |
| ItemFromName(Property) | **Description**<br>양식 개체 ComboBox의 이름으로 원소를 찾는다.<br>**Declaration**<br>`LPDISPATCH ItemFormName(BSTR name)`<br>**Parameters**<br>name : 양식 개체 ComboBox의 이름<br>return : 양식 개체 ComboBox Object(IXHwpFormComboBox) |

---

## IXHwpFormComboBox

**양식 개체 콤보 박스 오브젝트**
(ComboBox 개체 - 사용 편이를 위해 제공됨)

| Item Name | Description |
|-----------|-------------|
| Application(Property) | **Description**<br>최상위 오브젝트를 얻어옴(IHwpObject) |
| Name(Property) | 이름 |
| ForeColor(Property) | 전경색 |
| BackColor(Property) | 배경색 |
| GroupName(Property) | 그룹 이름 |
| TabStop(Property) | Tap Stop |
| TabOrder(Property) | 탭 순서 |
| Width(Property) | 너비 |
| Height(Property) | 높이 |
| Left(Property) | 왼쪽 좌표 |
| Top(Property) | 위쪽 좌표 |
| CharShapeID(Property) | 그리기 개체 Control ID |
| FollowContext(Property) | 주위의 글자 속성을 따를지의 여부 |
| AutoSize(Property) | 글자 크기에 맞게 개체 크기가 바뀜 |
| BorderType(Property) | 테두리 타입 |
| DrawFrame(Property) | 틀을 그릴지의 여부 |
| EditEnable(Property) | 에디트 상태 활성, 비활성의 여부 |
| Enabled(Property) | 활성, 비활성의 여부 |
| ListBoxRows(Property) | 리스트 박스 열 |
| ListBoxWidth(Property) | 리스트 박스 너비 |
| Text(Property) | 선택된 값 |
| WordWrap(Property) | 자동 줄 바꿈 |
| Count(Property) | 아이템 개수 |
| CurSel(Property) | 현재 선택된 인덱스 |
| LBText(Property) | 인덱스에 해당하는 값 |
| InsertString(Method) | **Description**<br>양식 개체 ComboBox에 문자열을 채워넣는다.<br>**Declaration**<br>`void InsertString(BSTR itemvalue, long index)`<br>**Parameters**<br>itemvalue : 리스트에 채워넣기 위한 값<br>index : 리스트의 특정 위치 |
| DeleteString(Method) | **Description**<br>리스트의 지정한 위치에 있는 값을 지운다.<br>**Declaration**<br>`void DeleteString(unsigned long index)`<br>**Parameters**<br>index : 리스트의 특정 위치 |
| FindStringExact(Method) | **Description**<br>지정한 문자열이 리스트에 있는지를 찾는다.<br>**Declaration**<br>`long FindStringExact(long index, BSTR itemvalue)`<br>**Parameters**<br>index : 리스트의 특정 위치<br>itemvalue : 리스트에서 찾기 위한 값 |
| ResetContent(Method) | **Description**<br>리스트 내용을 초기화 한다.<br>**Declaration**<br>`void ResetContent(void)` |

---

## IXHwpFormEdits

**IXHwpFormEdit 오브젝트를 관리하는 오브젝트**
(양식개체 Edit을 관리하는 Collection 개체 - 사용 편이를 위해 제공됨)

| Item Name | Description |
|-----------|-------------|
| Application(Property) | **Description**<br>최상위 오브젝트를 얻어옴(IHwpObject) |
| Item(Property) | **Description**<br>지정한 원소의 양식개체 Edit 오브젝트를 얻어온다.<br>**Declaration**<br>`LPDISPATCH Item(long index)`<br>**Parameters**<br>index : 원소의 인덱스<br>return : 원소의 값(IXHwpFormEdit) |
| Count(Property) | **Description**<br>원소의 총개수 |
| ItemFromName(Property) | **Description**<br>양식 개체 Edit의 이름으로 원소를 찾는다.<br>**Declaration**<br>`LPDISPATCH ItemFormName(BSTR name)`<br>**Parameters**<br>name : 양식 개체 Edit의 이름<br>return : 양식 개체 Edit Object(IXHwpFormEdit) |

---

## IXHwpFormEdit

**양식 개체 에디트 오브젝트**
(Edit 개체 - 사용 편이를 위해 제공됨)

| Item Name | Description |
|-----------|-------------|
| Application(Property) | **Description**<br>최상위 오브젝트를 얻어옴(IHwpObject) |
| Name(Property) | 이름 |
| ForeColor(Property) | 전경색 |
| BackColor(Property) | 배경색 |
| GroupName(Property) | 그룹 이름 |
| TabStop(Property) | Tap Stop |
| TabOrder(Property) | 탭 순서 |
| Width(Property) | 너비 |
| Heigh(Property) | 높이 |
| Left(Property) | 왼쪽 좌표 |
| Top(Property) | 위쪽 좌표 |
| CharShapeID(Property) | 그리기 개체 Control ID |
| FollowContext(Property) | 주위의 글자 속성을 따를지의 여부 |
| AutoSize(Property) | 글자 크기에 맞게 개체 크기가 바뀜 |
| BorderType(Property) | 테두리 타입 |
| DrawFrame(Property) | 틀을 그릴지의 여부 |
| Enabled(Property) | 활성, 비활성의 여부 |
| MaxLength(Property) | 에디트 가능한 총 길이 |
| MultiLine(Property) | 멀티 라인 지원 |
| Number(Property) | 숫자만 입력 가능 |
| PasswordChar(Property) | 패스워드 표시에 사용할 글자 |
| ReadOnly(Property) | 읽기만 가능 |
| ScrollBars(Property) | 스크롤바 표시 |
| TabKeyBehavior(Property) | 탭 키를 눌렀을 때 반응 |
| Text(Property) | 에디트 텍스트 |
| WordWrap(Property) | 자동 줄 바꿈 |
| LineCount(Property) | 에디트 라인 줄 수 |

---

## IXHwpWindows

**IXHwpWindow오브젝트를 관리하는 오브젝트**
(Window를 관리하는 Collection 개체 - 사용 편이를 위해 제공됨)

| Item Name | Description |
|-----------|-------------|
| Application(Property) | **Description**<br>최상위 오브젝트를 얻어옴(IHwpObject - 읽기 전용) |
| Active_XHwpWindow(Property) | **Description**<br>현재 활성화 상태인 윈도우 Object를 얻어온다.(IXHwpWindow) |
| Item(Property) | **Description**<br>지정한 원소의 윈도우 오브젝트를 얻어온다.<br>**Declaration**<br>`LPDISPATCH Item(long index)`<br>**Parameters**<br>index : 원소의 인덱스<br>return : 원소의 값(IXHwpWindow) |
| Count(Property) | **Description**<br>원소의 총개수 |
| Add(Method) | **Description**<br>윈도우를 하나 추가한다.(새창으로 열기 기능과 동일)<br>**Declaration**<br>`LPDISPATCH Add(void)`<br>**Parameters**<br>return : 추가된 윈도우 오브젝트(IXHwpWindow) |
| Close(Method) | **Description**<br>윈도우를 모두 닫는다.<br>**Declaration**<br>`BOOL Close(BOOL isDirty)`<br>**Parameters**<br>isDirty : TRUE이면 문서 내용이 변경된 경우 닫지 않는다. FALSE이면 문서 내용이 변경된 것과 상관없이 강제로 닫는다. |

---

## IXHwpWindow

**윈도우 오브젝트**
(Window 개체 - 사용 편이를 위해 제공됨)

| Item Name | Description |
|-----------|-------------|
| Application(Property) | **Description**<br>최상위 오브젝트를 얻어옴(IHwpObject - 읽기 전용) |
| XHwpDocuments(Property) | **Description**<br>도큐먼트 관리 오브젝트를 얻어옴(IXHwpDocuments - 읽기 전용) |
| XHwpTabs(Property) | **Description**<br>탭 관리 오브젝트를 얻어옴(IXHwpTabs - 읽기 전용) |
| Left(Property) | **Description**<br>윈도우의 좌측 위치 좌표를 설정/얻음 |
| Top(Property) | **Description**<br>윈도우의 맨위 위치 좌표를 설정/얻음 |
| Width(Property) | **Description**<br>윈도우의 너비를 설정/얻음 |
| Height(Property) | **Description**<br>윈도우의 높이를 설정/얻음 |
| Visible(Property) | **Description**<br>윈도우 보이기/보이지 않기 설정/얻음 |
| Close(Method) | **Description**<br>윈도우를 닫음<br>**Declaration**<br>`void Close(BOOL isDirty)`<br>**Parameters**<br>isDirty : TRUE이면 문서 내용이 변경된 경우 닫지 않는다. FALSE이면 문서 내용이 변경된 것과 상관없이 강제로 닫는다. |

---

## IXHwpTabs

**IXHwpTab오브젝트를 관리하는 오브젝트**
(Tab를 관리하는 Collection 개체 - 사용 편이를 위해 제공됨)

| Item Name | Description |
|-----------|-------------|
| Count(Property) | **Description**<br>윈도우에 열려있는 탭의 개수(읽기 전용) |
| Item(Property) | **Description**<br>지정한 원소의 탭 오브젝트를 얻어온다.<br>**Declaration**<br>`LPDISPATCH Item(long index)`<br>**Parameters**<br>index : 원소의 인덱스<br>return : 원소의 값(IXHwpTab) |
| Add(Method) | **Description**<br>지정한 원소의 탭 오브젝트를 추가한다.(문서를 새 탭으로 열기)<br>**Declaration**<br>`LPDISPATCH Add(void)`<br>**Parameters**<br>return : 추가된 원소의 값(IXHwpTab) |
| Close(Method) | **Description**<br>탭을 모두 닫는다.<br>**Declaration**<br>`BOOL Close(BOOL isDirty)`<br>**Parameters**<br>isDirty : TRUE이면 문서 내용이 변경된 경우 닫지 않는다. FALSE이면 문서 내용이 변경된 것과 상관없이 강제로 닫는다. |

---

## IXHwpTab

**탭 오브젝트**
(Tab 개체 - 사용 편이를 위해 제공됨)

| Item Name | Description |
|-----------|-------------|
| Close(Method) | **Description**<br>탭을 닫는다.<br>**Declaration**<br>`void Close(BOOL isDirty)`<br>**Parameters**<br>isDirty : TRUE이면 문서 내용이 변경된 경우 닫지 않는다. FALSE이면 문서 내용이 변경된 것과 상관없이 강제로 닫는다. |

---

## HSet

**ParameterSet Item 데이터들의 집합**

예) `HParameterSet.HSecDef.HSet` - HSecDef에 저장된 모든 Item 데이터의 집합을 얻음

| Item Name | Description |
|-----------|-------------|
| SetItem(Method) | **Description**<br>지정된 아이템이름을 가진 데이터에 VARIANT값을 대입한다.<br>**Declaration**<br>`SetItem(BSTR itemid, VARIANT value)`<br>**Parameters**<br>itemid : 아이템 이름<br>value : 데이터 |

---

## HAction

**한/글에서 특정 기능을 수행하기 위한 액션 오브젝트**

### 사용 예제

아래와 같은 형태로 사용하는 것이 HAction을 사용하는 올바른 사용법이다.

```javascript
HAction.GetDefault("Print", HParameterSet.HPrint.HSet); // 액션 초기화
HParameterSet.HPrint.NumCopy = 3; //인쇄 매수를 3장으로 지정
HAction.Execute("Print", HParameterSet.HPrint.HSet); // 액션 수행
```

| Item Name | Description |
|-----------|-------------|
| GetDefault(Method) | **Description**<br>지정된 액션이름과 LPDISPATCH로 HSet을 받아 액션을 초기화 한다.<br>**Declaration**<br>`BOOL GetDefault(BSTR actname, LPDISPATCH object)`<br>**Parameters**<br>actname : 액션 이름 |
