# HwpAutomation_2504_part05

## Page 41

### FileTranslate(Method) (계속)

| 항목 | 내용 |
|-----|-----|
| Parameters | curLang : 현재 문서에 적용된 언어<br>transLang : 번역할 대상 언어 |
| Remark | 한국어에서 다른 언어로 번역은 가능하지만, 현재 다른언어에서의 번역은 한국어밖에 지원하지 않음. |

### FileTranslate curLang 표

| curLang | Description |
|---------|-------------|
| ko | 한국어(대한민국) |
| en | 영어(미국) |
| vi | 베트남어(베트남) |
| ja | 일본어(일본) |
| zh-cn | 중국어(간체,PRC) |

### GetTranslateLangList(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 번역 가능한 언어의 목록을 반환한다 |
| Declaration | `BSTR GetTranslateLangList(BSTR curLang);` |
| Parameters | curLang : 현재 문서에 적용된 언어<br>return : 현재 언어로 번역 가능한 목록을 문자열로 반환 |

### SetCurMetatagName(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 선택된 개체나 블록, 셀 또는 문서정보에 메타태그 이름을 설정한다 |
| Declaration | `bool SetCurMetatagName(BSTR tag);` |
| Parameters | tag : 설정할 태그 이름 |
| Remark | 한/글 2022 부터 지원 |

### GetCurMetatagName(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 현재 메타태그의 이름을 가져온다 |
| Declaration | `BSTR GetCurMetatagName();` |
| Remark | 한/글 2022 부터 지원 |

### RenameMetatag(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 현재 메타태그의 이름을 변경한다 |
| Declaration | `bool SetCurMetatagName(BSTR tag);` |
| Parameters | tag : 설정할 태그 이름 |
| Remark | 한/글 2022 부터 지원 |

### MetatagExist(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 메타태그가 등록되어 있는지 확인한다. |
| Declaration | `bool MetatagExist(BSTR tag);` |
| Parameters | tag : 태그 이름 |
| Remark | 한/글 2022 부터 지원 |

### MoveToMetatag(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 지정한 메타태그로 캐럿을 이동한다. |
| Declaration | `bool MoveToMetatag(BSTR tag, VARIANT text, VARIANT start, VARIANT select);` |

## Page 42

| 항목 | 내용 |
|-----|-----|
| Parameters | tag : 태그 이름. GetMetatagNameText/PutMetatagNameText과 같은 형식으로 이름 뒤에 '{{#}}'로 번호를 지정할 수 있다.<br>text : true/false. 메타태그가 내부에 텍스트가 있는 경우 텍스트 안으로 이동할지 지정<br>start : true/false. 메타태그의 시작으로 이동할지 끝으로 이동할지 지정<br>select : true/false. 메타태그의 내용을 블록으로 선택할지, 캐럿만 이동할지 지정 |
| Remark | 한/글 2022 부터 지원 |

### GetMetatagNameText(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 지정한 이름의 메타태그 내용을 가져온다 |
| Declaration | `BSTR GetMetatagNameText(BSTR tag);` |
| Parameters | tag : 텍스트를 구할 메타태그 이름의 리스트. 다음과 같이 태그 사이를 문자 코드 0x02로 구분하여 한 번에 여러 개의 태그를 지정할 수 있다.<br>- "태그이름#1\x2태그이름#2\x2...태그이름#n"<br>- 지정한 태그 이름이 문서 중에 두 개 이상 존재할 때의 표현 방식은 다음과 같다. |
| Return | - 텍스트 데이터가 돌아온다. 텍스트에서 탭은 '\t'(0x9), 문단 바뀜은 CR/LF(0x0D/0x0A)로 표현되며, 이외의 특수 코드는 포함되지 않는다.<br>- 태그 텍스트의 끝은 0x02로 표현되며, 그 이후 다음 태그의 텍스트가 연속해서 지정한 태그 리스트의 개수만큼 위치한다. 지정한 이름의 태그가 없거나 사용자가 해당 태그에 아무 텍스트도 입력하지 않았으면 해당 텍스트에는 빈 문자열이 돌아온다. |
| Remark | 한/글 2022 부터 지원 |

### GetMetatagNameText 태그 표현 방식

| 표현 | 설명 |
|-----|-----|
| 태그이름 | 이름의 태그 중 첫 번째 |
| 태그이름{{n}} | 지정한 이름의 태그 중 n 번째 |

예를 들어 "제목{{1}}\x2본문\x2이름{{0}}" 과 같이 지정하면 '제목'이라는 이름의 태그 중 두 번째, '본문'이라는 이름의 태그 중 첫 번째, '이름'이라는 태그 중 첫 번째를 각각 지정한다. 즉, '태그이름'과 '태그이름{{0}}'은 동일한 의미로 해석된다.

### PutMetatagNameText(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 지정한 메타태그의 내용을 채운다. |
| Declaration | `void PutMetatagNameText(BSTR tag, BSTR text);` |
| Parameters | tag : 내용을 채울 메타태그 이름의 리스트. 한 번에 여러 개의 태그를 지정할 수 있으며, 형식은 GetMetatagNameText와 동일하다. 다만 태그 이름 뒤에 '{{#}}'로 번호를 지정하지 않으면 해당 이름을 가진 모든 태그에 동일한 텍스트를 채워 넣는다. 즉, PutMetatagNameText에서는 '태그이름'과 '태그이름{{0}}'의 의미가 다르다.<br>text : 메타태그에 채워 넣을 문자열의 리스트. 형식은 tag와 동일하게 태그의 개수만큼 텍스트를 0x02로 구분하여 지정한다. |
| Remark | 한/글 2022 부터 지원 |

### GetMetatagList(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 문서에 적용된 메타태그 목록을 가져온다 |
| Declaration | `BSTR GetMetatagList(VARIANT number, VARIANT option);` |

## Page 43

| 항목 | 내용 |
|-----|-----|
| Parameters | number : 문서 중에서 동일한 이름의 태그가 여러 개 존재할 때 이를 구별하기 위한 식별 방법을 지정한다.<br>option : 태그를 가져올 범위. 아래 값을 더해서 범위를 지정할 수 있다.<br>return : 각 태그 사이를 문자코드 0x02로 구분하여 다음과 같은 형식으로 리턴 한다. (가장 마지막 태그에는 0x02가 붙지 않는다.)<br>"태그이름#1\x2태그이름#2\x2...태그이름#n" |
| Remark | 문서 중에 동일한 이름의 태그가 여러 개 존재할 때는 number에 지정한 타입에 따라 3 가지의 서로 다른 방식을 중에서 선택할 수 있다.<br>한/글 2022 부터 지원 |

### GetMetatagList number 표

| ID | 값 | 설명 |
|----|---|-----|
| hwpMetatagPlain | 0 | 아무 기호 없이 순서 대로 태그 이름이 나열된다. |
| hwpMetatagNumber | 1 | 태그 이름 뒤에 일련 번호가 {{#}}와 같은 형식으로 붙는다 |
| hwpMetatagCount | 2 | 태그 이름 뒤에 그 이름의 태그가 몇 개 있는지 {{#}}와 같은 형식으로 붙는다. |

### GetMetatagList option 표

| ID | 값 | 설명 |
|----|---|-----|
| hwpMetatagPara | 1 | 셀렉션 내에 부여된 태그 리스트만을 구한다. |
| hwpMetatagCell | 2 | 셀에 부여된 태그 리스트만을 구한다. |
| hwpMetatagTable | 4 | 표에 부여된 태그 리스트만을 구한다. |
| hwpMetatagClickHere | 8 | 누름틀에 부여된 태그 리스트만을 구한다. |
| hwpMetatagShapeObject | 16 | 도형에 부여된 태그 리스트만을 구한다. |
| hwpMetatagPicture | 32 | 그림에 부여된 태그 리스트만을 구한다. |

### CurMetatagState(Property)

| 항목 | 내용 |
|-----|-----|
| Description | 캐럿이 위치한 메타태그의 상태 정보를 구한다.(읽기전용) |
| Declaration | `long CurMetatagState();` |
| Parameters | return : 얻어지는 내용은 일종의 flag값으로 비트가 set되어 있으면 다음과 같은 의미를 가진다. |
| Remark | 한/글 2022 부터 지원 |

### CurMetatagState 반환값 표

| value | Description |
|-------|-------------|
| 0x01 | 셀 |
| 0x02 | 누름틀 |
| 0x04 | 글상자 |
| 0x08 | 도형 |
| 0x10 | 표 |
| 0x20 | 태그명 없음 |
| 0x40 | 문단 |

### ModifyMetatagProperties(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 지정한 메타태그의 속성을 바꾼다. |
| Declaration | `long ModifyMetatagProperties(BSTR tag, long remove, long add)` |

## Page 44

| 항목 | 내용 |
|-----|-----|
| Parameters | tag : 속성을 바꿀 태그 이름의 리스트. 형식은 PutMetatagNameText과 동일.<br>remove : 제거될 속성<br>add : 추가될 속성<br>return : 음수가 리턴되면 에러임을 나타낸다. |
| Remark | 속성의 값은 아래와 같다.<br>remove와 add에 둘다 0이 입력되면 현재 속성을 돌려준다.<br>리턴값이 음수인지 확인하여 쉽게 에러임을 판별할 수 있으며 자세한 에러내용은 bit mask로 and 연산하여 알아 낼 수 있다.<br>한/글 2022 부터 지원 |

### ModifyMetatagProperties 속성 값

| long value | 설명 |
|------------|-----|
| 0x00000001 | 양식모드에서 편집가능 속성 (0: 편집 불가, 1: 편집 가능) |

### ModifyMetatagProperties return 값의 bit field

| long value bit mask | 설명 |
|---------------------|-----|
| 0x00000001 | 양식모드에서 편집가능 속성 (0: 편집 불가, 1: 편집 가능) |
| 0x80000000 | 에러 |
| 0x40000000 | 태그를 찾을 수 없음 |

### IsTrackChange(Property)

| 항목 | 내용 |
|-----|-----|
| Description | 변경내용추적 상태를 확인하고 설정한다 |
| Parameters | 0 – 해제<br>1 - 설정 |

### IsTrackChangePassword(Property)

| 항목 | 내용 |
|-----|-----|
| Description | 변경내용추적 암호가 걸려 있는지 확인한다(읽기전용) |

### GetUserProperty(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 문서 정보 – 사용자 지정의 속성을 가져온다. |
| Declaration | `BSTR GetUserProperty(BSTR name, long option);` |
| Parameters | name : 속성을 가져올 이름<br>option : 0으로 고정(사용하지 않음) |
| Remark | 가져온 속성은 형식 + 값으로 되어 있으며, 문자코드 0x02로 구분된다.<br>한/글 2022 부터 지원 |

### SetUserProperty(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 문서 정보 - 사용자 지정 속성을 설정한다 |
| Declaration | `long SetUserProperty(BSTR name, long format, BSTR value, long option);` |
| Parameters | name : 사용자가 지정할 속성의 이름을 직접 입력한다.<br>format : 사용자가 지정할 속성의 데이터 형식을 목록에서 선택한다. 선택한 형식에 따라 입력할 수 있는 값이 달라진다. |

### SetUserProperty format 표

| value | Description |
|-------|-------------|
| 0 | 텍스트+숫자 |
| 1 | 날짜 |

## Page 45

| value | Description |
|-------|-------------|
| 2 | 숫자 |
| 3 | 예 또는 아니요 |

**Parameters (계속)**
- value : 사용자가 선택한 형식에 맞는 속성 값을 입력한다. 선택한 형식에 따라 텍스트, 숫자, 또는 날짜를 입력하거나 '예' 또는 '아니요' 값을 할당할 수도 있다.
- option : 0으로 고정(사용하지 않음)

**Remark**: 한/글 2022 부터 지원

### GetUserPropertyNameList(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 문서 정보 – 사용자 지정 속성에 등록된 이름 목록을 가져온다 |
| Declaration | `BSTR GetUserPropertyNameList(long option);` |
| Parameters | option : 0으로 고정(사용하지 않음) |
| Remark | 한/글 2022 부터 지원 |

### GetCtrlToPicture(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 선택한 개체를 이미지로 저장한다 |
| Declaration | `bool GetCtrlToPicture(BSTR path, long format, long fullSave);` |
| Parameters | path : 이미지가 저장 될 경로 및 파일명(확장자를 포함해서 작성해야한다)<br>format : 이미지를 저장할 형식<br>fullSave : 표가 2페이지 이상 넘어갔을 경우 모두 저장할지 여부 |
| Remark | 한/글 2024 부터 지원 |

### GetCtrlToPicture Format 표

| Format | 값 | 설명 |
|--------|---|-----|
| EMF | 16 | EMF형식으로 저장한다 |
| BMP | 20 | BMP형식으로 저장한다 |
| GIF | 22 | GIF형식으로 저장한다 (그림일 경우에만 사용가능) |
| JPG | 25 | JPG형식으로 저장한다 (그림일 경우에만 사용가능) |
| PNG | 27 | PNG형식으로 저장한다 (그림일 경우에만 사용가능) |
| TIF | 29 | TIF형식으로 저장한다 (그림일 경우에만 사용가능) |

### SelectCtrl(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 캐럿이 위치한 메타태그의 상태 정보를 구한다.(읽기전용) |
| Declaration | `bool SelectCtrl(BSTR ctrlList, long option);` |
| Parameters | ctrlList: 선택할 컨트롤 ID. 문자코드 0x02를 이용해서 여러 개의 컨트롤ID를 지정할 수 있다.<br>- 컨트롤 ID는 HwpCtrl.CtrlID가 리턴하는 ID와 동일하다. 자세한 것은 Ctrl 오브젝트 Properties인 CtrlID를 참조.<br>option : 현재 선택된 컨트롤에서 추가로 선택을 할지, 취소 후 다시 선택할지에 대한 옵션.<br>- 0 = 추가로 선택<br>- 1 = 취소 후 ctrlList에 등록된 컨트롤만 선택 |
| Remark | 한/글 2024 부터 지원 |

## Page 46

## IHwpObjectEvents(_DIHwpObjectEvents) : 한/글에서부터 발생되는 이벤트

| Item Name | Return | Description |
|-----------|--------|-------------|
| Quit | 없음 | 한/글을 종료할 때 발생 |
| CreateXHwpWindow | 없음 | 한/글에서 새 문서 창을 열었을 때 발생 |
| CloseXHwpWindow | 없음 | 한/글에서 문서 창을 닫았을 때 발생 |
| NewDocument | long | 새 문서를 생성할 경우 발생(Document ID를 반환) |
| DocumentBeforeClose | long | 문서를 닫기 직전에 발생(Document ID를 반환) |
| DocumentBeforeOpen | long | 문서를 열기 직전에 발생(Document ID를 반환) |
| DocumentAfterOpen | long | 문서를 열고 난 후에 발생(Document ID를 반환) |
| DocumentBeforeSave | long | 문서를 저장하기 직전에 발생(Document ID를 반환) |
| DocumentAfterSave | long | 문서를 저장한 후에 발생(Document ID를 반환) |
| DocumentAfterClose | long | 문서를 닫고 난 후에 발생(Document ID를 반환) |
| DocumentChange | long | 문서가 변경됐을 경우에 발생(Document ID를 반환) |
| DocumentBeforePrint | long | 문서를 인쇄하기 직전에 발생(Document ID를 반환) |
| DocumentAfterPrint | long | 문서를 인쇄하고 난 후에 발생(Document ID를 반환) |
| DocumentClickedHyperlink | | 하이퍼링크를 클릭했을 때 발생 |
| DocumentModifiedHyperlink | | 하이퍼링크를 수정했을 때 발생 |
| BeforeQuit | 없음 | 한/글을 종료하기 직전에 발생 |

## IXHwpDocuments: IXHwpDocument(도큐먼트) 오브젝트를 관리하는 오브젝트
(Document를 관리하는 Collection 개체 - 사용 편이를 위해 제공됨)

| Item Name | Description |
|-----------|-------------|
| Application(Property) | **Description**<br>최상위 오브젝트를 얻어옴(IHwpObject) |
| Item(Property) | **Description**<br>지정한 원소의 도큐먼트 오브젝트를 얻어온다.<br>**Declaration**<br>`LPDISPATCH Item(long index)`<br>**Parameters**<br>index : 원소의 인덱스<br>return : 원소의 값(IXHwpDocument) |
| Count(Property) | **Description**<br>원소의 총개수 |
| Active_XHwpDocument(Property) | **Description**<br>현재 활성화 상태인 도큐먼트 Object를 얻어온다.(IXHwpDocument) |
| Add(Method) | **Description**<br>도큐먼트 오브젝트를 추가한다.<br>**Declaration**<br>`LPDISPATCH Add(BOOL isTab)`<br>**Parameters**<br>isTab : TRUE = 새탭으로 열리는 도큐먼트, FALSE = 새창으로 열리는 도큐먼트<br>return : 열리게 되는 도큐먼트(IXHwpDocument) |

## Page 47

| Item Name | Description |
|-----------|-------------|
| Close(Method) | **Description**<br>관리하고 있는 도큐먼트 오브젝트를 삭제한다.<br>**Declaration**<br>`void Close(BOOL isDirty)`<br>**Parameters**<br>isDirty : TRUE이면 변경된 문서는 닫지 않는다. FALSE이면 변경된 문서도 닫는다. |
| FindItem(Method) | **Description**<br>도큐먼트 아이디로 지정된 도큐먼트 오브젝트를 얻는다.<br>**Declaration**<br>`LPDISPATCH FindItem(long Docid)`<br>**Parameters**<br>Docid : 도큐먼트의 고유 ID<br>return : 도큐먼트 ID에 해당하는 도큐먼트 오브젝트(IXHwpDocument) |

## IXHwpDocument: 도큐먼트 오브젝트
(Document 개체 - 사용 편이를 위해 제공됨)

| Item Name | Description |
|-----------|-------------|
| Application(Property) | **Description**<br>최상위 오브젝트를 얻어옴(IHwpObject - 읽기 전용) |
| Path(Property) | **Description**<br>도큐먼트의 Path를 얻어옴(읽기 전용) |
| FullName(Property) | **Description**<br>도큐먼트의 전체 경로를 얻어옴(읽기 전용) |
| EditMode(Property) | **Description**<br>도큐먼트의 에디트 모드를 설정하거나/얻어옴 |
| Modified(Property) | **Description**<br>도큐먼트의 변경 여부를 설정하거나/얻어옴 |
| Format(Property) | **Description**<br>도큐먼트의 저장된 포맷을 얻어옴(읽기 전용) |
| Password(Property) | **Description**<br>도큐먼트의 패스워드를 설정(쓰기 전용) |
| XHwpSummaryInfo(Property) | **Description**<br>IXHwpSummaryInfo 문서 요약 정보 오브젝트 (읽기 전용) |
| XHwpDocumentInfo(Property) | **Description**<br>IXHwpDocumentInfo 문서 정보 오브젝트 (읽기 전용) |
| XHwpPrint(Property) | **Description**<br>IXHwpPrint 프린트 오브젝트 (읽기 전용) |
| XHwpRange(Property) | **Description**<br>IXHwpRange Range 오브젝트 (읽기 전용) |
| XHwpFind(Property) | **Description**<br>IXHwpFind 찾기 오브젝트 (읽기 전용) |
| XHwpSelection(Property) | **Description**<br>IXHwpSelection 블록 선택 오브젝트 (읽기 전용) |
| XHwpFormPushButtons(Property) | **Description**<br>IXHwpFormPushButtons 양식개체 푸쉬버튼을 관리하는 오브젝트(읽기 전용) |

## Page 48

| Item Name | Description |
|-----------|-------------|
| XHwpFormCheckButtons(Property) | **Description**<br>IXHwpFormCheckButtons 양식개체 체크박스를 관리하는 오브젝트(읽기 전용) |
| XHwpFormRadioButtons(Property) | **Description**<br>IXHwpFormRadioButtons 양식개체 라디오버튼을 관리하는 오브젝트(읽기 전용) |
| XHwpFormComboBoxs(Property) | **Description**<br>IXHwpFormComboBoxs 양식개체 콤보박스를 관리하는 오브젝트(읽기 전용) |
| XHwpFormEdits(Property) | **Description**<br>IXHwpFormEdits 양식개체 에디트를 관리하는 오브젝트(읽기 전용) |
| XHwpCharacterShape(Property) | **Description**<br>IXHwpCharacterShape 글자 모양 속성 오브젝트(읽기 전용) |
| XHwpParagraphShape(Property) | **Description**<br>IXHwpParagraphShape 문단 모양 속성 오브젝트(읽기 전용) |
| XHwpSendMail(Property) | **Description**<br>IXHwpSendMail 메일 보내기 오브젝트 (읽기 전용) |
| DocumentID(Property) | **Description**<br>도큐먼트의 고유 ID(읽기 전용) |
| Close(Method) | **Description**<br>문서를 닫는다.<br>**Declaration**<br>`BOOL Close(BOOL isDirty)`<br>**Parameters**<br>isDirty : TRUE = 문서 내용이 변경된 상태면 문서를 닫지 않는다./ FALSE = 문서 내용이 변경되었어도 강제로 문서를 닫는다. |
| Save(Method) | **Description**<br>문서를 저장한다.<br>**Declaration**<br>`BOOL Save(Variant save_if_dirty)`<br>**Parameters**<br>save_if_dirty : True를 지정하면 문서가 변경된 경우에만 저장한다. False를 지정하면 변경 여부에 관계없이 무조건 저장한다. 생략하면 True가 지정된다.<br>**Remark**<br>문서의 경로가 지정되어 있지 않으면 "새이름으로 저장" 대화상자가 떠서 사용자에게 경로를 묻는다. |
| SaveAs(Method) | **Description**<br>문서를 지정한 이름으로 저장한다.<br>**Declaration**<br>`BOOL Save(BSTR path, VARIANT format, VARIANT arg)`<br>**Parameters**<br>path : 문서 파일의 경로<br>format : 문서 형식. 별도 설명 참조. 생략하면 "HWP"가 지정된다.<br>arg : 세부 옵션. 의미는 format에 지정한 파일 형식에 따라 다르다. 생략하면 빈 문자열이 지정된다.<br>**Remark**<br>format, arg의 일반적인 개념에 대해서는 Open 참조. |

## Page 49

### "HWP" 형식으로 파일 저장시 arg 옵션

| 함수 | 인자 타입 | 기본 값 (Default) | 설명 |
|-----|---------|-----------------|-----|
| lock | boolean | TRUE | 저장한 후 해당 파일을 계속 오픈한 상태로 lock을 걸지 여부 |
| backup | boolean | FALSE | 백업 파일 생성 여부 |
| compress | boolean | TRUE | 압축 여부 |
| fullsave | boolean | FALSE | 스토리지 파일을 완전히 새로 생성하여 저장 |
| prvimage | int | 2 | 미리보기 이미지 (0=off, 1=BMP, 2=GIF) |
| prvtext | int | 1 | 미리보기 텍스트(0=off, 1=0) |

### Undo(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 문서에 기록된 Undo Item을 실행한다. |
| Declaration | `BOOL Undo(long count)` |
| Parameters | count : 아이템의 count까지 Undo한다. |

### Redo(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 문서에 기록된 Redo Item을 실행한다. |
| Declaration | `BOOL Redo(long count)` |
| Parameters | count : 아이템의 count까지 Redo한다. |

### Open(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 문서 파일을 연다. |
| Declaration | `Open(BSTR path, BSTR format, BSTR arg)` |
| Parameters | path : 문서 파일의 경로(URL 사용 가능)<br>format : 문서 형식. 별도 설명 참조. 빈 문자열을 지정하면 자동으로 인식한다. 생략하면 빈 문자열이 지정된다.<br>arg : 세부 옵션. 의미는 format에 지정한 파일 형식에 따라 다르다. 생략하면 빈 문자열이 지정된다. |
| Remark | format에 지정할 수 있는 문서 형식은 현재 시스템에 설치된 문서 필터(*.dft)의 종류에 따라 달라진다. |

### Open 문서 형식 종류

| Format | 설명 |
|--------|-----|
| HWP | 워디안 native format |
| HWP30 | 한글 3.X/96/97 |
| HTML | 인터넷 문서 |
| TEXT | 아스키 텍스트 문서 |
| UNICODE | 유니코드 텍스트 문서 |
| HWP20 | 한글 2.0 |
| HWP21 | 한글 2.1/2.5 |
| HWP15 | 한글 1.X |
| HWPML1X | HWPML 1.X 문서 (Open만 가능) |
| HWPML2X | HWPML 2.X 문서 (Open / SaveAs 가능) |
| RTF | 서식있는 텍스트 문서 |
| DBF | DBASE II/III 문서 |
| HUNMIN | 훈민정음 3.0/2000 |

## Page 50

| Format | 설명 |
|--------|-----|
| MSWORD | 마이크로소프트 워드 문서 |
| HANA | 하나워드 문서 |
| ARIRANG | 아리랑 문서 |
| ICHITARO | 一太郞 문서 (일본 워드프로세서) |
| WPS | WPS 문서 |
| DOCIMG | 인터넷 프레젠테이션 문서(SaveAs만 가능) |
| SWF | Macromedia Flash 문서(SaveAs만 가능) |

### arg 옵션 문법

arg에 지정할 수 있는 옵션의 의미는 필터가 정의하기에 따라 다르지만, 신텍스는 다음과 같이 공통된 형식을 사용한다.

```
key:value;key:value;...
```

- key는 A-Z, a-z, 0-9, _ 로 구성된다.
- value는 타입에 따라 다음과 같은 3 종류가 있다.
  - boolean: ex) fullsave:true (== fullsave)
  - integer: ex) type:20
  - string: ex) prefix:_This_
- value는 생략 가능하며, 이때는 콜론도 생략한다.

### arg에 지정할 수 있는 옵션

| Format | 옵션 |
|--------|-----|
| "모든 파일" | setcurdir(boolean, FALSE) : 로드한 후 해당 파일이 존재하는 폴더로 현재 위치를 변경한다. hyperlink 정보가 상대적인 위치로 되어 있을 때 유용하다. |
| "HWP" | - lock (boolean, TRUE) : 로드한 후 해당 파일을 계속 오픈한 상태로 lock을 걸지 여부<br>- notext (boolean, FALSE) : 텍스트 내용을 읽지 않고 헤더 정보만 읽을지 여부. (스타일 로드 등에 사용)<br>- template (boolean, FALSE): 새로운 문서를 생성하기 위해 템플릿 파일을 오픈한다. 이 옵션이 주어지면 lock은 무조건 FALSE로 처리된다.<br>- suspendpassword (boolean, FALSE): TRUE로 지정하면, 암호가 있는 파일일 경우 암호를 묻지 않고 무조건 읽기에 실패한 것으로 처리한다.<br>- forceopen (boolean, FALSE): TRUE로 지정하면, 읽기 전용으로 읽어야 하는 경우 대화상자를 띄우지 않는다. |
| "HTML" | - code(string, codepage) : 문서 변환시 사용되는 코드 페이지를 지정할 수 있이며 code키가 존재할 경우 필터 사용시 사용자 다이얼로그를 띄우지 않는다.<br>- textunit(boolean, pixel) : Export될 Text의 크기의 단위 결정(pixel, point, mili 지정 가능.)<br>- formatunit(boolean, pixel) : Export될 문서 포맷 관련 (마진, Object 크기 등) 단위 결정. pixel, point, mili 지정 가능 |
| "DOCIMG" | - asimg(boolean, FALSE) : 저장할 때 페이지를 image로 저장<br>- ashtml(boolean, FALSE) : 저장할 때 페이지를 html로 저장 |
| "TEXT" | - code(string, codepage): 문서 변환시 사용되는 코드 페이지를 지정할 수 있이며 code키가 존재할 경우 필터 사용시 사용자 다이얼로그를 띄우지 않는다. |

### codepage 종류

ks : 한글 KS 완성형 | kssm : 한글 조합형 | sjis : 일본 | utf8 : UTF8 | unicode : 유니코드 | gb : 중국 간체 | big5 : 중국 번체 | acp : Active Codepage 현재 시스템의 코드 페이지
