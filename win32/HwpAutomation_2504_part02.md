# HwpAutomation_2504_part02

## Page 11

### 생성된 스크립트 코드를 인터넷 익스플로러에 적용한 예제

```html
<HTML>
<SCRIPT LANGUAGE="JSCRIPT">
var App = new ActiveXObject("HwpFrame.HwpObject.2");
function OnScriptMacro_script1()
{
    App.HAction.GetDefault("InsertText", App.HParameterSet.HInsertText.HSet);
    App.HParameterSet.HInsertText.Text = "글자입력";
    App.HAction.Execute("InsertText", App.HParameterSet.HInsertText.HSet);
    App.HAction.Run("MoveSelLineBegin");
    App.HAction.GetDefault("CharShape", App.HParameterSet.HCharShape.HSet);
    App.HParameterSet.HCharShape.TextColor = 16750848;
    App.HAction.Execute("CharShape", App.HParameterSet.HCharShape.HSet);
    App.HAction.Run("Cancel");
}
</SCRIPT>
<BODY>
<BUTTON OnClick="OnScriptMacro_script1()"> 글자 테스트</BUTTON>
</BODY>
</HTML>
```

### 생성된 스크립트 코드를 MFC에 적용한 예제

```cpp
IHwpObject pHwpObject;
void CTestText::OnButtonClick()
{
    BOOL bres = pHwpObject.CreateDispatch("HwpFrame.HwpObject.2");
    if (bres == FALSE)
        return ;
    HAction haction;
    haction.AttachDispatch(pHwpObject.GetHAction());
    HParameterSet hparameterset;
    hparameterset.AttachDispatch(pHwpObject.GetHParameterSet());
    HInsertText hinserttext;
    hinserttext.AttachDispatch(hparameterset.GetHInsertText);
    HSet hset1;
    hset1.AttachDispatch(hinserttext.GetHSet);
    haction.GetDefault("InsertText", hset1);
    hinserttext.Text = "글자입력";
    haction.Execute("InsertText", hset1);
    haction.Run("MoveSelLineBegin");
    HCharShape hcharshape;
    hcharshape.AttachDispatch(hparameterset.GetHCharShape);
    HSet hset2;
    hset2.AttachDispatch(hcharshape.GetHSet);
    haction.GetDefault("CharShape", hset2);
    hcharshape.TextColor = 16750848;
    haction.Execute("CharShape", hset2);
    haction.Run("Cancel");
}
```

## Page 12

## 1.7. OLE Automation 오브젝트

### IHwpObject : 최상위 개체 (모든 Automation Object의 최상위 Object이다.

| Item Name | Description |
|-----------|-------------|
| IsModified(Property) | **Description**<br>문서가 변경되어있는지 나타낸다. (읽기 전용)<br>**Remark**<br>0 = 변경되지 않은 깨끗한 상태<br>1 = 변경된 상태<br>2 = 변경되었으나 자동 저장된 상태 |
| IsEmpty(Property) | **Description**<br>아무 내용도 들어있지 않은 빈 문서인지 여부를 나타낸다.(읽기 전용) |
| EditMode(Property) | **Description**<br>현재 편집 모드<br>**Remark**<br>0 : 읽기 전용<br>1 : 일반 편집모드<br>2 : 양식 모드(양식 사용자 모드) : Cell과 누름틀 중 양식 모드에서 편집 가능 속성을 가진 것만 편집 가능하다.<br>16 : 배포용 문서 (SetEditMode로 지정 불가능) |
| SelectionMode(Property) | **Description**<br>문서의 내용이 어떤 Selection 상태인가를 알려준다.(읽기 전용)<br>**Remark**<br>일반 블록이 아닌 F3키나 F4키에 의해 블록이 지정된 경우, HWPSEL_STRICT_MODE ( = 0x10, 십진수 16)으로 OR 마스크되어 오기때문에, 항상 0x0F(십진수15)로 AND 마스크한 결과로 판단하도록 한다. |
| CurFieldState(Property) | **Description**<br>캐럿이 위치한 필드의 상태 정보를 구한다.(읽기 전용)<br>**Remark**<br>bit 5 - 31 = 예약<br>bit 4 = 필드명의 존재 여부 (1 = 있음, 0 = 없음)<br>bit 0 - 3 = 필드의 종류 (0 = 없음, 1 = 셀, 2 = 누름틀) |
| PageCount(Property) | **Description**<br>문서 페이지 수 (읽기 전용)<br>**Remark**<br>문서의 전체 페이지 수를 나타낸다. 문서 전체에 대한 pagination이 수행되지 않은 상태에서 이 property를 참조하면 먼저 문서 전체의 pagination을 먼저 수행하므로 긴 문서에 대해 문서 내용 변경과 참조를 반복하면 속도가 심각하게 느려질 수 있다. |
| CellShape(Property) | **Description**<br>현재 선택되어있는 표와 셀의 모양 정보를 나타낸다.<br>**Remark**<br>ParameterSet/Table로 표의 속성에 대한 기본 정보를 나타내며, 이 가운데 "Cell" 아이템이 ParameterSet/Cell로 셀의 속성을 나타낸다. 셀 블록이 잡혀있지 않은 상태이면 현재 캐럿이 위치한 셀 하나만을 대상으로 한다. 현재 표 내부에 캐럿이 위치하지 않으면 에러가 발생한다. |
| CharShape(Property) | **Description**<br>현재 Selection의 글자 모양을 나타낸다.<br>**Remark**<br>property get을 수행하면 현재 selection 내의 글자 모양을 구할 수 있다. selection이 존재하지 않으면 현재 캐럿이 위치한 곳의 글자 모양을 돌려준다. 글자 모양 중 특정 항목이 selection 내에서 서로 다른 속성을 가지고 있으면 아예 아이템 자체가 존재하지 않는다.<br>property set을 수행하면 아이템이 존재하는 항목에 대해서만 속성을 설정한다. (ParameterSet의 형식은 ParameterSet/CharShape 참조.) |

## Page 13

| Item Name | Description |
|-----------|-------------|
| HeadCtrl(Property) | **Description**<br>문서 중 첫 번째 컨트롤(읽기 전용)<br>**Remark**<br>문서 중의 모든 컨트롤(표, 그림 등의 특수 문자들)은 linked list로 서로 연결되어 있는데, 그 list의 시작 컨트롤을 나타낸다. 이 컨트롤로부터 시작, Ctrl.Next를 이용해 forward iteration을 수행할 수 있다. |
| LastCtrl(Property) | **Description**<br>문서 중 마지막 컨트롤 (읽기 전용)<br>**Remark**<br>문서 중의 모든 컨트롤(표, 그림 등의 특수 문자들)은 linked list로 서로 연결되어 있는데, 그 list의 마지막 컨트롤을 나타낸다. 이 컨트롤로부터 시작, Ctrl.Prev를 이용해 backward iteration을 수행할 수 있다. |
| CurSelectedCtrl(Property) | **Description**<br>현재 선택되어 있는 컨트롤(읽기 전용) |
| ParaShape(Property) | **Description**<br>현재 Selection의 문단 모양을 나타낸다.<br>**Remark**<br>개념과 사용법은 CharShape과 동일하다. (ParameterSet의 형식은 ParameterSet/ParaShape 참조.) |
| ParentCtrl(Property) | **Description**<br>현재 캐럿의 상위 컨트롤(읽기 전용)<br>**Remark**<br>상위 컨트롤은 현재 캐럿이 위치한 리스트를 보유한 컨트롤이다. 예를 들어 셀 내부에 위치하면 표, 각주 내용에 위치하면 각주, 바탕쪽이면 구역 컨트롤이 상위 컨트롤이다. 현재 캐럿이 본문 레벨에 위치해 상위 컨트롤이 없을 때는 NULL이 리턴된다. |
| ViewProperties(Property) | **Description**<br>뷰의 상태 정보<br>**Remark**<br>조판 부호, 화면 확대 비율과 같은 view에 관련된 정보를 나타낸다. (ParameterSet의 형식은 ParameterSet/ViewProperties 참조.) |
| Path(Property) | **Description**<br>문서 파일의 경로 |
| IsPrivateInfoProtected(Property) | **Description**<br>개인정보가 보호되어 있는지 여부를 알 수 있는 속성(읽기 전용)<br>**Remark**<br>문서 내에 개인정보가 보호된 문서인지 아닌지를 알 수 있다. |
| Open(Method) | **Description**<br>문서 파일을 연다.<br>**Declaration**<br>`Open(BSTR path, BSTR format, BSTR arg)`<br>**Parameters**<br>path : 문서 파일의 경로(URL 사용 가능)<br>format : 문서 형식. 별도 설명 참조. 빈 문자열을 지정하면 자동으로 인식한다. 생략하면 빈 문자열이 지정된다.<br>arg : 세부 옵션. 의미는 format에 지정한 파일 형식에 따라 다르다. 생략하면 빈 문자열이 지정된다.<br>**Remark**<br>format에 지정할 수 있는 문서 형식은 현재 시스템에 설치된 문서 필터(*.dft)의 종류에 따라 달라진다. |

## Page 14

### 문서 형식 종류

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
| "HWP" | - lock (boolean, TRUE) : 로드한 후 해당 파일을 계속 오픈한 상태로 lock을 걸지 여부<br>- notext (boolean, FALSE) : 텍스트 내용을 읽지 않고 헤더 정보만 읽을지 여부. (스타일 로드 등에 사용)<br>- template (boolean, FALSE): 새로운 문서를 생성하기 위해 템플릿 파일을 오픈한다. 이 옵션이 주어지면 lock은 무조건 FALSE로 처리된다.<br>- suspendpassword (boolean, FALSE): TRUE로 지정하면, 암호가 있는 파일일 경우 암호를 묻지 않고 무조건 읽기에 실패한 것으로 처리한다.<br>- forceopen (boolean, FALSE): TRUE로 지정하면, 읽기 전용으로 읽어야 하는 경우 대화상자를 띄우지 않는다.<br>- versionwarning (boolean, TRUE): FALSE로 지정하면 "상위 버전에서 작성한 문서입니다" 메세지가 안나옴. |
| "HTML" | - code(string, codepage) : 문서 변환시 사용되는 코드 페이지를 지정할 수 있이며 code키가 존재할 경우 필터 사용시 사용자 다이얼로그를 띄우지 않는다.<br>- textunit(boolean, pixel) : Export될 Text의 크기의 단위 결정(pixel, point, mili 지정 가능.) |

## Page 15

| Format | 옵션 |
|--------|-----|
| "HTML" (계속) | - formatunit(boolean, pixel) : Export될 문서 포맷 관련 (마진, Object 크기 등) 단위 결정. pixel, point, mili 지정 가능 |
| "DOCIMG" | - asimg(boolean, FALSE) : 저장할 때 페이지를 image로 저장<br>- ashtml(boolean, FALSE) : 저장할 때 페이지를 html로 저장 |
| "TEXT" | - code(string, codepage): 문서 변환시 사용되는 코드 페이지를 지정할 수 있이며 code키가 존재할 경우 필터 사용시 사용자 다이얼로그를 띄우지 않는다. |

### codepage 종류

| 코드 | 설명 |
|-----|-----|
| ks | 한글 KS 완성형 |
| kssm | 한글 조합형 |
| sjis | 일본 |
| utf8 | UTF8 |
| unicode | 유니코드 |
| gb | 중국 간체 |
| big5 | 중국 번체 |
| acp | Active Codepage 현재 시스템의 코드 페이지 |

### Save(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 현재 편집중인 문서를 저장한다. |
| Declaration | `BOOL Save(VARIANT save_if_dirty)` |
| Parameters | save_if_dirty : True를 지정하면 문서가 변경된 경우에만 저장한다. False를 지정하면 변경 여부에 관계없이 무조건 저장한다. 생략하면 True가 지정된다. |
| Remark | 문서의 경로가 지정되어 있지 않으면 "새이름으로 저장" 대화상자가 떠서 사용자에게 경로를 묻는다. |

### SaveAs(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 현재 편집중인 문서를 지정한 이름으로 저장한다. |
| Declaration | `BOOL Save(BSTR path, VARIANT format, VARIANT arg)` |
| Parameters | path : 문서 파일의 경로<br>format : 문서 형식. 별도 설명 참조. 생략하면 "HWP"가 지정된다.<br>arg : 세부 옵션. 의미는 format에 지정한 파일 형식에 따라 다르다. 생략하면 빈 문자열이 지정된다. |
| Remark | format, arg의 일반적인 개념에 대해서는 Open 참조. |

### "HWP" 형식으로 파일 저장시 arg 옵션

| 함수 | 인자 타입 | 기본 값 (Default) | 설명 |
|-----|---------|-----------------|-----|
| lock | boolean | TRUE | 저장한 후 해당 파일을 계속 오픈한 상태로 lock을 걸지 여부 |
| backup | boolean | FALSE | 백업 파일 생성 여부 |
| compress | boolean | TRUE | 압축 여부 |
| fullsave | boolean | FALSE | 스토리지 파일을 완전히 새로 생성하여 저장 |
| prvimage | int | 2 | 미리보기 이미지 (0=off, 1=BMP, 2=GIF) |
| prvtext | int | 1 | 미리보기 텍스트(0=off, 1=0) |

## Page 16

### Insert(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 현재 캐럿 위치에 문서 파일을 삽입한다. |
| Declaration | `Insert(BSTR path, VARIANT format, VARIANT arg)` |
| Parameters | path : 문서 파일의 경로, URL 사용 가능<br>format : 문서 형식. 별도 설명 참조. 빈 문자열을 지정하면 자동으로 디텍트한다. 생략하면 빈 문자열이 지정된다.<br>arg : 세부 옵션. 의미는 format에 지정한 파일 형식에 따라 다르다. 생략하면 빈 문자열이 지정된다. |
| Remark | format, arg에 대해서는 Open 참조 |

### SelectText(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 블록을 설정 한다. |
| Declaration | `BOOL SelectText(long spara, long spos, long epara, long epo)` |
| Parameters | spara : 블록 시작 위치의 문단 번호.<br>spos : 블록 시작 위치의 문단 중에서 문자의 위치.<br>epara : 블록 끝 위치의 문단 번호.<br>epos : 블록 끝 위치의 문단 중에서 문자의 위치. |
| Remark | epos가 가리키는 문자는 포함되지 않는다. |

### CreateField(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 캐럿의 현재 위치에 누름틀을 생성한다. |
| Declaration | `BOOL CreateField(BSTR direction, VARIANT memo, VARIANT name)` |
| Parameters | direction : 누름틀에 입력이 안된 상태에서 보여지는 안내문/지시문.<br>memo : 누름틀에 대한 설명/도움말<br>name : 누름틀 필드에 대한 필드 이름 |

### MoveToField(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 지정한 필드로 캐럿을 이동한다. |
| Declaration | `BOOL MoveToField(BSTR field, VARIANT text, VARIANT start, VARIANT select)` |
| Parameters | field : 필드 이름. GetFieldText/PutFieldText과 같은 형식으로 이름 뒤에 '{{#}}'로 번호를 지정할 수 있다.<br>text : 필드가 누름틀일 경우 누름틀 내부의 텍스트로 이동할지(True) 누름틀 코드로 이동할지(False)를 지정한다. 누름틀이 아닌 필드일 경우 무시된다. 생략하면 True가 지정된다.<br>start : 필드의 처음(True)으로 이동할지 끝(False)으로 이동할지 지정한다. select를 True로 지정하면 무시된다. 생략하면 True가 지정된다.<br>select : 필드 내용을 블록으로 선택할지(True), 캐럿만 이동할지(False) 지정한다. 생략하면 False가 지정된다. |
| Remark | 누름틀은 한글97 메뉴 중 입력 메뉴에 문서마당 정보를 선택하면 누름틀을 만드실 수 있습니다. |

### FieldExist(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 문서 중에 지정한 데이터 필드가 존재하는지 검사한다. |
| Declaration | `BOOL FieldExist(BSTR field)` |
| Parameters | field : 필드 이름 |

## Page 17

### GetFieldText(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 지정한 필드에서 문자열을 구한다. |
| Declaration | `BSTR GetFieldText(BSTR fieldlist)` |
| Parameters | fieldlist : 텍스트를 구할 필드 이름의 리스트. 다음과 같이 필드 사이를 문자 코드 0x02로 구분하여 한 번에 여러 개의 필드를 지정할 수 있다.<br>"필드이름#1\x2필드이름#2\x2...필드이름#n" |
| Remark | 지정한 필드 이름이 문서 중에 두 개 이상 존재할 때의 표현 방식:<br>- 필드이름 : 이름의 필드 중 첫 번째<br>- 필드이름{{n}} : 지정한 이름의 필드 중 n 번째<br><br>예를 들어 "제목{{1}}\x2본문\x2이름{{0}}" 과 같이 지정하면 '제목'이라는 이름의 필드 중 두 번째, '본문'이라는 이름의 필드 중 첫 번째, '이름'이라는 이름의 필드 중 첫 번째를 각각 지정한다. 즉, '필드이름'과 '필드이름{{0}}'은 동일한 의미로 해석된다.<br><br>return : 텍스트 데이터가 돌아온다. 텍스트에서 탭은 '\t'(0x9), 문단 바뀜은 CR/LF(0x0D/0x0A)로 표현되며, 이외의 특수 코드는 포함되지 않는다. 필드 텍스트의 끝은 0x02로 표현되며, 그 이후 다음 필드의 텍스트가 연속해서 지정한 필드 리스트의 개수만큼 위치한다. 지정한 이름의 필드가 없거나 사용자가 해당 필드에 아무 텍스트도 입력하지 않았으면 해당 텍스트에는 빈 문자열이 돌아온다. |

### PutFieldText(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 지정한 필드의 내용을 채운다. |
| Declaration | `void PutFieldText(BSTR fieldlist, BSTR textlist)` |
| Parameters | fieldlist : 내용을 채울 필드 이름의 리스트. 한 번에 여러 개의 필드를 지정할 수 있으며, 형식은 GetFieldText와 동일하다. 다만 필드 이름 뒤에 '{{#}}'로 번호를 지정하지 않으면 해당 이름을 가진 모든 필드에 동일한 텍스트를 채워 넣는다. 즉, PutFieldText에서는 '필드이름'과 '필드이름{{0}}'의 의미가 다르다.<br>textlist : 필드에 채워 넣을 문자열의 리스트. 형식은 필드 리스트와 동일하게 필드의 개수만큼 텍스트를 0x02로 구분하여 지정한다. |
| Remark | 현재 필드에 입력되어 있는 내용은 지워진다. 채워진 내용의 글자모양은 필드에 지정해 놓은 글자모양을 따라간다.<br>fieldlist의 필드 개수와, textlist의 텍스트 개수는 동일해야 한다. 존재하지 않는 필드에 대해서는 무시한다. |

### RenameField(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 지정한 필드의 이름을 바꾼다. |
| Declaration | `void RenameField(BSTR oldname, BSTR newname)` |
| Parameters | oldname : 이름을 바꿀 필드 이름의 리스트. 형식은 PutFieldText과 동일하다.<br>newname : 새로운 필드 이름의 리스트. oldname과 동일한 개수의 필드 이름을 0x02로 구분하여 지정한다. |
| Remark | 예를 들어 oldname에 "title{{0}}\x2title{{1}}", newname에 "tt1\x2tt2"로 지정하면 첫 번째 title은 tt1로, 두 번째 title은 tt2로 변경된다.<br>oldname의 필드 개수와, newname의 필드 개수는 동일해야 한다. 존재하지 않는 필드에 대해서는 무시한다. |

### GetCurFieldName(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 현재 캐럿 위치의 데이터 필드 이름을 구한다. |
| Declaration | `BSTR GetCurFieldName([HwpFieldOption option])` |
| Parameters | option : 다음과 같은 옵션을 지정할 수 있다. 0을 지정하면 모두 off이다. 생략하면 0이 지정된다. (HwpFieldOption : short type) |
| Return | 필드 이름이 돌아온다. 필드 이름이 없는 경우 빈 문자열이 돌아온다. |
| Remark | GetFieldList()의 option 중에 hwpFieldSelection (= 4) 옵션은 사용하지 않는다. |

## Page 18

### GetCurFieldName Option 표

| ID | 값 | 설명 |
|----|---|-----|
| hwpFieldCell | 1 | 셀에 부여된 필드 리스트만을 구한다. hwpFieldClickHere와는 함께 지정할 수 없다. |
| hwpFieldClickHere | 2 | 누름틀에 부여된 필드 리스트만을 구한다. hwpFieldCell과는 함께 지정할 수 없다. |

### SetCurFieldName(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 현재 캐럿 위치의 데이터 필드 이름을 설정한다. |
| Declaration | `BOOL SetCurFieldName(BSTR fieldname, [HwpFieldOption option], [BSTR direction], [BSTR memo])` |
| Parameters | fieldname : 데이터 필드 이름.<br>option : 다음과 같은 옵션을 지정할 수 있다. 0을 지정하면 모두 off이다. 생략하면 0이 지정된다. (HwpFieldOption : short type)<br>direction : 누름틀 필드의 안내문. 누름클 필드일 때만 유효하다.<br>memo : 누름틀 필드의 메모. 누름클 필드일 때만 유효하다. |
| Remark | GetFieldList()의 option 중에 hwpFieldSelection (= 4) 옵션은 사용하지 않는다. |

### ModifyFieldProperties(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 지정한 필드의 속성을 바꾼다. |
| Declaration | `long ModifyFieldProperties(LPCTSTR field, long remove, long add)` |
| Parameters | field : 속성을 바꿀 필드 이름의 리스트. 형식은 PutFieldText과 동일.<br>remove : 제거될 속성<br>add : 추가될 속성<br>return : 음수가 리턴되면 에러임을 나타낸다. |

### 속성 값

| long value | 설명 |
|------------|-----|
| 0x00000001 | 양식모드에서 편집가능 속성 (0: 편집 불가, 1: 편집 가능) |

### return 값의 bit field

| long value bit mask | 설명 |
|---------------------|-----|
| 0x00000001 | 양식모드에서 편집가능 속성 (0: 편집 불가, 1: 편집 가능) |
| 0x80000000 | 에러 |
| 0x40000000 | 필드를 찾을 수 없음 |

**Remark**: remove와 add에 둘다 0이 입력되면 현재 속성을 돌려준다. 리턴값이 음수인지 확인하여 쉽게 에러임을 판별할 수 있으며 자세한 에러내용은 bit mask로 and 연산하여 알아 낼 수 있다. 리턴값은 여러 가지 추가 정보가 같이 올 수 있으므로 반드시 bit mask를 사용하여 비교해야 한다.

## Page 19

### SetFieldViewOption(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 양식 모드와 읽기 전용모드일때 현재 열린 문서의 필드의 겉보기 속성(『』표시)을 바꾼다. |
| Declaration | `long SetFieldViewOption(long option)` |
| Parameters | option : 겉보기 속성 bit<br>return : 설정된 속성이 리턴 된다. 에러일 경우 0 이 리턴 된다. |
| Remark | EditMode와 비슷하게 현재 열려있는 문서에 대한 속성이다. 따라서 저장되지 않는다. |

### SetFieldViewOption 옵션 표

| option | 누름틀 | 개인정보/문서요약/날짜시간/파일경로 | 비고 |
|--------|-------|--------------------------|-----|
| 1 | 『』을 표시하지 않음 | 『』을 표시하지 않음 | |
| 2 | 『』을 빨간색으로 표시 | 『』을 흰색으로 표시 | 설정하지 않았을 때 기본 값 |
| 3 | 『』을 흰색으로 표시 | 『』을 흰색으로 표시 | |

### GetFieldList(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 문서 중의 필드 리스트를 구한다. |
| Declaration | `BSTR GetFieldList([HwpFieldNumber number], [HwpFieldOption option])` |
| Parameters | number : 문서 중에서 동일한 이름의 필드가 여러 개 존재할 때 이를 구별하기 위한 식별 방법을 지정한다. 생략하면 hwpFieldPlain이 지정된다. (HwpFieldNumber : short type)<br>option : 다음과 같은 옵션을 조합할 수 있다. 0을 지정하면 모두 off이다. 생략하면 0이 지정된다. (HwpFieldOption : unsigned short type)<br>return : 각 필드 사이를 문자코드 0x02로 구분하여 다음과 같은 형식으로 리턴 한다. (가장 마지막 필드에는 0x02가 붙지 않는다.)<br>"필드이름#1\x2필드이름#2\x2...필드이름#n" |

### HwpFieldNumber 표

| ID | 값 | 설명 |
|----|---|-----|
| hwpFieldPlain | 0 | 아무 기호 없이 순서대로 필드 이름이 나열된다. |
| hwpFieldNumber | 1 | 필드 이름 뒤에 일련번호가 {{#}}와 같은 형식으로 붙는다. |
| hwpFieldCount | 2 | 필드 이름뒤에 그 이름의 필드가 몇 개 있는지 {{#}}와 같은 형식으로 붙는다. |

### HwpFieldOption 표

| ID | 값 | 설명 |
|----|---|-----|
| hwpFieldCell | 1 | 셀에 부여된 필드 리스트만을 구한다. hwpFieldClickHere와는 함께 지정할 수 없다. |
| hwpFieldClickHere | 2 | 누름틀에 부여된 필드 리스트만을 구한다. hwpFieldCell과는 함께 지정할 수 없다. |
| hwpFieldSelection | 4 | 셀렉션 내에 존재하는 필드 리스트를 구한다. |
| hwpFieldTextbox | 8 | 글상자에 부여된 필드 리스트만을 구한다. hwpFieldCell, hwpFieldClickHere와는 함께 지정할 수 없다. |

**Remark**: 문서 중에 동일한 이름의 필드가 여러 개 존재할 때는 number에 지정한 타입에 따라 3 가지의 서로 다른 방식을 중에서 선택할 수 있다.

## Page 20

### GetFieldList 예시

예를 들어 문서 중 title, body, title, body, footer 순으로 5 개의 필드가 존재할 때, hwpFieldPlain, hwpFieldNumber, HwpFieldCount 세 가지 형식에 따라 다음과 같은 내용이 돌아온다.

| 형식 | 결과 |
|-----|-----|
| hwpFieldPlain | "title\x2body\x2title\x2body\x2footer" |
| hwpFieldNumber | "title{{0}}\x2body{{0}}\x2title{{1}}\x2body{{1}}\x2footer{{0}}" |
| hwpFieldCount | "title{{2}}\x2body{{2}}\x2footer{{1}}" |

### MovePos(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 캐럿의 위치를 옮긴다. |
| Declaration | `BOOL MovePos([unsigned int moveID], [unsigned int para], [unsigned int pos])` |
| Parameters | moveID : 다음과 같은 값을 지정할 수 있다. 생략하면 moveCurList가 지정된다. |

### MovePos moveID 표

| ID | 값 | 설명 |
|----|---|-----|
| moveMain | 0 | 루트 리스트의 특정 위치.(para pos로 위치 지정) |
| moveCurList | 1 | 현재 리스트의 특정 위치.(para pos로 위치 지정) |
| moveTopOfFile | 2 | 문서의 시작으로 이동. |
| moveBottomOfFile | 3 | 문서의 끝으로 이동. |
| moveTopOfList | 4 | 현재 리스트의 시작으로 이동 |
| moveBottomOfList | 5 | 현재 리스트의 끝으로 이동 |
| moveStartOfPara | 6 | 현재 위치한 문단의 시작으로 이동 |
| moveEndOfPara | 7 | 현재 위치한 문단의 끝으로 이동 |
| moveStartOfWord | 8 | 현재 위치한 단어의 시작으로 이동. (현재 리스트만을 대상으로 동작한다.) |
| moveEndOfWord | 9 | 현재 위치한 단어의 끝으로 이동. (현재 리스트만을 대상으로 동작한다.) |
| moveNextPara | 10 | 다음 문단의 시작으로 이동. (현재 리스트만을 대상으로 동작한다.) |
| movePrevPara | 11 | 앞 문단의 끝으로 이동. (현재 리스트만을 대상으로 동작한다.) |
| moveNextPos | 12 | 한 글자 앞으로 이동. (서브 리스트를 옮겨다닐 수 있다.) |
| movePrevPos | 13 | 한 글자 뒤로 이동. (서브 리스트를 옮겨다닐 수 있다.) |
| moveNextPosEx | 14 | 한 글자 앞으로 이동. (서브 리스트를 옮겨다닐 수 있다.) (머리말/꼬리말, 각주/미주, 글상자 포함.) |
| movePrevPosEx | 15 | 한 글자 뒤로 이동. (서브 리스트를 옮겨다닐 수 있다.) (머리말/꼬리말, 각주/미주, 글상자 포함.) |
| moveNextChar | 16 | 한 글자 앞으로 이동. (현재 리스트만을 대상으로 동작한다.) |
