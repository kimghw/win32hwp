# 한글 OLE Automation

최종 수정일 : 2025년 4월 15일

## 1.1. 한글 Objet Model

```
IHwpObject
├── IXHwpDocuments                    ├── IXHwpWindows
│   └── IXHwpDocument                 │   └── IXHwpWindow
│       ├── IXHwpFormPushButtons      │       ├── IXHwpTabs
│       │   └── IXHwpFormPushButton   │       │   └── IXHwpTab
│       ├── IXHwpFormCheckButtons     │       ├── IXHwpToobarLayout
│       │   └── IXHwpFormCheckButton  │       │   ├── IXHwpToolbar
│       ├── IXHwpFormRadioButtons     │       │   ├── IXHwpToolbarButton
│       │   └── IXHwpFormRadioButton  │       │   └── IXHwpToolBarMenuButton
│       ├── IXHwpFormComboBoxs        │
│       │   └── IXHwpFormComboBox     ├── IXHwpMessageBox
│       ├── IXHwpFormEdits            ├── IXHwpODBC
│       │   └── IXHwpFormEdit         ├── HAction
│                                     │   └── HParameterSet
│                                     │       └── H[SetObjects]
│                                     │           ├── HSet
│                                     │           └── HArray
│
│                                     ├── IDHwpAction
│                                     ├── IDHwpParameterSet
│                                     ├── IDHwpParameterArray
Interface : Single Object             └── IDHwpCtrlCode
Interface : Collection Object
```

## 1.2. 한글 Objets

### OLE Automation Object

Automation Object는 Automation Server가 Automation Client에게 제공하는 프로퍼티, 메소드, 이벤트로 구성된 개체이다. IHwpObject는 한글 Automation Server의 최상위 오브젝트이며 단일 오브젝트와 Collection 오브젝트로 구성된 하위의 오브젝트는 모두 IHwpObject로부터 생성된다.

#### JavaScript에서 IHwpObject 최상위 개체 생성

```javascript
var App = new ActiveXObject("HwpFrame.HwpObject.2"); //IXHwpObject 생성
var Docs = App.XHwpDocuments;
```

#### Visual C++에서 IHwpObject 최상위 개체 생성

```cpp
IHwpObject hwpobject;
if (hwpobject.CreateDispatch("HwpFrame.HwpObject.2")) {
    MessageBox("한글에 연결을 성공하였습니다.", "성공", MB_OK);
} else {
    MessageBox("한글에 연결을 실패하였습니다.", "오류", MB_OK);
    PostQuitMessage(0);
}
```

### OLE Automation Collection Object

OLE Automation 단일 Object를 그룹화 하는 Object

한글에서 새글 또는 새탭으로 열려있는 문서는 IXHwpDocument Object에서 제어 가능하며 열려있는 문서를 모두 관리 할 수 있는 것은 IXHwpDocuments Collection Object이다. IXHwpDocument Object는 IXHwpDocuments의 Add 메소드에 의하여 생성 될 수 있으며 생성된 Object는 Item 프로퍼티에 의하여 구할 수 있다.

### Property

OLE Automation Object의 상태에 대한 정보를 제어할 수 있는 기능을 제공

IXHwpDocument의 EditMode 프로퍼티는 지정된 문서가 읽기 모드인지 일반 편집 모드인지 부분 편집 모드인지를 얻을 수 있거나 설정할 수 있다.

### Method

OLE Automation Object가 특정한 동작을 수행하도록 하는 기능을 제공

IXHwpMessageBox의 DoModal 메소드는 메시지 박스를 보여준다.

### Event

OLE Automation Object에서 Client로 특정 사건이 발생했음을 알려주는 기능을 제공 하며 IHwpObject에서 이벤트를 얻을 수 있다.

#### Visual Basic에서 IHwpObject 최상위 개체로부터 종료 이벤트 얻기

```vb
Private WithEvents objHwpObject As HwpObject
Private Sub Form_Load()
    Set objHwpObject = New HwpObject
    MsgBox "한글에 접속"
End Sub

Private Sub objHwpObject_Quit()
    MsgBox "한글로부터 종료 이벤트를 받았습니다."
End Sub
```

## 1.3. 한글 내부 개체

### 내부 개체

한글은 Automation Client가 제어 할 수 있듯이 한글 내부에서도 스크립트로 제어할 수 있다. 스크립트는 JavaScript를 사용한다. 내부에서 사용되므로 IHwpObject는 이미 생성 되어있고 스크립트에서 최상위 IHwpObject 생성을 위한 코드는 필요가 없다.

#### 내부 스크립트에서 한글 제어 - JavaScript를 이용한 예제

```javascript
//IHwpObject는 이미 생성되어있으므로 IHwpObject의 하위 항목은
//XHwpDocuments와 XHwpMessageBox같이 바로 사용 가능하다.
function OnDocument_Open()
{
    var Docs = XHwpDocuments;
    Docs.Add(1); //새 창으로 문서 생성 (0 = 새 창, 1 = 새 탭)
    var count = Docs.Count; //생성된 Document의 개수 얻기
    var msgbox = XHwpMessageBox;
    msgbox.String = count;
    msgbox.DoModal();
}
```

### 양식 개체와 내부 개체

한글은 양식 개체와 내부 개체를 연동하여 사용할 수 있다. 양식 개체는 PushButton, CheckBox, ComboBox, RadioButton, Edit가 있다. 각 양식 개체는 Automation Object로 되어있으므로 스크립트로 제어가 가능하며 양식개체의 특정 이벤트에따라 스크립트가 수행되도록 할 수 있다.

#### 양식개체 스크립트 적용 예

```javascript
function OnPushButton_Click() //PushButton을 클릭 했을때 수행되는 함수
{
    //아래의 RadioButton, CheckBox, ComboBox등의 Name은
    //속성 창에 Name을 보면 알 수 있다.
    RadioButton.BackColor = 0xddaacc; //라이오버튼의 색깔을 변경
    CheckBox.Value = 1; //체크박스를 켜둔 상태
    ComboBox.InsertString("false", 0); //콤보박스에 false, true 문자 입력
    ComboBox.InsertString("true", 1);
    Edit.Text ="테스트 입니다."; //에디트에 문자 입력
}
```

## 1.4. 외부 클라이언트와 한글

한글은 외부 클라이언트(응용 프로그램)와 연결되어 외부 클라이언트는 한글을 제어할 수 있다. 외부 클라이언트가 한글에 접속을 요청할 때 한글은 자동으로 실행 된다.

#### 사용된 스크립트 - 스크립트 매크로로 생성

```javascript
HAction.Run("SelectAll");
HAction.GetDefault("CharShape", HParameterSet.HCharShape.HSet);
HParameterSet.HCharShape.FaceNameHangul = "궁서체";
HParameterSet.HCharShape.FaceNameLatin = "궁서체";
HParameterSet.HCharShape.FaceNameHanja = "궁서체";
HParameterSet.HCharShape.FaceNameJapanese = "궁서체";
HParameterSet.HCharShape.FaceNameOther = "궁서체";
HParameterSet.HCharShape.FaceNameSymbol = "궁서체";
HParameterSet.HCharShape.FaceNameUser = "궁서체";
HParameterSet.HCharShape.Height = 4000;
HParameterSet.HCharShape.TextColor = 16737792;
HParameterSet.HCharShape.UnderlineType = 1;
HParameterSet.HCharShape.ShadowType = 1;
HAction.Execute("CharShape", HParameterSet.HCharShape.HSet);
HAction.Run("Cancel");
```

#### 적용 코드 (Visual C++ MFC)

```cpp
void CTestHwpOleDlg::OnButton1()
{
    CDHwpParameterSet set;
    CDHwpAction dAction = myHwpObj.CreateAction(L"CharShape");
    set = dAction.CreateSet();
    dAction.GetDefault(set);

    myHwpObj.Run(L"SelectAll");
    set.SetItem(L"FaceNameHangul", COleVariant((LPCWSTR)"궁서체"));
    set.SetItem(L"FaceNameLatin", COleVariant((LPCWSTR)"궁서체"));
    set.SetItem(L"FaceNameHanja", COleVariant((LPCWSTR)"궁서체"));
    set.SetItem(L"FaceNameJapanese", COleVariant((LPCWSTR)"궁서체"));
    set.SetItem(L"FaceNameOther", COleVariant((LPCWSTR)"궁서체"));
    set.SetItem(L"FaceNameSymbol", COleVariant((LPCWSTR)"궁서체"));
    set.SetItem(L"FaceNameUser", COleVariant((LPCWSTR)"궁서체"));

    set.SetItem(L"FontTypeHangul", COleVariant((LPCWSTR)myHwpObj.FontType(L"TTF")));
    set.SetItem(L"FontTypeLatin", COleVariant((LPCWSTR)myHwpObj.FontType(L"TTF")));
    set.SetItem(L"FontTypeHanja", COleVariant((LPCWSTR)myHwpObj.FontType(L"TTF")));
    set.SetItem(L"FontTypeJapanese", COleVariant((LPCWSTR)myHwpObj.FontType(L"TTF")));
    set.SetItem(L"FontTypeOther", COleVariant((LPCWSTR)myHwpObj.FontType(L"TTF")));
    set.SetItem(L"FontTypeSymbol", COleVariant((LPCWSTR)myHwpObj.FontType(L"TTF")));
    set.SetItem(L"FontTypeUser", COleVariant((LPCWSTR)myHwpObj.FontType(L"TTF")));

    set.SetItem(L"Height", COleVariant((long)myHwpObj.PointToHwpUnit(40)));
    set.SetItem(L"TextColor", COleVariant((long)myHwpObj.RGBColor(255, 0, 0)));
    set.SetItem(L"UnderlineType", COleVariant((long)1));

    dAction.Execute(set);
    myHwpObj.Run(L"Cancel");
}
```

## 1.5. 외부 클라이언트와 실행 중인 한글

이미 실행 중인 한글 에 외부 클라이언트가 접속하여 한글을 제어할 수 있다. (Running Object Table을 액세스 할 수 있는 프로그래밍 언어로 작성 된 프로그램에서 가능함)

#### 예제 코드

```cpp
void CTestRunningHwpDlg::OnConnectButton()
{
    WCHAR szobjectname[] = L"!HwpObject.130";
    WCHAR* szname = NULL;

    IEnumMoniker* pEnumMoniker = NULL;
    IRunningObjectTable* pRot = NULL;
    IMoniker* iMoniker = NULL;
    IUnknown* iUnknown = NULL;
    IBindCtx* iBindCtx = NULL;
    ULONG uelement = 0;

    CHwpObject cActiveHwpObj;

    HRESULT hr = GetRunningObjectTable(0, &pRot);
    if (FAILED(hr)) {
        return;
    }

    hr = pRot->EnumRunning(&pEnumMoniker);
    if (FAILED(hr)) {
        pRot->Release();
        return;
    }

    hr = CreateBindCtx(0, &iBindCtx);
    if (FAILED(hr)) {
        pEnumMoniker->Release();
        pRot->Release();
        return;
    }

    while (SUCCEEDED(pEnumMoniker->Next(1, &iMoniker, &uelement))) {
        if (uelement == 0)
            return;

        pRot->GetObject(iMoniker, &iUnknown);
        iMoniker->GetDisplayName(iBindCtx, NULL, (OLECHAR**)&szname);

        CString strobject;
        strobject = szname;

        if (strobject.Find(szobjectname, 0) != -1) {
            LPDISPATCH pdisp;
            iUnknown->QueryInterface(IID_IDispatch, (void**)&pdisp);
            if (pdisp) {
                cActiveHwpObj.AttachDispatch(pdisp);
                cActiveHwpObj.Run(L"SelectAll");
            }
        }
    }
}
```

## 1.6. 스크립트 매크로를 이용한 코드 적용 예제

### 스크립트 매크로 기록으로 코드 생성

1. "글자입력" 문자열 입력
2. 전체 블록 설정
3. 글자모양에서 색깔을 파란색으로 설정
4. 블록 설정 해제

### 스크립트 매크로로 생성된 스크립트 코드

```javascript
function OnScriptMacro_script67()
{
    HAction.GetDefault("InsertText", HParameterSet.HInsertText.HSet);
    HParameterSet.HInsertText.Text = "글자입력";
    HAction.Execute("InsertText", HParameterSet.HInsertText.HSet);
    HAction.Run("SelectAll");
    HAction.GetDefault("CharShape", HParameterSet.HCharShape.HSet);
    with (HParameterSet.HCharShape)
    {
        FontTypeUser = FontType("TTF");
        FontTypeSymbol = FontType("TTF");
        FontTypeOther = FontType("TTF");
        FontTypeJapanese = FontType("TTF");
        FontTypeHanja = FontType("TTF");
        FontTypeLatin = FontType("TTF");
        FontTypeHangul = FontType("TTF");
        TextColor = RGBColor(0, 0, 255);
    }
    HAction.Execute("CharShape", HParameterSet.HCharShape.HSet);
    HAction.Run("Cancel");
}
```
