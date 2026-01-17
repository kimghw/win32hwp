# 한글 오토메이션 Event Handler 추가 방법 (MFC/ATL)

> 최종 수정일 : 2025년 4월 15일

---

## 1. 이벤트 핸들러 클래스 추가

### Visual Studio에서 클래스 추가

**경로:** 프로젝트 > 클래스 추가 > Visual C++ > ATL > ATL 단순 개체

### ATL 단순 개체 마법사 설정

| 항목 | 값 |
|------|-----|
| 약식 이름(S) | HwpObjectEventHandler |
| .h 파일(E) | HwpObjectEventHandler.h |
| 클래스(L) | CHwpObjectEventHandler |
| .cpp 파일(C) | HwpObjectEventHandler.cpp |
| CoClass(O) | HwpObjectEventHandler |
| 형식(T) | HwpObjectEventHandler Class |
| 인터페이스(N) | IHwpObjectEventHandler |
| ProgID(I) | (빈 값) |

### 옵션 설정

| 옵션 그룹 | 설정 |
|----------|------|
| 스레딩 모델 | 아파트(A) |
| 인터페이스 | 이중(D) |
| 집합체 | 예(Y) |

### 생성된 클래스 코드

```cpp
// CHwpEventHandler

class ATL_NO_VTABLE CHwpEventHandler :
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CHwpEventHandler, &CLSID_HwpEventHandler>,
    public IDispatchImpl<IHwpEventHandler, &IID_IHwpEventHandler,
           &LIBID_HwpAutomationEventHandlerLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{

BEGIN_COM_MAP(CHwpObjectEventHandler)
    COM_INTERFACE_ENTRY(IHwpObjectEventHandler)
    COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()

...
}

OBJECT_ENTRY_AUTO(__uuidof(HwpObjectEventHandler), CHwpObjectEventHandler)
```

---

## 2. IHwpObjectEvents 인터페이스 구현

1에서 생성한 클래스에 IHwpObjectEvents 인터페이스를 구현합니다.

```cpp
class ATL_NO_VTABLE CHwpEventHandler :
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CHwpEventHandler, &CLSID_HwpEventHandler>,
    public IDispatchImpl<IHwpEventHandler, &IID_IHwpEventHandler,
           &LIBID_HwpAutomationEventHandlerLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
    public IDispatchImpl<IHwpObjectEvents, &__uuidof(IHwpObjectEvents),
           &LIBID_HwpObjectLib, /* wMajor = */ 1>
{

BEGIN_COM_MAP(CHwpEventHandler)
    COM_INTERFACE_ENTRY(IHwpEventHandler)
    COM_INTERFACE_ENTRY2(IDispatch, IHwpObjectEvents)
    COM_INTERFACE_ENTRY(IHwpObjectEvents)
END_COM_MAP()

    // IHwpObjectEvents Methods
    STDMETHOD(Quit)()
    {
        OutputDebugString(_T("[HWP-EVENT] ObjectEvent Quit"));
        return E_NOTIMPL;
    }

    STDMETHOD(CreateXHwpWindow)()
    {
        return E_NOTIMPL;
    }

    STDMETHOD(CloseXHwpWindow)()
    {
        OutputDebugString(_T("[HWP-EVENT] ObjectEvent CloseXHwpWindow"));
        return E_NOTIMPL;
    }

    STDMETHOD(NewDocument)(long newVal)
    {
        return E_NOTIMPL;
    }

    STDMETHOD(DocumentBeforeClose)(long newVal)
    {
        OutputDebugString(_T("[HWP-EVENT] ObjectEvent DocumentBeforeClose"));
        return E_NOTIMPL;
    }

    STDMETHOD(DocumentBeforeOpen)(long newVal)
    {
        OutputDebugString(_T("[HWP-EVENT] ObjectEvent DocumentBeforeOpen"));
        return E_NOTIMPL;
    }

    STDMETHOD(DocumentAfterOpen)(long newVal)
    {
        OutputDebugString(_T("[HWP-EVENT] ObjectEvent DocumentAfterOpen"));
        return E_NOTIMPL;
    }

    STDMETHOD(DocumentBeforeSave)(long newVal)
    {
        OutputDebugString(_T("[HWP-EVENT] ObjectEvent DocumentBeforeSave"));
        return E_NOTIMPL;
    }

    STDMETHOD(DocumentAfterSave)(long newVal)
    {
        OutputDebugString(_T("[HWP-EVENT] ObjectEvent DocumentAfterSave"));
        return E_NOTIMPL;
    }

    STDMETHOD(DocumentAfterClose)(long newVal)
    {
        OutputDebugString(_T("[HWP-EVENT] ObjectEvent DocumentAfterClose"));
        return E_NOTIMPL;
    }

    STDMETHOD(DocumentChange)(long newVal)
    {
        OutputDebugString(_T("[HWP-EVENT] ObjectEvent DocumentChange"));
        return E_NOTIMPL;
    }

    ...
}
```

> **참고:** IHwpObjectEvents의 GUID value을 알기 위해서는 `#import` 속성에 `named_guids`를 추가해야 합니다.

```cpp
#import "progid:HWPFrame.HwpObject.2" raw_interfaces_only, raw_native_types, no_namespace, named_guids, auto_search
```

---

## 3. 이벤트 핸들러 연결

### 멤버 변수 선언

```cpp
CComObject<CHwpObjectEventHandler>* m_pHwpEventHandler;
IConnectionPoint* m_pCP;
DWORD m_dwHwpObjectEventHandlerCookie;
```

### 이벤트 핸들러 연결 코드

```cpp
COleException exception;

if (m_app.CreateDispatch(_T("HWPFrame.HwpObject.2"), &exception) != NULL) {
    CXHwpWindows wins = m_app.get_XHwpWindows();
    CXHwpWindow win = wins.get_Active_XHwpWindow();
    win.put_Visible(TRUE);

    IConnectionPointContainer* pCPC = NULL;
    m_app.m_lpDispatch->QueryInterface(IID_IConnectionPointContainer, (void**)&pCPC);

    if (pCPC) {
        pCPC->FindConnectionPoint(__uuidof(IHwpObjectEvents), &m_pCP);

        if (m_pCP) {
            CComObject<CHwpObjectEventHandler>::CreateInstance(&m_pHwpEventHandler);

            if (m_pHwpEventHandler) {
                m_pHwpEventHandler->SetContainerInfo(&m_app);

                LPUNKNOWN p;
                m_pHwpEventHandler->QueryInterface(IID_IDispatch, (void**)&p);
                m_pCP->Advise(p, &m_dwHwpObjectEventHandlerCookie);
            }
        }
    }
}
```

---

## 4. 이벤트 핸들러 연결 해제

```cpp
if (m_pCP) {
    m_pCP->Unadvise(m_dwHwpObjectEventHandlerCookie);
}

if (m_pHwpEventHandler) {
    m_pHwpEventHandler->Release();
}
```

---

## 참고 자료

- [#import Directive (C++) - Microsoft Docs](https://learn.microsoft.com/en-us/cpp/preprocessor/hash-import-directive-cpp?view=msvc-170)
- [IConnectionPoint Interface - Microsoft Docs](https://learn.microsoft.com/en-us/windows/win32/api/ocidl/nn-ocidl-iconnectionpoint)
