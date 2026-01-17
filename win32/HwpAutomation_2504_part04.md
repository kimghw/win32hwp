# HwpAutomation_2504_part04

## Page 31

### ImportStyle(Method)

| 항목 | 내용 |
|-----|-----|
| Description | HStyleTemplate 파라메터셋 오브젝트에 지정된 스타일을 Import한다. |
| Declaration | `BOOL ImportStyle(LPDISPATCH param)` |
| Parameters | param : HStyleTemplate |

### FindCtrl(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 현재 캐럿의 위치에서 Ctrl을 찾는다. |
| Declaration | `BSTR FindCtrl()` |
| Parameters | return : 컨트롤을 찾은 경우 CtrlID를 return 한다. |

### UnSelectCtrl(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 현재 Select된 Ctrl의 Selection을 해제한다. |

### IsActionEnable(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 액션이 실행 가능한 상태인지 확인한다. |
| Declaration | `bool IsActionEnable(BSTR actionID)` |
| Parameters | actionID: 컨트롤을 찾은 경우 CtrlID를 return 한다.<br>- Action ID는 ActionObject.hwp(별도문서)를 참고한다. |

### GetScriptSource(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 한글에서 사용하는 매크로의 소스코드를 가져온다. |
| Declaration | `BSTR GetScriptSource(BSTR FileName);` |
| Parameters | FileName : 매크로 소스를 가져올 한글문서 |

### GetFileInfo(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 파일 정보를 알아낸다. |
| Declaration | `LPDISPATCH* GetFileInfo(BSTR FileName);` |
| Parameters | FileName : 정보를 구하고자 하는 파일명<br>반환값(Return) : "FileInfo" ParameterSet이 반환된다. |

### GetFileInfo 반환 ParameterSet

| Item ID | Type | Description |
|---------|------|-------------|
| Format | PIT_BSTR | 파일 형식 |
| VersionStr | PIT_BSTR | 파일 버전 문자열 ex) 13.0.0.1 |
| VersionNum | PIT_UI4 | 파일 버전 ex) 0x0d000001 |
| Encrypted | PIT_UI1 | 암호화 여부<br>-1 : 판단 할 수 없음<br>0 : 암호가 걸려 있지 않음<br>양수: 암호가 걸려 있음. |
| Compressed | PIT_UI1 | 압축 여부 |

### RunScriptMacro(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 한글문서 내에 존재하는 매크로를 실행한다. |
| Declaration | `bool RunScriptMacro(BSTR FunctionName, long uMacroType, long uScriptType)` |
| Parameters | FunctionName : 실행할 매크로 함수이름 |

## Page 32

| 항목 | 내용 |
|-----|-----|
| Parameters (계속) | uMacroType : 매크로의 유형. 밑의 값 중 하나이다.<br>uScriptType : 스크립트의 유형. 현재는 JScript만을 유일하게 지원한다. reserved. |
| Remark | 한글은 매크로 기능을 지원하고 있다. 이렇게 작성된 매크로를 HwpAutomation으로 실행시킬 수 있도록 지원하는 함수이다. |

### RunScriptMacro uMacroType 표

| ID | 값 | Description |
|----|---|-------------|
| HWP_GLOBAL_MACRO_TYPE | 0 | 전역 매크로 |
| HWP_DOCUMENT_MACRO_TYPE | 1 | 문서에만 적용되는 매크로 |

### GetPageText(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 페이지 단위의 텍스트 추출 |
| Declaration | `BSTR GetPageText(long pgno, VARIANT option);` |
| Parameters | pgno : 텍스트를 추출 할 페이지의 번호(0부터 시작)<br>option : 추출 대상을 다음과 같은 옵션을 조합하여 지정할 수 있다. 생략(또는 0xffffffff)하면 모든 텍스트를 추출한다. |
| Remark | 일반 텍스트(글자처럼 취급 도형 포함)를 우선적으로 추출하고, 도형(표, 글상자) 내의 텍스트를 추출한다. |

### GetPageText option 표

| ID | 값 | Description |
|----|---|-------------|
| maskNormal | 0x00 | 본문 텍스트를 추출한다.(항상 기본) |
| maskTable | 0x01 | 표에대한 텍스트를 추출한다. |
| maskTextbox | 0x02 | 글상자 텍스트를 추출한다. |
| maskCaption | 0x04 | 캡션 텍스트를 추출한다. (표, ShapeObject) |

### SetBarCodeImage(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 바코드 이미지를 삽입한다. |
| Declaration | `bool SetBarCodeImage(BSTR lpImagePath, long pgno, long index, long x, long y, long width, long height)` |
| Parameters | ImagePath : 바코드 이미지 경로<br>pgno : 바코드 이미지를 삽입 할 페이지 번호(0부터 시작)<br>index : 페이지에 삽입될 바코드 이미지의 번호(0부터 시작)<br>- 같은 번호를 지정하면 이전에 삽입된 바코드의 이미지가 수정된다.<br>- 바코드 이미지를 공백으로 지정하면 해당 번호의 바코드 이미지가 삭제된다.<br>x : 바코드 이미지가 삽입 될 위치의 x좌표 (mm단위)<br>- 바코드 이미지의 좌측 상단 기준으로 페이지(종이 기준) 좌측 상단에서의 거리이다.<br>y : 바코드 이미지가 삽입 될 위치의 y좌표 (mm단위)<br>- 바코드 이미지의 좌측 상단 기준으로 페이지(종이 기준) 좌측 상단에서의 거리이다.<br>width : 바코드 이미지의 너비. 생략하면 원본 이미지의 너비가 적용된다.<br>height : 바코드 이미지의 높이. 생략하면 원본 이미지의 높이가 적용된다. |

### GetMessageBoxMode(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 현재 메시지 박스의 Mode를 long으로 얻어온다. |
| Parameters | return : 메시지 박스의 종류 |

## Page 33

### 메시지 박스 정의

```cpp
// 메시지 박스의 종류
#define MB_MASK                         0x00FFFFFF

// 1. 확인(MB_OK) : IDOK(1)
#define MB_OK_IDOK                      0x00000001
#define MB_OK_MASK                      0x0000000F

// 2. 확인/취소(MB_OKCANCEL) : IDOK(1), IDCANCEL(2)
#define MB_OKCANCEL_IDOK                0x00000010
#define MB_OKCANCEL_IDCANCEL            0x00000020
#define MB_OKCANCEL_MASK                0x000000F0

// 3. 종료/재시도/무시(MB_ABORTRETRYIGNORE) : IDABORT(3), IDRETRY(4), IDIGNORE(5)
#define MB_ABORTRETRYIGNORE_IDABORT     0x00000100
#define MB_ABORTRETRYIGNORE_IDRETRY     0x00000200
#define MB_ABORTRETRYIGNORE_IDIGNORE    0x00000400
#define MB_ABORTRETRYIGNORE_MASK        0x00000F00

// 4. 예/아니오/취소(MB_YESNOCANCEL) : IDYES(6), IDNO(7), IDCANCEL(2)
#define MB_YESNOCANCEL_IDYES            0x00001000
#define MB_YESNOCANCEL_IDNO             0x00002000
#define MB_YESNOCANCEL_IDCANCEL         0x00004000
#define MB_YESNOCANCEL_MASK             0x0000F000

// 5. 예/아니오(MB_YESNO) : IDYES(6), IDNO(7)
#define MB_YESNO_IDYES                  0x00010000
#define MB_YESNO_IDNO                   0x00020000
#define MB_YESNO_MASK                   0x000F0000

// 6. 재시도/취소(MB_RETRYCANCEL) : IDRETRY(4), IDCANCEL(2)
#define MB_RETRYCANCEL_IDRETRY          0x00100000
#define MB_RETRYCANCEL_IDCANCEL         0x00200000
#define MB_RETRYCANCEL_MASK             0x00F00000
```

### SetMessageBoxMode(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 한글에서 쓰는 다양한 메시지박스가 뜨지 않고, 자동으로 특정 버튼을 클릭한 효과를 주기 위해 사용한다. |
| Declaration | `long SetMessageBoxMode(long mode);` |
| Parameters | mode : GetMessageBoxMode 참고<br>return : 기존에 적용되어 있던 oldMode 값 |

### GetBinDataPath(Method)

| 항목 | 내용 |
|-----|-----|
| Description | Binary Data의 경로를 가져온다. |
| Declaration | `BSTR GetBinDataPath(long binid);` |
| Parameters | 바이너리 데이터의 ID 값 (1부터 시작) |

### SetDRMAuthority(Method)

| 항목 | 내용 |
|-----|-----|
| Description | DRM모듈에 권한을 설정한다. |
| Declaration | `bool SetDRMAuthority(long authority);` |
| Parameters | authority : |

### SetDRMAuthority authority 표

| 값 | Description |
|---|-------------|
| 0x0001 | 읽기 |
| 0x0002 | 편집 |

## Page 34

| 값 | Description |
|---|-------------|
| 0x0004 | 인쇄 |
| 0x0008 | 클립보드 |
| 0x0010 | 같은 프로세스 내에서도 시스템 클립보드를 사용하기 위한 권한 |
| 0x0020 | 클립보드 히스토리 사용하지 않도록 하기 위한 권한 |

### CreateSet(Method)

| 항목 | 내용 |
|-----|-----|
| Description | ParameterSet을 생성한다. |
| Declaration | `LPDISPATCH* CreateSet(BSTR setID);` |
| Parameters | setID : 생성할 ParameterSet의 ID |

### GetHeadingString(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 글머리표/문단번호/개요번호를 추출한다. |
| Parameters | return : 문자열이 반환된다.<br>- 글머리표/문단번호/개요번호가 있다면 해당 문자열이 반환된다. |
| Remark | - 글머리표/문단번호/개요번호가 있는 경우, 해당 문자열을 얻어올 수 있다.<br>- 문단에 글머리표/문단번호/개요번호가 없는 경우, 빈 문자열이 추출된다. |

### SetTitleName(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 사용자가 임의로 Title을 설정한다. |
| Declaration | `void SetTitleName(BSTR title);` |
| Parameters | title : 변경할 사용자 Title문자열 |

### GetSelectedPos(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 현재 설정된 블록의 위치정보를 얻어온다. |
| Declaration | `bool GetSelectedPos(long* slist, long* spara, long* spos, long* elist, long* epara, long* epos);` |
| Parameters | slist : 설정된 블록의 시작 리스트 아이디.<br>spara : 설정된 블록의 시작 문단 아이디.<br>spos : 설정된 블록의 문단 내 시작 글자 단위 위치.<br>elist : 설정된 블록의 끝 리스트 아이디.<br>epara : 설정된 블록의 끝 문단 아이디.<br>epos : 설정된 블록의 문단 내 끝 글자 단위 위치. |
| Remark | - 위의 리스트란, 문단과 컨트롤들이 연결된 한글문서 내 구조를 뜻한다.<br>- 리스트 아이디는 문서 내 위치정보중 하나로서 SelectText에 넘겨줄 때 사용한다.<br>- 매개변수로 포인터를 사용하므로, 포인터를 사용할 수 없는 언어에서는 사용이 불가능 하다.<br>- 포인터를 사용하지 않는 언어를 지원하기 위해서 ParameterSet을 사용하는 GetSelectedPosBySet()이 존재한다. |

### GetSelectedPosBySet(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 현재 설정된 블록의 위치정보를 얻어온다. (GetSelectedPos의 ParameterSet버전) |

## Page 35

| 항목 | 내용 |
|-----|-----|
| Declaration | `bool GetSelectedPosBySet(LPDISPATCH sset, LPDISPATCH eset);` |
| Parameters | sset : 설정된 블록의 시작 파라메터셋 (ListParaPos)<br>eset : 설정된 블록의 끝 파라메터셋 (ListParaPos) |

### SetPrivateInfoPassword(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 개인정보보호를 위한 암호를 등록한다. |
| Declaration | `bool SetPrivateInfoPassword(BSTR password);` |
| Parameters | password : 새 암호<br>return : true / false<br>- 정상적으로 암호가 설정되면 true를 반환한다.<br>- 암호설정에 실패하면 false를 반환한다. false를 반환하는 경우는 다음과 같다<br>- 암호의 길이가 너무 짧거나 너무 길 때 (영문: 5~44자, 한글: 3~22자)<br>- 암호가 이미 설정되었음. 또는 암호가 이미 설정된 문서임 |
| Remark | 개인정보 보호를 설정하기 위해서는 우선 개인정보 보호 암호를 먼저 설정해야 한다. 그러므로 개인정보 보호 함수를 실행하기 이전에 반드시 이 함수를 호출해야 한다. |

### RegisterPrivateInfoPattern(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 개인정보의 패턴을 등록한다. |
| Declaration | `bool RegisterPrivateInfoPattern(long PrivateType, BSTR PrivatePattern);` |
| Parameters | PrivateType : 등록할 개인정보 유형. 다음의 값 중 하나다.<br>PrivatePatterns : 등록할 개인정보 패턴. 예를 들면 이런 형태로 입력한다.<br>- (예) 주민등록번호 - "NNNNNN-NNNNNNN"<br>- 한/글이 이미 정의한 패턴은 정의하면 안 된다.<br>- 함수를 여러 번 호출하는 것을 피하기 위해 패턴을 ";"기호로 구분 반속해서 입력할 수 있도록 한다. |
| Remark | - 찾아낼 개인정보 패턴을 정의한다. (찾아서 보호 기능에서 사용됨)<br>- 한/글이 이미 정의하고 있는 패턴은 새로 등록할 수 없다. PrivateType으로 등록할 개인정보 유형을 선택한 뒤 PrivatePattern으로 패턴을 정의하는 식이다.<br>- 한/글이 기본적으로 등록한 패턴은 새로 등록할 수 없다. 또한 이미 등록한 패턴 역시 등록이 실패하며 false 반환한다.<br>- 한글은 패턴을 등록할 수 있는 단어를 제한해두고 있다. 사용할 수 있는 문자는 'N', '-', '.' 빈 칸 등이며, IP 주소 유형의 경우에만 '@'를 허용한다. 잘못된 문자를 사용하여 패턴을 등록하면 false를 반환한다.<br>- 마지막으로 동일 유형에는 여러 패턴을 반복적으로 입력할 수 있다. 이때 함수를 한번만 호출할 수 있도록 ';' 구분자를 두어 반복해서 등록한다. |

### RegisterPrivateInfoPattern PrivateType 표

| 값 | Description |
|---|-------------|
| 0x0001 | 전화번호 |
| 0x0002 | 주민등록번호 |
| 0x0004 | 외국인등록번호 |
| 0x0008 | 전자우편 |
| 0x0010 | 계좌번호 |
| 0x0020 | 신용카드번호 |
| 0x0040 | IP 주소 |
| 0x0080 | 생년월일 |
| 0x0100 | 주소 |
| 0x0200 | 사용자 정의 |
| 0x0400 | 기타 - 기타는 등록할 수 없다 |

## Page 36

### FindPrivateInfo(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 개인정보를 찾는다. |
| Declaration | `long FindPrivateInfo(long PrivateType, VARIANT PrivateString);` |
| Parameters | PrivateTypes : 찾을 개인정보 유형. 다음의 값을 하나이상 조합한다.<br>PrivateString : 기타 문자열. 0x0400 유형이 존재할 경우에만 유효하므로, 생략가능하다<br>return : 찾은 개인정보의 유형 값. 값의 의미는 위 표와 동일하다.<br>- 개인정보가 없는 경우에는 0을 반환한다.<br>- 또한, 검색 중 문서의 끝(end of document)을 만나면 -1을 반환한다. 이는 함수가 무한히 반복하는 것을 막아준다. |

### FindPrivateInfo PrivateTypes 표

| 값 | Description |
|---|-------------|
| 0x0001 | 전화번호 |
| 0x0002 | 주민등록번호 |
| 0x0004 | 외국인등록번호 |
| 0x0008 | 전자우편 |
| 0x0010 | 계좌번호 |
| 0x0020 | 신용카드번호 |
| 0x0040 | IP 주소 |
| 0x0080 | 생년월일 |
| 0x0100 | 주소 |
| 0x0200 | 사용자 정의 |
| 0x0400 | 기타 |

**Remark**
- 개인 정보 유형으로 등록된 정보를 찾아 이동한다. 동작방식은 한/글의 찾기와 비슷하다.
- 함수가 호출되면 현재 캐럿 위치에서부터 개인정보를 검색하기 시작한다. 개인정보가 검색되면 검색된 개인정보를 블록 선택한다. 만약 문서의 끝까지 개인정보가 검색되지 않을 경우에는 함수는 -1을 반환하여 무한으로 검색하는 것을 막는다.
- PrivateTypes는 검색할 개인정보 유형을 나타낸다. 단 하나의 유형을 검색할 수도 있지만 여러 유형을 bit flag로 조합해서 검색할 수 있다. 예를 들면, 주민등록번호와 외국인등록번호를 한꺼번에 검색할 수 있다.
- PrivateTypes에 0(Null)을 줄 수 있는데 이것은 기타유형을 제외한 모든 유형을 검색하게 한다. (0x03FF 값을 입력한 것과 동일함)
- PrivateTypes의 값 중 0x400 값은 단어를 직접 검색할 때 사용한다. 0x400 값은 뒤의 PrivateString과 조합되어 사용될 수 있으며, 다음과 같이 사용된다.
- 예) 신용카드번호 또는 단어 "신한카드"를 직접 검색
- `FindPrivateInfo(0x20+0x400, "신한카드");`
- FindPrivateInfo() 함수는 리턴값으로 현재 선택된 개인정보의 유형을 반환한다. 사용자는 리턴값을 확인함으로 선택된 개인정보의 유형을 확인할 수 있으며, 선택적으로 정보보호가 가능하다.
- 함수는 문서 내 개인정보가 없을 경우에는 0을 반환한다. 그리고, 문서의 끝에 도달한 경우에는 -1을 반환한다. 이 두 값은 이전 리턴값보다 의미상 중요한데, 이유는 루프문의 탈출코드로 사용되기 때문이다.

### ProtectPrivateInfo(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 개인정보를 보호한다. |

## Page 37

| 항목 | 내용 |
|-----|-----|
| Declaration | `bool ProtectPrivateInfo(BSTR ProtectingChar, VARIANT PrivatePatternType);` |
| Parameters | ProtectingChar : 보호문자. 개인정보는 해당문자로 가려진다.<br>ProtectingType : 보호유형. 개인정보 유형마다 설정할 수 있는 값이 다르다.<br>- 0값은 기본 보호유형으로 모든 개인정보를 보호문자로 보호한다. |

### 개인정보 유형에 따른 보호유형

| 개인정보 유형 | 보호 유형 | 보호유형 형태 |
|---------|--------|---------|
| 전화번호 (0x0001) | 0 | ************* |
| | 1 | NNN-****-NNNN |
| | 2 | NNN-NNN-**** |
| 주민등록번호 (0x0002) | 0 | ************** |
| | 1 | NNNNNN-******* |
| | 2 | NNNNNN-N****** |
| | 3 | NNNNNN-N*****N |
| 외국인등록번호 (0x0004) | 0 | ************** |
| | 1 | NNNNNN-******* |
| | 2 | NNNNNN-N****** |
| | 3 | NNNNNN-N*****N |
| 전자우편 (0x0008) | 0 | ************ |
| | 1 | ***@TT.TT.TT |
| | 2 | T**@TT.TT.TT |
| | 3 | TT*@TT.TT.TT |
| | 4 | TTT@***.***.**** |
| 계좌번호 (0x0010) | 0 | *************** |
| | 1 | NNN-**-****-NNN |
| | 2 | ***-NN-NNNN-*** |
| | 3 | NNN-**-NNNN-*** |
| | 4 | NNN-NN-****-*** |
| 신용카드번호 (0x0020) | 0 | ******************* |
| | 1 | NNNN-****-****-NNNN |
| | 2 | NNNN-****-NNNN-**** |
| | 3 | ****-NNNN-NNNN-**** |
| | 4 | NNNN-NNNN-****-**** |
| IP 주소 (0x0040) | 0 | ********* |
| | 1 | NNN.*.*.N |
| | 2 | NNN.*.N.* |
| | 3 | ***.N.N.* |
| 생년월일 (0x0080) | 0 | ********** |
| | 1 | NNNN-**-NN |
| | 2 | NNNN-**-** |

## Page 38

| 개인정보 유형 | 보호 유형 | 보호유형 형태 |
|---------|--------|---------|
| 생년월일 (0x0080) 계속 | 3 | NNNN-NN-** |

**Parameters (계속)**
- 이 값은 생략 가능하며, 생략할 경우 0(기본 보호유형)값을 설정한 것과 동일하게 동작한다.
- 사용자가 선택한 글자를 보호하는 경우(선택 글자 보호 기능) 이 값은 항상 0이다.
- return : true / false
  - 개인정보를 보호문자로 치환한 경우에 true를 반환한다.
  - 개인정보를 보호하지 못할 경우 false를 반환한다.
  - 문자열이 선택되지 않은 상태이거나, 개체가 선택된 상태에서는 실패한다.
  - 또한, 보호유형이 잘못된 설정된 경우에도 실패한다.

**Remark**
- 선택된 문자열을 보호한다.
- 한/글의 경우 "찾아서 보호"와 "선택 글자 보호"를 다른 기능으로 구현되었지만, API에서는 하나의 함수로 구현한다.
- ProtectingChar는 개인정보를 보호할 때 개인정보 대신 입력될 문자이다. 어떤 문자라도 상관없다.
- ProtectingType은 보호유형으로 개인정보의 전체를 보호할지 부분적으로 보호할지를 정해주는 값이다. 보호유형은 개인정보 유형에 따라 각각 다르다. (단, 기본 보호유형은 모든 보호유형에서 동일한 값(0)을 가진다)
- ProtectingType은 생략이 가능한데 생략할 경우에 기본 보호유형이 설정된다.
- ProtectPrivateInfo() 함수는 FindPrivateInfo() 함수로 선택된 문자열뿐만 아니라 사용자가 임의로 선택한 문자열도 보호할 수 있다. 예를 들어 "RepeatFind" 액션으로 찾아 선택한 문자열의 경우에도 보호가 가능하다.
- 개인정보 보호가 성공할 경우에는 true를 반환한다.
- 문자열이 선택된 경우에는 대부분 성공하지만 이전에 보호암호를 설정하지 않은 경우에는 실패하게 된다. 또한, 문자열이 선택되지 않은 경우와 개인정보 유형과 맞지 않는 보호유형을 설정할 경우에도 실패하게 된다.

### SolarToLunar(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 양력을 음력으로 변환 |
| Declaration | `bool SolarToLunar(long sYear, long sMonth, long sDay, long* lYear, long* lMonth, long* lDay, VARIANT_BOOL* lLeap);` |
| Parameters | sYear : 양력 년<br>sMonth : 양력 월<br>sDay : 양력 일<br>lYear : 음력으로 반환된 년<br>lMonth : 음력으로 반환된 월<br>lDay : 음력으로 반환된 일<br>lLeap : 윤달 |
| Remark | - 1841~2043년 사이만 변환 가능<br>- 1841년 1월 23일 이전 날짜는 변환 불가 |

### SolarToLunarBySet(Method)

| 항목 | 내용 |
|-----|-----|
| Description | SolarToLunar의 ParameterSet버전 |
| Declaration | `LPDISPATCH* SolarToLunarBySet(long sYear, long sMonth, long sDay);` |
| Parameters | sYear : 양력 년<br>sMonth : 양력 월<br>sDay : 양력 일 |

## Page 39

| 항목 | 내용 |
|-----|-----|
| Return | ParameterSet형식으로 반환 |
| Remark | - 1841~2043년 사이만 변환 가능<br>- 1841년 1월 23일 이전 날짜는 변환 불가 |

### SolarToLunarBySet 반환 ParameterSet

| Item ID | Type | Description |
|---------|------|-------------|
| Year | PIT_UI4 | 년 |
| Month | PIT_UI4 | 월 |
| Day | PIT_UI4 | 일 |
| Leap | PIT_UI1 | 윤달인지 아닌지 |

### LunarToSolar(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 음력을 양력으로 변환 |
| Declaration | `bool LunarToSolar(long lYear, long lMonth, long lDay, VARIANT_BOOL lLeap, long* sYear, long* sMonth, long* sDay);` |
| Parameters | lYear : 음력 년<br>lMonth : 음력 월<br>lDay : 음력 일<br>lLeap : 윤달<br>sYear : 양력으로 반환된 년<br>sMonth : 양력으로 반환된 월<br>sDay : 양력으로 반환된 일 |
| Remark | - 1841~2043년 사이만 변환 가능<br>- 입력한 날이 그 달에 포함되지 않으면 실패 |

### LunarToSolarBySet(Method)

| 항목 | 내용 |
|-----|-----|
| Description | LunarToSolar의 ParameterSet버전 |
| Declaration | `LPDISPATCH* LunarToSolarBySet(long lYear, long lMonth, long lDay, VARIANT_BOOL lLeap);` |
| Parameters | lYear : 음력 년<br>lMonth : 음력 월<br>lDay : 음력 일<br>lLeap : 윤달<br>return : ParameterSet형식으로 반환 |
| Remark | - 1841~2043년 사이만 변환 가능<br>- 입력한 날이 그 달에 포함되지 않으면 실패 |

### LunarToSolarBySet 반환 ParameterSet

| Item ID | Type | Description |
|---------|------|-------------|
| Year | PIT_UI4 | 년 |
| Month | PIT_UI4 | 월 |
| Day | PIT_UI4 | 일 |

### ScanFont(Method)

| 항목 | 내용 |
|-----|-----|
| Description | GetFontList를 호출하기 이전에 필수적으로 호출해야한다 |

### GetFontList(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 현재 문서에서 사용된 글꼴 목록을 문자열 형태로 반환 |
| Declaration | `BSTR GetFontList(VARIANT langid);` |
| Parameters | langid : 글꼴 언어 |

## Page 40

### GetFontList LangID 표

| LangID | Description |
|--------|-------------|
| 0 | 한글 |
| 1 | 영문 |
| 2 | 한자 |
| 3 | 일어 |
| 4 | 외국어 |
| 5 | 기호 |
| 6 | 사용자 |

**Remark**: 얻어온 글꼴 목록은 \x02 구분되어 있다.

### ReplaceFont(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 문서에 적용된 글꼴을 변경한다 |
| Declaration | `bool ReplaceFont(long langid, BSTR desFontName, long desFontType, BSTR newFontName, long newFontType);` |
| Parameters | langid : 글꼴 언어(위 GetFontList참고)<br>desFontName : 대상 글꼴 이름<br>desFontType : 대상 글꼴 타입 (TTF or HFT)<br>newFontName : 변경 할 글꼴 이름<br>newFontType : 변경 할 글꼴 타입 (TTF or HFT) |

### ReleaseAction(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 생성한 액션을 명시적으로 릴리즈할 필요가 있을 경우에 사용. |
| Declaration | `void ReleaseAction(LPDISPATCH action);` |
| Parameters | action : ActionObject.hwp 참고 |

### SetUserInfo(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 환경설정 – 일반 - 사용자 정보 변경 |
| Declaration | `void SetUserInfo(long userInfoId, BSTR value);` |
| Parameters | userInfoId : 변경할 정보<br>value : 입력할 정보 |

### SetUserInfo userInfoId 표

| userInfoId | Description |
|------------|-------------|
| 0 | 사용자 이름 |
| 1 | 회사 이름 |
| 2 | 직책 이름 |
| 3 | 부서 이름 |

### GetUserInfo(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 사용자 정보에 해당하는 값을 얻어온다 |
| Declaration | `BSTR GetUserInfo(long userInfoId);` |
| Parameters | userInfoId : 변경할 정보. 위 SetUserInfo참고<br>return : userInfoId에 해당하는 정보 |

### FileTranslate(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 문서를 번역해서 적용시켜준다 |
| Declaration | `bool FileTranslate(BSTR curLang, BSTR transLang);` |
| Parameters | (다음 페이지 참조) |
