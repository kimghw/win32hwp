# 2. 기타 정보

## 1) 선 종류

| 값 | Description |
|---|-------------|
| 0 | 없음. (NULL) |
| 1 | 실선. (SOLID) |
| 2 | 긴 점선. (DASH) |
| 3 | 점선. (DOT) |
| 4 | -.-.-.-. |
| 5 | -..-..-.. |
| 6 | HNCDR_LS_DASH보다 긴 선분의 반복. (LONGDASH) |
| 7 | HNCDR_LS_DOT보다 큰 동그라미의 반복. (CIRCLE) |
| 8 | 2중선. (DOUBLESLIM) |
| 9 | 가는 선 + 굵은 선 2중선. (SLIMTHICK) |
| 10 | 굵은 선 + 가는 선 2중선. (THICKSLIM) |
| 11 | 가는 선 + 굵은 선 + 가는 선 3중선. (SLIMTHICKSLIM) |
| 12 | 물결. (WAVE) |
| 13 | 물결 2중선. (DOUBLEWAVE) |
| 14 | 두꺼운 3D. (THICK3D) |
| 15 | 두꺼운 3D. 광원 반대. (THICKREV3D) |
| 16 | 3D 단선. (3D) |
| 17 | 3D 단선. 광원 반대. (REV3D) |

---

## 2) 선 굵기

| 값 | Description |
|---|-------------|
| -1 | 최소값 (=0.1 mm) |
| 0 | 0.1 mm |
| 1 | 0.12 mm |
| 2 | 0.15 mm |
| 3 | 0.2 mm |
| 4 | 0.25 mm |
| 5 | 0.3 mm |
| 6 | 0.4 mm |
| 7 | 0.5 mm |
| 8 | 0.6 mm |
| 9 | 0.7 mm |
| 10 | 1.0 mm |
| 11 | 1.5 mm |
| 12 | 2.0 mm |
| 13 | 3.0 mm |
| 14 | 4.0 mm |
| 15 | 5.0 mm |
| 16 | 최대값 (=5.0 mm) |

---

## 3) 번호모양

| 값 | Description |
|---|-------------|
| 0 | 1, 2, 3 |
| 1 | 동그라미 쳐진 1, 2, 3 |
| 2 | I, II, III |
| 3 | i, ii, iii |
| 4 | A, B, C |
| 5 | a, b, c |
| 6 | 동그라미 쳐진 A, B, C |
| 7 | 동그라미 쳐진 a, b, c |
| 8 | 가, 나, 다 |
| 9 | 동그라미 쳐진 가, 나, 다 |
| 10 | ㄱ, ㄴ, ㄷ |
| 11 | 동그라미 쳐진 ㄱ, ㄴ, ㄷ |
| 12 | 일, 이, 삼 |
| 13 | 一, 二, 三 |
| 14 | 동그라미 쳐진 一, 二, 三 |
| 각주/미주 전용 (0x80부터 시작) | 4가지 문자가 차례로 반복 |
| 0x81 | 사용자 지정 문자 반복 |

---

## 4) 하이퍼링크 Command 문자열

많은 정보를 한 줄의 문자열에 포함하고 있으므로 상당히 복잡한 구조를 가지고 있습니다. 가장 빠르게 익힐 수 있는 방법은 원하는 형식의 하이퍼링크를 직접 만들고 해당 문서를 HWPML(*.hml)형식으로 저장한 후 XML문서를 볼 수 있는 프로그램(예:Microsoft Internet Explorer)에서 열어보면 자세한 내용을 알 수 있습니다. (IE에서는 hml확장자를 xml로 변경하신 후 봐야합니다.)

예) HyperLink정보가 hml문서에 저장되어 있는 모습

```xml
<FIELDBEGIN Type="Hyperlink" InstId="2118971508" Editable="false" Dirty="false" Property="0"
Command="http://www.haansoft.com;1;0;0" />
```

**하이퍼링크의 문자열**은 ";"을 구분자로 하는 다음과 같은 구조를 가집니다.

```
[TARGET];[LINK_TYPE];[OBJ_TYPE];[OPTION]
```

TARGET은 하이퍼링크의 대상을 뜻하며, 연결 유형에 따라 다음과 같은 형태를 가집니다.

| Link Type | Syntax | Example |
|-----------|--------|---------|
| 글개체 | 글문서?개체ID | ParameterSet.hwp?#2043988344 |
| 웹 주소 | URL | http://www.haansoft.com |
| 이메일 주소 | mailto:메일주소 | mailto:swlab@haansoft.com |
| 외부 프로그램 | file path | c:\hnc\hnctt\hnctt.exe |

※ 연결 유형이 글개체이고, 동일문서상의 개체일 경우에는 구문(Syntax)의 글문서를 제외할 수 있습니다. (예: ?#2043988344)

TARGET에서 사용되는 개체ID는 hml문서를 참조하여 얻을 수 있습니다. 개체 Element안의 "InstId"속성이 그 개체의 ID를 나타냅니다. TARGET에서 사용할 때에는 앞에 #을 붙여 개체ID임을 나타냅니다.

TARGET 다음으로 표현되는 데이터들은 다음과 같은 의미를 가집니다.

| Item | Decription |
|------|------------|
| LINK_TYPE | 연결 유형 : 0 = 글개체, 1 = 웹 주소, 2 = 이메일, 3 = 외부 프로그램 |
| OBJ_TYPE | 연결할 글개체의 유형. LINK_TYPE이 글개체가 아니면 이 값은 무시된다. 0 = 책갈피, 1 = 개요, 2 = 표, 3 = 그림, 4 = 수식 |
| OPTION | 하이퍼링크 이동시 옵션. 외부의 글문서와 연결된 경우에만 적용된다. 0 = 현재창에 외부문서를 연다. (현재문서는 닫힘) 1 = 현재창에 새 탭을 띄워 외부문서를 연다. 2 = 새 창을 띄워 외부문서를 연다. |

다음은 위 내용을 종합하여 작성한 하이퍼링크 Command 문자열의 예제입니다.

| Command 문자열 | 설명 |
|-------------|-----|
| ParameterSet.hwp?#2043988344;0;0;0 | 외부문서"ParameterSet.hwp"의 책갈피와 연결. 현재문서에 연다. |
| C:\Hnc\Hwp70\Readme.Hwp?#204399566;0;0;0 | 외부문서"Readme.hwp"의 책갈피와 연결(절대경로). 현재문서에 연다. |
| ?#2043988345;0;1;1 | 현재문서의 "개요"와 연결. 새 탭에 연다. |
| ?#2043988347;0;3;2 | 현재문서의 "그림"과 연결. 새 창에 연다. |
| http://www.haansoft.com;1;0;0 | 해당 웹 주소로 연결한다. |
| mailto:swlab@haansoft.com;2;0;0 | 해당 메일주소로 연결한다. (연결된 프로그램 자동로딩) |
| c:\hnc\hnctt\hnctt.exe;3;0;0 | 한글의 타자연습프로그램을 로딩시킨다. |

**상호참조**는 하이퍼링크의 확장된 형태로 하이퍼링크와 비슷한 형태의 Command 문자열을 가집니다. **상호참조의 문자열**은 다음과 같은 구조를 가집니다.

```
[TARGET];[OBJ_TYPE];[REF_STRING];[HYPERLINK];[OPTION]
```

각 항목은 다음과 같습니다.

| Item | Decription |
|------|------------|
| TARGET | HyperLink와 동일 |
| OBJ_TYPE | 참조 대상의 유형. 0 = 표, 1 = 그림, 2 = 수식, 3 = 각주, 4 = 미주, 5 = 개요, 6 = 책갈피 |
| REF_STRING | 참조 내용. 0 = 개체가 위치한 쪽, 1 = 개체번호, 2 = 개체내용, 3 = 위/아래 존재여부 (일반) 0 = 개체가 위치한 쪽, 1 = 책갈피이름, 2 = 책갈피내용, 3 = 위/아래 존재여부 (책갈피) 0 = 개체가 위치한 쪽, 1 = 개체번호, 3 = 위/아래 존재여부 (각주/미주) |
| HYPERLINK | 하이퍼링크 여부 : 0 = 연결 안 함, 1 = 연결함 |
| OPTION | 하이퍼링크 이동시 옵션. 외부의 글문서와 연결된 경우에만 적용된다. 0 = 현재창에 외부문서를 연다. (현재문서는 닫힘) 1 = 현재창에 새 탭을 띄워 외부문서를 연다. 2 = 새 창을 띄워 외부문서를 연다. |

TARGET에서 사용되는 개체ID는 하이퍼링크와 마찬가지로 hml문서를 참조하여 얻을 수 있습니다. 개체 Element안의 "InstId"속성이 그 개체의 ID를 나타냅니다. TARGET에서 사용할 때에는 앞에 #을 붙여 개체ID임을 나타냅니다.

"각주","미주"개체는 "InstId"속성이 존재하지 않습니다. 이런 경우에는 개체가 존재하는 순서(INDEX)로 개체를 구분합니다. TARGET에서 사용할 때는 마찬가지로 앞에 #을 붙입니다.

"책갈피" 개체의 경우에는 "InstId"속성 대신 "Name"속성을 사용합니다. 이 경우에는 #을 붙이지 않습니다.

다음은 위 내용을 종합하여 작성한 상호참조 Command 문자열 예제입니다.

| Command 문자열 | 설명 |
|-------------|-----|
| ?#2043988345;0;0;0;0 | 표를 상호참조하여 화면에 표의 Page를 표시한다. |
| ?#1;3;1;0;0 | 각주를 상호참조하여 화면에 각주번호를 표시한다. |
| C:\Hnc\Hwp70\Readme.Hwp?#204399566;1;2;1;2 | 외부문서"ReadMe.hwp"의 그림을 상호참조한다. 클릭시 새 창에 띄움 |
| ?책갈피;6;2;0;0 | 책갈피를 상호참조하여 화면에 책갈피 내용을 표시한다. |

※ 상호참조는 하이퍼링크와 다르게 절대경로만으로 외부문서를 참조할 수 있다.

---

## 5) 적용범위

적용범위란 현재 실행한 액션이 적용될 범위를 말하는 것이다. 예를 들면 표의 테두리를 변경하는 액션을 수행할 때 이것을 전체 셀에 할 것인지, 선택된 셀에만 적용할 것인지를 가늠하는 용도로 쓰인다.

**ApplyTo**는 위에서 말한 적용범위를 나타낸 아이템으로 적용범위가 필요한 모든 액션에 들어가 있다.

ApplyTo에 들어가는 값은 각각의 액션에 따라 다르며 들어갈 수 있는 값은 액션이 직접 정의하여 사용한다. (일반적으로 동일한 파라메터셋을 사용하는 액션은 동일한 ApplyTo 값을 가진다.)

캐럿의 상태에 따라 ApplyTo에 들어갈 수 있는 값이 제한적일 수 있다. 이런 경우 대화상자에서 선택될 수 있는 적용범위를 제한해 주는 것이 좋은데 이것을 정의한 아이템이 바로 **ApplyClass**이다. 마찬가지로 적용범위의 제한이 필요한 모든 액션에 들어가 있다. (적용범위를 가지나 캐럿의 상태에 따라 적용범위를 제한할 필요가 없으면 ApplyClass가 필요 없다.)

ApplyClass는 ApplyTo와 마찬가지로 각각의 액션이 들어가는 값을 정의하며, ApplyTo와 다르게 정의한 값을 조합하여 사용한다. (일종의 Mask시스템으로 지정된 bit flag가 존재하면 ApplyTo로 적용 가능한 값이다.)

---

## 6) Type 변환 표

| HNC Type | C++(MFC) Type | HNC Type | C++(MFC) Type |
|----------|---------------|----------|---------------|
| PIT_NULL | NULL | PMT_INT8 | char |
| PIT_BSTR | BSTR (OLE Automation string) | PMT_INT16 | short |
| PIT_I1 | char (1byte signed integer) | PMT_INT32 | long |
| PIT_I2 | short (2byte signed integer) | PMT_INT | int |
| PIT_I4 | long (4byte signed integer) | PMT_BYTE | BYTE (unsigned char) |
| PIT_I | int (machine dependent integer) | PMT_UINT16 | unsigned short |
| PIT_UI1 | unsigned char | PMT_UINT32 | unsigned long |
| PIT_UI2 | unsigned short | PMT_UINT | unsigned int |
| PIT_UI4 | unsigned long | PMT_CHAR | char |
| PIT_UI | unsigned int | PMT_UCHAR | BYTE (unsigned char) |
| PIT_SET | HwpParameterSet 내부에서 해당 ParameterSet을 생성한 뒤 그 객체의 Dispatch ID를 돌려준다. | PMT_WCHAR | unsigned short |
|  |  | PMT_SHORT | short |
|  |  | PMT_LONG | long |
|  |  | PMT_ULONG | unsigned long |
| PIT_ARRAY | HwpParameterArray 내부에서 해당 ParameterArray를 생성한 뒤 그 객체의 Dispatch ID를 돌려준다. | PMT_WORD | unsigned short |
|  |  | PMT_DWORD | unsigned long |
|  |  | PMT_BOOL | int |
|  |  | PMT_ASTR | char * |
| PIT_BINDATA | any (따로 타입을 체크하지 않음) | PMT_WSTR | wchar_t * |
| PIT_UI64 | 리눅스용 64bit UINT | PMT_BSTR | BSTR |

※ PIT_* or PMT_* Type은 각 시스템의 호환성을 위해 내부적으로 지정한 기호상수이다. 각 시스템에서 이 기호상수를 해석해서 시스템에 잘 호환되는 내부 Data Type으로 변환하여 사용한다. 그러므로, 해당 타입이 꼭 C++(MFC) Type과 100% 호환되는 것은 아니며, 또한 액션에 따라 내부적으로 해석하는 방법이 다른 수도 있다. 해당 타입을 사용할 때에는 의미론적인 Data Type으로 인지하여 사용한다.

---

## 7) HWPUNIT

글이 사용하는 기본 단위. 1mm는 283.465HWPUNIT 이며, 1inch는 7200HWPUNIT 이다.

---

## 8) URC

32bit 정수값. HWPUNIT 또는 Relative Character Position을 나타낸다.

Bit0 = 0인 경우, HWPUNIT을 나타내며, Bit1~31에 HWPUNIT에 해당하는 값이 저장된다.

BIT0 = 1인 경우, Relative Character이며 Bit1~31에는 n*100의 값을 가진다.

```c
#define HWPURC_MAKE(type, value)    (((value) << 1) | ((type) & 1))
#define HWPURC_TYPE(data)           ((data) & 1)
#define HWPURC_VALUE(data)          ((data) >> 1)
```
