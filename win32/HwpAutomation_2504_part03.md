# HwpAutomation_2504_part03

## Page 21

### MovePos moveID 표 (계속)

| ID | 값 | 설명 |
|----|---|-----|
| movePrevChar | 17 | 한 글자 뒤로 이동. (현재 리스트만을 대상으로 동작한다.) |
| moveNextWord | 18 | 한 단어 앞으로 이동. (현재 리스트만을 대상으로 동작한다.) |
| movePrevWord | 19 | 한 단어 뒤로 이동. (현재 리스트만을 대상으로 동작한다.) |
| moveNextLine | 20 | 한 줄 위로 이동. |
| movePrevLine | 21 | 한 줄 아래로 이동. |
| moveStartOfLine | 22 | 현재 위치한 줄의 시작으로 이동. |
| moveEndOfLine | 23 | 현재 위치한 줄의 끝으로 이동. |
| moveParentList | 24 | 한 레벨 상위로 이동한다. |
| moveTopLevelList | 25 | 탑레벨 리스트로 이동한다. |
| moveRootList | 26 | 루트 리스트로 이동한다.<br>추가 설명:<br>현재 루트 리스트에 위치해 있어 더 이상 상위 리스트가 없을 때는 위치 이동 없이 리턴한다.<br>이동한 후의 위치는 상위 리스트에서 서브리스트가 속한 컨트롤 코드가 위치한 곳이다.<br>위치 이동시 셀렉션은 무조건 풀린다. |
| moveLeftOfCell | 100 | 현재 캐럿이 위치한 셀의 왼쪽 |
| moveRightOfCell | 101 | 현재 캐럿이 위치한 셀의 오른쪽 |
| moveUpOfCell | 102 | 현재 캐럿이 위치한 셀의 위쪽 |
| moveDownOfCell | 103 | 현재 캐럿이 위치한 셀의 아래쪽 |
| moveStartOfCell | 104 | 현재 캐럿이 위치한 셀에서 행(row)의 시작 |
| moveEndOfCell | 105 | 현재 캐럿이 위치한 셀에서 행(row)의 끝 |
| moveTopOfCell | 106 | 현재 캐럿이 위치한 셀에서 열(column)의 시작 |
| moveBottomOfCell | 107 | 현재 캐럿이 위치한 셀에서 열(column)의 끝 |
| moveScrPos | 200 | 한/글 문서장에서의 screen 좌표로서 위치를 설정 한다. |
| moveScanPos | 201 | GetText() 실행 후 위치로 이동한다. |

**Parameters (계속)**
- para : 이동할 문단의 번호. moveMain 또는 moveCurList가 지정되었을 때만 사용된다. moveScrPos가 지정되었을 때는 문단번호가 아닌 스크린 좌표로 해석된다. (스크린 좌표 : LOWORD = x 좌표, HIWORD = y 좌표)
- pos : 이동할 문단 중에서 문자의 위치. moveMain 또는 moveCurList가 지정되었을 때만 사용된다.

**Remark**
- moveScrPos가 지정되었을 때는 스크린 좌표는 마우스 커서의 (x,y) 좌표를 그대로 넘겨 주면 된다.
- moveScanPos는 문서를 검색하는 중에 캐럿을 이동 시키려 할 경우에만 사용이 가능하다.

### InitScan(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 문서 검색을 위한 준비 작업을 한다. |

## Page 22

| 항목 | 내용 |
|-----|-----|
| Declaration | `BOOL InitScan([unsigned int option], [unsigned int rang], [unsigned int spara], [unsigned int spos], [unsigned int epara], [unsigned int epos])` |
| Parameters | option : 찾을 대상을 다음과 같은 옵션을 조합하여 지정할 수있다. 생략하면 모든 컨트롤을 찾을 대상으로 한다.<br>range : 검색의 범위를 다음과 같은 옵션을 조합하여 지정할 수 있다. 생략하면 "문서 시작부터 - 문서의 끝까지" 검색 범위가 지정된다. |

### InitScan option 표

| ID | 값 | 설명 |
|----|---|-----|
| maskNormal | 0x00 | 본문을 대상으로 검색한다. (서브리스트를 검색하지 않는다.) |
| maskChar | 0x01 | char 타입 컨트롤 마스크를 대상으로 한다. (강제줄나눔, 문단끝, 하이픈, 묶움빈칸, 고정폭빈칸, 등...) |
| maskInline | 0x02 | inline 타입 컨트롤 마스크를 대상으로 한다. (누름틀 필드 끝, 등...) |
| maskCtrl | 0x04 | extende 타입 컨트롤 마스크를 대상으로 한다. (바탕쪽, 프리젠테이션, 다단, 누름틀 필드 시작, Shape Object, 머리말, 꼬리말, 각주, 미주, 번호관련 컨트롤, 새번호 관련 컨트롤, 감추기, 찾아보기, 글자겹침, 등...) |

### InitScan range 표

| ID | 값 | 설명 |
|----|---|-----|
| scanSposCurrent | 0x0000 | 캐럿 위치부터. (시작 위치) |
| scanSposSpecified | 0x0010 | 특정 위치부터. (시작 위치) |
| scanSposLine | 0x0020 | 줄의 시작부터. (시작 위치) |
| scanSposParagraph | 0x0030 | 문단의 시작부터. (시작 위치) |
| scanSposSection | 0x0040 | 구역의 시작부터. (시작 위치) |
| scanSposList | 0x0050 | 리스트의 시작부터. (시작 위치) |
| scanSposControl | 0x0060 | 컨트롤의 시작부터. (시작 위치) |
| scanSposDocument | 0x0070 | 문서의 시작부터. (시작 위치) |
| scanEposCurrent | 0x0000 | 캐럿 위치까지. (끝 위치) |
| scanEposSpecified | 0x0001 | 특정 위치까지. (끝 위치) |
| scanEposLine | 0x0002 | 줄의 끝까지. (끝 위치) |
| scanEposParagraph | 0x0003 | 문단의 끝까지. (끝 위치) |
| scanEposSection | 0x0004 | 구역의 끝까지. (끝 위치) |
| scanEposList | 0x0005 | 리스트의 끝까지. (끝 위치) |
| scanEposControl | 0x0006 | 컨트롤의 끝까지. (끝 위치) |
| scanEposDocument | 0x0007 | 문서의 끝까지. (끝 위치) |
| scanWithinSelection | 0x00ff | 검색의 범위를 블록으로 제한. |
| scanForward | 0x0000 | 정뱡향. (검색 방향) |
| scanBackward | 0x0100 | 역방향. (검색 방향) |

## Page 23

**Parameters (계속)**
- spara : 검색 시작 위치의 문단 번호. scanSposSpecified 옵션이 지정되었을 때만 유효하다.
- spos : 검색 시작 위치의 문단 중에서 문자의 위치. scanSposSpecified 옵션이 지정되었을 때만 유효하다.
- epara : 검색 끝 위치의 문단 번호. scanEposSpecified 옵션이 지정되었을 때만 유효하다.
- epos : 검색 끝 위치의 문단 중에서 문자의 위치. scanEposSpecified 옵션이 지정되었을 때만 유효하다.

**Remark**
문서의 검색 과정은 InitScan()으로 검색위한 준비 작업을 하고 GetText()를 호출하여 본문의 텍스트를 얻어온다. GetText()를 반복호출하면 연속하여 본문의 텍스트를 얻어올 수 있다. 검색이 끝나면 ReleaseScan()을 호출하여 관련 정보를 초기화 해 주면 된다.

### ReleaseScan(Method)

| 항목 | 내용 |
|-----|-----|
| Description | InitScan()으로 설정된 정보를 초기화 한다. |
| Declaration | `void ReleaseScan()` |

### GetText(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 문서 중에서 텍스트를 얻는다. |
| Declaration | `long GetText(BSTR FAR* text)` |
| Parameters | text : 텍스트 데이터가 돌아온다. 텍스트에서 탭은 '\t'(0x9), 문단 바뀜은 CR/LF(0x0D/0x0A)로 표현되며, 이외 특수 코드는 포함되지 않는다. |
| Return | 0 = 텍스트 정보 없음.<br>1 = 리스트의 끝.<br>2 = 일반 텍스트.<br>3 = 다음 문단.<br>4 = 제어문자 내부로 들어감.<br>5 = 제어 문자를 빠져 나옴.<br>101 = 초기화 안됨. (InitScan() 실패 또는 InitScan()를 실행하지 않음)<br>102 = 텍스트 변환 실패. |
| Remark | GetText()의 사용이 끝나면 ReleaseScan()을 반드시 호출하여 관련 정보를 초기화 해주어야 한다. 텍스트가 있는 문단으로 캐럿을 이동 시키려면 moveScanPos를 주고 MovePos()를 호출하면 된다. |

### GetPos(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 현재 캐럿의 위치 정보를 얻어온다. |
| Declaration | `void GetPos(long FAR* list, long FAR* para, long FAR* pos)` |
| Parameters | list : 캐럿이 위치한 문서 내 리스트 아이디.<br>para : 캐럿이 위치한 문단 아이디.<br>pos : 캐럿이 위치한 문단 내 글자 단위 위치. |
| Remark | 위의 리스트란, 문단과 컨트롤들이 연결된 한/글 문서 내 구조를 뜻한다. 리스트 아이디는 문서내 위치 정보 중 하나로서 SelectText에 넘겨줄 때 사용한다. |

### SetPos(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 캐럿을 문서 내 특정 위치로 위치시킨다. |
| Declaration | `void SetPos(long list, long para, long pos)` |
| Parameters | list : 캐럿이 위치할 문서 내 리스트 아이디.<br>para : 캐럿이 위치할 문단 아이디.<br>pos : 캐럿이 위치할 문단 내 글자 단위 위치. |

## Page 24

### KeyIndicator(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 상태바에 나타날 정보를 알아낸다. |
| Declaration | `BOOL KeyIndicator(long FAR* seccnt, long FAR* secno, long FAR* prnpageno, long FAR* colno, long FAR* line, long FAR* pos, short FAR* over, BSTR FAR* ctrlname)` |
| Parameters | seccnt : 총 구역<br>secno : 현재 구역<br>prnpageno : 쪽<br>colno : 단<br>line : 줄<br>pos : 칸<br>over : (true:수정, false:삽입) |
| Remark | 컨트롤 바깥쪽에서 상태바를 만들어서 상태바에 표시할 정보들의 내용을 알아낼 때 유용하다.<br>주의: 이 함수는 빠른 속도가 요구되므로 parameter로 포인터를 받는다. 따라서 포인터를 사용할 수 없는 언어에서는 사용이 불가능하다. |

### GetTextFile(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 현재 열린 문서를 문자열로 넘겨준다. |
| Declaration | `VARIANT GetTextFile(BSTR format, BSTR option)` |
| Parameters | format : 파일의 형식<br>option<br>return : 저장된 텍스트로 파일을 문자열로 바꿔서 리턴한다. |
| Remark | 이 함수는 JScript나 VBScript와 같이 직접적으로 local disk를 접근하기 힘든 언어를 위해 만들어졌으므로 disk를 접근할 수 있는 언어에서는 사용하지 않기를 권장합니다. disk를 접근할 수 있다면, Save나 SaveBlockAction을 사용하십시오. 이 함수 역시 내부적으로는 save나 SaveBlockAction을 호출하도록 되어있고 텍스트로 저장된 파일이 메모리에서 3~4번 복사되기 때문에 느리고, 메모리를 낭비합니다. |

### GetTextFile format 표

| format | 설명 | 비고 |
|--------|-----|-----|
| HWP | HWP native format | BASE64로 인코딩되어 있다. 저장된 내용을 다른 곳에서 보여줄 필요가 없다면 이 포맷을 사용하기를 권장합니다. |
| HWPML2X | HWP 형식과 호환 | 문서의 모든 정보를 유지 |
| HTML | 인터넷 문서 HTML 형식 | 글 고유의 서식은 손실된다. |
| UNICODE | 유니코드 텍스트 | 서식정보가 없는 텍스트만 저장 |
| TEXT | 일반 텍스트 | 유니코드에만 있는 정보(한자, 고어, 특수문자 등)는 모두 손실된다. |

### GetTextFile option 표

| option | 설명 | 비고 |
|--------|-----|-----|
| saveblock | 선택된 블록만 저장 | 개체 선택 상태에서는 동작하지 않는다. |

### SetTextFile(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 문서를 문자열로 지정한다. |

## Page 25

| 항목 | 내용 |
|-----|-----|
| Declaration | `long SetTextFile(VARIANT data, BSTR format, BSTR option)` |
| Parameters | data : 문자열로 변경된 text 파일<br>format : 파일의 형식<br>option |
| Remark | 이 함수는 JScript나 VBScript와 같이 직접적으로 local disk를 접근하기 힘든 언어를 위해 만들어졌으므로 disk를 접근할 수 있는 언어에서는 사용하지 않기를 권장합니다. disk를 접근할 수 있다면, Open이나 Insert를 사용하십시오. 이 함수 역시 내부적으로는 Open이나 Insert를 호출하도록 되어있고 텍스트로 저장된 파일이 메모리에서 3~4번 복사되기 때문에 느리고, 메모리를 낭비합니다. |

### SetTextFile format 표

| format | 설명 | 비고 |
|--------|-----|-----|
| HWP | HWP native format | BASE64 로 인코딩되어 있어야 한다. 저장된 내용을 다른 곳에서 보여줄 필요가 없다면 이 포맷을 사용하기를 권장합니다. |
| HWPML2X | HWP 형식과 호환 | 문서의 모든 정보를 유지 |
| HTML | 인터넷 문서 HTML 형식 | 글 고유의 서식은 손실된다. |
| UNICODE | 유니코드 텍스트 | 서식정보가 없는 텍스트만 저장 |
| TEXT | 일반 텍스트 | 유니코드에만 있는 정보(한자, 고어, 특수문자 등)는 모두 손실된다. |

### SetTextFile option 표

| option | 설명 | 비고 |
|--------|-----|-----|
| insertfile | 현재커서 이후에 삽입 | |

### CreatePageImage(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 지정한 페이지의 이미지를 파일로 생성한다. |
| Declaration | `boolean CreatePageImage(BSTR path, [long pgno], [short resolution], [short depth], [BSTR format])` |
| Parameters | path : 생성할 이미지 파일의 경로<br>pgno : 페이지 번호. 0부터 PageCount - 1까지. 생략하면 0이 사용된다.<br>resolution : 이미지 해상도. DPI 단위(96, 300, 1200)로 지정한다. 생략하면 96이 사용된다.<br>depth : 이미지 파일의 color depth(1, 4, 8, 24)를 지정한다.<br>format : 이미지 파일의 포맷. "bmp", "gif" 중의 하나. 생략하면 "bmp"가 사용된다. |

### Run(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 액션을 실행한다. |
| Declaration | `void Run(BSTR action)` |
| Parameters | action : 액션 ID (별도 문서 참조) |

### LockCommand(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 특정 액션이 실행되지 않도록 잠근다. |
| Declaration | `void LockCommand(BSTR actionID, boolean lock)` |
| Parameters | actionID : 액션 ID<br>lock : True이면 액션의 실행을 lock시키고, False이면 unlock시킨다. |

## Page 26

### IsCommandLock(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 특정 액션이 잠금 상태인지 여부를 조사한다. |
| Declaration | `boolean IsCommandLock(BSTR actionID)` |
| Parameters | actionID : 액션 ID |

### InsertPicture(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 현재 캐럿의 위치에 그림을 삽입한다. |
| Declaration | `Ctrl InsertPicture(BSTR path, [boolean embeded], [short sizeoption], [boolean reverse], [boolean watermark], [short effect], [long width], [long height])` |
| Parameters | path : 삽입할 이미지 파일, URL 사용 가능<br>embeded : 이미지 파일을 문서내에 포함할지 여부 (True/False). 생략하면 True<br>sizeoption : 삽입할 그림의 크기를 지정하는 옵션<br>reverse : 이미지의 반전 유무 (True/False)<br>watermark : watermark효과 유무 (True/False)<br>effect : 그림 효과<br>width : 그림의 가로 크기 지정. 단위는 mm<br>height : 그림의 높이 크기 지정. 단위는 mm |

### HwpSizeOption enum

```cpp
typedef enum {
    // 이미지 원래의 크기로 삽입한다.
    // width와 height를 지정할 필요없다.
    realSize = 0,
    // width와 height에 지정한 크기로 그림을 삽입한다.
    specificSize = 1,
    // 현재 캐럿이 표의 셀안에 있을 경우,
    // 셀의 크기에 맞게 자동 조절하여 삽입한다.
    // width는 셀의 width만큼,
    // height는 셀의 height만큼 확대/축소된다.
    // 캐럿이 셀안에 있지 않으면
    // 이미지의 원래 크기대로 삽입된다.
    cellSize = 2,
    // 현재 캐럿이 표의 셀안에 있을 경우,
    // 셀의 크기에 맞추어 원본 이미지의 가로 세로의
    // 비율이 동일하게 확대/축소하여 삽입한다.
    cellSizeWithSameRatio = 3
} HwpSizeOption;
```

### HwpPictureEffect enum

```cpp
typedef enum {
    RealPicture = 0,  // 실제 이미지 그대로
    GrayScale = 1,    // 그레이스케일
    BlackWhite = 2    // 흑백효과
} HwpPictureEffect;
```

### InsertBackgroundPicture(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 배경이미지를 넣는다. |
| Declaration | `VARIANT InsertBackgroundPicture(BSTR bordertype, BSTR path, [bool embedded], [long filloption], [bool watermark], [long effect], [long brightness], [long contrast]);` |
| Parameters | bordertype : 배경 유형을 지정<br>path : 삽입할 이미지 파일, URL 사용 가능<br>embeded : 이미지 파일을 문서내에 포함할지 여부 (True/False). 생략하면 True<br>filloption : 삽입할 그림의 크기를 지정하는 옵션<br>effect : 이미지효과<br>watermark : watermark효과 유무 (True/False) 이 옵션이 true이면 brightness 와 contrast 옵션이 무시된다.<br>brightness : 밝기 지정(-100 ~ 100), 기본값 : 0<br>contrast : 선명도 지정(-100 ~ 100), 기본값 : 0 |
| Remark | CellBorderFill의 SetItem 중 FillAttr 의 SetItem FileName 에 이미지의 binary data를 지정해 줄 수가 없어서 만든 함수다. 기타 배경에 대한 다른 조정은 Action과 ParameterSet의 조합으로 가능하다. |

## Page 27

### InsertBackgroundPicture bordertype 표

| bordertype | 설명 | 비고 |
|------------|-----|-----|
| "SelectedCell" | 현재 선택된 표의 셀의 배경을 변경한다. | 반드시 셀이 선택되어 있어야 함. 커서가 위치하는 것만으로는 동작하지 않음. |
| "SelectedCellDelete" | 현재 선택된 표의 셀의 배경을 지운다. | |

### InsertBackgroundPicture filloption 표

| filloption | 설명 | 비고 |
|------------|-----|-----|
| 0 | 바둑판식으로 - 모두 | |
| 1 | 바둑판식으로 - 가로/위 | |
| 2 | 바둑판식으로 - 가로/아로 | |
| 3 | 바둑판식으로 - 세로/왼쪽 | |
| 4 | 바둑판식으로 - 세로/오른쪽 | |
| 5 | 크기에 맞추어 | 설정하지 않았을 때 기본 값 |
| 6 | 가운데로 | |
| 7 | 가운데 위로 | |
| 8 | 가운데 아래로 | |
| 9 | 왼쪽 가운데로 | |
| 10 | 왼쪽 위로 | |
| 11 | 왼쪽 아래로 | |
| 12 | 오른쪽 가운데로 | |
| 13 | 오른쪽 위로 | |
| 14 | 오른쪽 아래로 | |

### InsertBackgroundPicture effect 표

| effect | 설명 | 비고 |
|--------|-----|-----|
| 0 | 원래 그림 | 설정하지 않았을 때 기본 값 |
| 1 | 그래이 스케일 | |
| 2 | 흑백으로 | |

### CreateAction(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 오브젝트를 생성한다. |
| Declaration | `Action CreateAction(BSTR action)` |
| Parameters | action : 액션 ID (별도 문서 참조) |
| Remark | 액션에 대한 세부적인 제어가 필요할 때 사용한다. 예를 들어 기능을 수행하지 않고 대화상자만을 띄운다든지, 대화상자 없이 지정한 옵션에 따라 기능을 수행하는 등에 사용할 수 있다. |

## Page 28

### InsertCtrl(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 현재 캐럿 위치에 컨트롤을 삽입한다. |
| Declaration | `Ctrl InsertCtrl(BSTR ctrlid, [ParameterSet initparam])` |
| Parameters | ctrlid : 삽입할 컨트롤의 ID<br>initparam : 컨트롤의 초기 속성. 생략하면 디폴트 속성으로 생성한다. |
| Remark | - ctrlid에 지정할 수 있는 컨트롤 ID는 HwpCtrl.CtrlID가 리턴하는 ID와 동일하다. 자세한 것은 Ctrl 오브젝트 Properties인 CtrlID를 참조.<br>- initparam에는 컨트롤의 초기 속성을 지정한다. 대부분의 컨트롤은 Ctrl.Properties와 동일한 포맷의 parameter set을 사용하지만, 컨트롤 생성시에는 다른 포맷을 사용하는 경우도 있다. 예를 들어 표의 경우 Ctrl.Properties에는 "Table" 셋을 사용하지만, 생성시 initparam에 지정하는 값은 "TableCreation" 셋이다. |

### DeleteCtrl(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 문서내 컨트롤을 삭제한다. |
| Declaration | `boolean DeleteCtrl(HwpCtrlCode ctrl)` |
| Parameters | ctrl : 삭제할 문서내 컨트롤 |

### GetMousePos(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 마우스의 현재 위치를 구한다 |
| Declaration | `ParameterSet GetMousePos(long Xrelto, long Yrelto)` |
| Parameters | Xrelto : X좌표계의 기준 위치<br>Yrelto : Y좌표계의 기준 위치 |
| Return | "MousePos" ParameterSet 이 리턴된다. |
| Remark | 단위가 HWPUNIT임을 주의하십시오. |

### GetMousePos Xrelto 표

| value | 설명 | 비고 |
|-------|-----|-----|
| 0 | 종이 기준 | 종이 기준으로 좌표를 가져온다. |
| 1 | 쪽 기준 | 쪽 기준으로 좌표를 가져온다. |

### GetMousePos Yrelto 표

| value | 설명 | 비고 |
|-------|-----|-----|
| 0 | 종이 기준 | 종이 기준으로 좌표를 가져온다. |
| 1 | 쪽 기준 | 쪽 기준으로 좌표를 가져온다. |

### [Set ID] MousePos

| Item ID | Type | Description |
|---------|------|-------------|
| XRelTo | PIT_UI4 | 가로 상대적 기준<br>0 : 종이<br>1 : 쪽 |
| YRelTo | PIT_UI4 | 세로 상대적 기준<br>0 : 종이<br>1 : 쪽 |
| Page | PIT_UI4 | 페이지 번호 ( 0 based) |
| X | PIT_I4 | 가로 클릭한 위치 (HWPUNIT) |
| Y | PIT_I4 | 세로 클릭한 위치 (HWPUNIT) |

### Clear(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 현재 편집중인 문서의 내용을 닫고 빈문서 편집 상태로 돌아간다. |

## Page 29

| 항목 | 내용 |
|-----|-----|
| Declaration | `void Clear([HwpSaveOption option])` |
| Parameters | option : 편집중인 문서의 내용에 대한 처리 방법. 생략하면 hwpAskSave가 선택된다. (HwpSaveOption : short type) |
| Remark | format, arg에 대해서는 Open 참조. hwpSaveIfDirty, hwpSave가 지정된 경우 현재 문서 경로가 지정되어 있지 않으면 "새이름으로 저장" 대화상자를 띄워 사용자에게 경로를 묻는다. |

### HwpSaveOption 표

| ID | 값 | 설명 |
|----|---|-----|
| hwpAskSave | 0 | 문서의 내용이 변경되었을 때 사용자에게 저장할지 묻는 대화상자를 띄운다. |
| hwpDiscard | 1 | 문서 내용을 버린다. |
| hwpSaveIfDirty | 2 | 문서가 변경된 경우 저장 한다. |
| hwpSave | 3 | 무조건 저장한다. |

### RegisterModule(Method)

| 항목 | 내용 |
|-----|-----|
| Remark | 보안 모듈 관련 문서 참조 |

### ReplaceAction(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 특정 Action을 다른 Action으로 대체한다. |
| Declaration | `bool ReplaceAction(BSTR OldActionID, BSTR NewActionID);` |
| Parameters | OldActionID : 변경될 원본 Action ID. HwpAutomation 사용할 수 있는 Action ID는 ActionObject.hwp(별도문서)를 참고한다.<br>NewActionID : 변경할 대체 Action ID. |
| Remark | - 특정 Action을 다른 Action으로 대체한다.<br>- 이는 메뉴나 단축키로 호출되는 Action을 대체할 뿐, CreateAction()이나, Run() 등의 함수를 이용할 때에는 아무런 영향을 주지 않는다.<br>- 즉, ReplaceAction("Cut", "Copy")을 호출하여 "오려내기" Action을 "복사하기" Action으로 교체하면 Ctrl+X 단축키나 오려내기 메뉴/툴바 기능을 수행하더라도 복사하기 기능이 수행되지만, 코드 상에서 Run("Cut")을 실행하면 오려내기 Action이 실행된다.<br>- ReplaceAction()을 사용할 때에는 대체되는 Action들이 서로 Loop를 형성하지 않도록 조심해야 한다. |

### InitHParameterSet(Method)

| 항목 | 내용 |
|-----|-----|
| Description | HParameterSet Object를 초기화한다. |

### GetPosBySet(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 현재 캐럿의 위치 정보를 ParameterSet으로 얻어온다. |
| Remark | 캐럿의 위치를 ParameterSet으로 얻어온다. 포인터를 사용할 수 없는 언어에서도 사용가능하다. |

### SetPosBySet(Method)

| 항목 | 내용 |
|-----|-----|
| Description | 캐럿을 ParameterSet으로 얻어지는 위치로 옮긴다. |
| Declaration | `boolean SetPosBySet(LPDISPATCH pos);` |
| Parameters | pos : 캐럿을 옮길 위치에 대한 ParameterSet 정보 |

### Application(Property)

| 항목 | 내용 |
|-----|-----|
| Description | 최상위 오브젝트 (IHwpObject interface)를 얻는다. |

### XHwpMessageBox(Property)

| 항목 | 내용 |
|-----|-----|
| Description | 메시지 박스를 제어하는 XHwpMessageBox Object를 얻는다. |

## Page 30

### XHwpDocuments(Property)

| 항목 | 내용 |
|-----|-----|
| Description | 도큐먼트를 관리하는 XHwpDocuments Object를 얻는다. |

### XHwpWindows(Property)

| 항목 | 내용 |
|-----|-----|
| Description | 윈도우를 관리하는 XHwpWindows Object를 얻는다. |

### HParameterSet(Property)

| 항목 | 내용 |
|-----|-----|
| Description | 파라메터셋 오브젝트를 관리하는 HParameterSet Object를 얻는다. |

### HAction(Property)

| 항목 | 내용 |
|-----|-----|
| Description | Action을 제어하는 HAction Object를 얻는다. |

### XHwpODBC(Property)

| 항목 | 내용 |
|-----|-----|
| Description | ODBC로 제어할 수 있는 Object를 얻는다. |

### Version(Property)

| 항목 | 내용 |
|-----|-----|
| Description | 글 과 글 OCX의 버젼 정보를 구한다. 읽기 전용.<br>- byte 3 = 글의 major version.<br>- byte 2 = 글의 minor version.<br>- byte 1 = 글 OCX의 major version.<br>- byte 0 = 글 OCX의 minor version. |

### EngineProperties(Property)

| 항목 | 내용 |
|-----|-----|
| Description | 환경설정 정보를 설정한다. |
| Declaration | `void EngineProperties(LPDISPATCH param)` - set<br>`LPDISPATCH* EngineProperties()` - get |
| Parameters | param : HParameterSet |
| Remark | 환경설정에서 지정할 수 있는 옵션 값을 설정할 수 있다. |

### EngineProperties Item 표

| Item ID | Type | Description |
|---------|------|-------------|
| DoAnyCursorEdit | PIT_UI1 | 마우스로 두 번 누르기 한곳에 입력가능 |
| DoOutLineStyle | PIT_UI1 | 개요 번호 삽입 문단에 개요 스타일 자동 적용 |
| InsertLock | PIT_UI1 | 삽입 잠금 |
| TabMoveCell | PIT_UI1 | 표 안에서 <Tab>으로 셀 이동 |
| FaxDriver | PIT_BSTR | 팩스 드라이버 |
| PDFDriver | PIT_BSTR | PDF 드라이버 |
| EnableAutoSpell | PIT_UI1 | 맞춤법 도우미 작동 |
| OpenNewWindow | PIT_UI1 | 새창으로 불러오기 |
| CtrlMaskAs2002 | PIT_UI1 | 한글 2002 방식으로 조판 부호 표시하기 (한글 버전 : 7.5.11.603) |
| ShowGuildLines | PIT_UI1 | 투명선 보이기 (한글 버전 : 7.5.11.603) |
| ImageAutoCheck | PIT_UI1 | 이미지 파일의 경로 검사 다이얼로그 띄우기 유무. |
| OptimizeWebHwpCopy | PIT_UI1 | 웹한글로 복사 최적화 끄고/켜기 |

### ExportStyle(Method)

| 항목 | 내용 |
|-----|-----|
| Description | HStyleTemplate 파라메터셋 오브젝트에 지정된 스타일을 Export한다. |
| Declaration | `BOOL ExportStyle(LPDISPATCH param)` |
| Parameters | param : HStyleTemplate |

### ImportStyle(Method)

| 항목 | 내용 |
|-----|-----|
| Description | (다음 페이지 참조) |
