# 한글(HWP) 자동화 프로젝트 규칙

## 필수: 보안 모듈 등록

한글 자동화 코드 작성 시 **반드시** 보안 모듈을 등록해야 합니다.

```python
import win32com.client as win32

hwp = win32.Dispatch('HWPFrame.HwpObject')
hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')  # 필수!
```

---

## 기존 인스턴스 연결 (ROT)

이미 실행 중인 한글에 연결하려면 **Running Object Table(ROT)** 사용:

```python
from cursor import get_hwp_instance

hwp = get_hwp_instance()
if not hwp:
    print("한글이 실행 중이 아닙니다")
```

### ROT 직접 구현 (Python)

```python
import pythoncom
import win32com.client as win32

def get_hwp_from_rot():
    """ROT에서 실행 중인 한글 인스턴스 찾기"""
    context = pythoncom.CreateBindCtx(0)
    rot = pythoncom.GetRunningObjectTable()

    for moniker in rot:
        name = moniker.GetDisplayName(context, None)
        if "HwpObject" in name:
            obj = rot.GetObject(moniker)
            hwp = win32.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
            return hwp
    return None
```

**ROT 방식의 특징:**
- 이미 실행 중인 한글 프로세스에 연결
- 새 인스턴스를 생성하지 않음
- 한글이 먼저 열려 있어야 함

---

## 새 파일 열기

```python
import win32com.client as win32

hwp = win32.Dispatch('HWPFrame.HwpObject')
hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')

# 파일 열기
hwp.Open("C:\\path\\to\\file.hwp")

# 또는 옵션과 함께
hwp.Open("C:\\path\\to\\file.hwp", "HWP", "lock:false")
```

### Open 메서드 arg 옵션 (HWP 형식)

| 옵션 | 타입 | 기본값 | 설명 |
|------|------|--------|------|
| `lock` | boolean | TRUE | 파일 lock 유지 여부 |
| `notext` | boolean | FALSE | 헤더만 읽기 (스타일 로드용) |
| `template` | boolean | FALSE | 템플릿으로 열기 (lock 자동 해제) |
| `suspendpassword` | boolean | FALSE | 암호 묻지 않고 실패 처리 |
| `forceopen` | boolean | FALSE | 읽기 전용 대화상자 숨김 |

---

## WSL에서 Windows Python 실행

WSL 환경에서는 `win32com`이 작동하지 않습니다. Windows 측 Python을 호출해야 합니다.

```bash
# WSL에서 Windows Python 실행
cmd.exe /c "cd /d C:\\win32hwp && python script.py"

# 또는
cmd.exe /c "python C:\\win32hwp\\script.py"
```

**동작 원리:**
1. WSL에서 Windows의 `cmd.exe` 실행
2. `cmd.exe`가 Windows 경로로 이동
3. Windows에 설치된 Python이 스크립트 실행
4. Python이 `win32com`으로 한글 프로그램 제어

**경로 주의:**
- WSL 경로: `/mnt/c/win32hwp/` (파일 읽기/쓰기용)
- Windows 경로: `C:\win32hwp\` (실행용)

---

## 참고 문서

HWP API 관련 작업 시 **반드시** 아래 문서를 먼저 확인:

### 로컬 문서 (`/mnt/d/hwp_docs/win32/`)
- `ActionTable_2504_part01~06.md` - 액션 테이블 (사용 가능한 액션 목록)
- `ParameterSetTable_2504_part01~18.md` - 파라미터셋 (HCharShape, HTableCreation 등)
- `HwpAutomation_2504_part01~07.md` - 메서드/속성 API

### 웹 검색
API 사용법이 불확실하거나 로컬 문서에 없으면 웹 검색으로 확인

---

## 컨트롤 속성 및 주요 CtrlID

### HeadCtrl로 컨트롤 순회

```python
def get_all_ctrls(hwp):
    """문서 내 모든 컨트롤 조회"""
    ctrls = []
    ctrl = hwp.HeadCtrl
    while ctrl:
        try:
            ctrls.append({
                'ctrl': ctrl,
                'id': ctrl.CtrlID,
                'desc': ctrl.UserDesc,
            })
        except:
            pass
        ctrl = ctrl.Next
    return ctrls
```

### 특정 셀 내부 컨트롤 찾기 (GetAnchorPos 필터링)

```python
def get_ctrls_in_cell(hwp, target_list_id):
    """특정 셀(list_id)에 속한 컨트롤 찾기"""
    ctrls = []
    ctrl = hwp.HeadCtrl
    while ctrl:
        try:
            anchor = ctrl.GetAnchorPos(0)
            ctrl_list_id = anchor.Item("List")
            ctrl_para_id = anchor.Item("Para")

            if ctrl_list_id == target_list_id:
                ctrls.append({
                    'ctrl': ctrl,
                    'id': ctrl.CtrlID,
                    'desc': ctrl.UserDesc,
                    'para': ctrl_para_id
                })
        except:
            pass
        ctrl = ctrl.Next
    return ctrls
```

### 컨트롤 속성 (Properties)

| 속성 | 설명 |
|------|------|
| `ctrl.CtrlID` | 컨트롤 종류 ("tbl", "gso", "eqed" 등) |
| `ctrl.UserDesc` | 사용자 친화적 설명 ("그림", "표", "수식" 등) |
| `ctrl.Properties` | 속성 객체 (Width, Height, TreatAsChar 등) |
| `ctrl.GetAnchorPos(0)` | 컨트롤 앵커 위치 (List, Para 등) |
| `ctrl.Next` / `ctrl.Prev` | 다음/이전 컨트롤 |

### GetAnchorPos 반환값

```python
anchor = ctrl.GetAnchorPos(0)
list_id = anchor.Item("List")   # 컨트롤이 속한 list_id (0=본문, 1+=셀)
para_id = anchor.Item("Para")   # 컨트롤이 속한 문단 번호
```

### 주요 CtrlID

| CtrlID | UserDesc | 설명 |
|--------|----------|------|
| `tbl` | 표 | 테이블 |
| `gso` | 그림 | 그리기 개체 (이미지 포함) |
| `eqed` | 수식 | 수식 편집기 |
| `secd` | 구역 정의 | 섹션 (필터링 시 제외 권장) |
| `cold` | 단 정의 | 단 (필터링 시 제외 권장) |
| `toc` | 차례/목차 | 목차 컨트롤 |
| `fn` | 각주 | 각주 |
| `en` | 미주 | 미주 |
| `hn` | 머리말 | 머리말 |
| `ft` | 꼬리말 | 꼬리말 |

### TreatAsChar (글자처럼 취급)

```python
props = ctrl.Properties
treat_as_char = props.Item("TreatAsChar")  # 1=글자취급, 0=아님
```

**주의:** TreatAsChar=1인 개체는 `FindCtrl()`로 찾기 어렵습니다. `HeadCtrl` 순회 방식을 사용하세요.

---

## 주요 IHwpObject 속성

| 속성 | 설명 |
|------|------|
| `IsModified` | 문서 변경 여부 (0=깨끗, 1=변경됨, 2=자동저장됨) |
| `IsEmpty` | 빈 문서 여부 |
| `EditMode` | 편집 모드 (0=읽기전용, 1=일반, 2=양식, 16=배포용) |
| `PageCount` | 전체 페이지 수 |
| `HeadCtrl` | 첫 번째 컨트롤 (linked list 시작) |
| `LastCtrl` | 마지막 컨트롤 |
| `CurSelectedCtrl` | 현재 선택된 컨트롤 |
| `ParentCtrl` | 현재 캐럿의 상위 컨트롤 |
| `CharShape` | 현재 Selection의 글자 모양 |
| `ParaShape` | 현재 Selection의 문단 모양 |
| `CellShape` | 현재 선택된 표/셀의 모양 정보 |
| `Path` | 문서 파일 경로 |

---

## 자주 쓰는 API 패턴

### 텍스트 삽입
```python
hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
hwp.HParameterSet.HInsertText.Text = "텍스트"
hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
```

### 서식 변경 (블록 지정 필수!)
```python
hwp.HAction.Run("SelectAll")  # 또는 hwp.SelectText(...)
# 서식 변경 코드
hwp.HAction.Run("Cancel")  # 선택 해제
```

### 파일 저장
```python
hwp.Save()  # 현재 파일에 저장
hwp.SaveAs("C:\\path\\to\\new.hwp", "HWP", "compress:true")  # 다른 이름으로 저장
```

### 커서 이동 (MovePos)
```python
hwp.MovePos(2)   # 문서 시작 (moveTopOfFile)
hwp.MovePos(3)   # 문서 끝 (moveBottomOfFile)
hwp.MovePos(1, para, pos)  # 현재 리스트의 특정 위치로
```

| moveID | 값 | 설명 |
|--------|---|------|
| moveMain | 0 | 루트 리스트의 특정 위치 |
| moveCurList | 1 | 현재 리스트의 특정 위치 |
| moveTopOfFile | 2 | 문서 시작 |
| moveBottomOfFile | 3 | 문서 끝 |
| moveTopOfList | 4 | 현재 리스트 시작 |
| moveBottomOfList | 5 | 현재 리스트 끝 |
| moveStartOfPara | 6 | 문단 시작 |
| moveEndOfPara | 7 | 문단 끝 |
| moveNextPara | 10 | 다음 문단 |
| movePrevPara | 11 | 이전 문단 |
| moveNextPos | 12 | 한 글자 앞으로 (서브리스트 이동 가능) |
| movePrevPos | 13 | 한 글자 뒤로 (서브리스트 이동 가능) |

### 위치 정보
```python
pos = hwp.GetPos()       # (list_id, para_id, char_pos)
hwp.SetPos(list_id, para_id, char_pos)

key = hwp.KeyIndicator()
page = key[3]            # 페이지 번호
line = key[4]            # 줄 번호
```

---

## Git 규칙

**사용자가 명시적으로 요청한 경우에만** git 명령어를 실행합니다.

- 커밋, 푸시, 브랜치 생성 등 git 작업은 사용자의 직접 요청이 있을 때만 수행
- 코드 작성/수정 후 자동으로 커밋하지 않음
- PR 생성도 사용자 요청 시에만 진행
