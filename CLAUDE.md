# 한글(HWP) 자동화 프로젝트 규칙

## 필수: 보안 모듈 등록

한글 자동화 코드 작성 시 **반드시** 보안 모듈을 등록해야 합니다.
이를 통해 파일 열기/저장 시 보안 승인 팝업이 나타나지 않습니다.

```python
import win32com.client as win32

hwp = win32.Dispatch('HWPFrame.HwpObject')
hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')  # 필수!
```

## 기존 인스턴스 연결 방법 (ROT 사용)

이미 실행 중인 한글에 연결하려면 **Running Object Table(ROT)**을 사용해야 합니다.
`win32.Dispatch()`는 새 인스턴스를 생성하므로, 기존 인스턴스에 연결하려면 아래 방법을 사용하세요.

```python
import pythoncom
import win32com.client as win32

def get_hwp_instance():
    """실행 중인 한글 인스턴스에 연결하거나, 없으면 새로 생성"""
    context = pythoncom.CreateBindCtx(0)
    rot = pythoncom.GetRunningObjectTable()

    for moniker in rot:
        name = moniker.GetDisplayName(context, None)
        if 'HwpObject' in name:
            obj = rot.GetObject(moniker)
            hwp = win32.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
            return hwp

    # 실행 중인 한글이 없으면 새로 생성
    hwp = win32.Dispatch('HWPFrame.HwpObject')
    hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')
    hwp.XHwpWindows.Item(0).Visible = True
    return hwp

hwp = get_hwp_instance()
```

## 테이블 생성 시 주의사항

- 액션 이름: `TableCreate` (TableCreation 아님)
- 파라미터 전달 시 `.HSet` 사용 필수

```python
ctrl = hwp.HParameterSet.HTableCreation
hwp.HAction.GetDefault("TableCreate", ctrl.HSet)
ctrl.Rows = 5
ctrl.Cols = 5
hwp.HAction.Execute("TableCreate", ctrl.HSet)
```

## 텍스트 삽입

```python
hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
hwp.HParameterSet.HInsertText.Text = "텍스트"
hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
```

## 글자/문단 서식 변경 시 주의사항 (필수!)

글자체, 글자 크기, 색상, 자간, 문단 정렬 등 **모든 서식 변경 시 반드시 먼저 블록(범위)을 지정**해야 합니다.
블록 지정 없이 서식 변경 시 적용되지 않습니다.

```python
# 방법 1: 전체 선택
hwp.HAction.Run("SelectAll")

# 방법 2: SelectText로 정확한 범위 선택 (가장 확실한 방법)
# SelectText(시작문단, 시작위치, 끝문단, 끝위치)
hwp.SelectText(0, 0, 0, 10)  # Para=0의 0~10번째 글자 선택

# 방법 3: 현재 문단 선택
hwp.HAction.Run("SelectPararaph")

# 방법 4: 선택 모드로 커서 이동
hwp.HAction.Run("MoveSelRight")  # 오른쪽으로 선택하며 이동
hwp.HAction.Run("MoveSelNextWord")  # 다음 단어까지 선택

# 블록 지정 후 서식 변경 예시
hwp.SelectText(2, 0, 2, 50)  # 3번째 문단의 0~50 위치 선택
hwp.HAction.GetDefault("CharShape", hwp.HParameterSet.HCharShape.HSet)
hwp.HParameterSet.HCharShape.TextColor = hwp.RGBColor(255, 0, 0)  # 빨간색
hwp.HAction.Execute("CharShape", hwp.HParameterSet.HCharShape.HSet)
hwp.HAction.Run("Cancel")  # 선택 해제
```

## 참고 문서 (필수 참조)

HWP API 관련 작업 시 **반드시** `C:\hwp_docs\win32\` 폴더의 문서를 먼저 확인할 것:

- `ActionTable_2504_part01~03.md` - 액션 테이블 (사용 가능한 액션 목록)
- `ParameterSetTable_2504_part01~05.md` - 파라미터셋 (HCharShape, HTableCreation 등)
- `OnlyMethod_2504.md` - 메서드 전용 API
- `OnlyProperty_2504.md` - 속성 전용 API

API 사용법이 불확실할 때는 위 문서를 먼저 검색해서 정확한 사용법을 확인한다.

## 커서 이벤트 처리 전략

### 문제점
- 한글은 사용자 커서 이동과 API 커서 이동을 구분하지 않음
- 둘 다 `DocumentChange` 이벤트 발생
- API로 위치 정보를 읽으려면 커서를 이동해야 함 → 무한 루프 위험

### 해결 전략
1. **전역 플래그 `IS_API_MOVING`**: API가 커서 이동 중일 때 True
2. **시간 기반 판단**: 마지막 이벤트로부터 2초 이상 지나면 사용자 이벤트로 간주

```python
# 전역 플래그
IS_API_MOVING = False
LAST_EVENT_TIME = 0
EVENT_THRESHOLD = 2  # 2초

def is_user_event():
    """사용자 이벤트인지 판단"""
    global IS_API_MOVING, LAST_EVENT_TIME

    if IS_API_MOVING:
        return False  # API가 움직이는 중

    current_time = time.time()
    if current_time - LAST_EVENT_TIME >= EVENT_THRESHOLD:
        LAST_EVENT_TIME = current_time
        return True  # 2초 이상 지남 → 사용자 이벤트

    return False

def GetPosWithinPara(hwp):
    """문단 시작/끝 위치 추출 (커서 이동 필요)"""
    global IS_API_MOVING

    IS_API_MOVING = True  # 플래그 ON

    current_pos = hwp.GetPos()
    hwp.HAction.Run("MoveParaBegin")  # → DocumentChange 발생 → 무시됨
    para_start = hwp.GetPos()
    hwp.HAction.Run("MoveParaEnd")    # → DocumentChange 발생 → 무시됨
    para_end = hwp.GetPos()
    hwp.SetPos(*current_pos)          # → DocumentChange 발생 → 무시됨

    IS_API_MOVING = False  # 플래그 OFF

    return current_pos, para_start, para_end

def on_document_change():
    """이벤트 핸들러"""
    if not is_user_event():
        return  # API 이벤트거나 2초 이내면 무시

    # 사용자 이벤트일 때만 처리
    result = GetPosWithinPara(hwp)
```

참고: `pos_monitoring.py` 파일에 전체 구현 있음
