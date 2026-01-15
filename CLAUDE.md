# 한글(HWP) 자동화 프로젝트 규칙

## 참고 문서 (필수!)

HWP API 관련 작업 시 **반드시** 아래 문서를 먼저 확인:

### 로컬 문서 (D:\hwp_docs\win32\)
- `ActionTable_2504_part01~06.md` - 액션 테이블 (사용 가능한 액션 목록)
- `ParameterSetTable_2504_part01~18.md` - 파라미터셋 (HCharShape, HTableCreation 등)
- `HwpAutomation_2504_part01~07.md` - 메서드/속성 API

### 웹 검색
API 사용법이 불확실하거나 로컬 문서에 없으면 웹 검색으로 확인

---

## 필수: 보안 모듈 등록

한글 자동화 코드 작성 시 **반드시** 보안 모듈을 등록해야 합니다.

```python
import win32com.client as win32

hwp = win32.Dispatch('HWPFrame.HwpObject')
hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')  # 필수!
```

## 기존 인스턴스 연결 (ROT)

이미 실행 중인 한글에 연결하려면 **Running Object Table(ROT)** 사용:

```python
import pythoncom
import win32com.client as win32

def get_hwp_instance():
    context = pythoncom.CreateBindCtx(0)
    rot = pythoncom.GetRunningObjectTable()

    for moniker in rot:
        name = moniker.GetDisplayName(context, None)
        if 'HwpObject' in name:
            obj = rot.GetObject(moniker)
            return win32.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
    return None
```

## 주요 API 패턴

### 테이블 생성
- 액션 이름: `TableCreate` (TableCreation 아님)
- 파라미터 전달 시 `.HSet` 사용 필수

### 텍스트 삽입
```python
hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
hwp.HParameterSet.HInsertText.Text = "텍스트"
hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
```

### 서식 변경 (필수: 블록 지정 먼저!)
```python
hwp.HAction.Run("SelectAll")  # 또는 hwp.SelectText(para, pos, para, pos)
# 서식 변경 후
hwp.HAction.Run("Cancel")  # 선택 해제
```

## 커서 위치 관련

### 위치 정보 API
- `hwp.GetPos()` → (list_id, para_id, char_pos)
- `hwp.SetPos(list, para, pos)` → 커서 이동
- `hwp.KeyIndicator()` → (총구역, 현재구역, 페이지, 단, 줄, 칸, 수정모드)

### 커서 이동 액션
- `MoveParaBegin` / `MoveParaEnd` - 문단 시작/끝
- `MoveRight` / `MoveLeft` - 한 글자 이동
- `MoveDown` / `MoveUp` - 한 줄 이동
- `MoveSelRight` - 선택하며 이동

### 텍스트 가져오기
```python
hwp.HAction.Run("MoveSelRight")  # 선택
text = hwp.GetTextFile("TEXT", "saveblock")  # 선택 영역만
hwp.HAction.Run("Cancel")  # 선택 해제
```

## 주의사항

### 이벤트 기반 vs 폴링
- `DocumentChange` 이벤트: API 커서 이동도 발생 → 무한 루프 위험
- **폴링 방식 권장**: 0.1초 간격으로 `GetPos()` 비교

### 위치 정보 읽을 때
- 문단/줄 범위를 알려면 커서를 이동해야 함
- 이동 후 반드시 원래 위치로 복원: `hwp.SetPos(list, para, pos)`
