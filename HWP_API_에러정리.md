# HWP API 에러 및 해결 방법 정리

## 1. ROT 연결 관련

### 문제: `win32.GetActiveObject()` 실패
```
pywintypes.com_error: (-2147221021, 'Operation unavailable', None, None)
```

### 원인
`GetActiveObject()`는 한글에서 지원하지 않음

### 해결
**Running Object Table(ROT)**을 사용하여 기존 인스턴스에 연결:
```python
import pythoncom
import win32com.client as win32

def get_hwp_instance():
    pythoncom.CoInitialize()
    context = pythoncom.CreateBindCtx(0)
    rot = pythoncom.GetRunningObjectTable()

    for moniker in rot:
        name = moniker.GetDisplayName(context, None)
        if 'HwpObject' in name:
            obj = rot.GetObject(moniker)
            return win32.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
    return None
```

---

## 2. 속성명 오류

### 문제: `CharWidthRatio` 속성 없음
```
AttributeError: 'HCharShape' object has no attribute 'CharWidthRatio'
```

### 해결
장평 속성명은 언어별로 다름:
```python
# 올바른 속성명
hwp.HParameterSet.HCharShape.RatioHangul = 100   # 한글 장평 (50~200%)
hwp.HParameterSet.HCharShape.RatioLatin = 100    # 영문 장평
hwp.HParameterSet.HCharShape.RatioHanja = 100    # 한자 장평
# ... 기타 언어별 속성
```

---

## 3. 자간 설정

### 올바른 속성명
```python
# 자간 설정 (-50% ~ 50%)
hwp.HParameterSet.HCharShape.SpacingHangul = 0   # 한글 자간
hwp.HParameterSet.HCharShape.SpacingLatin = 0    # 영문 자간
hwp.HParameterSet.HCharShape.SpacingHanja = 0    # 한자 자간
hwp.HParameterSet.HCharShape.SpacingJapanese = 0 # 일본어 자간
hwp.HParameterSet.HCharShape.SpacingOther = 0    # 기타 자간
hwp.HParameterSet.HCharShape.SpacingSymbol = 0   # 심벌 자간
hwp.HParameterSet.HCharShape.SpacingUser = 0     # 사용자 자간
```

---

## 4. 블록 선택 없이 서식 변경 시도

### 문제
서식 변경 코드 실행해도 적용되지 않음

### 원인
블록(범위)을 지정하지 않고 서식 변경 시도

### 해결
**반드시 블록을 먼저 선택한 후 서식 변경**:
```python
# 방법 1: SelectText 사용 (가장 확실)
hwp.SelectText(시작문단, 시작위치, 끝문단, 끝위치)

# 방법 2: 전체 선택
hwp.HAction.Run("SelectAll")

# 서식 변경 후 선택 해제
hwp.HAction.Run("Cancel")
```

---

## 5. SelectText 범위 오류

### 문제
마지막 글자가 선택되지 않음

### 원인
`SelectText`의 끝 위치가 exclusive (해당 위치 미포함)

### 해결
끝 위치를 +1 또는 +2 해서 원하는 글자까지 포함:
```python
# "안녕하세요." 전체 선택 (마침표 포함)
# 마침표가 5번 위치라면
hwp.SelectText(0, 0, 0, 7)  # 끝 위치를 +2
```

---

## 6. MoveNextSentence 미작동

### 문제
`hwp.HAction.Run('MoveNextSentence')` 실행해도 커서가 이동하지 않음

### 원인
한글 API에 문장 단위 이동 액션이 없음

### 해결
텍스트를 읽어서 마침표(.) 위치를 직접 찾아 이동:
```python
# 문단 텍스트 읽기
hwp.SetPos(0, para_num, 0)
hwp.HAction.Run('MoveParaEnd')
hwp.HAction.Run('MoveSelParaBegin')
text = hwp.GetTextFile('TEXT', '')
hwp.HAction.Run('Cancel')

# 마침표 위치 찾기
first_dot = text.find('.')
second_dot = text.find('.', first_dot + 1)

# 2번째 문장 선택
start_pos = first_dot + 1
end_pos = second_dot + 2  # 마침표 포함
hwp.SelectText(para_num, start_pos, para_num, end_pos)
```

---

## 7. FindReplace 속성 오류

### 문제
```
AttributeError: 'HFindReplace' object has no attribute 'FindFlag'
```

### 해결
`FindFlag` 대신 다른 방법으로 검색 후 위치 확인:
```python
hwp.HAction.GetDefault('FindReplace', hwp.HParameterSet.HFindReplace.HSet)
hwp.HParameterSet.HFindReplace.FindString = '.'
hwp.HParameterSet.HFindReplace.Direction = 0  # 앞으로
hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
hwp.HAction.Execute('FindReplace', hwp.HParameterSet.HFindReplace.HSet)
hwp.HAction.Run('Cancel')

# 현재 위치 확인
pos = hwp.GetPos()
```

---

## 8. 포트 충돌 (웹서버)

### 문제
```
OSError: [Errno 10048] 액세스 권한에 의해 소켓에 액세스가 금지됨
```

### 원인
해당 포트가 이미 사용 중

### 해결
다른 포트 사용:
```python
app.run(host='127.0.0.1', port=9999, debug=False)
```

---

## 요약: 자주 발생하는 실수

| 실수 | 해결 |
|------|------|
| 블록 선택 안 함 | `SelectText()` 또는 `SelectAll` 먼저 실행 |
| 잘못된 속성명 | API 문서에서 정확한 속성명 확인 |
| SelectText 범위 | 끝 위치 +1~2 추가 |
| 문장 이동 액션 | 직접 마침표 찾아서 이동 |
| 새 인스턴스 생성 | ROT로 기존 인스턴스 연결 |
