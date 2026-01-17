# HWP API 에러 및 해결 방법 정리

> 최종 수정일: 2026-01-16

---

## 목차
1. [COM 연결 관련 에러](#1-com-연결-관련-에러)
2. [보안 모듈 관련](#2-보안-모듈-관련)
3. [속성명/메서드 오류](#3-속성명메서드-오류)
4. [블록 선택 관련](#4-블록-선택-관련)
5. [파일 열기/저장 관련](#5-파일-열기저장-관련)
6. [테이블 관련](#6-테이블-관련)
7. [텍스트 검색/조작 관련](#7-텍스트-검색조작-관련)
8. [시스템/환경 관련](#8-시스템환경-관련)
9. [에러 코드 레퍼런스](#9-에러-코드-레퍼런스)

---

## 1. COM 연결 관련 에러

### 1.1 `GetActiveObject()` 실패
```
pywintypes.com_error: (-2147221021, 'Operation unavailable', None, None)
```

**원인**: `GetActiveObject()`는 한글에서 지원하지 않음

**해결**: **Running Object Table(ROT)**을 사용하여 기존 인스턴스에 연결
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

### 1.2 라이브러리 등록 오류
```
pywintypes.com_error: (-2147319779, '라이브러리가 등록되지 않았습니다.', None, None)
```

**원인**: 한글 오토메이션이 시스템에 등록되지 않음

**해결**: 관리자 권한 CMD에서 오토메이션 등록
```batch
# 관리자 권한 CMD 실행 후
cd "C:\Program Files\HNC\Hwp\2024"
hwp.exe -regserver
```

---

### 1.3 매개 변수 개수 오류
```
pywintypes.com_error: (-2147352562, '매개 변수의 개수가 잘못되었습니다.', None, None)
```

**원인**: `win32.dynamic.Dispatch` 사용 시 기본값이 적용되지 않음

**해결**: `gencache.EnsureDispatch` 사용
```python
# 잘못된 방식
hwp = win32.dynamic.Dispatch('HWPFrame.HwpObject')

# 올바른 방식
hwp = win32.gencache.EnsureDispatch('HWPFrame.HwpObject')
```

또는 MakePy Utility로 타입 라이브러리 재생성:
1. `pythonwin` 실행
2. Tools > COM MakePy utility
3. HWP Type Library 선택 후 OK

---

### 1.4 32비트/64비트 호환성 문제
```
pywintypes.com_error: (-2147221021, 'Operation unavailable', None, None)
```

**원인**: 파이썬 비트와 한글 프로그램 비트가 다름

**해결**:
1. 한글 프로그램의 비트 확인 (보통 32비트)
2. 동일한 비트의 Python 사용
```python
import platform
print(platform.architecture())  # 비트 확인
```

---

### 1.5 한글 프로그램 종료 후 오류
```
pywintypes.com_error: (-2147023174, 'RPC 서버를 사용할 수 없습니다.', None, None)
```

**원인**: hwp.Quit() 실행 후 또는 한글 창이 닫힌 후 hwp 객체 사용

**해결**: 연결 상태 확인 후 사용
```python
def is_hwp_connected(hwp):
    try:
        _ = hwp.Path
        return True
    except:
        return False

if not is_hwp_connected(hwp):
    hwp = get_hwp_instance()  # 재연결
```

---

## 2. 보안 모듈 관련

### 2.1 보안 승인 팝업 발생
**증상**: 파일 열기/저장 시 보안 확인 대화상자 표시

**해결**: 보안 모듈 등록
```python
import win32com.client as win32

hwp = win32.gencache.EnsureDispatch('HWPFrame.HwpObject')
hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')  # 필수!
```

**사전 작업** (레지스트리 등록):
1. `regedit` 실행
2. `HKEY_CURRENT_USER\SOFTWARE\HNC\HwpAutomation\Modules` 경로로 이동
3. 문자열 값 추가:
   - 이름: `FilePathCheckerModule`
   - 값: `C:\경로\FilePathCheckerModuleExample.dll`

---

### 2.2 2024 버전 RegisterModule 문제
**증상**: 한글 2024에서 RegisterModule 사용해도 팝업 지속 발생

**원인**: 2024부터 ActiveX 버전 사용 불가, Object 버전으로 변경됨

**해결**:
1. 최신 보안 모듈 DLL 사용
2. 한컴디벨로퍼 포럼에서 업데이트된 보안 모듈 다운로드
3. SetMessageBoxMode로 대화상자 자동 처리
```python
# 메시지 박스 자동 처리 (예/확인 자동 클릭)
hwp.SetMessageBoxMode(0x00010000)  # MB_YESNO_IDYES
```

---

## 3. 속성명/메서드 오류

### 3.1 `CharWidthRatio` 속성 없음
```
AttributeError: 'HCharShape' object has no attribute 'CharWidthRatio'
```

**해결**: 장평 속성명은 언어별로 다름
```python
# 올바른 속성명 (장평 50~200%)
hwp.HParameterSet.HCharShape.RatioHangul = 100   # 한글 장평
hwp.HParameterSet.HCharShape.RatioLatin = 100    # 영문 장평
hwp.HParameterSet.HCharShape.RatioHanja = 100    # 한자 장평
hwp.HParameterSet.HCharShape.RatioJapanese = 100 # 일본어 장평
hwp.HParameterSet.HCharShape.RatioOther = 100    # 기타 장평
hwp.HParameterSet.HCharShape.RatioSymbol = 100   # 심벌 장평
hwp.HParameterSet.HCharShape.RatioUser = 100     # 사용자 장평
```

---

### 3.2 자간 설정 속성
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

### 3.3 글꼴 크기 설정
**주의**: Height는 HWPUNIT 단위 (1pt = 100 HWPUNIT)
```python
# 12pt로 설정
hwp.HParameterSet.HCharShape.Height = 1200  # 12 * 100

# 또는 API 함수 사용
hwp.PointToHwpUnit(12)  # 1200 반환
```

---

### 3.4 `FindFlag` 속성 없음
```
AttributeError: 'HFindReplace' object has no attribute 'FindFlag'
```

**해결**: 다른 속성들을 개별적으로 설정
```python
hwp.HAction.GetDefault('FindReplace', hwp.HParameterSet.HFindReplace.HSet)
hwp.HParameterSet.HFindReplace.FindString = '검색어'
hwp.HParameterSet.HFindReplace.Direction = 0      # 0: 앞으로, 1: 뒤로
hwp.HParameterSet.HFindReplace.IgnoreMessage = 1  # 메시지 무시
hwp.HParameterSet.HFindReplace.MatchCase = 0      # 대소문자 구분
hwp.HParameterSet.HFindReplace.WholeWordOnly = 0  # 온전한 낱말
hwp.HAction.Execute('FindReplace', hwp.HParameterSet.HFindReplace.HSet)
```

---

### 3.5 개체 속성/메소드 미지원 (438 에러)
```
"개체 속성 또는 메소드를 지원하지 않습니다" (438 에러)
```

**원인**: hwpctrl.ocx를 사용하지 않는 환경에서 해당 API 사용

**해결**:
1. API 문서에서 지원 여부 확인
2. HwpCtrl과 HwpObject 차이 확인
3. 대체 API 사용

---

## 4. 블록 선택 관련

### 4.1 서식 변경이 적용되지 않음
**문제**: 서식 변경 코드 실행해도 적용되지 않음

**원인**: 블록(범위)을 지정하지 않고 서식 변경 시도

**해결**: 반드시 블록을 먼저 선택한 후 서식 변경
```python
# 방법 1: SelectText 사용 (가장 확실)
hwp.SelectText(시작문단, 시작위치, 끝문단, 끝위치)

# 방법 2: 전체 선택
hwp.HAction.Run("SelectAll")

# 서식 변경 코드...

# 서식 변경 후 선택 해제
hwp.HAction.Run("Cancel")
```

---

### 4.2 SelectText 범위 오류
**문제**: 마지막 글자가 선택되지 않음

**원인**: `SelectText`의 끝 위치가 exclusive (해당 위치 미포함)

**해결**: 끝 위치를 +1 또는 +2
```python
# "안녕하세요." 전체 선택 (마침표 포함)
# 한글 문자 = pos 2, 영문/공백/기호 = pos 1
# 예: "안녕" = pos 4, "안녕." = pos 5
hwp.SelectText(0, 0, 0, 7)  # 끝 위치를 +2
```

**위치 계산 팁**:
- 한글 1글자 = pos 2
- 영문/숫자/공백/기호 1글자 = pos 1
- 문단 끝 = pos 1

---

### 4.3 현재 블록 위치 가져오기
```python
# 포인터 사용 가능한 언어
slist, spara, spos, elist, epara, epos = [0]*6
hwp.GetSelectedPos(slist, spara, spos, elist, epara, epos)

# 포인터 사용 불가한 언어 (Python 등)
sset = hwp.CreateSet("ListParaPos")
eset = hwp.CreateSet("ListParaPos")
hwp.GetSelectedPosBySet(sset, eset)
```

---

## 5. 파일 열기/저장 관련

### 5.1 파일 열기 실패
**증상**: `hwp.Open()` 호출 시 실패

**가능한 원인들**:
1. 파일 경로 오류
2. 파일이 다른 프로그램에서 사용 중
3. 암호가 걸린 파일
4. 손상된 파일

**해결**:
```python
# 1. 절대 경로 사용 (역슬래시 주의)
path = r"C:\Users\Documents\test.hwp"
# 또는
path = "C:/Users/Documents/test.hwp"

# 2. 파일 존재 확인
import os
if os.path.exists(path):
    result = hwp.Open(path)
    if not result:
        print("파일 열기 실패")

# 3. 암호 파일 처리 옵션
hwp.Open(path, "HWP", "suspendpassword:true")  # 암호 파일 건너뛰기
```

---

### 5.2 HWPX 파일 열기 오류
**증상**: `.hwpx` 확장자 파일 열기 실패

**해결**:
1. 한글 버전 확인 (2014 이상 필요)
2. 형식 명시적 지정
```python
hwp.Open(path, "HWPX", "")
```

---

### 5.3 저장 시 옵션 설정
```python
# 압축 저장
hwp.SaveAs(path, "HWP", "compress:true")

# 백업 파일 생성
hwp.SaveAs(path, "HWP", "backup:true")

# 미리보기 이미지 포함
hwp.SaveAs(path, "HWP", "prvimage:2")  # 0=off, 1=BMP, 2=GIF
```

---

### 5.4 읽기 전용 파일 문제
```python
# 읽기 전용 경고 대화상자 없이 열기
hwp.Open(path, "HWP", "forceopen:true")
```

---

## 6. 테이블 관련

### 6.1 테이블 생성 액션명 오류
**문제**: `TableCreation` 액션이 동작하지 않음

**원인**: 액션명은 `TableCreate` (TableCreation 아님)

**해결**:
```python
# 방법 1: HParameterSet 사용
pset = hwp.HParameterSet.HTableCreation
hwp.HAction.GetDefault("TableCreate", pset.HSet)  # 액션명: TableCreate
pset.Rows = 5
pset.Cols = 3
pset.WidthType = 0  # 0=단에맞춤, 1=문단에맞춤, 2=임의값
pset.HeightType = 0  # 0=자동, 1=임의값
hwp.HAction.Execute("TableCreate", pset.HSet)

# 방법 2: CreateAction 사용
act = hwp.CreateAction("TableCreate")
pset = act.CreateSet()
act.GetDefault(pset)
pset.SetItem("Cols", 3)
pset.SetItem("Rows", 5)
act.Execute(pset)
```

---

### 6.2 셀 테두리/배경 설정
```python
# 셀이 선택된 상태에서만 동작
hwp.HAction.Run("TableCellBlock")  # 셀 블록 모드

act = hwp.CreateAction("CellBorderFill")
pset = act.CreateSet()
act.GetDefault(pset)

# 테두리 설정
pset.SetItem("BorderTypeLeft", 1)   # 실선
pset.SetItem("BorderTypeRight", 1)
pset.SetItem("BorderTypeTop", 1)
pset.SetItem("BorderTypeBottom", 1)
pset.SetItem("BorderWidthLeft", 1)  # 0.1mm
pset.SetItem("BorderColorLeft", 0)  # 검정 (BGR)

act.Execute(pset)
```

---

### 6.3 표 내 셀 이동
```python
# 현재 셀 기준 이동 (MovePos 사용)
hwp.MovePos(100)  # moveLeftOfCell: 왼쪽 셀
hwp.MovePos(101)  # moveRightOfCell: 오른쪽 셀
hwp.MovePos(102)  # moveUpOfCell: 위쪽 셀
hwp.MovePos(103)  # moveDownOfCell: 아래쪽 셀
hwp.MovePos(104)  # moveStartOfCell: 행 시작
hwp.MovePos(105)  # moveEndOfCell: 행 끝
hwp.MovePos(106)  # moveTopOfCell: 열 시작
hwp.MovePos(107)  # moveBottomOfCell: 열 끝
```

---

## 7. 텍스트 검색/조작 관련

### 7.1 MoveNextSentence 미작동
**문제**: `hwp.HAction.Run('MoveNextSentence')` 실행해도 커서가 이동하지 않음

**원인**: 한글 API에 문장 단위 이동 액션이 없음

**해결**: 텍스트를 읽어서 마침표(.) 위치를 직접 찾아 이동
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

### 7.2 GetText 초기화 오류
```
GetText() Return: 101 (초기화 안됨)
```

**원인**: InitScan() 호출하지 않음 또는 실패

**해결**:
```python
# 올바른 사용 순서
hwp.InitScan()  # 1. 초기화
text = ""
while True:
    state, txt = hwp.GetText()  # 2. 텍스트 가져오기
    if state <= 1:
        break
    text += txt
hwp.ReleaseScan()  # 3. 반드시 해제!
```

---

### 7.3 GetText 반환값
| 값 | 의미 |
|----|------|
| 0 | 텍스트 정보 없음 |
| 1 | 리스트의 끝 |
| 2 | 일반 텍스트 |
| 3 | 다음 문단 |
| 4 | 제어문자 내부로 들어감 |
| 5 | 제어 문자를 빠져 나옴 |
| 101 | 초기화 안됨 (InitScan 필요) |
| 102 | 텍스트 변환 실패 |

---

### 7.4 InitScan 옵션
```python
# 검색 대상 옵션
maskNormal = 0x00   # 본문만 (기본값)
maskChar = 0x01     # char 타입 컨트롤 포함
maskInline = 0x02   # inline 타입 컨트롤 포함
maskCtrl = 0x04     # extended 타입 컨트롤 포함

# 검색 범위 옵션
scanSposDocument = 0x0070  # 문서 시작부터
scanEposDocument = 0x0007  # 문서 끝까지

# 예: 문서 전체 검색
hwp.InitScan(0x07, 0x0077)
```

---

## 8. 시스템/환경 관련

### 8.1 포트 충돌 (웹서버)
```
OSError: [Errno 10048] 액세스 권한에 의해 소켓에 액세스가 금지됨
```

**원인**: 해당 포트가 이미 사용 중

**해결**: 다른 포트 사용
```python
app.run(host='127.0.0.1', port=9999, debug=False)
```

---

### 8.2 스택 오버플로우 오류
```
'Hwp.exe- 시스템 오류'
'스택에 사용할 새 보호 페이지를 만들 수 없습니다.'
```

**원인**: 재귀 호출 또는 대용량 문서 처리 중 메모리 부족

**해결**:
1. 대용량 문서는 분할 처리
2. 루프에서 불필요한 객체 해제
3. 문서 처리 후 `hwp.Clear()` 호출

---

### 8.3 이벤트 기반 vs 폴링
**주의**: `DocumentChange` 이벤트는 API 커서 이동도 발생 -> 무한 루프 위험

**권장 방식**: 폴링
```python
import time

last_pos = None
while True:
    current_pos = hwp.GetPos()
    if current_pos != last_pos:
        # 위치 변경 처리
        last_pos = current_pos
    time.sleep(0.1)  # 0.1초 간격
```

---

## 9. 에러 코드 레퍼런스

### COM 에러 코드 (pywintypes.com_error)

| 에러 코드 | HEX | 메시지 | 원인 및 해결 |
|-----------|-----|--------|--------------|
| -2147221021 | 0x800401E3 | Operation unavailable | ROT 연결 사용, 32/64비트 확인 |
| -2147352562 | 0x80020006 | Invalid number of parameters | gencache.EnsureDispatch 사용 |
| -2147319779 | 0x8002802D | Library not registered | hwp.exe -regserver 실행 |
| -2147352573 | 0x80020003 | Member not found | 속성/메서드명 확인 |
| -2147467262 | 0x80004002 | No such interface supported | COM 인터페이스 확인 |
| -2147023174 | 0x800706BA | RPC server unavailable | 한글 프로그램 실행 확인 |

---

### MovePos moveID 값

| ID | 값 | 설명 |
|----|---|-----|
| moveMain | 0 | 루트 리스트의 특정 위치 |
| moveCurList | 1 | 현재 리스트의 특정 위치 |
| moveTopOfFile | 2 | 문서의 시작 |
| moveBottomOfFile | 3 | 문서의 끝 |
| moveTopOfList | 4 | 현재 리스트의 시작 |
| moveBottomOfList | 5 | 현재 리스트의 끝 |
| moveStartOfPara | 6 | 문단의 시작 |
| moveEndOfPara | 7 | 문단의 끝 |
| moveStartOfWord | 8 | 단어의 시작 |
| moveEndOfWord | 9 | 단어의 끝 |
| moveNextPara | 10 | 다음 문단 시작 |
| movePrevPara | 11 | 이전 문단 끝 |
| moveNextPos | 12 | 한 글자 앞으로 |
| movePrevPos | 13 | 한 글자 뒤로 |
| moveNextChar | 16 | 한 글자 앞으로 (현재 리스트만) |
| movePrevChar | 17 | 한 글자 뒤로 (현재 리스트만) |
| moveNextWord | 18 | 한 단어 앞으로 |
| movePrevWord | 19 | 한 단어 뒤로 |
| moveNextLine | 20 | 한 줄 위로 |
| movePrevLine | 21 | 한 줄 아래로 |
| moveStartOfLine | 22 | 줄의 시작 |
| moveEndOfLine | 23 | 줄의 끝 |

---

### 메시지 박스 모드 값

| 모드 | 값 | 설명 |
|------|-----|------|
| MB_OK_IDOK | 0x00000001 | 확인 버튼 자동 클릭 |
| MB_OKCANCEL_IDOK | 0x00000010 | 확인/취소 중 확인 |
| MB_OKCANCEL_IDCANCEL | 0x00000020 | 확인/취소 중 취소 |
| MB_YESNO_IDYES | 0x00010000 | 예/아니오 중 예 |
| MB_YESNO_IDNO | 0x00020000 | 예/아니오 중 아니오 |
| MB_YESNOCANCEL_IDYES | 0x00001000 | 예/아니오/취소 중 예 |

---

## 요약: 자주 발생하는 실수

| 실수 | 해결 |
|------|------|
| 블록 선택 안 함 | `SelectText()` 또는 `SelectAll` 먼저 실행 |
| 잘못된 속성명 | API 문서에서 정확한 속성명 확인 (언어별 구분) |
| SelectText 범위 | 끝 위치 +1~2 추가 |
| 문장 이동 액션 | 직접 마침표 찾아서 이동 |
| 새 인스턴스 생성 | ROT로 기존 인스턴스 연결 |
| 보안 팝업 | RegisterModule로 보안 모듈 등록 |
| TableCreate 액션명 | TableCreation이 아닌 TableCreate |
| GetText 전 InitScan | InitScan() 반드시 먼저 호출, ReleaseScan()으로 해제 |
| 매개변수 오류 | gencache.EnsureDispatch 사용 |
| 단위 변환 | Height는 HWPUNIT (1pt = 100) |

---

## 참고 자료

### 공식 문서
- [한컴디벨로퍼](https://developer.hancom.com/hwpautomation)
- [한컴디벨로퍼 포럼](https://forum.developer.hancom.com/c/hwp-automation/52)

### 커뮤니티/패키지
- [pyhwpx](https://pypi.org/project/pyhwpx/) - 파이썬 HWP 자동화 래퍼
- [hwpapi](https://pypi.org/project/hwpapi/) - 파이썬 HWP API 패키지
- [pyhwpx Cookbook](https://wikidocs.net/book/8956)

### 로컬 문서 (win32/)
- `ActionTable_2504_part01~06.md` - 액션 테이블
- `ParameterSetTable_2504_part01~18.md` - 파라미터셋
- `HwpAutomation_2504_part01~07.md` - 메서드/속성 API
