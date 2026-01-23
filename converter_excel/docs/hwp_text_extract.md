# HWP API 텍스트 추출 가이드

HWP 자동화 API를 사용하여 문서에서 텍스트를 추출하는 다양한 방법을 정리합니다.

---

## 목차

1. [InitScan/GetText - 문서 검색 기반 추출](#1-initscangettext---문서-검색-기반-추출)
2. [GetTextFile - 문서 전체 텍스트 추출](#2-gettextfile---문서-전체-텍스트-추출)
3. [GetPageText - 페이지 단위 추출](#3-getpagetext---페이지-단위-추출)
4. [GetFieldText - 필드 텍스트 추출](#4-getfieldtext---필드-텍스트-추출)
5. [셀 단위 텍스트 추출](#5-셀-단위-텍스트-추출)
6. [SelectText + GetTextFile 조합](#6-selecttext--gettextfile-조합)
7. [실제 사용 예제](#7-실제-사용-예제)

---

## 1. InitScan/GetText - 문서 검색 기반 추출

문서를 순차적으로 스캔하면서 텍스트를 추출하는 방법입니다.

### API 설명

#### InitScan(Method)
문서 검색을 위한 준비 작업을 수행합니다.

```
BOOL InitScan([unsigned int option], [unsigned int range],
              [unsigned int spara], [unsigned int spos],
              [unsigned int epara], [unsigned int epos])
```

**option 값:**
| ID | 값 | 설명 |
|----|---|------|
| maskNormal | 0x00 | 본문만 검색 (서브리스트 제외) |
| maskChar | 0x01 | char 타입 컨트롤 (강제줄나눔, 문단끝 등) |
| maskInline | 0x02 | inline 타입 컨트롤 (누름틀 필드 끝 등) |
| maskCtrl | 0x04 | extended 타입 컨트롤 (표, 머리말, 각주 등) |

**range 값 (시작 위치):**
| ID | 값 | 설명 |
|----|---|------|
| scanSposCurrent | 0x0000 | 캐럿 위치부터 |
| scanSposDocument | 0x0070 | 문서 시작부터 |

**range 값 (끝 위치):**
| ID | 값 | 설명 |
|----|---|------|
| scanEposCurrent | 0x0000 | 캐럿 위치까지 |
| scanEposDocument | 0x0007 | 문서 끝까지 |

**range 값 (검색 방향):**
| ID | 값 | 설명 |
|----|---|------|
| scanForward | 0x0000 | 정방향 |
| scanBackward | 0x0100 | 역방향 |

#### GetText(Method)
문서에서 텍스트를 얻습니다.

```
long GetText(BSTR FAR* text)
```

**반환값:**
| 값 | 설명 |
|----|------|
| 0 | 텍스트 정보 없음 |
| 1 | 리스트의 끝 |
| 2 | 일반 텍스트 |
| 3 | 다음 문단 |
| 4 | 제어문자 내부로 들어감 |
| 5 | 제어문자를 빠져나옴 |
| 101 | 초기화 안됨 |
| 102 | 텍스트 변환 실패 |

#### ReleaseScan(Method)
InitScan으로 설정된 정보를 초기화합니다.

### 예제 코드

```python
import win32com.client as win32

def extract_all_text_scan(hwp):
    """InitScan/GetText로 문서 전체 텍스트 추출"""
    texts = []

    # 검색 범위 설정: 문서 전체
    # option=0x07 (모든 컨트롤 포함), range=0x0077 (문서 시작~끝)
    hwp.InitScan(0x07, 0x0077)

    while True:
        state, text = hwp.GetText()

        if state == 0:  # 텍스트 없음
            continue
        elif state == 1:  # 리스트 끝
            break
        elif state == 2:  # 일반 텍스트
            if text:
                texts.append(text)
        elif state == 3:  # 다음 문단
            texts.append('\n')
        elif state == 4:  # 제어문자 내부 진입 (표 등)
            pass
        elif state == 5:  # 제어문자 탈출
            pass
        elif state >= 100:  # 오류
            break

    hwp.ReleaseScan()
    return ''.join(texts)


def extract_text_current_list(hwp):
    """현재 리스트(셀 등)의 텍스트만 추출"""
    texts = []

    # 현재 리스트만 검색: scanSposList(0x0050) | scanEposList(0x0005)
    hwp.InitScan(0x00, 0x0055)

    while True:
        state, text = hwp.GetText()
        if state <= 1 or state >= 100:
            break
        if state == 2 and text:
            texts.append(text)

    hwp.ReleaseScan()
    return ''.join(texts)


# 테스트
if __name__ == "__main__":
    hwp = win32.Dispatch('HWPFrame.HwpObject')
    hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')

    hwp.Open("C:\\test\\sample.hwp")

    full_text = extract_all_text_scan(hwp)
    print("문서 전체 텍스트:")
    print(full_text)
```

---

## 2. GetTextFile - 문서 전체 텍스트 추출

현재 열린 문서를 문자열로 반환합니다.

### API 설명

```
VARIANT GetTextFile(BSTR format, BSTR option)
```

**format 값:**
| format | 설명 |
|--------|------|
| HWP | HWP native format (BASE64 인코딩) |
| HWPML2X | HWP 형식과 호환 (모든 정보 유지) |
| HTML | 인터넷 문서 형식 |
| UNICODE | 유니코드 텍스트 (서식 없음) |
| TEXT | 일반 텍스트 (유니코드 정보 손실) |

**option 값:**
| option | 설명 |
|--------|------|
| saveblock | 선택된 블록만 저장 |

### 예제 코드

```python
import win32com.client as win32

def get_document_text(hwp):
    """문서 전체 텍스트 추출"""
    # UNICODE 형식으로 추출 (서식 없는 순수 텍스트)
    text = hwp.GetTextFile("UNICODE", "")
    return text


def get_selected_text(hwp):
    """선택된 블록의 텍스트만 추출"""
    # saveblock 옵션으로 선택 영역만 추출
    text = hwp.GetTextFile("TEXT", "saveblock")
    return text


def get_document_as_html(hwp):
    """문서를 HTML로 추출"""
    html = hwp.GetTextFile("HTML", "")
    return html


# 테스트
if __name__ == "__main__":
    hwp = win32.Dispatch('HWPFrame.HwpObject')
    hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')

    hwp.Open("C:\\test\\sample.hwp")

    # 전체 텍스트
    text = get_document_text(hwp)
    print("전체 텍스트:")
    print(text)

    # 일부 선택 후 추출
    hwp.MovePos(2)  # 문서 시작
    hwp.SelectText(0, 0, 0, 100)  # 첫 문단 일부 선택
    selected = get_selected_text(hwp)
    print("\n선택된 텍스트:")
    print(selected)
```

**주의사항:**
- 이 함수는 JScript/VBScript 등 디스크 접근이 어려운 언어를 위해 만들어짐
- 디스크 접근 가능한 언어에서는 Save/SaveBlockAction 사용 권장
- 내부적으로 파일 저장 후 메모리 복사를 수행하므로 느리고 메모리 낭비 있음

---

## 3. GetPageText - 페이지 단위 추출

특정 페이지의 텍스트를 추출합니다.

### API 설명

```
BSTR GetPageText(long pgno, VARIANT option)
```

**Parameters:**
- pgno: 텍스트를 추출할 페이지 번호 (0부터 시작)
- option: 추출 대상 옵션

**option 값:**
| ID | 값 | 설명 |
|----|---|------|
| maskNormal | 0x00 | 본문 텍스트 (항상 기본) |
| maskTable | 0x01 | 표 텍스트 추출 |
| maskTextbox | 0x02 | 글상자 텍스트 추출 |
| maskCaption | 0x04 | 캡션 텍스트 추출 |

**동작 방식:**
- 일반 텍스트(글자처럼 취급 도형 포함)를 우선 추출
- 이후 도형(표, 글상자) 내의 텍스트 추출

### 예제 코드

```python
import win32com.client as win32

def get_page_text_all(hwp, page_no):
    """특정 페이지의 모든 텍스트 추출"""
    # 모든 옵션 조합: 0xFFFFFFFF 또는 생략
    text = hwp.GetPageText(page_no, 0xFFFFFFFF)
    return text


def get_page_text_body_only(hwp, page_no):
    """특정 페이지의 본문 텍스트만 추출"""
    # 본문만: maskNormal = 0x00
    text = hwp.GetPageText(page_no, 0x00)
    return text


def get_page_text_with_tables(hwp, page_no):
    """특정 페이지의 본문 + 표 텍스트 추출"""
    # 본문 + 표: maskTable = 0x01
    text = hwp.GetPageText(page_no, 0x01)
    return text


def get_all_pages_text(hwp):
    """모든 페이지 텍스트 추출"""
    page_count = hwp.PageCount
    all_texts = []

    for page_no in range(page_count):
        text = hwp.GetPageText(page_no, 0xFFFFFFFF)
        all_texts.append(f"=== Page {page_no + 1} ===\n{text}")

    return '\n\n'.join(all_texts)


# 테스트
if __name__ == "__main__":
    hwp = win32.Dispatch('HWPFrame.HwpObject')
    hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')

    hwp.Open("C:\\test\\sample.hwp")

    # 첫 페이지 전체 텍스트
    text = get_page_text_all(hwp, 0)
    print("첫 페이지 텍스트:")
    print(text)

    # 모든 페이지
    all_text = get_all_pages_text(hwp)
    print("\n모든 페이지 텍스트:")
    print(all_text)
```

---

## 4. GetFieldText - 필드 텍스트 추출

누름틀 필드나 셀 필드에서 텍스트를 추출합니다.

### API 설명

```
BSTR GetFieldText(BSTR fieldlist)
```

**Parameters:**
- fieldlist: 텍스트를 구할 필드 이름의 리스트
  - 여러 필드는 문자 코드 0x02로 구분
  - 예: "필드이름#1\x02필드이름#2\x02...필드이름#n"

**필드 이름 표현:**
- 필드이름: 해당 이름의 첫 번째 필드
- 필드이름{{n}}: 해당 이름의 n번째 필드 (0부터 시작)

**반환값:**
- 탭: '\t' (0x9)
- 문단 바뀜: CR/LF (0x0D/0x0A)
- 필드 구분: 0x02

### 예제 코드

```python
import win32com.client as win32

def get_single_field_text(hwp, field_name):
    """단일 필드 텍스트 추출"""
    text = hwp.GetFieldText(field_name)
    return text


def get_multiple_field_texts(hwp, field_names):
    """여러 필드 텍스트 추출"""
    # 필드 이름들을 0x02로 구분하여 연결
    fieldlist = '\x02'.join(field_names)
    result = hwp.GetFieldText(fieldlist)

    # 결과를 분리하여 반환
    texts = result.split('\x02')
    return dict(zip(field_names, texts))


def get_all_field_list(hwp):
    """문서 내 모든 필드 목록 조회"""
    # hwpFieldNumber=1: 필드 이름 뒤에 일련번호 추가
    field_list = hwp.GetFieldList(1, 0)
    if field_list:
        fields = field_list.split('\x02')
        return fields
    return []


def get_cell_field_list(hwp):
    """셀 필드 목록만 조회"""
    # option=1: hwpFieldCell (셀 필드만)
    field_list = hwp.GetFieldList(1, 1)
    if field_list:
        return field_list.split('\x02')
    return []


def get_all_fields_with_text(hwp):
    """모든 필드와 텍스트를 딕셔너리로 반환"""
    fields = get_all_field_list(hwp)
    if not fields:
        return {}

    result = {}
    for field in fields:
        text = hwp.GetFieldText(field)
        result[field] = text
    return result


# 테스트
if __name__ == "__main__":
    hwp = win32.Dispatch('HWPFrame.HwpObject')
    hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')

    hwp.Open("C:\\test\\form.hwp")

    # 모든 필드 목록
    fields = get_all_field_list(hwp)
    print("필드 목록:", fields)

    # 모든 필드 텍스트
    field_texts = get_all_fields_with_text(hwp)
    for name, text in field_texts.items():
        print(f"  {name}: {text}")
```

---

## 5. 셀 단위 텍스트 추출

표의 개별 셀에서 텍스트를 추출하는 방법입니다.

### 방법 1: SetPos + SelectAll + GetTextFile

```python
def get_cell_text(hwp, list_id):
    """특정 셀의 텍스트 추출"""
    try:
        # 셀로 이동
        hwp.SetPos(list_id, 0, 0)

        # 셀 전체 선택
        hwp.HAction.Run("SelectAll")

        # 선택된 텍스트 가져오기
        text = hwp.GetTextFile("TEXT", "saveblock")

        # 선택 해제
        hwp.HAction.Run("Cancel")

        # 텍스트 정리
        if text:
            text = text.strip().replace('\r\n', ' ').replace('\r', ' ').replace('\n', ' ')
            return text
        return ""
    except:
        return ""
```

### 방법 2: InitScan으로 현재 리스트 텍스트 추출

```python
def get_cell_text_scan(hwp, list_id):
    """InitScan으로 셀 텍스트 추출"""
    try:
        hwp.SetPos(list_id, 0, 0)
        texts = []

        # scanSposList(0x0050) | scanEposList(0x0005) = 현재 리스트만
        hwp.InitScan(0x00, 0x0055)

        while True:
            state, text = hwp.GetText()
            if state <= 1 or state >= 100:
                break
            if state == 2 and text:
                texts.append(text)

        hwp.ReleaseScan()
        return ''.join(texts).strip()
    except:
        return ""
```

### 방법 3: 셀 필드 이름으로 추출

셀에 필드 이름이 지정된 경우 GetFieldText로 추출할 수 있습니다.

```python
def get_cell_text_by_field(hwp, cell_address):
    """셀 주소(A1, B2 등)로 필드 텍스트 추출"""
    # 셀 이름이 필드로 설정된 경우
    text = hwp.GetFieldText(cell_address)
    return text


def get_cell_field_name(hwp):
    """현재 위치의 셀 필드 이름 조회"""
    # option=1: hwpFieldCell (셀 필드)
    name = hwp.GetCurFieldName(1)
    return name
```

### 표 전체 셀 순회 예제

```python
def extract_all_table_cells(hwp):
    """표 내 모든 셀 텍스트 추출"""
    # 표 안에 있는지 확인
    parent = hwp.ParentCtrl
    if not parent or parent.CtrlID != 'tbl':
        print("표 안에 있지 않습니다")
        return []

    results = []

    # 첫 셀로 이동
    hwp.MovePos(104)  # moveStartOfCell
    first_list_id = hwp.GetPos()[0]

    visited = set()
    queue = [first_list_id]
    visited.add(first_list_id)

    while queue:
        list_id = queue.pop(0)

        # 셀 텍스트 추출
        text = get_cell_text(hwp, list_id)
        results.append({
            'list_id': list_id,
            'text': text
        })

        # 인접 셀 탐색
        for move_id in [101, 103]:  # moveRightOfCell, moveDownOfCell
            hwp.SetPos(list_id, 0, 0)
            hwp.MovePos(move_id, 0, 0)
            next_id = hwp.GetPos()[0]

            if next_id != list_id and next_id not in visited:
                visited.add(next_id)
                queue.append(next_id)

    return results


# 테스트
if __name__ == "__main__":
    hwp = win32.Dispatch('HWPFrame.HwpObject')
    hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')

    hwp.Open("C:\\test\\table.hwp")

    # 첫 표로 이동 후 셀 추출
    hwp.MovePos(2)  # 문서 시작
    ctrl = hwp.HeadCtrl
    while ctrl:
        if ctrl.CtrlID == 'tbl':
            # 표 선택 후 내부로 진입
            hwp.SetPosBySet(ctrl.GetAnchorPos(0))
            hwp.MovePos(12)  # moveNextPos

            cells = extract_all_table_cells(hwp)
            for cell in cells:
                print(f"셀 {cell['list_id']}: {cell['text']}")
            break
        ctrl = ctrl.Next
```

---

## 6. SelectText + GetTextFile 조합

특정 범위를 선택하고 텍스트를 추출하는 방법입니다.

### API 설명 - SelectText

```
BOOL SelectText(long spara, long spos, long epara, long epos)
```

**Parameters:**
- spara: 블록 시작 위치의 문단 번호
- spos: 블록 시작 위치의 문단 내 문자 위치
- epara: 블록 끝 위치의 문단 번호
- epos: 블록 끝 위치의 문단 내 문자 위치

**주의:** epos가 가리키는 문자는 포함되지 않음

### 예제 코드

```python
def get_text_range(hwp, spara, spos, epara, epos):
    """특정 범위의 텍스트 추출"""
    # 범위 선택
    hwp.SelectText(spara, spos, epara, epos)

    # 선택된 텍스트 추출
    text = hwp.GetTextFile("TEXT", "saveblock")

    # 선택 해제
    hwp.HAction.Run("Cancel")

    return text


def get_paragraph_text(hwp, para_no):
    """특정 문단의 텍스트 추출"""
    # 문단 전체 선택 (0부터 끝까지)
    hwp.SelectText(para_no, 0, para_no + 1, 0)
    text = hwp.GetTextFile("TEXT", "saveblock")
    hwp.HAction.Run("Cancel")
    return text.strip()


def get_text_from_current_to_end(hwp):
    """현재 위치부터 문서 끝까지 텍스트 추출"""
    # 현재 위치 저장
    pos = hwp.GetPos()

    # 문서 끝까지 선택
    hwp.HAction.Run("MoveSelDocEnd")
    text = hwp.GetTextFile("TEXT", "saveblock")
    hwp.HAction.Run("Cancel")

    # 원위치 복귀
    hwp.SetPos(pos[0], pos[1], pos[2])

    return text


# 테스트
if __name__ == "__main__":
    hwp = win32.Dispatch('HWPFrame.HwpObject')
    hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')

    hwp.Open("C:\\test\\sample.hwp")

    # 첫 문단 텍스트
    text = get_paragraph_text(hwp, 0)
    print("첫 문단:", text)

    # 0~50자 범위
    text = get_text_range(hwp, 0, 0, 0, 50)
    print("0~50자:", text)
```

---

## 7. 실제 사용 예제

### 전체 예제: 다양한 추출 방법 비교

```python
# -*- coding: utf-8 -*-
"""HWP 텍스트 추출 종합 예제

WSL에서 실행 시:
  cmd.exe /c "cd /d C:\\win32hwp && python text_extract_example.py"
"""

import win32com.client as win32


def init_hwp():
    """HWP 인스턴스 생성 및 보안 모듈 등록"""
    hwp = win32.Dispatch('HWPFrame.HwpObject')
    hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModuleExample')
    return hwp


def example_1_scan_method(hwp, file_path):
    """예제 1: InitScan/GetText 방식"""
    print("=== 방법 1: InitScan/GetText ===")

    hwp.Open(file_path)
    texts = []

    # 전체 문서 검색 (모든 컨트롤 포함)
    hwp.InitScan(0x07, 0x0077)

    while True:
        state, text = hwp.GetText()
        if state <= 1 or state >= 100:
            break
        if state == 2 and text:
            texts.append(text)
        elif state == 3:
            texts.append('\n')

    hwp.ReleaseScan()

    result = ''.join(texts)
    print(f"추출 결과 ({len(result)}자):\n{result[:200]}...")
    return result


def example_2_textfile_method(hwp, file_path):
    """예제 2: GetTextFile 방식"""
    print("\n=== 방법 2: GetTextFile ===")

    hwp.Open(file_path)

    # TEXT 형식으로 추출
    text = hwp.GetTextFile("TEXT", "")

    print(f"추출 결과 ({len(text)}자):\n{text[:200]}...")
    return text


def example_3_page_method(hwp, file_path):
    """예제 3: GetPageText 방식"""
    print("\n=== 방법 3: GetPageText ===")

    hwp.Open(file_path)

    page_count = hwp.PageCount
    all_text = []

    for page in range(page_count):
        # 모든 요소 포함
        text = hwp.GetPageText(page, 0xFFFFFFFF)
        all_text.append(f"[Page {page + 1}]\n{text}")

    result = '\n'.join(all_text)
    print(f"페이지 수: {page_count}")
    print(f"추출 결과 ({len(result)}자):\n{result[:200]}...")
    return result


def example_4_field_method(hwp, file_path):
    """예제 4: GetFieldText 방식 (양식 문서용)"""
    print("\n=== 방법 4: GetFieldText ===")

    hwp.Open(file_path)

    # 모든 필드 목록 조회
    field_list = hwp.GetFieldList(1, 0)  # number=1: 일련번호 포함

    if not field_list:
        print("필드가 없습니다.")
        return {}

    fields = field_list.split('\x02')
    print(f"발견된 필드 ({len(fields)}개): {fields[:5]}...")

    result = {}
    for field in fields[:10]:  # 처음 10개만
        text = hwp.GetFieldText(field)
        result[field] = text
        print(f"  {field}: {text[:50]}..." if len(text) > 50 else f"  {field}: {text}")

    return result


def example_5_table_cell_method(hwp, file_path):
    """예제 5: 표 셀 텍스트 추출"""
    print("\n=== 방법 5: 표 셀 텍스트 ===")

    hwp.Open(file_path)

    # 첫 번째 표 찾기
    ctrl = hwp.HeadCtrl
    table_ctrl = None
    while ctrl:
        if ctrl.CtrlID == 'tbl':
            table_ctrl = ctrl
            break
        ctrl = ctrl.Next

    if not table_ctrl:
        print("표가 없습니다.")
        return []

    print("표를 찾았습니다.")

    # 표로 이동
    anchor = table_ctrl.GetAnchorPos(0)
    hwp.SetPosBySet(anchor)
    hwp.MovePos(12)  # 표 내부로 진입

    # 첫 셀로 이동
    hwp.MovePos(104)  # moveStartOfCell
    first_list_id = hwp.GetPos()[0]

    # BFS로 모든 셀 순회
    cells = []
    visited = set()
    queue = [first_list_id]
    visited.add(first_list_id)

    while queue and len(cells) < 100:
        list_id = queue.pop(0)

        # 셀 텍스트 추출
        hwp.SetPos(list_id, 0, 0)
        hwp.HAction.Run("SelectAll")
        text = hwp.GetTextFile("TEXT", "saveblock")
        hwp.HAction.Run("Cancel")

        if text:
            text = text.strip()
        cells.append({'list_id': list_id, 'text': text or ''})

        # 인접 셀 탐색
        for move_id in [101, 103]:
            hwp.SetPos(list_id, 0, 0)
            hwp.MovePos(move_id, 0, 0)
            next_id = hwp.GetPos()[0]
            if next_id != list_id and next_id not in visited:
                visited.add(next_id)
                queue.append(next_id)

    print(f"셀 수: {len(cells)}")
    for i, cell in enumerate(cells[:10]):
        print(f"  셀 {i}: {cell['text'][:30]}..." if len(cell['text']) > 30 else f"  셀 {i}: {cell['text']}")

    return cells


def main():
    """메인 함수"""
    hwp = init_hwp()

    # 테스트 파일 경로
    test_file = "C:\\test\\sample.hwp"

    try:
        example_1_scan_method(hwp, test_file)
        example_2_textfile_method(hwp, test_file)
        example_3_page_method(hwp, test_file)
        example_4_field_method(hwp, test_file)
        example_5_table_cell_method(hwp, test_file)
    except Exception as e:
        print(f"오류 발생: {e}")
    finally:
        hwp.Clear(1)  # 문서 닫기 (hwpDiscard)


if __name__ == "__main__":
    main()
```

---

## 참고: MovePos 셀 이동 관련 상수

| ID | 값 | 설명 |
|----|---|------|
| moveLeftOfCell | 100 | 셀 왼쪽으로 이동 |
| moveRightOfCell | 101 | 셀 오른쪽으로 이동 |
| moveUpOfCell | 102 | 셀 위쪽으로 이동 |
| moveDownOfCell | 103 | 셀 아래쪽으로 이동 |
| moveStartOfCell | 104 | 행의 시작 셀로 이동 |
| moveEndOfCell | 105 | 행의 끝 셀로 이동 |
| moveTopOfCell | 106 | 열의 시작 셀로 이동 |
| moveBottomOfCell | 107 | 열의 끝 셀로 이동 |

---

## 요약

| 방법 | 용도 | 장점 | 단점 |
|------|-----|------|------|
| InitScan/GetText | 문서 순차 스캔 | 세밀한 제어, 위치 정보 | 복잡함 |
| GetTextFile | 전체/블록 추출 | 간단함 | 느림, 메모리 낭비 |
| GetPageText | 페이지별 추출 | 페이지 단위 작업 | 옵션 제한적 |
| GetFieldText | 필드 추출 | 양식 문서에 최적 | 필드 설정 필요 |
| 셀 순회 + GetTextFile | 표 셀 추출 | 표 데이터 추출 | 복잡함 |

---

## 참고 문서

- `/mnt/d/hwp_docs/win32/HwpAutomation_2504_part03.md` - InitScan, GetText, GetTextFile
- `/mnt/d/hwp_docs/win32/HwpAutomation_2504_part04.md` - GetPageText
- `/mnt/d/hwp_docs/win32/HwpAutomation_2504_part02.md` - GetFieldText, PutFieldText
