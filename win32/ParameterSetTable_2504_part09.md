# Parameter Set Table - Part 09

---

## 72) InsertText : 텍스트 삽입

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Text | PIT_BSTR | | 삽입할 텍스트 |

---

## 73) KeyMacro : 키매크로

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Index | PIT_I | | 정의(or 실행)할 매크로의 인덱스. |
| RepeatCount | PIT_I | | 실행 반복 횟수 |
| Name | PIT_BSTR | | 매크로 이름 |

---

## 74) Label : 라벨

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| TopMargin | PIT_I4 | | 용지 여백 : 위쪽 |
| LeftMargin | PIT_I4 | | 용지 여백 : 왼쪽 |
| BoxWidth | PIT_I4 | | 이름표 크기 : 폭 |
| BoxLength | PIT_I4 | | 이름표 크기 : 길이 |
| MarginHor | PIT_I4 | | 이름표 크기 : 좌우 |
| BoxMarginVer | PIT_I4 | | 이름표 크기 : 상하 |
| LabelCols | PIT_I4 | | 이름표 개수 : 줄 수 (or 세로) |
| LabelRows | PIT_I4 | | 이름표 개수 : 칸 수 (or 가로) |
| LandScape | PIT_I4 | | 용지 방향 |
| PageWidth | PIT_I4 | | 문서의 폭 |
| PageLen | PIT_I4 | | 문서의 길이 |
| TextGap | PIT_I4 | | |
| PageCnt | PIT_I4 | | |

---

## 75) LinkDocument : 문서 연결

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Name | PIT_BSTR | | 연결 문서 FILE NAME |
| PageInherit | PIT_UI1 | | TRUE = 쪽 번호 잇기, FALSE = 쪽 번호 잇지 않기. |
| FootnoteInherit | PIT_UI1 | | TRUE = 각주 번호 잇기, FALSE = 각주 번호 잇지 않기. |

---

## 76) ListParaPos : 커서의 위치

HwpCtrl.GetPosBySet. SetPosBySet, HwpCtrlCode.GetAnchorPos에서 사용, 해당 액션은 존재하지 않음.

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| List | PIT_UI | | 현재 위치한 리스트 |
| Para | PIT_UI | | 현재 위치한 문단 |
| Pos | PIT_UI | | 현재 위치한 글자 |

---

## 77) ListProperties : 서브 리스트의 속성

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| TextDirection | PIT_UI1 | | 글자 방향. (세부 스펙은 미정) |
| LineWrap | PIT_UI1 | | 경계에서 줄 나눔 방식<br>0 = 일반적인 줄 바꿈<br>1 = 줄을 바꾸지 않음<br>2 = 자간을 조정하여 한 줄을 유지 |
| VertAlign | PIT_UI1 | | 세로 정렬<br>0 = 위로 정렬<br>1 = 가운데 정렬<br>2 = 아래로 정렬 |
| MarginLeft | PIT_I4 | | 각 방향 여백 : 왼쪽 |
| MarginRight | PIT_I4 | | 각 방향 여백 : 오른쪽 |
| MarginTop | PIT_I4 | | 각 방향 여백 : 위 |
| MarginBottom | PIT_I4 | | 각 방향 여백 : 아래 |

---

## 78) MailMergeGenerate : 메일 머지 만들기

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Input | PIT_UI1 | | 자료 종류<br>0 = WAB<br>1 = OAB<br>2 = HWP<br>3 = DBF |
| HwpPath | PIT_BSTR | | Hwp 문서 경로. |
| HwpId | PIT_UI | | Hwp 문서 ID |
| DbfPath | PIT_BSTR | | dbf file path |
| DbfCode | PIT_UI1 | | dbf file codepage<br>0 = KS<br>1 = KSSM<br>2 = GB<br>3 = BIG5<br>4 = SJIS |
| Output | PIT_UI1 | | 출력 방향<br>0 = PRINTER<br>1 = PREVIEW<br>2 = FILE<br>3 = MAIL |
| FileName | PIT_BSTR | | 파일 이름 |
| Continue | PIT_UI1 | | 쪽번호 잇기 |
| PrintSet | PIT_SET | Print | 인쇄 선택 사항 |
| Subject | PIT_BSTR | | 메일 제목 |
| Type | PIT_UI1 | | 메일 종류<br>0 = 본문<br>1 = 첨부파일 |
| Field | PIT_BSTR | | 메일 주소 필드 |
| FieldUpdate | PIT_UI1 | | 필드 단위 업데이트 |
| NxlPath | PIT_BSTR | | 넥셀 파일 경로 |

---

## 79) MakeContents : 차례 만들기

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Make | PIT_UI | | 생성할 차례의 종류, 다음의 값들의 조합이다<br>0x01: 제목 차례<br>0x02: 표 차례<br>0x04: 그림 차례<br>0x08: 수식 차례<br>제목 차례를 지정한 경우에는 다음의 값을 추가로 지정할 수 있다.<br>0x10: 개요 문단으로 모으기<br>0x20: 스타일로 모으기<br>0x40: 차례코드로 모으기 |
| Level | PIT_I | | 개요 수준 |
| AutoTabRight | PIT_I1 | | 문단 오른쪽 끝 자동 탭 여부 : 0 = 자동 탭 사용안함, 1 = 자동 탭 사용 |
| Leader | PIT_UI | | 오른쪽 끝 탭 채울 모양(선 종류) |
| Styles | PIT_ARRAY | PIT_UI | 모을 스타일 목록 |
| StyleName | PIT_BSTR | | 모을 스타일 이름들 |
| OutFileName | PIT_BSTR | | 만들 파일 이름. ""이면 현재 문서에 생성 |
| Position | PIT_BSTR | | 만들 위치. 반드시 0이어야 한다. (글 컨트롤은 탭이 없으므로)<br>0 = 현재 문서<br>1 = 새 탭으로 |
| Type | PIT_UI | | 만들 차례 형식<br>0 : 필드로 넣기<br>1 : 문자열로 넣기 |
| Hyperlink | PIT_UI | | 하이퍼링크 연결 |

---

## 80) MarkpenShape : 형광펜 모양

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Color | PIT_UI4 | | 형광펜색 (COLORREF) |

---

## 81) MasterPage : 바탕쪽

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Type | PIT_I4 | | 바탕쪽 종류<br>0 = 양쪽<br>1 = 짝수쪽<br>2 = 홀수쪽 |
| Duplicate | PIT_UI1 | | 기존 바탕쪽과 겹침 (On/Off) |
| Front | PIT_UI1 | | 바탕쪽과 앞으로 보내기 (On/Off) |
| ApplyTo | PIT_UI1 | | 적용대상<br>0 = 현재구역<br>1 = 문서 전체 |
| CopySectionNumber | PIT_UI4 | | 바탕쪽 가져오기의 구역 번호 |
| CopyMasterPageTypes | PIT_ARRAY | PIT_UI4 | 바탕쪽 가져우기의 바탕쪽 종류들<br>0 = 양 쪽<br>1 = 짝수 쪽<br>2 = 홀수 쪽<br>3 = 구역 마지막쪽 |

---

*Page 81-90*
