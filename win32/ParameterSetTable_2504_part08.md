# ParameterSetTable_2504_part08

---

## 62) GotoE : 찾아가기

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| SetSelectionIndex | PIT_UI1 | | 현재 선택되어 있는 라디오 값 |
| DialogResult | PIT_UI | | 대화상자의 반환값 |

---

## 63) GridInfo : 격자 정보

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Method | PIT_UI1 | | 격자 방식<br>0 = 격자와 상관없이<br>1 = 격자 자석효과<br>2 = 격자에만 붙음 |
| Align | PIT_UI1 | | 격자 기준(쪽 = 0 / 종이 = 1) |
| HorzAlign | PIT_I | | 격자 기준 가로 offset (단위 HWPUNIT) |
| VertAlign | PIT_I | | 격자 기준 세로 offset (단위 HWPUNIT) |
| Type | PIT_UI1 | | 격자 모양<br>0 = 점 격자<br>1 = 선 격자 |
| HorzSpan | PIT_UI | | 가로 간격 (단위 HWPUNIT) |
| VertSpan | PIT_UI | | 세로 간격 (단위 HWPUNIT) |
| HorzRange | PIT_U | | 가로 자석 범위 (단위 HWPUNIT) |
| VertRange | PIT_U | | 세로 자석 범위 (단위 HWPUNIT) |
| Show | PIT_UI1 | | 격자 보이기 ( on / off ) |
| ZOrder | PIT_UI1 | | 격자 위치(글 위/글 아래) (ZOrder)<br>0 = 글 아래, 1 = 글 위 |
| ViewLine | PIT_UI1 | | 선격자 보이기 종류<br>0 = 모두<br>1 = 수평격자만<br>2 = 수직격자만 |

---

## 64) HeaderFooter : 머리말/꼬리말

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| HeaderFooterCtrlType | PIT_UI | | 머리말/꼬리말 종류 : 0 = 머리말, 1 = 꼬리말 |
| HeaderFooterStyle | PIT_UI | | 머리말/꼬리말 마당 스타일 |
| Type | PIT_UI1 | | 머리말/꼬리말 위치 : 0 = 양쪽, 1 = 짝수쪽, 2 = 홀수쪽 |

---

## 65) HyperLink : 하이퍼링크 삽입 / 고치기

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Text | PIT_BSTR | | 하이퍼링크가 표시되는 문자열 |
| Command | PIT_BSTR | | Command String 참조 |
| NoLInk | PIT_I | | "연결 안함" 여부 |
| ShapeObject | PIT_I | | 그림 및 그리기객체가 Selection되어 있는지 여부 |
| DirectInsert | PIT_I | | 현재 캐럿 위치에 무조건 하이퍼링크 삽입 여부 (블록지정 상태면 블록해제 후 삽입) |

---

## 66) HyperlinkJump : 하이퍼링크 이동

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Source | PIT_BSTR | | Source에 대한 Object Path |
| Target | PIT_BSTR | | 이동할 Target에 대한 Object Path |

---

## 67) Idiom : 상용구

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| InputText | PIT_BSTR | | 삽입될 스트링/끼워 넣을 파일 |
| InputType | PIT_UI1 | | 입력기 상용구/한글 상용구 |

---

## 68) IndexMark : 찾아보기 표식

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| First | PIT_BSTR | | 첫 번째 키 |
| Second | PIT_BSTR | | 두 번째 키 |

---

## 69) InputDateStyle : 날짜/시간 표시 형식

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| DateStyleType | PIT_UI1 | | 문자열로 넣기/코드로 넣기 |
| DateStyleDataForm | PIT_BSTR | | 필드 컨트롤의 안내문/지시문 |

---

## 70) InsertFieldTemplate : 문서마당 정보

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| ShowSingle | PIT_I | | 문서마당 정보 대화상자에서 페이지(탭) 보이기 옵션 :<br>-1 = 모든 페이지 보이기<br>0 = 누름틀 페이지만 보이기<br>1 = 개인 정보 페이지만 보이기<br>2 = 문서 요약 페이지만 보이기<br>3 = 만든 날짜 페이지만 보이기<br>4 = 파일 경로 페이지만 보이기 |
| TemplateDirection | PIT_BSTR | | 필드 컨트롤의 안내문/지시문 |
| TemplateHelp | PIT_BSTR | | 필드 컨트롤의 도움말 |
| TemplateName | PIT_BSTR | | 필드 이름 (name) |
| TemplateType | PIT_UI1 | | 필드의 종류.<br>TemplateDirection/Help/Name의 값이 실제로 적용되는 위치 :<br>0 = 누름틀, 1 = 개인 정보, 2 = 문서 요약, 3 = 만든 날짜, 4 = 파일 경로 |
| Editable | PIT_UI1 | | 필드의 양식모드에서 편집여부<br>0 = 편집 불가능<br>1 = 편집 가능 |

※ 필드(Field)는 꺽쇠(『』)로 둘러싸인 정보를 말하며, 문서마당의 모든 정보는 필드로 구성된다.

※ ShowSingle 아이템은 OnPopupDialog() 수행 시, 원하는 페이지만 대화상자에 표시하고 싶을 때 사용한다.

---

## 71) InsertFile : 파일 삽입

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| FileName | PIT_BSTR | | 삽입할 파일의 이름 |
| FileFormat | PIT_BSTR | | 삽입할 파일의 확장자 |
| FileArg | PIT_BSTR | | 삽입할 파일의 Argument |
| KeepSection | PIT_UI1 | | 끼워 넣을 문서를 구역으로 나누어 쪽 모양을 유지할지 여부 on / off (ver:0x0605010E) |
| KeepCharshape | PIT_UI1 | | 끼워 넣을 문서의 글자 모양을 유지할지 여부 on / off |
| KeepParashape | PIT_UI1 | | 끼워 넣을 문서의 문단 모양을 유지할지 여부 on / onff |
| KeepStyle | PIT_UI1 | | 끼워 넣을 문서의 스타일을 유지할지 여부 on / onff |
| MoveNextPos | PIT_UI1 | | 삽입하고 삽입된 파일 다음 위치로 이동할지 여부 |

---
