# ParameterSetTable_2504_part17

## 132) PronounceInfo : 한자/일어 발음 표시

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Show | PIT_UI1 | | 한자/일어 발음 표시 보이기 |
| Position | PIT_UI1 | | 표시 위치 |
| Hanja | PIT_UI1 | | 한자 표시 |
| Japanese | PIT_UI1 | | 일어 표시 |
| FontName | PIT_BSTR | | 글꼴 |
| TextSize | PIT_UI1 | | 글자 크기 |
| TextColor | PIT_UI4 | | 글자 색 |
| Heterography | PIT_UI1 | | 동자이음 문자 표시 여부 |

---

## 133) StyleItem : 스타일 - 바로 편집하기 대화상자

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Type | PIT_UI1 | | 스타일 종류 |
| NameLocal | PIT_BSTR | | 스타일 이름(로컬) |
| NameEng | PIT_BSTR | | 스타일 이름(영문) |
| Next | PIT_I | | 다음 스타일 인덱스 |
| LockForm | PIT_UI1 | | 양식스타일 보호 |
| CharShape | PIT_SET | CharShape | 글자 모양 |
| ParaShape | PIT_SET | ParaShape | 문단 모양 |

---

## 134) InsertFieldTemplate : 상호 참조 넣기

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| ShowSingle | PIT_UI | | 문서 마당 정보 PropertySheet 대화상자에서 원하는 페이지(탭)만 보이기 |
| TemplateDirection | PIT_BSTR | | 필드 컨트롤의 안내문/지시문 |
| TemplateHelp | PIT_BSTR | | 필드 컨트롤의 도움말 |
| TemplateName | PIT_BSTR | | 필드 이름 (name) |
| TemplateType | PIT_UI1 | | 필드의 종류<br>0 = 누름틀<br>1 = 개인 정보<br>2 = 문서 요약<br>3 = 만든 날짜<br>4 = 파일 이름/경로 |
| Editable | PIT_UI1 | | 부분편집 모드에서 편집속성 |

---

## 135) SaveAsImage : 바이너리 그림을 다른 형태로 저장하는 옵션을 설정

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| ResizeImage | PIT_UI1 | | 문서에 삽입된 그림 오브젝트의 크기를 고려하여 최소 크기로 저장 여부. |
| DelCutting | PIT_UI1 | | 잘라내기 이미지 불필요한 부분 삭제 여부 |
| SaveAsFormat | PIT_I | | 다른이름으로 저장한 포맷 |
| SaveDpiX | PIT_UI | | 변경될 이미지 X축 DPI |
| SaveDpiY | PIT_UI | | 변경될 이미지 Y축 DPI |
| SaveType | PIT_UI1 | | 저장 액션 타입(None = 0x00, insert = 0x01, SaveTime = 0x03) |
| MinWidth | PIT_UI | | 최소 너비 |
| MinHeight | PIT_UI | | 최소 높이 |
| ExtendFormat | PIT_UI1 | | 지원 포맷 확장, (한/글 2024부터 지원) |
| Resolution | PIT_UI | | 제한 해상도 설정(예 : 3840 * 2160), (한/글 2024부터 지원) |
| AllFormat | PIT_UI1 | | 모든 이미지 포맷 변환, (한/글 2024부터 지원) |

---

## 136) BrailleConvert : 점자 변환

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| ResultType | PIT_UI1 | | 결과 문자코드 형식 |
| CharHeight | PIT_I4 | | 글자 모양-크기 |
| FontName | PIT_BSTR | | 글자 모양-글꼴 |
| FontType | PIT_UI1 | | 글자 모양-글꼴타입 |
| LineCharApply | PIT_UI1 | | 줄/글자 수 - 적용 여부 |
| LineCharType | PIT_UI1 | | 줄/글자 수 - 종류 |
| LineSpacing | PIT_I4 | | 줄/글자 수 - 줄 간격 |
| CharSpacing | PIT_I4 | | 줄/글자 수 - 글자 간격 |
| PaperApply | PIT_UI1 | | 용지 - 적용 여부 |
| PaperType | PIT_I4 | | 용지 - 설정 용지 타입 |
| TargetView | PIT_UI1 | | 결과 새창/새탭 0:새창으로, 1:새탭으로 |

---

## 137) PictureChange : 그림 바꾸기

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| PicturePath | PIT_BSTR | | 그림 경로 |
| PictureEmbed | PIT_UI1 | | Embed 여부 |

---

## 138) PresentationRange : 문서 전체 프레젠테이션 설정

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| PresentationDefault | PIT_SET | Presentation | 기본 설정값 |
| ExistPresentation | PIT_UI1 | | section에 프레젠테이션 설정 유무. 1 = 설정, 0 = 비설정. |

---

## 139) DeletePage : 쪽 지우기 <span style="color:blue">글2007</span>

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Range | PIT_UI1 | | 범위 |
| RangeCustom | PIT_BSTR | | 사용자가 직접 입력한 범위 |
| UsingPagenum | PIT_UI1 | | 문서의 쪽 번호로 입력 |

---

## 140) TrackChange : 변경 추적

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| InsertShape | PIT_UI1 | | 삽입 모양 |
| InsertColor | PIT_UI1 | | 삽입 색 |
| DeleteShape | PIT_UI1 | | 삭제 모양 |
| DeleteColor | PIT_UI1 | | 삭제 색 |
| ChangeShape | PIT_UI1 | | 변경 모양 |
| ChangeColor | PIT_UI1 | | 변경 색 |
| Format | PIT_UI1 | | 서식 추적 여부 |
| FormatShape | PIT_UI1 | | 서식 추적 모양 |
| FormatColor | PIT_UI1 | | 서식 추적 색 |
| MemoWidth | PIT_I4 | | 메모 너비 |
| MemoLine | PIT_UI1 | | 메모 선 표시 |
| MemoColor | PIT_UI1 | | 메모 색 |
| Tooltip | PIT_UI1 | | 툴팁 표시 |

---

## 141) PrivateInfoSecurity : 개인 정보 보안

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Command | PIT_BSTR | | command string |
| Password | PIT_BSTR | | 개인 정보 암호 |
| ChangePassword | PIT_BSTR | | 바꿀 개인 정보 암호 |
| Pattern | PIT_BSTR | | 개인 정보 패턴 |
| Telephone | PIT_UI1 | | 전화번호 (찾을 개인 정보 on/off) |
| Resident | PIT_UI1 | | 주민등록번호 (찾을 개인 정보 on/off) |
| Email | PIT_UI1 | | 전자우편 (찾을 개인 정보 on/off) |
| Account | PIT_UI1 | | 계좌 번호 (찾을 개인 정보 on/off) |
| Credit | PIT_UI1 | | 신용카드 번호 (찾을 개인 정보 on/off) |
| Address | PIT_UI1 | | 주소 (찾을 개인 정보 on/off) |
| Birthday | PIT_UI1 | | 생년월일 (찾을 개인 정보 on/off) |
| IPAddress | PIT_UI1 | | 인터넷주소 (찾을 개인 정보 on/off) |
| ForeignerNo | PIT_UI1 | | 외국인등록번호 (찾을 개인 정보 on/off) |
| UserDef | PIT_UI1 | | 사용자 정의 (찾을 개인 정보 on/off) |
| Etc | PIT_UI1 | | 기타 (찾을 개인 정보 on/off) |
| NoMessageBox | PIT_UI1 | | 메세지 박스를 띄우지 않을지 여부 |
| DelHyperlink | PIT_UI1 | | 하이퍼링크(이메일 연결)를 지울지 여부 |
| MarkChar | PIT_UI | | 개인 정보 표시 문자 |
| MarkCharType | PIT_UI | | 표시 문자 선택 사항 |
| PasswordOnOff | PIT_UI1 | | 개인 정보 보안 암호 설정/취소 |
| InfoType | PIT_UI | | 개인 정보 Type(0:정규표현식-일반, 1:일반문자열, 2:정규표현식-주민번호, 3:주소검색, 4:외국인번호, 5이상: 나머지 개인정보의 정규표현식) |
| License | PIT_UI1 | | 운전면허 (찾을 개인 정보 on/off) |
| Passport | PIT_UI1 | | 여권번호 (찾을 개인 정보 on/off) |
| Encrypted | PIT_UI1 | | 문서에 암호화된 개인 정보 필드가 있을 때 여부 |
| UnregistPattern | PIT_UI1 | | |
