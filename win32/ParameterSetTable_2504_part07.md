# ParameterSetTable_2504_part07

---

## 52) FileOpen : 파일 오픈

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| OpenFileName | PIT_BSTR | | 파일 이름. (OpenFileName과 SaveFileName은 같은 아이템을 공유한다. 즉 OpenFileName과 SaveFileName은 이름만 다를 뿐 동일한 아이템이다) |
| SaveFileName | PIT_BSTR | | 파일 이름. (OpenFileName과 SaveFileName은 같은 아이템을 공유한다. 즉 OpenFileName과 SaveFileName은 이름만 다를 뿐 동일한 아이템이다) |
| OpenFormat | PIT_BSTR | | 파일 형식. (OpenFileName과 마찬가지로 동일 아이템 지칭) |
| SaveFormat | PIT_BSTR | | 파일 형식. (OpenFileName과 마찬가지로 동일 아이템 지칭) |
| OpenReadOnly | PIT_UI1 | | 읽기 전용 |
| OpenFlag | PIT_UI1 | | 옵션<br>0x00 = 새 창으로 열기<br>0x01 = 현재 창의 새 탭에 열기<br>0x02 = 현재 창의 현재 탭에 열기<br>0x03 = 위 세 값의 mask<br>0x04 =이미 열려진 문서일 때 다시 load할지 뭍을 것인지<br>0x08 = 초기 모드를 최근 작업 문서 상태로<br>0x10 = 문서 마당 |
| SaveOverWrite | PIT_UI1 | | 덮어 쓰기 |
| ModifiedFlag | PIT_UI1 | | Modify 플래그 |
| Argument | PIT_BSTR | | Argument |
| SaveCMFDoc30 | PIT_UI1 | | 97 호환 저장 |
| SaveHwp97 | PIT_UI1 | | 97 문서로 저장 |
| SaveDistribute | PIT_UI1 | | 배포용 문서로 저장 |
| SaveDRMDoc | PIT_UI1 | | 보안 문서로 저장 |

---

## 53) FileSaveAs : 파일 저장

FileOpen과 멤버가 동일

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| OpenFileName | PIT_BSTR | | 파일 이름. (OpenFileName과 SaveFileName은 같은 아이템을 공유한다. 즉 OpenFileName과 SaveFileName은 이름만 다를 뿐 동일한 아이템이다) |
| SaveFileName | PIT_BSTR | | 파일 이름. (OpenFileName과 SaveFileName은 같은 아이템을 공유한다. 즉 OpenFileName과 SaveFileName은 이름만 다를 뿐 동일한 아이템이다) |
| OpenFormat | PIT_BSTR | | 파일 형식. (OpenFileName과 마찬가지로 동일 아이템 지칭) |
| SaveFormat | PIT_BSTR | | 파일 형식. (OpenFileName과 마찬가지로 동일 아이템 지칭) |
| OpenReadOnly | PIT_UI1 | | 읽기 전용 |
| OpenFlag | PIT_UI1 | | 옵션<br>0x00 = 새 창으로 열기<br>0x01 = 현재 창의 새 탭에 열기<br>0x02 = 현재 창의 현재 탭에 열기<br>0x03 = 위 세 값의 mask<br>0x04 =이미 열려진 문서일 때 다시 load할지 뭍을 것인지<br>0x08 = 초기 모드를 최근 작업 문서 상태로<br>0x10 = 문서 마당 |
| SaveOverWrite | PIT_UI1 | | 덮어 쓰기 |
| ModifiedFlag | PIT_UI1 | | Modify 플래그 |
| Argument | PIT_BSTR | | Argument |
| SaveCMFDoc30 | PIT_UI1 | | 97 호환 저장 |
| SaveHwp97 | PIT_UI1 | | 97 문서로 저장 |
| SaveDistribute | PIT_UI1 | | 배포용 문서로 저장 |
| SaveDRMDoc | PIT_UI1 | | 보안 문서로 저장 |

---

## 54) FileSaveBlock : 블록 지정된 부분을 저장

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| FileName | PIT_BSTR | | 파일 이름 |
| Format | PIT_BSTR | | 파일 포맷 |
| Argument | PIT_BSTR | | argument |

---

## 55) FileSendMail : 메일 보내기

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| To | PIT_BSTR | | 받는 사람 |
| Type | PIT_UI1 | | 메일 보내기 유형: 0 = 본문, 1 = 첨부파일 |
| Subject | PIT_BSTR | | 제목 |
| Filepath | PIT_BSRT | | 사용자가 직접 입력한 파일 (이 아이템은 Type 아이템이 1(첨부파일)로 설정되어 있을 때만 유효하다) |

---

## 56) FileSetSecurity : 배포용 문서

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Password | PIT_BSTR | | 배포용 문서 암호(7자리 이상 가능) |
| NoPrint | PIT_UI1 | | 프린트 가능한 배포용 문서를 만들지의 여부<br>0 : 프린트 가능<br>1 : 프린트 가능하지 않음 |
| NoCopy | PIT_UI1 | | 문서 내용 복사가 가능한 배포용 문서를 만들지의 여부<br>0 : 복사 가능<br>1 : 복사 가능하지 않음 |

---

## 57) FindReplace : 찾기/찾아 바꾸기

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| FindString | PIT_BSTR | | 찾을 문자열 |
| ReplaceString | PIT_BSTR | | 바꿀 문자열 |
| Direction | PIT_UI1 | | 찾을 방향 :<br>0 = 아래쪽, 1 = 위쪽, 2 = 문서 전체 |
| MatchCase | PIT_UI1 | | 대소문자 구별 (on / off) |
| AllWordForms | PIT_UI1 | | 문자열 결합 (on / off) |
| SeveralWords | PIT_UI1 | | 여러 단어 찾기 (on / off) |
| UseWildCards | PIT_UI1 | | 아무개 문자 (on / off) |
| WholeWordOnly | PIT_UI1 | | 온전한 낱말 (on / off) |
| AutoSpell | PIT_UI1 | | 토씨 자동 교정 (on / off) |
| ReplaceMode | PIT_UI1 | | 찾아 바꾸기 모드 (on / off) |
| IgnoreFindString | PIT_UI1 | | 찾을 문자열 무시 (on / off) |
| IgnoreReplaceString | PIT_UI1 | | 바꿀 문자열 무시 (on / off) |
| FindCharShape | PIT_SET | CharShape | 찾을 글자 모양 |
| FindParaShape | PIT_SET | ParaShape | 찾을 문단 모양 |
| ReplaceCharShape | PIT_SET | CharShape | 바꿀 글자 모양 |
| ReplaceParaShape | PIT_SET | ParaShape | 바꿀 문단 모양 |
| FindStyle | PIT_BSTR | | 찾을 스타일 |
| ReplaceStyle | PIT_BSTR | | 바꿀 스타일 |
| IgnoreMessage | PIT_UI1 | | 메시지박스 표시 안함. (on / off) |
| HanjaFromHangul | PIT_UI1 | | 한글임으로 한자 차기 |
| FindJaso | PIT_UI1 | | 자소로 찾기 (on / off) |
| FindRegExp | PIT_UI1 | | 정규식(조건식)으로 찾기 (on / off) (ver:0x06050107) |
| FindType | PIT_UI1 | | 다시 찾기를 할 때 마지막으로 실행한 [찾기 TRUE]를 할 것인가 [찾아가기 FALSE]할 것인가의 여부 (ver:0x06050110) |

---

## 58) FlashProperties : 플래시 파일 삽입 시 필요한 옵션

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Base | PIT_BSTR | | 경로의 Base |
| Qulaity | PIT_BSTR | | 재생 품질 |
| Scale | PIT_BSTR | | 스케일 |
| WMode | PIT_BSTR | | 윈도우 모드 |
| AutoPlay | PIT_UI1 | | 자동 재생 여부 : 0 = off, 1 = on |
| LoopPlay | PIT_UI1 | | 반복 재생 여부 : 0 = off, 1 = on |
| ShowMenu | PIT_UI1 | | 메뉴 보이기 : 0 = Hide, 1 = Show |
| BgColor | PIT_UI4 | | 배경색 (COLORREF) |

---

## 59) FootnoteShape / EndnoteShape : 미주 / 각주 모양

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| NumberFormat | PIT_UI1 | | 번호모양 |
| UserChar | PIT_UI2 | | 사용자 기호 (WCHAR) |
| PrefixChar | PIT_UI2 | | 앞 장식 문자 (WCHAR) |
| Suffix | PIT_UI2 | | 뒤 장식 문자 (WCHAR) |
| PlaceAt | PIT_UI1 | | 위치<br>- 각주용 옵션 (한 페이지 내에서 각주를 다단에 어떻게 위치시킬지)<br>0 = 각 단마다 따로 배열<br>1 = 통단으로 배열<br>2 = 가장 오른쪽 단에 배열<br><br>- 미주용 옵션 (문서 내에서 미주를 어디에 위치시킬지)<br>0 = 문서의 마지막<br>1 = 구역의 마지막 |
| Restart | PIT_UI1 | | 번호 매기기<br>0 = 앞 구역에 이어서<br>1 = 현재 구역부터 새로 시작<br>2 = 쪽마다 새로 시작 (각주 전용) |
| NewNumber | PIT_UI2 | | 시작 번호 (1 .. n)<br>번호 매기기 값이 '쪽마다 새로 시작' 일 때만 사용된다. |
| LineLength | PIT_I4 | | 구분선 길이 (HWPUNIT) |
| LineType | PIT_UI1 | | 선 종류 |
| LineWidth | PIT_UI1 | | 선 굵기 |
| SpaceAboveLine | PIT_I4 | | 구분선 위 여백 (HWPUNIT) |
| SpaceBelowLine | PIT_I4 | | 구분선 아래 여백 (HWPUNIT) |
| SpaceBetweenNotes | PIT_I4 | | 주석 사이 여백 (HWPUNIT) |
| SuperScript | PIT_UI1 | | 각주 내용 중 번호 코드의 모양을 위첨자 형식으로 할지 |
| BeneathText | PIT_UI1 | | 텍스트에 이어 바로 출력할지 여부 (on / off) |
| LineColor | PIT_UI4 | | 선 색깔 (COLORREF) |

---

## 60) FtpUpload : 웹서버로 올리기

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Server | PIT_BSTR | | Ftp 서버 내임 |
| UserName | PIT_BSTR | | 사용자 이름 |
| Password | PIT_BSTR | | 사용자 패스워드 |
| Directory | PIT_BSTR | | 디렉터리 |
| FileName | PIT_BSTR | | 파일 명 |
| SiteName | PIT_ARRAY | PIT_BSTR | 사이트 이름 |
| SaveType | PIT_UI | | 저장할 포맷. 0 = HTML 1 = HWP |

---

## 61) FtpDownload : 웹서버로 올리기

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Server | PIT_BSTR | | Ftp 서버 내임 |
| UserName | PIT_BSTR | | 사용자 이름 |
| Password | PIT_BSTR | | 사용자 패스워드 |
| Directory | PIT_BSTR | | 디렉터리 |
| FileName | PIT_BSTR | | 파일 명 |
| SaveType | PIT_UI | | 저장할 포맷. 0 = HTML 1 = HWP 2: OOXML |
