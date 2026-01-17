# ParameterSetTable_2504_part16

---

## 122) TableSwap : 표 뒤집기

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Type | PIT_UI1 | | 표 뒤집기 형식<br>0 = 상하 뒤집기<br>1 = 좌우 뒤집기<br>2 = X와 Y를 바꿈<br>3 = 반시계 방향으로 90도 회전<br>4 = 180도 회전<br>5 = 시계 방향으로 90도 회전 |
| SwapMargin | PIT_UI1 | | 여백 뒤집기 지원여부 |

---

## 123) TableTblToStr : 표를 문자열로

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| DelimiterType | PIT_UI1 | | 분리 문자(탭, 쉼표, 공백) |
| UserDefine | PIT_BSTR | | 사용자 정의 필드 구분 기호 |

---

## 124) TableTemplate : 표 마당 정보

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Format | PIT_UI | | 적용할 서식. 다음 값의 조합으로 구성된다.<br>0x0001 = 테두리<br>0x0002 = 글자 모양과 문단 모양<br>0x0004 = 셀 배경<br>0x0008 = 그레이 스케일 |
| ApplyTarger | PIT_UI | | 적용 대상. 다음 값의 조합으로 구성된다.<br>0x0001 = 제목 줄<br>0x0002 = 마지막 줄<br>0x0004 = 첫째 칸<br>0x0008 = 마지막 칸 |
| FileName | PIT_BSTR | | 표 마당 파일 이름 |
| CreateMode | PIT_UI1 | | 표 만들기 모드 (표 만들기에서 제목줄에 제목 속성 넣기 위해) |
| ThemeColor | PIT_UI | | 테마색 설정. |

---

## 125) TextCtrl : TEXT 컨트롤의 공통 데이터

CtrlCode.Properties에서 사용된다.

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| CtrlData | PIT_SET | CtrlData | 컨트롤 이름 저장을 위한 영역 |

---

## 126) TextVertical : 세로쓰기

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Landscope | PIT_UI1 | | 용지 방향. 0 = 좁게, 1 = 넓게 |
| TextDirection | PIT_UI2 | | 글자 방향.<br>0 = 보통 (왼쪽에서 오른쪽)<br>1 = 세로쓰기 (라틴 문자 회전)<br>2 = 세로쓰기 (라틴 문자 포함) |
| TextVerticalWidthHead | PIT_I | | 머리말/꼬리말 세로쓰기 여부 |
| ApplyTo | PIT_UI1 | | 적용 대상<br>0 = 선택된 구역<br>1 = 선택된 문자열<br>2 = 현재 구역<br>3 = 문서전체<br>4 = 새 구역 : 현재 위치부터 새로<br>5 = no items (적용대상 없음) |
| ApplyClass | PIT_UI1 | | 적용 대상 분류.<br>적용 대상 분류는 현재 캐럿의 상태에 따라 ApplyTo에 적용 가능한 대상을 한정짓는 역할을 한다. 내부적으로 값이 계산되므로, 값을 참조하는 용도로만 사용하도록 한다. 다음의 값의 조합으로 구성된다.<br>0x0001 = 선택된 구역<br>0x0002 = 선택된 문자열<br>0x0004 = 현재 구역<br>0x0008 = 문서 전체<br>0x0010 = 새 구역 : 현재 위치부터 새로 |

---

## 127) UserQCommandFile : 사용자 자동 명령 파일 저장/로드

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Save | PIT_I4 | | 저장 (TRUE = Save, FALSE = Open) |
| FileName | PIT_BSTR | | 파일명 |
| LoadType | PIT_I4 | | 로드 방법 (TRUE = Overwrite, FALSE = Merge) |

---

## 128) VersionInfo : 버전 정보

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| SourcePath | PIT_BSTR | | 버전 비교용 소스 패스 |
| TargetPath | PIT_BSTR | | 버전 비교용 타겟 패스 |
| ItemStartIndex | PIT_UI1 | | 버전 비교를 보여줄 시작 히스토리 인덱스 |
| ItemEndIndex | PIT_UI1 | | 버전 비교를 보여줄 마지막 히스토리 인덱스 |
| ItemOverWrite | PIT_UI1 | | 히스토리 정보를 저장할 때 마지막 버전으로 덮어쓰는 플랙 (on/off) |
| ItemSaveDescription | PIT_UI1 | | 히스토리 정보를 저장할 때 설명을 입력하는 대화상자를 띄우는 플랙 (on/off) |
| TempFilePath | PIT_ARRAY | PIT_BSTR | 버전 비교용 임시파일 경로 |
| ItemInfoIndex | PIT_UI4 | | 버전 정보 얻어오기 및 삭제 시 사용될 인덱스 |
| SaveFilePath | PIT_BSTR | | 버전 저장 파일 경로(OCX 컨트롤용) |
| ItemInfoWriter | PIT_BSTR | | 작성자 정보 |
| ItemInfoDescription | PIT_BSTR | | 해당 버전에 대한 설명 |
| ItemInfoTimeHi | PIT_UI4 | | 날짜 정보, FILETIME의 HIWORD |
| ItemInfoTimeLo | PIT_UI4 | | 날짜 정보, FILETIME의 LOWORD |
| ItemInfoLock | PIT_UI1 | | 히스토리 정보 수정 플랙 |
| VersionAutoSave | PIT_UI1 | | 새 버전으로 자동 저장 on/off |
| VersionDiffSplitView | PIT_UI1 | | 버전 비교 방식 (한 화면에서 비교 : 0, 두 화면에서 비교 : 1) |
| UsedStanTime | PIT_UI1 | | 표준시 사용 플랙 |
| UsedCert | PIT_UI1 | | 공인인증서 인증 사용 플랙 |
| FileDiff | PIT_UI1 | | 문서 비교 액션 플랙 |
| ResultSourcePath | PIT_BSTR | | 문서 비교 SRC 결과 파일 경로 (OCX 컨트롤용) |
| ResultTargetPath | PIT_BSTR | | 문서 비교 TGT 결과 파일 경로 (OCX 컨트롤용) |
| ResultMergedPath | PIT_BSTR | | 문서 비교 Merged 결과 파일 경로 (OCX 컨트롤용) |
| ResultOption | PIT_UI1 | | 문서 비교 결과 옵션 (0:메모, 1:교정부호) (OCX 컨트롤용) |
| ResultShowMemo | PIT_UI1 | | 문서 비교 결과 메모를 화면에 보이게 할 지 여부(문서 비교 결과 옵션이 메모일 경우에 동작) (OCX 컨트롤용) |

---

## 129) ViewProperties : 뷰의 속성

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| OptionFlag | PIT_UI | | 뷰 옵션 플랙. 여러 개를 OR연산하여 지정할 수 있음.<br>0x0001 = off : 쪽윤곽, on : 기본 보기<br>0x0002 = 공백과 폭이 없는 컨트롤을 기호로<br>0x0004 = 문단 마크 기호로<br>0x0008 = 안내선<br>0x0010 = 그리기 격자<br>0x0020 = 그림 감춤 |
| ZoomType | PIT_UI1 | | 화면 확대 종류.<br>0 = 사용자 정의<br>1 = 쪽 맞춤<br>2 = 폭 맞춤<br>3 = 여러 쪽 |
| ZoomRatio | PIT_UI2 | | 화면 확대 종류가 "사용자 정의"인 경우 화면 확대 비율. 10% ~ 500% |
| ZoomCntX | PIT_UI1 | | 화면 확대 종류가 "여러 쪽"인 경우 가로 페이지 수. 1 ~ 8 |
| ZoomCntY | PIT_UI1 | | 화면 확대 종류가 "여러 쪽"인 경우 세로 페이지 수. 1 ~ 8 |
| ZoomMirror | PIT_UI1 | | 맞쪽 보기. 페이지 수가 2의 배수일 때만 동작 |
| PageDir | PIT_UI1 | | 쪽 방향(HWPPAGE_DIR_VERT : 0, HWPPAGE_DIR_HORZ : 1) |
| MouseWheelDir | PIT_UI1 | | 마우스 휠 방향(HWPWHEEL_DIR_VERT : 0, HWPWHEEL_DIR_HORZ : 1) |
| DragDrop | PIT_UI1 | | 드래그 앤 드롭 지원 |

---

## 130) ViewStatus : 뷰 상태 정보 ver:0x06000101

HwpCtrl.GetViewStatus에서 사용, 해당 액션은 존재하지 않음.

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Type | PIT_UI | | 0 (현재 View의 절대 Pos값만 지원함) |
| ViewPosX | PIT_I4 | | 현재 뷰의 X값 |
| ViewPosY | PIT_I4 | | 현재 뷰의 Y값 |

---

## 131) CompatibleDocument : 호환 문서

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| TargetProgram | PIT_UI | | 대상 프로그램 |
| Default | PIT_UI | | 대화상자 기본 제공 여부 |
| CurrentVersion | PIT_UI | | 현재 버전 여부 |
