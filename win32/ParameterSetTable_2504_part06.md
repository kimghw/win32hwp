# ParameterSetTable_2504_part06

## 42) DrawShear : 그리기 개체 기울이기 정보

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| XFactor | PIT_I | | X축 기울이기 각도 |
| YFactor | PIT_I | | Y축 기울이기 각도 |

---

## 43) DrawTextart : 글맵시 속성

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| String | PIT_BSTR | | 글맵시에 넣을 문자열 내용 |
| FontName | PIT_BSTR | | 글꼴 |
| FontStyle | PIT_BSTR | | 속성. 값은 항상 0(Regular)이다. |
| FontType | PIT_UI1 | | 폰트 종류 : 0 = don't care, 1 = TTF, 2 = HFT |
| LineSpacing | PIT_I4 | | 줄 간격 (50 ~ 500) |
| CharSpacing | PIT_I4 | | 글자간격 (50 ~ 500) |
| AlignType | PIT_UI1 | | 정렬 방식 |
| Shape | PIT_UI1 | | 형태 (0 ~ 54) |
| ShadowType | PIT_UI1 | | 그림자 종류.<br>0 = none, 1 = drop, 2 = continuous |
| ShadowOffsetX | PIT_I1 | | 그림자 X축 간격 (-48% ~ 48%) |
| ShadowOffsetY | PIT_I1 | | 그림자 Y축 간격 (-48% ~ 48%) |
| ShadowColor | PIT_UI4 | | 그림자 색<br>RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR) |
| NumberOfLines | PIT_UI1 | | 글맵시에 넣을 문자열 내용의 줄 수 |

---

## 44) DropCap : 문단 첫 글자 장식

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Style | PIT_UI | | 글자 장식 모양<br>0=없음<br>1=2줄차지<br>2=3줄차지<br>4=여백 |
| FaceName | PIT_BSTR | | 글꼴 |
| LineStyle | PIT_I | | 선 스타일 |
| LineWeight | PIT_UI | | 선 굵기 |
| LineColor | PIT_UI4 | | 선 색<br>RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR) |
| FaceColor | PIT_UI4 | | 글꼴 색<br>RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR) |
| Spacing | PIT_I | | 본문과의 간격 |

---

## 45) Dutmal : 덧말

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| ResultText | PIT_BSTR | | 본말 |
| SubText | PIT_BSTR | | 덧말 |
| PosType | PIT_UI1 | | 덧말 위치. 0 = 위, 1 = 아래. |
| FsizeRatio | PIT_UI1 | | 덧말 크기 Percent. 0=이면 기본 50%. |
| Option | PIT_UI1 | | 옵션 |
| StyleNo | PIT_UI1 | | 스타일번호 (옵션이 켜있을 때) |
| Align | PIT_UI1 | | 덧말의 좌우 Align.<br>0 = 양쪽 정렬<br>1 = 왼쪽 정렬<br>2 = 오른쪽 정렬<br>3 = 가운데 정렬<br>4 = 배분 정렬<br>5 = 나눔 정렬<br>기본은 가운데 정렬이며 옵션이 켜있을 때만 유효 |
| Delete | PIT_UI1 | | 덧말 지움 |
| Modify | PIT_UI1 | | 덧말 편집 모드 여부 |

---

## 46) EngineProperties : 환경 설정 옵션

HwpCtrl의 EngineProperties에서 사용된다. 해당 액션은 존재하지 않음

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| DoAnyCursorEdit | PIT_UI1 | | 마우스로 두 번 누르기 한곳에 입력가능 |
| DoOutLineStyle | PIT_UI1 | | 개요 번호 삽입 문단에 개요 스타일 자동 적용 |
| InsertLock | PIT_UI1 | | 삽입 잠금 |
| TabMoveCell | PIT_UI1 | | 표 안에서 <Tab>으로 셀 이동 |
| FaxDriver | PIT_BSTR | | 팩스 드라이버 |
| PDFDriver | PIT_BSTR | | PDF 드라이버 |
| EnableAutoSpell | PIT_UI1 | | 맞춤법 도우미 작동 |
| OpenNewWindow | PIT_UI1 | | 새창으로 불러오기 |
| CtrlMaskAs2002 | PIT_UI1 | | 한글 2002 방식으로 조판 부호 표시하기 |
| ShowGuildLines | PIT_UI1 | | 개체 투명선 보이기 |
| ImageAutoCheck | PIT_UI1 | | 이미지 파일의 경로 검사 다이얼로그 띄우기 유무. |
| OptimizeWebHwpCopy | PIT_UI1 | | 웹한글로 복사 최적화 끄고/켜기 |

---

## 47) EqEdit : 수식

EqEdit는 ShapeObject로부터 계승받았으므로 위 표에 정리된 EqEdit의 아이템들 이외에 ShapeObject의 아이템들을 사용할 수 있다.

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| String | PIT_BSTR | | 수식 스크립트. |
| BaseUnit | PIT_I4 | | 수식이 삽입되는 앞의 글자와 같은 높이 (기본 값은 POINT 10 ) |
| Color | PIT_I4 | | 수식이 삽입되는 글자 색과 같은 색 (기본 값은 WINDOWTEXT 색)<br>RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR) |
| LineMode | PIT_I4 | | 줄 단위를 사용할지의 여부 |
| Version | PIT_BSTR | | 수식 스크립트 버전 정보 |
| ApplyTo | PIT_UI4 | | 수식 속성 적용 범위<br>0 : 선택된 수식 개체<br>1 : 문서 전체 |
| VisualString | PIT_BSTR | | 수식 비주얼 스크립트 |
| EqFontName | PIT_BSTR | | 폰트명 |

---

## 48) ExchangeFootnoteEndNote : 각주/미주 변환

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Flag | PIT_UI1 | | 옵션<br>0 : 모든 각주를 미주로 바꾸기<br>1 : 모든 미주를 각주로 바꾸기<br>2 : 각주와 미주를 서로 바꾸기 |

---

## 49) FieldCtrl : 필드 컨트롤의 공통 데이터

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Command | PIT_BSTR | | 필드의 명령문 |
| Editable | PIT_UI1 | | 일부분 readonly mode에서 편집 가능한 필드인지 여부 |
| FieldDirty | PIT_UI1 | | 필드가 초기화 상태인지 수정되어 있는 상태인지 여부 |
| CtrlData | PIT_SET | CtrlData | 필드 이름 저장을 위한 영역 |
| User | PIT_BSTR | | 컨트롤 데이터 |
| MemoType | PIT_UI1 | | 메모 타입, (한/글 2022부터 지원) |

---

## 50) FileConvert : 여러 파일을 동시에 특정 포맷으로 변환하여 저장 (관련 Action/API 존재하지 않음)

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Format | PIT_BSTR | | 변환 포맷 |
| SrcFileList | PIT_ARRAY | PIT_BSTR | Source 파일 리스트 |
| DestFileList | PIT_ARRAY | PIT_BSTR | Destination 파일 리스트 |

---

## 51) FileInfo : 파일 정보

HwpCtrl.GetFileInfo에서 사용, 해당 액션은 존재하지 않음.

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Format | PIT_BSTR | | 파일의 형식.<br>HWP : 한글 파일<br>UNKNOWN : 알 수 없음. |
| VersionStr | PIT_BSTR | | 파일의 버전 문자열<br>ex)5.0.0.3 |
| VersionNum | PIT_UI4 | | 파일의 버전<br>ex) 0x05000003 |
| Encrypted | PIT_I4 | | 암호 여부 (현재는 파일 버전 3.0.0.0 이후 문서-한글97, 한글 워디안, 한글 2002-에 대해서만 판단한다.)<br>-1 : 판단 할 수 없음<br>0 : 암호가 걸려 있지 않음<br>양수: 암호가 걸려 있음. |
| Compressed | PIT_I4 | | 압축 여부 |
