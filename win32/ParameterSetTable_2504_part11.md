# ParameterSetTable_2504_part11

---

## 88) PageHiding : 감추기

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Fields | PIT_UI | | 감출 대상 비트 필드.<br>0x01 = 머리말<br>0x02 = 꼬리말<br>0x04 = 바탕쪽<br>0x08 = 테두리<br>0x10 = 배경<br>0x20 = 쪽번호 위치 |

---

## 89) PageNumCtrl : 페이지번호 (97의 홀수 쪽에서 시작)

CtrlCode.Properties에서 사용된다.

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| PageStartsOn | PIT_UI1 | | 페이지 번호 적용 옵션.<br>0 = 양쪽<br>1 = 짝수쪽<br>2 = 홀수쪽 |

---

## 90) PageNumPos : 쪽 번호 위치

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| NumberFormat | PIT_UI1 | | 번호 모양 :<br>0 = 1, 2, 3<br>1 = ①, ②, ③<br>2 = Ⅰ, Ⅱ, Ⅲ<br>3 = ⅰ, ⅱ, ⅲ<br>4 = A, B, C<br>8 = 가, 나, 다<br>13 = 一, 二, 三<br>15 = 갑, 을, 병<br>16 = 甲, 乙, 丙<br>※ 중간에 빈 번호에도 문자포맷이 존재하나 이곳에서 사용하지 않아 생략함 |
| UserChar | PIT_UI2 | | 사용자 기호(WCHAR). 한글2007에선 더 이상 사용하지 않는다. |
| PrefixChar | PIT_UI2 | | 앞 장식 문자(WCHAR). 한글2007에선 더 이상 사용하지 않는다. |
| SuffixChar | PIT_UI2 | | 뒤 장식 문자(WCHAR). 한글2007에선 더 이상 사용하지 않는다. |
| SideChar | PIT_UI2 | | 양쪽 옆 장식 문자(WCHAR). L'-'만 사용할 수 있다. |
| DrawPos | PIT_UI1 | | 번호 위치<br>0 = 쪽 번호 없음<br>1 = 왼쪽 위,<br>2 = 가운데 위,<br>3 = 오른쪽 위,<br>4 = 왼쪽 아래,<br>5 = 가운데 아래,<br>6 = 오른쪽 아래,<br>7 = 바깥쪽 위,<br>8 = 바깥쪽 아래,<br>9 = 안쪽 위,<br>10 = 안쪽 아래 |
| NewNumber | PIT_UI2 | | 새 시작 번호 (1 .. n) |

---

## 91) ParaShape : 문단 모양

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| LeftMargin | PIT_I4 | | 왼쪽 여백 (URC) |
| RightMargin | PIT_I4 | | 오른쪽 여백 (URC) |
| Indentation | PIT_I4 | | 들여쓰기/내어 쓰기 (URC) |
| PrevSpacing | PIT_I4 | | 문단 간격 위 (URC) |
| NextSpacing | PIT_I4 | | 문단 간격 아래 (URC) |
| LineSpacingType | PIT_UI1 | | 줄 간격 종류 (HWPUNIT)<br>0 = 글자에 따라<br>1 = 고정 값<br>2 = 여백만 지정 |
| LineSpacing | PIT_I4 | | 줄 간격 값.<br>줄 간격 종류(LineSpacingType)에 따라 :<br>- "글자에 따라"일 경우(0 - 500%)<br>- "고정 값"일 경우(URC)<br>- "여백만 지정"일 경우(URC) |
| AlignType | PIT_UI1 | | 정렬 방식<br>0 = 양쪽 정렬<br>1 = 왼쪽 정렬<br>2 = 오른쪽 정렬<br>3 = 가운데 정렬<br>4 = 배분 정렬<br>5 = 나눔 정렬 (공백에만 배분) |
| BreakLatinWord | PIT_UI1 | | 줄 나눔 단위 (라틴 문자)<br>0 = 단어<br>1 = 하이픈<br>2 = 글자 |
| BreakNonLatinWord | PIT_UI1 | | 단위 (비 라틴 문자) TRUE = 글자, FALSE = 어절 |
| SnapToGrid | PIT_UI1 | | 편집 용지의 줄 격자 사용 (on / off) |
| Condense | PIT_UI1 | | 공백 최소값 (0 - 75%) |
| WidowOrphan | PIT_UI1 | | 외톨이줄 보호 (on / off) |
| KeepWithNext | PIT_UI1 | | 다음 문단과 함께 (on / off) |
| KeepLinesTogether | PIT_UI1 | | 문단 보호 (on / off) |
| PagebreakBefore | PIT_UI1 | | 문단 앞에서 항상 쪽 나눔 (on / off) |
| TextAlignment | PIT_UI1 | | 세로 정렬<br>0 = 글꼴기준<br>1 = 위<br>2 = 가운데<br>3 = 아래 |
| FontLineHeight | PIT_UI1 | | 글꼴에 어울리는 줄 높이 (on / off) |
| HeadingType | PIT_UI1 | | 문단 머리 모양<br>0 = 없음<br>1 = 개요<br>2 = 번호<br>3 = 불릿 |
| Level | PIT_UI1 | | 단계 (0 - 6) |
| BorderConnect | PIT_UI1 | | 문단 테두리/배경 - 테두리 연결 (on / off) |
| BorderText | PIT_UI1 | | 문단 테두리/배경 - 여백 무시 (0 = 단, 1 = 텍스트) |
| BorderOffsetLeft | PIT_I | | 문단 테두리/배경 - 4방향 간격 (HWPUNIT) : 왼쪽 |
| BorderOffsetRight | PIT_I | | 문단 테두리/배경 - 4방향 간격 (HWPUNIT) : 오른쪽 |
| BorderOffsetTop | PIT_I | | 문단 테두리/배경 - 4방향 간격 (HWPUNIT) : 위 |
| BorderOffsetBottom | PIT_I | | 문단 테두리/배경 - 4방향 간격 (HWPUNIT) : 아래 |
| TailType | PIT_UI1 | | 문단 꼬리 모양 (마지막 꼬리 줄 적용) on/off |
| LineWrap | PIT_UI1 | | 글꼴에 어울리는 줄 높이 (on/off) |
| TabDef | PIT_SET | TabDef | 탭 정의 |
| Numbering | PIT_SET | NumberingShape | 문단 번호<br>문단 머리 모양(HeadingType)이 '개요', '번호' 일 때 사용 |
| Bullet | PIT_SET | BulletShape | 불릿 모양<br>문단 머리 모양(HeadingType)이 '불릿'(글머리표) 일 때 사용 |
| BorderFill | PIT_SET | BorderFill | 테두리/배경 |
| AutoSpaceEAsianEng | PIT_UI1 | | 한글과 영어 간격을 자동 조절 (on/off) |
| AutoSpaceEAsianNum | PIT_UI1 | | 한글과 숫자 간격을 자동 조절 (on/off) |
| SuppressLineNum | PIT_UI1 | | 줄 번호 표시 여부 |
| Checked | PIT_UI1 | | 체크 글머리표 사용시 체크 여부 (on/off) |
| TextDir | PIT_UI1 | | 텍스트 방향<br>0 = 자동<br>1 = 오른편에서 왼편<br>2 = 왼편에서 오른편<br>(한/글 2022부터 지원) |

---

## 92) Password : 문서 암호

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| String | PIT_BSTR | | 암호 문자열 |
| FullRange | PIT_UI1 | | TRUE = 유니코드 모든 문자를 사용, FALSE = 영문자만 사용 |
| Ask | PIT_UI1 | | TRUE = 문서 암호를 확인, FALSE = 문서 암호를 설정 |
| Level | PIT_UI1 | | 0 = 보안 수준 낮음, 1 = 보안 수준 높음 |
| RWAsk | PIT_UI1 | | 읽기/쓰기 암호 설정을 위해 새로 추가 |
| ReadString | PIT_BSTR | | 열기 암호 문자열 |
| WriteString | PIT_BSTR | | 쓰기 암호 문자열 |
| ReadOnly | PIT_UI1 | | 0 = 읽기 전용 아님, 1 = 읽기 전용 열기 |

---

## 93) Preference : 환경 설정

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| ShowSinglePage | PIT_I | | 환경 설정 PropertySheet에 표시할 PropertyPage 번호 (하나만 선택) |
| ApplyLinkAttr | PIT_UI1 | | 하이퍼링크 글자 속성 문서 전체에 적용하기 여부 (on/off) |
| ApplyForbidden | PIT_BSTR | | (금칙 처리) 새 문서에 기본 값으로 설정 (on/off) |
| StartForbiddenStr | PIT_BSTR | | (금칙 처리) 새 문서에 적용할 줄 앞 금칙 문자열 |
| EndForbiddenStr | PIT_BSTR | | (금칙 처리) 새 문서에 적용할 줄 뒤 금칙 문자열 |
| UsePageLayout | PIT_UI1 | | 새 문서 속성을 변경하면 유효화된 값으로 판단하고 사용함 |
| PasteObjectAsPicture | PIT_UI1 | | 그림(메타파일) 형식으로 개체 붙이기 여부 |
| HwpxFormatDefaults | PIT_UI1 | | hwpx 기본 포맷 설정 여부 |

---

## 94) Presentation : 프레젠테이션

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Effect | PIT_UI | | 화면 전환 효과 |
| Sound | PIT_BINDATA | | 효과음 |
| InvertText | PIT_I | | 검은색 글자를 흰색으로 |
| ShowMode | PIT_I | | 자동 전환 모드 |
| ShowPage | PIT_UI | | 현재 쪽 |
| ShowTime | PIT_UI | | 전환 시간 (초) |

---

## 95) Print : 인쇄

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Range | PIT_UI1 | | 인쇄 범위<br>0 = 문서전체 (연결된 문서 포함)<br>1 = 현재 쪽<br>2 = 현재부터<br>3 = 현재까지<br>4 = 일부분<br>5 = 선택한 쪽만<br>6 = 현재 문서 (연결된 문서 미포함)<br>7 = 현재 구역 |
| RangeCustom | PIT_BSTR | | 사용자가 직접 입력한 인쇄 범위 |
| RangeIncludeLinkedDoc | PIT_UI1 | | 연결된 문서 포함 |
| NumCopy | PIT_UI2 | | 인쇄 매수 |
| Collate | PIT_UI1 | | 한 부씩 찍기 |
| PrintMethod | PIT_UI1 | | 인쇄 방법<br>0 = 자동 인쇄<br>1 = 공급 용지에 맞추어<br>2 = 나눠 찍기<br>3 = 자동으로 모아 찍기<br>4 = 2쪽씩 모아 찍기<br>5 = 3쪽씩 모아 찍기<br>6 = 4쪽씩 모아 찍기<br>7 = 6쪽씩 모아 찍기<br>8 = 8쪽씩 모아 찍기<br>9 = 9쪽씩 모아 찍기<br>10 = 16쪽씩 모아 찍기 |
| PrinterPaperSize | PIT_I | | 공급용지 종류(DEVMODE.dmPaperSize) |
| PrinterPaperWidth | PIT_I | | 공급용지 종류(DEVMODE.dmPaperWidth) |
| PrinterPaperLength | PIT_I | | 공급용지 종류(DEVMODE.dmPaperLength) |
| PrintAutoHeadNote | PIT_UI1 | | 머리말 자동 인쇄 |
| PrintAutoFootNote | PIT_UI1 | | 꼬리말 자동 인쇄 |
| PrintAutoHeadnoteLtext | PIT_BSTR | | 자동 머리말의 왼쪽 String |
| PrintAutoHeadnoteCtext | PIT_BSTR | | 자동 머리말의 가운데 String |
| PrintAutoHeadnoteRtext | PIT_BSTR | | 자동 머리말의 오른쪽 String |
