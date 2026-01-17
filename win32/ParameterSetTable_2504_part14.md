# ParameterSetTable_2504_part14

---

## 106) ShapeObjComment : 개체 설명문개체 설명문

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| EditShapeObjCommentStr | PIT_BSTR | | 개체 설명 문자열 |
| EditShapeObjCommentFlag | PIT_UI | | 개체 설명문 데이터 작성을 위한 추가 데이터 전달. 0 = 사용안함, 1 = 문단띠) |

---

## 107) ShapeObjectCopyPaste : 그리기 개체 모양 복사/붙여 넣기

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Type | PIT_I | | 그리기 모양 복사/붙여 넣기 종류 (예약.. 현재 사용하지 않음) |
| ShapeObjectLine | PIT_UI | | 그리기 선 모양 복사 |
| ShapeObjectFill | PIT_UI1 | | 그리기 채우기 복사 |
| ShapeObjectSize | PIT_UI1 | | 그리기 개체 크기 복사 |
| ShapeObjectShadow | PIT_UI1 | | 그리기 개체 그림자 복사 |
| ShapeObjectPicEffect | PIT_UI1 | | 그림 효과 복사 |

---

## 108) Sort : 소트

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| KeyOption | PIT_ARRAY | PIT_I4 | 키 콤보에서 선택된 키를 저장함. |
| CheckJasoReverse | PIT_UI1 | | 자소 단위 비교 Flag - 종, 중, 초 |
| DelimiterType | PIT_UI1 | | 필드 구분 기호 형식 : 0 = 탭(Tab), 1 = 콤마(,), 2 = 빈칸(Space), 3 = 사용자 정의 |
| DelimiterChars | PIT_BSTR | | 필드 구분 기호들. DelimiterType이 3(사용자 정의)일 경우에만 유효 |
| IgnoreMultiDelimiter | PIT_UI1 | | 연속되는 구분기호 무시 Flag |
| CheckFromRear | PIT_UI1 | | 단어 뒤에서 부터 비교 Flag |
| CheckExtendYear | PIT_UI1 | | 두 자리 년도 확장 check Flag |
| YearBase | PIT_UI | | 두 자리 년도 시작 년도 |
| LangOrderType | PIT_UI1 | | 사전언어순서 값 |
| CheckJaso | PIT_UI1 | | 자소 단위 비교 Flag - 초, 중, 종 |
| EachPara | PIT_UI1 | | 정렬시 각 문단별 정렬 여부 Flag – 키 옵션이 "단어" 인 경우만 사용됨 |

---

## 109) Style : 스타일

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Apply | PIT_I | | 적용할 스타일 인덱스 |

---

## 110) StyleDelete : 스타일 지우기

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Target | PIT_I | | 지워야할 스타일 인덱스 |
| Alternation | PIT_I | | 대체할 스타일 인덱스 |

---

## 111) StyleTemplate : 스타일 마당

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| FileName | PIT_BSTR | | 파일 이름 |

---

## 112) Sum : 블록 계산 (합계/평균/줄 수)

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Sum | PIT_BSTR | | 합 |
| Average | PIT_BSTR | | 평균 |
| LineCount | PIT_BSTR | | 줄 수 |
| Comma | PIT_UI1 | | 세 자리마다 쉼표로 자리 구분 (on / off) |
| Option | PIT_I4 | | 형식 옵션 |
| Method | PIT_I4 | | 동작 옵션 (0:copy, 1: paste) |

---

## 113) SummaryInfo : 문서 정보

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Title | PIT_BSTR | | 제목 |
| Subject | PIT_BSTR | | 주제 |
| Author | PIT_BSTR | | 지은이 |
| Date | PIT_BSTR | | 날짜 |
| KeyWords | PIT_BSTR | | 키워드 |
| Comments | PIT_BSTR | | 기타 |
| CreationTimeLow | PIT_UI4 | | 작성한 날짜 (low) |
| CreationTimeHigh | PIT_UI4 | | 작성한 날짜 (high) |
| ModifiedTimeLow | PIT_UI4 | | 마지막 수정한 날짜 (low) |
| ModifiedTimeHigh | PIT_UI4 | | 마지막 수정한 날짜 (high) |
| PrintedTimeLow | PIT_UI | | 마지막 인쇄한 날짜 (low) |
| PrintedTimeHigh | PIT_UI4 | | 마지막 인쇄한 날짜 (high) |
| LastSavedBy | PIT_BSTR | | 마지막 저장한 사람 |
| Characters | PMT_INT | | 문서분량 (글자) |
| Words | PMT_INT | | 문서분량 (낱말) |
| Lines | PMT_INT | | 문서분량 (줄) |
| Paragraphs | PMT_INT | | 문서분량 (문단) |
| Pages | PMT_INT | | 문서분량 (쪽) |
| CopyPapers | PMT_INT | | 문서분량 (원고지) |
| Etcetera | PMT_INT | | 문서분량 (표, 그림 등) |
| DocVersion | PIT_BSTR | | 문서 파일 버전 |
| HwpVersion | PIT_BSTR | | 문서를 생성한 한글 워드프로그램의 버전 |
| HanjaChar | PIT_I1 | | 문서분량 (한자 수) |
| CharactersExceptSpace | PIT_I | | 문서분량 (글자-공백 제외) |
| ExtractImagePath | PIT_BSTR | | 그림 파일 연결 경로 - 한/글 외부에서 사용함 |
| ExtractImageBaseFileName | PIT_BSTR | | 삽입 그림을 연결 그림으로 변경할 때, 파일의 기본 이름 |
| ExtractImageExtName | PIT_BSTR | | 연결 그림으로 변경할 삽입 그림 파일의 확장자 |
| ChangeImageExtFrom | PIT_BSTR | | 변경할 그림 파일의 확장자(FROM) |
| ChangeImageExtTo | PIT_BSTR | | 변경할 그림 파일의 확장자(TO) |
| EmbedImagePath | PIT_BSTR | | 삽입할 연결 그림 파일의 경로 |
| SelectedSummaryInfo | PIT_SET | SummaryInfo | 선택된 영역의 정보 |
| LicenseInfo | PIT_SET | SummaryInfo | 저작권 CCL, 공공누리 정보 |
| LicenseFlag | PIT_UI4 | SummaryInfo | 공공누리, CCL 플래그 |
| UserPropertyName | PIT_ARRAY | PIT_BSTR | 사용자 속성 – 이름, (한/글 2024부터 지원) |
| UserPropertyFormat | PIT_ARRAY | PIT_BSTR | 사용자 속성 - 형식, (한/글 2024부터 지원) |
| UserPropertyValue | PIT_ARRAY | PIT_BSTR | 사용자 속성 - 값, (한/글 2024부터 지원) |

---

## 114) TabDef : 탭 정의

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| AutoTabLeft | PIT_UI1 | | 문단 왼쪽 끝 탭 (on / off) |
| AutoTabRight | PIT_UI1 | | 문단 오른쪽 끝 탭 (on / off) |
| TabItem | PIT_ARRAY | PIT_I | 각각의 탭 정의. 하나의 탭 아이템은 세 개의 인수로 표현되어 있음. (n * 3 + 0) - PIT_I : 탭 위치 (URC), (n * 3 + 1) - PIT_I : 채울 모양 (아래참조), (n * 3 + 2) - PIT_I : 탭 종류 (아래참조.) 채울 모양 : 선 종류, 탭 종류 : 0 = 왼쪽, 1 = 오른쪽, 2 = 가운데, 3 = 소수점 |
| DeleteTab | PIT_ARRAY | PIT_I | 지운 탭 위치 (URC) 첫 번째 값이 -1 이면 모두 지웠음을 의미한다. |

---

*Page 131-140*
