# ParameterSet Table (Part 02)

### 8) Caption : 캡션 속성

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Side | PIT_UI1 | | 방향. 0 = 왼쪽, 1 = 오른쪽, 2 = 위, 3 = 아래 |
| Width | PIT_I | | 캡션 폭 (가로 방향일 때만 사용됨. 단위 HWPUNIT) |
| Gap | PIT_I | | 캡션과 틀 사이 간격(HWPUNIT) |
| CapFullSize | PIT_UI1 | | 캡션 폭에 여백을 포함할지 여부 (세로 방향일 때만 사용됨) |

---

### 9) CaptureEnd : 갈무리 끝

갈무리 때 코어엔진에서 액션에게 보내는 정보(시작/끝점의 좌표). 내부에서만 사용된다.

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| BeginX | PIT_I | | 시작점과 X 좌표(페이지 X좌표) |
| BeginY | PIT_I | | 시작점과 Y 좌표(페이지 Y좌표) |
| EndX | PIT_I | | 끝점의 X 좌표(페이지 X좌표) |
| EndY | PIT_I | | 끝점의 Y 좌표(페이지 Y좌표) |
| PageNum | PIT_UI | | 페이지 번호 |
| Path | PIT_BSTR | | 파일 경로, (한/글 2022부터 지원) |
| Format | PIT_I | | 포맷 번호, (한/글 2022부터 지원) |

---

### 10) Cell : 셀

Cell은 ListProperties로부터 계승받았으므로 위 표에 정리된 Cell의 아이템들 이외에 ListProperties의 아이템들을 사용할 수 있다.

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| HasMargin | PIT_UI1 | | 테이블의 기본 셀 여백 대신 자체 셀 여백을 적용할지 여부. (on / off) |
| Protected | PIT_UI1 | | 사용자 편집을 막을지 여부 : 0 = off, 1 = on |
| Header | PIT_UI1 | | 제목 셀인지 여부 : 0 = off, 1 = on |
| Width | PIT_I4 | | 셀의 폭 (HWPUNIT) |
| Height | PIT_I4 | | 셀의 높이 (HWPUNIT) |
| Editable | PIT_UI1 | | 양식모드에서 편집 가능 여부 : 0 = off, 1 = on |
| Dirty | PIT_UI1 | | 초기화 상태인지 수정된 상태인지 여부 : 0 = off, 1 = on |
| CellCtrlData | PIT_SET | CtrlData | 셀 데이터 |

---

### 11) CellBorderFill : 셀 테두리/배경

CellBorderFill은 BorderFillExt로부터 계승받았으므로 위 표에 정리된 CellBorderFill의 아이템들 이외에 BorderFillExt의 아이템들을 사용할 수 있다.

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| ApplyTo | PIT_UI1 | | 적용 대상 : 0 = 선택된 셀, 1 = 전체 셀, 2 = 여러 셀에 걸쳐 적용 |
| NoNeighborCell | PIT_UI1 | | 주변 셀에 선 모양을 적용하지 않을지 여부 (1이면 적용하지 않는다) |
| TableBorderFill | PIT_SET | BorderFill | 표 테두리/배경 |
| AllCellsBorderFill | PIT_SET | BorderFill | 전체 셀의 테두리/배경 |
| SelCellsBorderFill | PIT_SET | BorderFill | 선택된 셀의 테두리/배경 |

---

### 12) ChangeRome : 로마자 변환

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Option | PIT_UI1 | | 변환옵션 :<br>0 = 일반<br>1 = 주소<br>2 = 사람이름<br>3 = 한글복원 |
| HanString | PIT_BSTR | | 변경시킬 또는 변경된 한글문자 |

---

### 13) CharShape : 글자 모양

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| FaceNameHangul | PIT_BSTR | | 글꼴 이름 (한글) |
| FaceNameLatin | PIT_BSTR | | 글꼴 이름 (영문) |
| FaceNameHanja | PIT_BSTR | | 글꼴 이름 (한자) |
| FaceNameJapanese | PIT_BSTR | | 글꼴 이름 (일본어) |
| FaceNameOther | PIT_BSTR | | 글꼴 이름 (기타) |
| FaceNameSymbol | PIT_BSTR | | 글꼴 이름 (심벌) |
| FaceNameUser | PIT_BSTR | | 글꼴 이름 (사용자) |
| FontTypeHangul | PIT_UI1 | | 폰트 종류 (한글) : 0 = don't care, 1 = TTF, 2 = HFT |
| FontTypeLatin | PIT_UI1 | | 폰트 종류 (영문) : 0 = don't care, 1 = TTF, 2 = HFT |
| FontTypeHanja | PIT_UI1 | | 폰트 종류 (한자) : 0 = don't care, 1 = TTF, 2 = HFT |
| FontTypeJapanese | PIT_UI1 | | 폰트 종류 (일본어) : 0 = don't care, 1 = TTF, 2 = HFT |
| FontTypeOther | PIT_UI1 | | 폰트 종류 (기타) : 0 = don't care, 1 = TTF, 2 = HFT |
| FontTypeSymbol | PIT_UI1 | | 폰트 종류 (심벌) : 0 = don't care, 1 = TTF, 2 = HFT |
| FontTypeUser | PIT_UI1 | | 폰트 종류 (사용자) : 0 = don't care, 1 = TTF, 2 = HFT |
| SizeHangul | PIT_UI1 | | 각 언어별 크기 비율. (한글) 10% - 250% |
| SizeLatin | PIT_UI1 | | 각 언어별 크기 비율. (영문) 10% - 250% |
| SizeHanja | PIT_UI1 | | 각 언어별 크기 비율. (한자) 10% - 250% |
| SizeJapanese | PIT_UI1 | | 각 언어별 크기 비율. (일본어) 10% - 250% |
| SizeOther | PIT_UI1 | | 각 언어별 크기 비율. (기타) 10% - 250% |
| SizeSymbol | PIT_UI1 | | 각 언어별 크기 비율. (심벌) 10% - 250% |
| SizeUser | PIT_UI1 | | 각 언어별 크기 비율. (사용자) 10% - 250% |
| RatioHangul | PIT_UI1 | | 각 언어별 장평 비율. (한글) 50% - 200% |
| RatioLatin | PIT_UI1 | | 각 언어별 장평 비율. (영문) 50% - 200% |
| RatioHanja | PIT_UI1 | | 각 언어별 장평 비율. (한자) 50% - 200% |
| RatioJapanese | PIT_UI1 | | 각 언어별 장평 비율. (일본어) 50% - 200% |
| RatioOther | PIT_UI1 | | 각 언어별 장평 비율. (기타) 50% - 200% |
| RatioSymbol | PIT_UI1 | | 각 언어별 장평 비율. (심벌) 50% - 200% |
| RatioUser | PIT_UI1 | | 각 언어별 장평 비율. (사용자) 50% - 200% |
| SpacingHangul | PIT_I1 | | 각 언어별 자간. (한글) -50% - 50% |
| SpacingLatin | PIT_I1 | | 각 언어별 자간. (영문) -50% - 50% |
| SpacingHanja | PIT_I1 | | 각 언어별 자간. (한자) -50% - 50% |
| SpacingJapanese | PIT_I1 | | 각 언어별 자간. (일본어) -50% - 50% |
| SpacingOther | PIT_I1 | | 각 언어별 자간. (기타) -50% - 50% |
| SpacingSymbol | PIT_I1 | | 각 언어별 자간. (심벌) -50% - 50% |
| SpacingUser | PIT_I1 | | 각 언어별 자간. (사용자) -50% - 50% |
| OffsetHangul | PIT_I1 | | 각 언어별 오프셋. (한글) -100% - 100% |
| OffsetLatin | PIT_I1 | | 각 언어별 오프셋. (영문) -100% - 100% |
| OffsetHanja | PIT_I1 | | 각 언어별 오프셋. (한자) -100% - 100% |
| OffsetJapanese | PIT_I1 | | 각 언어별 오프셋. (일본어) -100% - 100% |
| OffsetOther | PIT_I1 | | 각 언어별 오프셋. (기타) -100% - 100% |
| OffsetSymbol | PIT_I1 | | 각 언어별 오프셋. (심벌) -100% - 100% |
| OffsetUser | PIT_I1 | | 각 언어별 오프셋. (사용자) -100% - 100% |
| Bold | PIT_UI1 | | Bold : 0 = off, 1 = on |
| Italic | PIT_UI1 | | Italic : 0 = off, 1 = on |
| SmallCaps | PIT_UI1 | | Small Caps : 0 = off, 1 = on |
| Emboss | PIT_UI1 | | Emboss : 0 = off, 1 = on |
| Engrave | PIT_UI1 | | Engrave : 0 = off, 1 = on |
| SuperScript | PIT_UI1 | | Superscript : 0 = off, 1 = on |
| SubScript | PIT_UI1 | | Subscript : 0 = off, 1 = on |
| UnderlineType | PIT_UI1 | | 밑줄 종류 :<br>0 = none, 1 = bottom, 2 = center, 3 = top |
| OutlineType | PIT_UI1 | | 외곽선 종류 :<br>0 = none, 1 = solid, 2 = dot, 3 = thick,<br>4 = dash, 5 = dashdot, 6 = dashdotdot |
| ShadowType | PIT_UI1 | | 그림자 종류 :<br>0 = none, 1 = drop, 2 = continuous |
| TextColor | PIT_UI4 | | 글자색. (COLORREF)<br>RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR) |
| ShadeColor | PIT_UI4 | | 음영색. (COLORREF)<br>RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR) |
| UnderlineShape | PIT_UI1 | | 밑줄 모양 : 선 종류 |
| UnderlineColor | PIT_UI4 | | 밑줄 색 (COLORREF)<br>RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR) |
| ShadowOffsetX | PIT_I1 | | 그림자 간격 (X 방향) -100% - 100% |
| ShadowOffsetY | PIT_I1 | | 그림자 간격 (Y 방향) -100% - 100% |
| ShadowColor | PIT_UI4 | | 그림자 색 (COLORREF)<br>RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR) |
| StrikeOutType | PIT_UI1 | | 취소선 종류 :<br>0 = none, 1 = red single, 2 = red double,<br>3 = text single, 4 = text double |
| DiacSymMark | PIT_UI1 | | 강조점 종류 :<br>0 = none, 1 = 검정 동그라미, 2 = 속 빈 동그라미 |
| UseFontSpace | PIT_UI1 | | 글꼴에 어울리는 빈칸 : 0 = off, 1 = on |
| UseKerning | PIT_UI1 | | 커닝 : 0 = off, 1 = on |
| Height | PIT_I4 | | 글자 크기 (HWPUNIT) |
| StrikeOutShape | PIT_UI1 | | 취소선 모양 (아래 표 참조) |
| StrikeOutColor | PIT_UI4 | | 취소선색 (COLORREF) |
| BorderFill | PIT_SET | BorderFill | 테두리/배경 |

#### StrikeOutShape 값

| 값 | 설명 |
|---|------|
| 0 | 실선 |
| 1 | 파선 |
| 2 | 점선 |
| 3 | 일점쇄선 |
| 4 | 이점쇄선 |
| 5 | 긴 파선 |
| 6 | 원형 점선 |
| 7 | 이중 실선 |
| 8 | 얇고 굵은 이중선 |
| 9 | 굵고 얇은 이중선 |
| 10 | 얇고 굵고 얇은 삼중선 |
| 11 | 물결선 |
| 12 | 이중 물결선 |
| 13 | 3D 굵은선 |
| 14 | 3D 굵은선 (광원 반대) |
| 15 | 3D 실선 |
| 16 | 3D 실선 (광원 반대) |

---

### 14) ChCompose : 글자 겹침

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Chars | PIT_BSTR | | 겹쳐질 글자들 |
| CircleType | PIT_UI | | 테두리 타입 (아래 표 참조) |
| CharSize | PIT_I | | 테두리 내부 문자의 크기 비율<br>-6 ~ +2 (40%~120%) |
| CheckCompose | PIT_UI1 | | 테두리 내부의 문자 겹치기 여부 TRUE=겹치기:FALSE=펼치기 |
| CharShapes | PIT_SET | ChComposeShapes | 겹쳐지는 문자들의 속성셋 |
| TextDir | PIT_UI1 | | 텍스트 방향 |

#### CircleType 값

| 값 | 설명 |
|---|------|
| 1 | ○ |
| 3 | □ |
| 5 | △ |
| 8 | ◇ |

---

### 15) ChComposeShapes : 글자 겹치기 글자 속성셋

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| CircleCharShape | PIT_SET | CharShape | 겹쳐지는 문자들의 속성. |
| InnerCharShape1-9 | PIT_SET | CharShape | 겹쳐지는 문자들의 속성 1~9 |
