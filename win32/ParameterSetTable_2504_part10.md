# Parameter Set Table - Part 10

---

## 82) MemoShape : 메모 모양

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Width | PIT_I4 | | 너비 (HWPUNIT) |
| LineType | PIT_UI1 | | 선 종류 |
| LineWidth | PIT_UI1 | | 선 굵기 |
| LineColor | PIT_UI4 | | 선 색깔 (COLORREF) |
| FillColor | PIT_UI4 | | 채우기 색깔 (COLORREF) |
| ActiveFillColor | PIT_UI4 | | 활성화된 채우기 색깔 (COLORREF) |
| MemoType | PIT_UI4 | | 메모 종류 - 현재 사용안 함<br>1 = 메모 넣기, 2 = 메모 지우기, 3 = 메모 고치기 |

---

## 83) MousePos : 마우스 위치

HwpCtrl.GetMousePos에서 사용, 해당 액션은 존재하지 않음.

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| XRelTo | PIT_UI4 | | 가로 상대적 기준<br>0 : 종이<br>1 : 쪽 |
| YRelTo | PIT_UI4 | | 세로 상대적 기준<br>0 : 종이<br>1 : 쪽 |
| Page | PIT_UI4 | | 페이지 번호 ( 0 based) |
| X | PIT_I4 | | 가로 클릭한 위치 (HWPUNIT) |
| Y | PIT_I4 | | 세로 클릭한 위치 (HWPUNIT) |

---

## 84) MovieProperties : 동영상 파일 삽입 시 필요한 옵션

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Base | PIT_BSTR | | 동영상 파일의 경로 |
| Caption | PIT_BSTR | | 자막 파일의 경로 |
| AutoPlay | PIT_UI1 | | 자동 재생 여부 : 0 = off, 1 = on |
| AutoRewind | PIT_UI1 | | 되감기 여부 : 0 = off, 1 = on |
| ShowMenu | PIT_UI1 | | 메뉴 보이기 : 0 = Hide, 1 = Show |
| ShowCtrlPanel | PIT_UI1 | | 컨트롤 패널 보이기 : 0 = Hide, 1 = Show |
| ShowPosCtrl | PIT_UI1 | | 위치 컨트롤 보이기 : 0 = Hide, 1 = Show |
| EnablePos | PIT_UI1 | | 위치 컨트롤 조절 여부 : 0 = Disable, 1 = Enable |
| ShowTrackBar | PIT_UI1 | | 음량 조절막대(Track Bar) 보이기 : 0 = Hide, 1 = Show |
| EnableTrack | PIT_UI1 | | 음량 조절 여부 : 0 = Disable, 1 = Enable |
| ShowChaption | PIT_UI1 | | 자막 보이기 : 0 = Hide, 1 = Show |
| ShowAudio | PIT_UI1 | | 오디오 설정 보이기 : 0 = Hide, 1 = Show |
| ShowStatus | PIT_UI1 | | 상태바 보기 (진행시간 등을 표시) : 0 = Hide, 1 = Show |

---

## 85) OleCreation : OLE 개체 생성

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Type | PIT_UI1 | | 생성 방식<br>0 = 새로 개체 생성<br>1 = 파일로부터 개체 생성<br>2 = 파일로 링크된 개체 생성 |
| Clsid | PIT_BSTR | | 클래스 ID ('새로 개체 생성'일 때 사용) |
| Path | PIT_BSTR | | 파일 경로<br>('파일로 링크된 개체 생성', '파일로부터 개체 생성'일 때 사용) |
| Aspect | PIT_I | | 생성된 OLE 개체를 아이콘으로 표시할지 여부 :<br>0 = 내용으로 표시, 1 = 아이콘으로 표시 |
| IconMetafile | PIT_BINDATA | | Aspect가 아이콘일 때 적용할 아이콘 데이터 |
| IconMM | PIT_I | | Aspect가 아이콘일 때 아이콘 매핑모드<br>1 = MM_TEXT<br>2 = MM_LOMETRIC<br>3 = MM_HIMETRIC<br>4 = MM_LOENGLISH<br>5 = MM_HIENGLISH<br>6 = MM_TWIPS<br>7 = MM_ISOTROPIC<br>8 = MM_ANISOTROPIC<br>※ MFC의 매핑모드와 값/의미가 동일하다. |
| IconXext | PIT_I | | Aspect가 아이콘일 때 X축 매핑단위 |
| IconYext | PIT_I | | Aspect가 아이콘일 때 Y축 매핑단위 |
| InnerOCX | PIT_I | | 글 내부에서 사용되는 OCX인지 여부 (예: 글의 Chart OLE) |
| SoProperties | PIT_SET | ShapeObject | 초기 shape object 속성 |
| FlashProperties | PIT_SET | FlashProperties | 플래시 파일 삽입 시 필요한 옵션 셋 |
| MovieProperties | PIT_SET | MovieProperties | 동영상 파일 삽입 시 필요한 옵션 셋 |

---

## 86) PageBorderFill : 구역의 쪽 테두리/배경

PageBorderFill은 BorderFill로부터 계승받았으므로 위 표에 정리된 PageBorderFill의 아이템들 이외에 BorderFill의 아이템들을 사용할 수 있다.

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| TextBorder | PIT_UI1 | | TRUE = 본문 기준, FALSE = 종이 기준 |
| HeaderInside | PIT_UI1 | | 머리말 포함 (on / off) |
| FooterInside | PIT_UI1 | | 꼬리말 포함 (on / off) |
| FillArea | PIT_UI1 | | 채울 영역<br>0 = 종이<br>1 = 쪽<br>2 = 테두리 |
| OffsetLeft | PIT_I | | 4방향 간격 (HWPUNIT) : 왼쪽 |
| OffsetRight | PIT_I | | 4방향 간격 (HWPUNIT) : 오른쪽 |
| OffsetTop | PIT_I | | 4방향 간격 (HWPUNIT) : 위 |
| OffsetBottom | PIT_I | | 4방향 간격 (HWPUNIT) : 아래 |
| BorderTypeLeft | PIT_UI2 | | 왼쪽 라인 스타일<br>0 실선<br>1 파선<br>2 점선<br>3 일점쇄선<br>4 이점쇄선<br>5 긴 파선<br>6 원형 점선<br>7 이중 실선<br>8 얇고 굵은 이중선<br>9 굵고 얇은 이중선<br>10 얇고 굵고 얇은 삼중선<br>11 물결선<br>12 이중 물결선<br>13 3D 굵은선<br>14 3D 굵은선 (광원 반대)<br>15 3D 실선<br>16 3D 실선 (광원 반대) |
| BorderTypeRight | PIT_UI2 | | 오른쪽 라인 스타일 |
| BorderTypeTop | PIT_UI2 | | 위쪽 라인 스타일 |
| BorderTypeBottom | PIT_UI2 | | 아래쪽 라인 스타일 |
| BorderWidthLeft | PIT_UI1 | | 왼쪽 테두리 두께<br>0 0.1mm<br>1 0.12mm<br>2 0.15mm<br>3 0.2mm<br>4 0.25mm<br>5 0.3mm<br>6 0.4mm<br>7 0.5mm<br>8 0.6mm<br>9 0.7mm<br>10 1.0mm<br>11 1.5mm<br>12 2.0mm<br>13 3.00mm<br>14 4.00mm<br>15 5.00mm |
| BorderWidthRight | PIT_UI1 | | 오른쪽 테두리 두께 |
| BorderWidthTop | PIT_UI1 | | 위쪽 테두리 두께 |
| BorderWidthBottom | PIT_UI1 | | 아래쪽 테두리 두께 |
| BorderCorlorLeft | PIT_UI4 | | 4방향 테두리 색깔(COLORREF) |
| BorderColorRight | PIT_UI4 | | 4방향 테두리 색깔(COLORREF) |
| BorderColorTop | PIT_UI4 | | 4방향 테두리 색깔(COLORREF) |
| BorderColorBottom | PIT_UI4 | | 4방향 테두리 색깔(COLORREF) |
| SlashFlag | PIT_UI2 | | 슬래쉬 대각선 플랙 (bit 0-2가 시계 방향으로 각각의 대각선 유무를 나타냄 |
| BackSlashFlag | PIT_UI2 | | 백슬래쉬 대각선 플랙 (bit 0-2가 반시계 방향으로 각각의 대각선 유무를 나타냄) |
| DiagonalType | PIT_UI2 | | 대각선 종류<br>선 종류 |
| DiagonalWidth | PIT_UI1 | | 대각선 두께<br>0 0.1mm<br>1 0.12mm<br>2 0.15mm<br>3 0.2mm<br>4 0.25mm<br>5 0.3mm<br>6 0.4mm<br>7 0.5mm<br>8 0.6mm<br>9 0.7mm<br>10 1.0mm<br>11 1.5mm<br>12 2.0mm<br>13 3.00mm<br>14 4.00mm<br>15 5.00mm |
| DiagonalColor | PIT_UI4 | | 대각선 색깔 (COLORREF) |
| BorderFill3D | PIT_UI1 | | 3차원 효과 on/off |
| Shadow | PIT_UI1 | | 그림자 효과 on/off |
| FillAttr | PIT_SET | DrawFillAttr | 배경 채우기 속성 |
| CrookedSlashFlag | PIT_UI2 | | 꺽어진 대각선 플랙 (bit 0, 1이 각각 slash, back slash의 가운데 대각선이 꺽어진 대각선임을 나타낸다.) |
| BreakCellSeparateLine | PIT_UI1 | | 자동으로 나뉜 표의 경계선 설정 TRUE/FALSE |
| ShadeFaceColorIncDec | PIT_UI1 | | 음영 비율 증가/감소.<br>음영 비율 증가/감소 엑션에서 전달 됨. 구현용으로만 쓰임.<br>이 아이템이 없으면(음영 비율 증가/감소는 없는 것이고 있다면 값이 TRUE이면 증가, FALSE이면 감소이다.) |
| CounterSlashFlag | PIT_UI1 | | 슬래쉬 대각선의 역방향 플랙(우상향 대각선) |
| CounterBackSlashFlag | PIT_UI1 | | 역슬래쉬 대각선의 역방향 플랙(좌상향 대각선) |
| CenterLineFlag | PIT_UI1 | | 중심선 |
| CrookedSlashFlag1 | PIT_UI1 | | 상향 꺾어진 대각선 |
| CrookedSlashFlag2 | PIT_UI1 | | 하향 꺽어진 대각선 |

---

## 87) PageDef : 구역 내의 용지 설정 속성

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| PaperWidth | PIT_I4 | | 용지 가로 크기 (HWPUNIT) |
| PaperHeight | PIT_I4 | | 용지 세로 크기 (HWPUNIT) |
| Landscape | PIT_I4 | | 용지 방향. 0 = 좁게, 1 = 넓게 |
| LeftMargin | PIT_I4 | | 왼쪽 여백 (HWPUNIT) |
| RightMargin | PIT_I4 | | 오른쪽 여백 (HWPUNIT) |
| TopMargin | PIT_I4 | | 위 여백 (HWPUNIT) |
| BottomMargin | PIT_I4 | | 아래 여백 (HWPUNIT) |
| HeaderLen | PIT_I4 | | 머리말 길이 (HWPUNIT) |
| FooterLen | PIT_I4 | | 꼬리말 길이 (HWPUNIT) |
| GutterLen | PIT_I4 | | 제본 여백 (HWPUNIT) |
| GutterType | PIT_UI1 | | 편집 방법.<br>0 = 한쪽 편집, 1 = 맞쪽 편집, 2 = 위로 넘기기 |
| ApplyTo | PIT_UI1 | | 적용범위<br>0 = 선택된 구역<br>1 = 선택된 문자열<br>2 = 현재 구역<br>3 = 문서전체<br>4 = 새 구역 : 현재 위치부터 새로 |
| ApplyClass | PIT_UI1 | | 적용범위의 분류 (대화상자를 호출할 경우 사용)<br>0x01 = 선택된 구역<br>0x02 = 선택된 문자열<br>0x04 = 현재 구역<br>0x08 = 문서전체<br>0x10 = 새 구역 : 현재 위치부터 새로 |

### Example : 편집용지 설정 예제

**C++**

```cpp
{
    DHwpAction dact = m_cHwpCtrl.CreateAction("PageSetup");
    DHwpParameterSet dset = dact.CreateSet();
    dact.GetDefault(dset);

    dset.SetItem ("ApplyTo", (COleVariant)(long)3);

    DHwpParameterSet _dset = dset.CreateItemSet ("PageDef", "PageDef");

    // 1mm = 283.465 HWPUNITs
    _dset.SetItem ("TopMargin", (COleVariant)(long)5668);
    _dset.SetItem ("BottomMargin", (COleVariant)(long)5668);
    _dset.SetItem ("LeftMargin", (COleVariant)(long)2834);
    _dset.SetItem ("RightMargin", (COleVariant)(long)2834);
    _dset.SetItem ("HeaderLen", (COleVariant)(long)1417);
    _dset.SetItem ("FooterLen", (COleVariant)(long)1417);
    _dset.SetItem ("GutterLen", (COleVariant)(long)0);

    dact.Execute(dset);
}
```
