# ParameterSet Table

최종 수정일 : 2025년 4월 15일

## 1. ParameterSet

### 1) ActionCrossRef : 상호참조 삽입

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Command | PIT_BSTR | | ※command string 참조 |

---

### 2) AutoFill : 자동 채우기

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| AutoFillSection | PIT_UI | | 자동 채우기 섹션 : 0 = 기본, 1 = 사용자 정의 |
| AutoFillItem | PIT_UI | | 섹션의 아이템 인덱스 : 0 ~ |

---

### 3) AutoNum : 번호 넣기

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| NumType | PIT_UI1 | | 번호 종류 :<br>0 = 쪽 번호<br>1 = 각주 번호<br>2 = 미주 번호<br>3 = 그림 번호<br>4 = 표 번호<br>5 = 수식 번호 |
| NewNumber | PIT_UI2 | | 새 시작 번호 (1 .. n) |
| NumFormat | PIT_UI2 | | 번호 모양 |

---

### 4) BookMark : 책갈피

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| Name | PIT_BSTR | | 책갈피 이름 |
| Type | PIT_UI | | 책갈피 종류 : 0 = 일반 책갈피, 1 = 블록 책갈피 |
| Command | PIT_UI | | 책갈피 명령의 종류 :<br>0 = 책갈피 생성, 1 = 책갈피로 이동, 2 = 책갈피 수정 |

---

### 5) BorderFill : 테두리/배경의 일반 속성

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| BorderTypeLeft | PIT_UI2 | | 4방향 테두리 종류 : 왼쪽 [선 종류] |
| BorderTypeRight | PIT_UI2 | | 4방향 테두리 종류 : 오른쪽 |
| BorderTypeTop | PIT_UI2 | | 4방향 테두리 종류 : 위 |
| BorderTypeBottom | PIT_UI2 | | 4방향 테두리 종류 : 아래 |
| BorderWidthLeft | PIT_UI1 | | 4방향 테두리 두께 : 왼쪽 [선 종류] |
| BorderWidthRight | PIT_UI1 | | 4방향 테두리 두께 : 오른쪽 |
| BorderWidthTop | PIT_UI1 | | 4방향 테두리 두께 : 위 |
| BorderWidthBottom | PIT_UI1 | | 4방향 테두리 두께 : 아래 |
| BorderColorLeft | PIT_UI4 | | 4방향 테두리 색깔 : 왼쪽<br>RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR) |
| BorderColorRight | PIT_UI4 | | 4방향 테두리 색깔 : 오른쪽<br>RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR) |
| BorderColorTop | PIT_UI4 | | 4방향 테두리 색깔 : 위<br>RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR) |
| BorderColorBottom | PIT_UI4 | | 4방향 테두리 색깔 : 아래<br>RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR) |
| SlashFlag | PIT_UI2 | | 슬래쉬 대각선 플래그 :<br>비트 플래그의 조합으로 표현되며 각 위치의 비트는 다음을 나타낸다.<br>bit 0 = 하단 대각선<br>bit 1 = 중앙 대각선<br>bit 2 = 상단 대각선<br>더 자세한 내용은 하단의 ※ 대각선의 형태를 참고한다. |
| BackSlashFlag | PIT_UI2 | | 백슬래쉬 대각선 플래그 :<br>비트 플래그의 조합으로 표현되며 각 위치의 비트는 다음을 나타낸다.<br>bit 0 = 하단 대각선<br>bit 1 = 중앙 대각선<br>bit 2 = 상단 대각선<br>더 자세한 내용은 하단의 ※ 대각선의 형태를 참고한다. |
| DiagonalType | PIT_UI2 | | 선 종류<br>셀에서는 대각선, 표에서는 자동으로 나눠진 경계선에서 사용 |
| DiagonalWidth | PIT_UI1 | | 선 종류<br>셀에서는 대각선, 표에서는 자동으로 나눠진 경계선에서 사용 |
| DiagonalColor | PIT_UI4 | | 선 색상<br>RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR)<br>셀에서는 대각선, 표에서는 자동으로 나눠진 경계선에서 사용 |
| BorderFill3D | PIT_UI1 | | 3차원 효과 : 0 = off, 1 = on |
| Shadow | PIT_UI1 | | 그림자 효과 : 0 = off, 1 = on |
| FillAttr | PIT_SET | DrawFillAttr | 배경 채우기 속성 |
| CrookedSlashFlag | PIT_UI2 | | 꺾인 대각선 플래그 (bit 0, 1이 각각 slash, back slash의 가운데 대각선이 꺾인 대각선임을 나타낸다.) |
| BreakCellSeparateLine | PIT_UI1 | | 자동으로 나뉜 표의 경계선 설정 :<br>0 = 경계선 설정을 기본 값에 따름, 1 = 사용자가 경계선 설정 |
| ShadeFaceColorIncDec | PIT_UI1 | | 음영 비율 증가/감소. 음영 비율 증가/감소 액션에서 전달 됨. 구현용으로만 쓰임.<br>이 아이템이 없으면(음영 비율 증가/감소는 없는 것이고 있다면 값이 TRUE이면 증가, FALSE이면 감소이다.) |
| CounterSlashFlag | PIT_UI1 | | 슬래쉬 대각선의 역방향 플래그(우상향 대각선) :<br>0 = 순방향, 1 = 역방향 |
| CounterBackSlashFlag | PIT_UI1 | | 역슬래쉬 대각선의 역방향 플래그(좌상향 대각선) :<br>0 = 순방향, 1 = 역방향 |
| CenterLineFlag | PIT_UI1 | | 중심선 : ( 0 = 중심선 없음, 1 = 중심선 있음)<br>더 자세한 내용은 하단의 ※ 중심선의 형태를 참고한다. |
| CrookedSlashFlag1 | PIT_UI1 | | Low Byte CrookedSlashFlag (슬래쉬 대각선의 꺾임 여부)<br>(CrookedSlashFlag를 쓰기 편하도록 CrookedSlashFlag1,2로 나눔) |
| CrookedSlashFlag2 | PIT_UI1 | | High Byte CrookedSlashFlag (역슬래쉬 대각선의 꺾임 여부)<br>(CrookedSlashFlag를 쓰기 편하도록 CrookedSlashFlag1,2로 나눔) |

#### ※ 대각선의 형태

대각선의 형태는 다음의 3가지의 아이템을 통해서 결정된다.

- SlashFlag(BackSlashFlag) -> 괄호 안은 역슬래쉬의 경우
- CrookedSlashFlag1(CrookedSlashFlag2)
- CounterSlashFlag(CounterBackSlashFlag)

**SlashFlag**는 대각선의 유형을 나타낸다. (BackSlashFlag는 반대방향)

**CrookedSlashFlag**는 대각선이 꺾임 여부를 나타내며, 오직 SlashFlag(BackSlashFlag)가 0x02인 경우에만 유효하다.

**CounterSlashFlag**는 역방향을 나타낸다.

**Example : 현재 선택된 셀에 꺾인 대각선 넣기**

```javascript
function OnInsertCrookedSlash()
{
    var vAct = vHwpCtrl.CreateAction("CellBorder");
    var vSet = act.CreateSet(); // Create CellBorder ParameterSet (drived BorderFill)
    vAct.GetDefault(vSet);

    vSet.SetItem("DiagonalType", 1);      // Line type is Solid
    vSet.SetItem("BackSlashFlag", 0x02);  // One Line Back-Slash
    vSet.SetItem("CrookedSlashFlag2", 1); // Slash is Crooked

    vAct.Execute(vSet);
}
```

#### ※ 중심선의 형태

중심선의 형태는 다음의 2가지의 아이템을 통해서 결정된다.

- CenterLineFlag
- CrookedSlashFlag1, CrookedSlashFlag2

**CenterLineFlag**는 단순히 중심선을 설정할 것인지 설정하지 않을 것인지를 나타낸다. CenterLineFlag가 1로 설정되었다면, 실제로 중심선의 가로 또는 세로를 설정하는 값은 CrookedSlashFlag가 결정하게 된다.

- CrookedSlashFlag1 = 가로 중심선 유/무(1/0)
- CrookedSlashFlag2 = 세로 중심선 유/무(1/0)

가로세로 모두 설정하려면 CrookedSlashFlag1/2 모두 값을 1로 설정하면 된다. 중심선과 대각선은 서로 배타적으로 적용된다. (대각선이 설정되면 중심선은 취소되며, 중심선이 설정되면 대각선은 취소된다.)

**Example : 현재 선택된 셀에 중심선 넣기**

```javascript
function OnInsertCrookedSlash()
{
    var vAct = vHwpCtrl.CreateAction("CellBorder");
    var vSet = act.CreateSet(); // Create CellBorder ParameterSet (drived BorderFill)
    vAct.GetDefault(vSet);

    vSet.SetItem("DiagonalType", 1);      // Line type is Solid
    vSet.SetItem("CenterLineFlag", 1);    // Set CenterLine
    vSet.SetItem("CrookedSlashFlag2", 1); // Vertical CenterLine

    vAct.Execute(vSet);
}
```

---

### 6) BorderFillExt : UI 구현을 위한 BorderFill 확장

BorderFillExt는 BorderFill로부터 계승받았으므로 위 표에 정리된 BorderFillExt의 아이템들 이외에 BorderFill의 아이템들을 사용할 수 있다.

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| TypeHorz | PIT_UI2 | | 중앙선 종류 : 가로 [선 종류] |
| TypeVert | PIT_UI2 | | 중앙선 종류 : 세로 |
| WidthHorz | PIT_UI1 | | 중앙선 두께 : 가로 [선 종류] |
| WidthVert | PIT_UI1 | | 중앙선 두께 : 세로 |
| ColorHorz | PIT_UI4 | | 중앙선 색깔 : 가로<br>RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR) |
| ColorVert | PIT_UI4 | | 중앙선 색깔 : 세로<br>RGB color를 나타내기 위한 32비트 값 (0x00BBGGRR) |

---

### 7) BulletShape : 불릿 모양(글머리표 모양)

| Item ID | Type | SubType | Description |
|---------|------|---------|-------------|
| HasCharShape | PIT_UI1 | | 자체적인 글자 모양을 사용할지 여부 :<br>0 = 스타일을 따라감, 1 = 자체 모양을 가짐 |
| CharShape | PIT_SET | CharShape | 글자 모양 (HasCharShape가 1일 경우에만 사용) |
| WidthAdjust | PIT_I | | 번호 너비 보정 값 (HWPUNIT) |
| TextOffset | PIT_I | | 본문과의 거리 (percent or HWPUNIT) |
| Alignment | PIT_UI1 | | 번호 정렬 :<br>0 = 왼쪽 정렬, 1 = 가운데 정렬, 2 = 오른쪽 정렬 |
| UseInstWidth | PIT_UI1 | | 번호 너비를 문서 내 문자열의 너비에 따를지 여부 on / off |
| AutoIndent | PIT_UI1 | | 번호 너비 자동 들여쓰기 여부 : 0 = 들여쓰기 안함, 1 = 들여쓰기 |
| TextOffsetType | PIT_UI1 | | 본문과의 거리 종류 : 0 = percent, 1 = HWPUNIT |
| BulletChar | PIT_UI2 | | 불릿 문자 코드 |
| HasImage | PIT_UI1 | | 그림 글머리표 여부 : 0 = 일반 글머리표, 1 = 그림 글머리표 |
| BulletImage | PIT_SET | DrawFillAttr | 글머리표 이미지 |
| CheckedBulletChar | PIT_UI2 | | 체크된 불릿 문자 코드 |
| Checkable | PIT_UI1 | | 체크 가능 여부 |
