# ActionTable_2504_part03

## Page 21

| Action ID | ParameterSet ID | Description | 비고 |
|-----------|-----------------|-------------|-----|
| MacroRepeat | - | 매크로 실행 | |
| MacroRepeatDlg | KeyMacro | 매크로 실행 | |
| MacroStop | - | 매크로 실행 중지 (정의/실행) | |
| MailMergeField | - | 메일 머지 필드(표시달기 or 고치기) | |
| MailMergeGenerate | MailMergeGenerate | 메일 머지 만들기 | |
| MailMergeInsert | FieldCtrl | 메일 머지 표시 달기 | |
| MailMergeModify | FieldCtrl | 메일 머지 고치기 | |
| MakeAllVersionDiffs | VersionInfo | 모든 버전비교 문서 만들기. 현재 문서 및 문서가 가지고 있는 버전정보를 HML 파일로 생성한다.(생성된 파일을 가지고 향 후 버전비교를 실행한다) | |
| MakeContents | MakeContents | 차례 만들기 | |
| MakeIndex | - | 찾아보기 만들기 | |
| ManualChangeHangul | - | 한영 수동 전환. 현재 커서위치 또는 문단나누기 이전에 입력된 내용에 대해서 강제적으로 한/영 전환을 한다. | |
| ManuScriptTemplate | FileOpen | 원고지 쓰기 | |
| MarkPenDelete | | 형광팬 삭제 | |
| MarkPenNext | | 형광팬 이동(다음) | |

## Page 22

| Action ID | ParameterSet ID | Description | 비고 |
|-----------|-----------------|-------------|-----|
| MarkPenPrev | | 형광팬 이동(이전) | |
| MarkPenShape | MarkpenShape* | 형광펜, 선택된 Text영역의 배경을 형광펜으로 칠해준다. Run() 실행불가, 반드시 MarkpenShape ParameterSet의 Color 아이템 값을 설정하고 사용해야 함 | |
| MarkPrivateInfo | PrivateInfoSecurity | 개인 정보 즉시 감추기(텍스트 블록 상태,암호화) | |
| MarkTitle | - | 제목 차례 표시([도구-차례/찾아보기-제목 차례 표시]메뉴에 대응) 차례 코드가 삽입되어 나중에 차례 만들기에서 사용할 수 있다.적용여부는 Ctrl+G,C를 이용해 조판부호를 확인하면 알 수 있다. | |
| MasterPage | MasterPage | 바탕쪽 | |
| MasterPageDelete | MasterPage* | 바탕쪽 삭제. 바탕쪽 편집모드일 경우에만 동작한다. | |
| MasterPageDuplicate | - | 기존 바탕쪽과 겹침. 바탕쪽 편집상태가 활성화되어 있으며 [구역 마지막쪽], [구역임의 쪽]일 경우에만 사용 가능하다. | |
| MasterPageEntry | MasterPage | 바탕쪽 편집모드. 바탕쪽이 존재할 때만 편집모드로 변환한다. | |
| MasterPageExcept | - | 첫 쪽 제외 | |
| MasterPageFront | - | 바탕쪽 앞으로 보내기. 바탕쪽 편집모드일 경우에만 동작한다. | |
| MasterPagePrevSection | - | 앞 구역 바탕쪽 사용 | |
| MasterPageToNext | - | 이후 바탕쪽 | |
| MasterPageToPrevious | - | 이전 바탕쪽 | |
| MasterPageTypeDlg | MasterPage* | 바탕쪽 종류 다이얼로그 띄움 | |
| MemoShape | SecDef | 메모 모양([입력-메모-메모 모양]메뉴와 동일함) | |
| MemoToNext | - | 다음 메모 | |
| MemoToPrev | - | 이전 메모 | |
| MessageBox | + | 메시지 박스 | |

## Page 23

| Action ID | ParameterSet ID | Description | 비고 |
|-----------|-----------------|-------------|-----|
| ModifyBookmark | BookMark | 책갈피 고치기 | |
| ModifyComposeChars | - | 고치기 - 글자 겹침 | |
| ModifyCrossReference | ActionCrossRef | 상호 참조 고치기 | |
| ModifyCtrl | - | 고치기 : 컨트롤 | |
| ModifyDutmal | - | 고치기 - 덧말 | |
| ModifyFieldClickhere | InsertFieldTemplate | 누름틀 정보 고치기 | |
| ModifyFieldDate | InsertFieldTemplate | 날짜 필드 고치기 | |
| ModifyFieldDateTime | InputDateStyle | 날짜/시간 넣기 고치기. 반드시 코드로 작성된 "날짜/시간"이어야하며 코드의 앞(누름틀 밖의)에 캐럿이 존재해야만 동작한다. | |
| ModifyFieldPath | InsertFieldTemplate | 문서 경로 필드 고치기 | |
| ModifyFieldSummary | InsertFieldTemplate | 문서 요약 필드 고치기 | |
| ModifyFieldUserInfo | InsertFieldTemplate | 개인 정보 필드 고치기 | |
| ModifyFillProperty | - | 고치기(채우기 속성 탭으로) 만약 Ctrl(ShapeObject,누름틀, 날짜/시간 코드 등)이 선택되지 않았다면 역방향탐색(SelectCtrlReverse)을 이용해서 개체를 탐색한다. 채우기 속성이 없는 Ctrl일 경우에는 첫 번째 탭이 선택된 상태로 고치기 창이 뜬다. | |
| ModifyHyperlink | HyperLink | 하이퍼링크 고치기 | |

## Page 24

| Action ID | ParameterSet ID | Description | 비고 |
|-----------|-----------------|-------------|-----|
| ModifyLineProperty | - | 고치기(선/테두리 속성 탭으로) 만약 Ctrl(ShapeObject,누름틀, 날짜/시간 코드 등)이 선택되지 않았다면 역방향탐색(SelectCtrlReverse)을 이용해서 개체를 탐색한다. 선/테두리 속성이 없는 Ctrl일 경우에는 첫 번째 탭이 선택된 상태로 고치기 창이 뜬다. | |
| ModifyRevision | RevisionDef | 교정 부호 고치기. ModifyFieldDateTime과 마찬가지로 정확히 교정부호(조판 부호)의 앞에 캐럿이 존재해야 실행된다. | |
| ModifyRevisionHyperlink | HyperLink* | 자료 연결(교정 부호) 고치기. ModifyRevision과 마찬가지로 정확히 교정부호(조판 부호)의 앞에 캐럿이 존재해야 실행된다. Run()으로 실행되지 않는다. | |
| ModifySecTextHorz | TextVertical | 가로 쓰기 | |
| ModifySecTextVert | TextVertical | 세로 쓰기(영문 눕힘) | |
| ModifySecTextVertAll | TextVertical | 세로 쓰기(영문 세움) | |
| ModifySection | SecDef | 구역 | |
| ModifyShapeObject | - | 고치기 - 개체 속성 | |
| MoveColumnBegin | - | 단의 시작점으로 이동한다. 단이 없을 경우에는 아무동작도 하지 않는다. 해당 리스트 안에서만 동작한다. | |
| MoveColumnEnd | - | 단의 끝점으로 이동한다. 단이 없을 경우에는 아무동작도 하지 않는다. 해당 리스트 안에서만 동작한다. | |
| MoveDocBegin | - | 문서의 시작으로 이동. 만약 셀렉션을 확장하는 경우에는 LIST_BEGIN/END와 동일하다. 현재 서브 리스트 내에 있으면 빠져나간다. | |
| MoveDocEnd | - | 문서의 끝으로 이동. 만약 셀렉션을 확장하는 경우에는 LIST_BEGIN/END와 동일하다. 현재 서브 리스트 내에 있으면 빠져나간다. | |

## Page 25

| Action ID | ParameterSet ID | Description | 비고 |
|-----------|-----------------|-------------|-----|
| MoveDown | - | 캐럿을 (논리적 개념의) 아래로 이동시킨다. | |
| MoveLeft | - | 캐럿을 (논리적 개념의) 왼쪽으로 이동시킨다. | |
| MoveLineBegin | - | 현재 위치한 줄의 시작/끝으로 이동 | |
| MoveLineDown | - | 한 줄 아래로 이동한다. | |
| MoveLineEnd | - | 현재 위치한 줄의 시작/끝으로 이동 | |
| MoveLineUp | - | 한 줄 위로 이동한다. | |
| MoveListBegin | - | 현재 리스트의 시작으로 이동 | |
| MoveListEnd | - | 현재 리스트의 끝으로 이동 | |
| MoveNextChar | - | 한 글자 뒤로 이동. 현재 리스트만을 대상으로 동작한다. | |
| MoveNextColumn | - | 뒤 단으로 이동 | |
| MoveNextParaBegin | - | 다음 문단의 시작으로 이동. 현재 리스트만을 대상으로 동작한다. | |
| MoveNextPos | - | 한 글자 뒤로 이동. 서브 리스트를 옮겨 다닐 수 있다. | |
| MoveNextPosEx | - | 한 글자 뒤로 이동. 서브 리스트를 옮겨 다닐 수 있다. (머리말, 꼬리말, 각주, 미주, 글상자 포함) | |
| MoveNextWord | - | 한 단어 뒤로 이동. 현재 리스트만을 대상으로 동작한다. | |
| MovePageBegin | - | 현재 페이지의 시작점으로 이동한다. 만약 캐럿의 위치가 변경되었다면 화면이 전환되어 쪽의 상단으로 페이지뷰잉이 맞춰진다. | |
| MovePageDown | - | 앞 페이지의 시작으로 이동. 현재 탑레벨 리스트가 아니면 탑레벨 리스트로 빠져나온다. | |
| MovePageEnd | - | 현재 페이지의 끝점으로 이동한다. 만약 캐럿의 위치가 변경되었다면 화면이 전환되어 쪽의 하단으로 페이지뷰잉이 맞춰진다. | |
| MovePageUp | - | 뒤 페이지의 시작으로 이동. 현재 탑레벨 리스트가 아니면 탑레벨 리스트로 빠져나온다. | |

## Page 26

| Action ID | ParameterSet ID | Description | 비고 |
|-----------|-----------------|-------------|-----|
| MoveParaBegin | - | 현재 위치한 문단의 시작/끝으로 이동 | |
| MoveParaEnd | - | 현재 위치한 문단의 시작/끝으로 이동 | |
| MoveParentList | - | 한 레벨 상위/탑레벨/루트 리스트로 이동한다. 현재 루트 리스트에 위치해 있어 더 이상 상위 리스트가 없을 때는 위치 이동 없이 리턴한다. 이동한 후의 위치는 상위 리스트에서 서브리스트가 속한 컨트롤 코드가 위치한 곳이다. 위치 이동시 셀렉션은 무조건 풀린다. | |
| MovePrevChar | - | 한 글자 앞 이동. 현재 리스트만을 대상으로 동작한다. | |
| MovePrevColumn | - | 앞 단으로 이동 | |
| MovePrevParaBegin | - | 앞 문단의 시작으로 이동. 현재 리스트만을 대상으로 동작한다. | |
| MovePrevParaEnd | - | 앞 문단의 끝으로 이동. 현재 리스트만을 대상으로 동작한다. | |
| MovePrevPos | - | 한 글자 앞으로 이동. 서브 리스트를 옮겨 다닐 수 있다. | |
| MovePrevPosEx | - | 한 글자 앞으로 이동. 서브 리스트를 옮겨 다닐 수 있다. (머리말, 꼬리말, 각주, 미주, 글상자 포함) | |
| MovePrevWord | - | 한 단어 앞으로 이동. 현재 리스트만을 대상으로 동작한다. | |
| MoveRight | - | 캐럿을 (논리적 개념의) 오른쪽으로 이동시킨다. | |
| MoveRootList | - | 한 레벨 상위/탑레벨/루트 리스트로 이동한다. 현재 루트 리스트에 위치해 있어 더 이상 상위 리스트가 없을 때는 위치 이동 없이 리턴한다. 이동한 후의 위치는 상위 리스트에서 서브리스트가 속한 컨트롤 코드가 위치한 곳이다. 위치 이동시 셀렉션은 무조건 풀린다. | |

## Page 27

| Action ID | ParameterSet ID | Description | 비고 |
|-----------|-----------------|-------------|-----|
| MoveScrollDown | - | 아래 방향으로 스크롤하면서 이동 | |
| MoveScrollNext | - | 다음 방향으로 스크롤하면서 이동 | |
| MoveScrollPrev | - | 이전 방향으로 스크롤하면서 이동 | |
| MoveScrollUp | - | 위 방향으로 스크롤하면서 이동 | |
| MoveSectionDown | - | 뒤 섹션으로 이동. 현재 루트 리스트가 아니면 루트 리스트로 빠져나온다. | |
| MoveSectionUp | - | 앞 섹션으로 이동. 현재 루트 리스트가 아니면 루트 리스트로 빠져나온다. | |
| MoveSelDocBegin | - | 셀렉션: 문서 처음 | |
| MoveSelDocEnd | - | 셀렉션: 문서 끝 | |
| MoveSelDown | - | 셀렉션: 캐럿을 (논리적 방향) 아래로 이동 | |
| MoveSelLeft | - | 셀렉션: 캐럿을 (논리적 방향) 왼쪽으로 이동 | |
| MoveSelLineBegin | - | 셀렉션: 줄 처음 | |
| MoveSelLineDown | - | 셀렉션: 한줄 아래 | |
| MoveSelLineEnd | - | 셀렉션: 줄 끝 | |
| MoveSelLineUp | - | 셀렉션: 한줄 위 | |
| MoveSelListBegin | - | 셀렉션: 리스트 처음 | |
| MoveSelListEnd | - | 셀렉션: 리스트 끝 | |
| MoveSelNextChar | - | 셀렉션: 다음 글자 | |
| MoveSelNextParaBegin | - | 셀렉션: 다음 문단 처음 | |
| MoveSelNextPos | - | 셀렉션: 다음 위치 | |
| MoveSelNextWord | - | 셀렉션: 다음 단어 | |
| MoveSelPageDown | - | 셀렉션: 페이지다운 | |
| MoveSelPageUp | - | 셀렉션: 페이지 업 | |
| MoveSelParaBegin | - | 셀렉션: 문단 처음 | |

## Page 28

| Action ID | ParameterSet ID | Description | 비고 |
|-----------|-----------------|-------------|-----|
| MoveSelParaEnd | - | 셀렉션: 문단 끝 | |
| MoveSelPrevChar | - | 셀렉션: 이전 글자 | |
| MoveSelPrevParaBegin | - | 셀렉션: 이전 문단 시작 | |
| MoveSelPrevParaEnd | - | 셀렉션: 이전 문단 끝 | |
| MoveSelPrevPos | - | 셀렉션: 이전 위치 | |
| MoveSelPrevWord | - | 셀렉션: 이전 단어 | |
| MoveSelRight | - | 셀렉션: 캐럿을 (논리적 방향) 오른쪽으로 이동 | |
| MoveSelTopLevelBegin | - | 셀렉션: 처음 | |
| MoveSelTopLevelEnd | - | 셀렉션: 끝 | |
| MoveSelUp | - | 셀렉션: 캐럿을 (논리적 방향) 위로 이동 | |
| MoveSelViewDown | - | 셀렉션: 아래 | |
| MoveSelViewUp | - | 셀렉션: 위 | |
| MoveSelWordBegin | - | 셀렉션: 단어 처음 | |
| MoveSelWordEnd | - | 셀렉션: 단어 끝 | |
| MoveTopLevelBegin | - | 탑레벨 리스트의 시작으로 이동 | |
| MoveTopLevelEnd | - | 탑레벨 리스트의 끝으로 이동 | |
| MoveTopLevelList | - | 한 레벨 상위/탑레벨/루트 리스트로 이동한다. 현재 루트 리스트에 위치해 있어 더 이상 상위 리스트가 없을 때는 위치 이동 없이 리턴한다. 이동한 후의 위치는 상위 리스트에서 서브리스트가 속한 컨트롤 코드가 위치한 곳이다. 위치 이동시 셀렉션은 무조건 풀린다. | |
| MoveUp | - | 캐럿을 (논리적 개념의) 위로 이동시킨다. | |
| MoveViewBegin | - | 현재 뷰의 시작에 위치한 곳으로 이동 | |

## Page 29

| Action ID | ParameterSet ID | Description | 비고 |
|-----------|-----------------|-------------|-----|
| MoveViewDown | - | 현재 뷰의 크기만큼 아래로 이동한다. PgDn 키의 기능이다. | |
| MoveViewEnd | - | 현재 뷰의 끝에 위치한 곳으로 이동 | |
| MoveViewUp | - | 현재 뷰의 크기만큼 위로 이동한다. PgUp 키의 기능이다. | |
| MoveWordBegin | - | 현재 위치한 단어의 시작으로 이동. 현재 리스트만을 대상으로 동작한다. | |
| MoveWordEnd | - | 현재 위치한 단어의 끝으로 이동. 현재 리스트만을 대상으로 동작한다. | |
| MPBreakNewSection | MasterPage | 새 구역 만들기–바탕쪽 편집 상태에서 | |
| MPCopyFromOtherSection | Masterpage | 바탕쪽 가져오기–다른 구역의 바탕쪽 종류와 내용을 복사 | |
| MPSectionToNext | - | 이후 구역으로 | |
| MPSectionToPrevious | - | 이전 구역으로 | |
| MPShowMarginBorder | - | 여백 보기–바탕쪽 편집 상태에서 | |
| MultiColumn | ColDef | 다단 | |
| NewNumber | AutoNum | 새 번호로 시작 | |
| NewNumberModify | AutoNum | 새 번호 고치기 | |
| NextTextBoxLinked | - | 연결된 글상자의 다음 글상자로 이동 | |
| NoneTextArtShadow | ShapeObject | 글맵시 그림자 없음 | |
| NoteDelete | - | 주석 지우기 | |
| NoteModify | - | 주석 고치기 | |
| NoteNoSuperscript | SecDef | 주석 번호 보통(윗 첨자 사용 안함) | |
| NoteNumProperty | - | 주석 번호 속성 | |

## Page 30

| Action ID | ParameterSet ID | Description | 비고 |
|-----------|-----------------|-------------|-----|
| NoteSuperscript | SecDef | 주석 번호 작게(윗 첨자) | |
| NoteToNext | - | 주석 다음으로 이동 | |
| NoteToPrev | - | 주석 앞으로 이동 | |
| OleCreateNew | OleCreation | 개체 삽입 | |
| OutlineNumber | SecDef | 개요번호 | |
| PageBorder | SecDef | 쪽 테두리/배경 | |
| PageBorderTab | SecDef | 쪽 테두리/배경(항상 테두리 탭이 선택되어 보인다.) | |
| PageFillTab | SecDef | 쪽 테두리/배경(항상 테두리 탭이 선택되어 보인다.) | |
| PageHiding | PageHiding | 감추기 | |
| PageHidingModify | PageHiding | 감추기 고치기 | |
| PageLandscape | SecDef | 용지 넓게 | |
| PageMarginSetup | SecDef | 편집 용지(쪽 여백 설정) | 한/글 2022 부터 지원 |
| PageNumPos | PageNumPos | 쪽 번호 매기기 | |
| PageNumPosModify | PageNumPos | 쪽 번호 매기기 | |
| PagePortrait | SecDef | 용지 좁게 | |
| PageSetup | SecDef | 편집 용지 | |
| PageSetupDL | SecDef | 편집 용지(쪽 여백 설정) | |
| ParagraphShape | ParaShape | 문단 모양 | |
| ParagraphShapeAlignCenter | - | 가운데 정렬 | |
| ParagraphShapeAlignDistribute | - | 배분 정렬 | |
| ParagraphShapeAlignDivision | - | 나눔 정렬 | |
| ParagraphShapeAlignJustify | - | 양쪽 정렬 | |
| ParagraphShapeAlignLeft | - | 왼쪽 정렬 | |
| ParagraphShapeAlignRight | - | 오른쪽 정렬 | |
