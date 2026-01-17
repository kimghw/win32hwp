# Action Table

**최종 수정일 : 2025년 4월 15일**

---

## 범례

- Action ID 중 **빨간색 밑줄 친 ID**는 Dummy Action으로서 내부적인 용도로만 쓰일 뿐 실제로 구현되지 않은 액션입니다.
- Action ID 중 *기울임 밑줄 친 ID*는 글 컨트롤이 글 내부의 기능을 모방한 액션으로써 글 내부에는 구현되어 있지 않지만, 필요에 의해 글 컨트롤에서 구현한 액션입니다.

### ParameterSet ID 기호 설명

| Symbol | Description |
|--------|-------------|
| `-` | ParameterSet 없음. 해당 Action에 대한 ParameterSet이 존재하지 않거나, 필요없음을 나타냄. |
| `+` | 추가 예정. 내부적으로는 ParameterSet을 사용하고 있지만, 사용하고 있는 ParameterSet이 아직 외부로 노출되지 않은 경우. |
| `*` | HwpCtrl.Run 불가능. ParameterSet이 있지만, 외부에서 ParameterSet을 만들어 주지 않으면 정상작동하지 않는 Action. 다른 Action에 종속된 Action이거나, DocSummaryInfo와 같이 값을 읽어오기만 하는 Action일 경우 해당. |

---

## Action 목록

| Action ID | ParameterSet ID | Description | 비고 |
|-----------|-----------------|-------------|-----|
| AddHanjaWord | + | 한자단어 등록 | |
| AllReplace | FindReplace* | 모두 바꾸기 | |
| AQcommandMerge | UserQCommandFile* | 입력 자동 명령 파일 저장/로드 (글메뉴의 [도구-빠른 교정-빠른 교정내용]에서 [입력 자동 명령 사용자 사전] 대화상자) | ParameterSet을 직접 조작하여 사용함. |
| AutoChangeHangul | - | 낱자모 우선 | |
| AutoChangeRun | - | 동작 | |
| AutoSpell Run | - | 맞춤법 - 메뉴에서 맞춤법 도우미 동작 On/Off | |
| AutoSpellSelect1 ~ 16 | - | 맞춤법 도우미를 통해 선택된 어휘로 변경 (어휘는 1에서 최대 16까지) | |
| Average | Sum | 블록 평균 | |
| BackwardFind | FindReplace* | 뒤로 찾기 | |
| Bookmark | BookMark | 책갈피 | |
| BookmarkEditDialog | | 북마크 편집 대화상자 호출 액션 - 책갈피 작업창에서 편집 대화상자를 호출하기 위한 액션 | |
| BreakColDef | - | 단 정의 삽입 | |
| BreakColumn | - | 단 나누기 | |
| BreakLine | - | line break | |
| BreakPage | - | 쪽 나누기 | |
| BreakPara | - | 문단 나누기 | |
| BreakSection | - | 구역 나누기 | |
| BulletDlg | ParaShape | 글머리표 대화상자 | |
| Cancel | - | ESC | |
| CaptionPosBottom | ShapeObject | 캡션 위치-아래 | |
| CaptionPosLeftBottom | ShapeObject | 캡션 위치-왼쪽 아래 | |
| CaptionPosLeftCenter | ShapeObject | 캡션 위치-왼쪽 가운데 | |
| CaptionPosLeftTop | ShapeObject | 캡션 위치-왼쪽 위 | |
| CaptionPosRightBottom | ShapeObject | 캡션 위치-오른쪽 아래 | |
| CaptionPosRightCenter | ShapeObject | 캡션 위치-오른쪽 가운데 | |
| CaptionPosRightTop | ShapeObject | 캡션 위치-오른쪽 위 | |
| CaptionPosTop | ShapeObject | 캡션 위치-위 | |
| CaptureDialog | - | 갈무리 끝 | |
| CaptureHandler | - | 갈무리 시작 | |
| CellBorder | CellBorderFill | 셀 테두리 | |
| CellBorderFill | CellBorderFill | 셀 테두리 | |
| CellFill | CellBorderFill | 셀 배경 | |
| CellZoneBorder | CellBorderFill | 셀 테두리(여러 셀에 걸쳐 적용) | 셀이 선택된 상태에서만 동작 |
| CellZoneBorderFill | CellBorderFill | 셀 테두리(여러 셀에 걸쳐 적용) | 셀이 선택된 상태에서만 동작 |
| CellZoneFill | CellBorderFill | 셀 배경(여러 셀에 걸쳐 적용) | 셀이 선택된 상태에서만 동작 |
| ChangeImageFileExtension | SummaryInfo | 연결 그림 확장자 바꾸기 | |
| ChangeObject | ShapeObject | 개체 변경하기 | |
| ChangeRome String | + | 로마자변환 - 입력받은 스트링 변환 | |
| ChangeRome User String | + | 로마자 사용자 데이터 추가 | |
| ChangeRome User | + | 로마자 사용자 데이터 | |
| ChangeRome | + | 로마자변환 | |
| CharShape | CharShape | 글자 모양 | |
| CharShapeBold | - | 단축키: Alt+L(글자 진하게) | |
| CharShapeCenterline | - | 취소선(CenterLine) | |
| CharShapeDialog | CharShape | 글자 모양 대화상자(내부 구현용) | |
| CharShapeDialogWithoutBorder | CharShape | 글자 모양 대화상자(내부 구현용, [글자 테두리] 탭 제외) | |
| CharShapeEmboss | - | 양각 | |
| CharShapeEngrave | - | 음각 | |
| CharShapeHeight | - | 글자 크기(글자 모양 대화상자에서 Focus이동용으로 사용) | |
| CharShapeHeightDecrease | - | 크기 작게 ALT+SHIFT+R | |
| CharShapeHeightIncrease | - | 크기 크게 ALT+SHIFT+E | |
| CharShapeItalic | - | 이탤릭 ALT + SHIFT + I | |
| CharShapeLang | - | 글꼴 언어(글자 모양 대화상자에서 Focus이동용으로 사용) | |
| CharShapeNextFaceName | - | 다음 글꼴 ALT+SHIFT+F | |
| CharShapeNormal | - | 보통모양 ALT+SHIFT+C | |
| CharShapeOutline | - | 외곽선 | |
| CharShapePrevFaceName | - | 이전 글꼴 ALT+SHIFT+G | |
| CharShapeShadow | - | 그림자 | |
| CharShapeSpacing | - | 자간(글자 모양 대화상자에서 Focus이동용으로 사용) | |
| CharShapeSpacingDecrease | - | 자간 좁게 ALT+SHIFT+N | |
| CharShapeSpacingIncrease | - | 자간 넓게 ALT+SHIFT+W | |
| CharShapeSubscript | - | 아래첨자 ALT+SHIFT+S | |
| CharShapeSuperscript | - | 위첨자 ALT+SHIFT+P | |
| CharShapeSuperSubscript | - | 첨자 : "위첨자 -> 아래첨자 -> 보통" 반복 | |
| CharShapeTextColorBlack | - | 글자색을 검정 | |
| CharShapeTextColorBlue | - | 글자색을 파랑 | |
| CharShapeTextColorBluish | - | 글자색을 청록 | |
| CharShapeTextColorGreen | - | 글자색을 초록 | |
| CharShapeTextColorRed | - | 글자색을 빨강 | |
| CharShapeTextColorViolet | - | 글자색을 자주 | |
| CharShapeTextColorWhite | - | 글자색을 흰색 | |
| CharShapeTextColorYellow | - | 글자색을 노랑 | |
| CharShapeTypeFace | - | 글꼴 이름(글자 모양 대화상자에서 Focus이동용으로 사용) | |
| CharShapeUnderline | - | 밑줄 ALT+SHIFT+U | |
| CharShapeWidth | - | 장평(글자 모양 대화상자에서 Focus이동용으로 사용) | |
| CharShapeWidthDecrease | - | 장평 좁게 ALT+SHIFT+J | |
| CharShapeWidthIncrease | - | 장평 넓게 ALT+SHIFT+K | |
| Close | - | 현재 리스트를 닫고 상위 리스트로 이동. | |
| CloseEx | - | 현재 리스트를 닫고 상위 리스트로 이동. Close의 확장 액션으로 전체화면 보기일 때 Root list로 빠져나온다. Shift+Esc | |
| Comment | - | 숨은 설명 | |
| CommentDelete | - | 숨은 설명 지우기 | |
| CommentModify | - | 숨은 설명 고치기 | |
| CompatibleDocument | CompatibleDocument | 호환 문서 | |
| ComposeChars | ChCompose | 글자 겹침 | |
| ConvertBrailleSetting | BrailleConvert | | |
| ConvertCase | ConvertCase | 대소문자 바꾸기 | |
| ConvertFullHalfWidth | ConvertFullHalf | 전각 반각 바꾸기 | |
| ConvertHiraGata | ConvertHiraToGata | 일어 바꾸기 | |
| ConvertJianFan | ConvertJianFan | 간/번체 바꾸기 | Text가 선택된 상태에서만 동작 |
| ConvertOptGugyulToHangul | ConvertToHangul | 한글로 옵션 - 구결을 한글로 | |
| ConvertOptHanjaToHangul | ConvertToHangul | 한글로 옵션 - 漢字를 한글로 | |
| ConvertOptHanjaToHanjaHangul | ConvertToHangul | 한글로 옵션 - 漢字를 漢字(한글)로 | |
| ConvertToBraille | BrailleConvert | 점자 변환 | |
| ConvertToBrailleSelected | BrailleConvert | | |
| ConvertToHangul | ConvertToHangul | 한글로 | |
| Copy | - | 복사하기 | |
| CopyPage | | 쪽 복사하기 | |
| Cut | - | 오려두기 | |
| Delete | - | Delete | |
| DeleteBack | - | Backspace | |
| DeleteCtrls | DeleteCtrls | 조판 부호 지우기 | |
| DeleteDocumentMasterPage | - | 문서 전체 바탕쪽 삭제 | |
| DeleteDutmal | + | 덧말 지우기 | |
| DeleteField | - | 누름틀/메모지우기 누름틀/메모 안의 내용은 지우지 않고, 단순히 누름틀만 지운다. | |
| DeleteFieldMemo | - | 메모 지우기 | |
| DeleteLine | - | CTRL-Y (한줄 지우기) | |
| DeleteLineEnd | - | ALT-Y (현재 커서에서 줄 끝까지 지우기) | |
| DeletePage | DeletePage | 쪽 지우기 | |
| DeletePrivateInfoMark | - | 개인 정보 감추기한 정보 다시보기 | |
| DeletePrivateInfoMarkAtCurrentPos | - | 현재 캐럿 위치의 감추기한 개인 정보 다시 보기 | |
| DeleteSectionMasterPage | - | 구역 바탕쪽 삭제 | |
| DeleteWord | - | 단어 지우기 CTRL-T | |
| DeleteWordBack | - | CTRL-BS(Back Space) | |
| DocFindEnd | FindReplace* | 문서 찾기 종료 | |
| DocFindInit | FindReplace* | 문서 찾기 초기화 | |
| DocFindNext | DocFindInfo* | 문서 찾기 계속 | |
| DocSummaryInfo | SummaryInfo | 문서 정보 | |
| DocumentInfo | DocumentInfo* | 현재 문서에 대한 정보 | |
| DocumentSecurity | DocSecurity | 문서 보안 설정 | |
| DrawObjCancelOneStep | - | 다각형(곡선) 그리는 중 이전 선 지우기 | |
| DrawObjCreatorArc | ShapeObject | 호 그리기 | |
| DrawObjCreatorCanvas | ShapeObject | 캔버스 그리기 | |
| DrawObjCreatorCurve | ShapeObject | 곡선 그리기 | |
| DrawObjCreatorEllipse | ShapeObject | 원 그리기 | |
| DrawObjCreatorFreeDrawing | ShapeObject | 펜 | |
| DrawObjCreatorHorzTextBox | ShapeObject | 가로 글상자 만들기 | |
| DrawObjCreatorLine | ShapeObject | 선 그리기 | |
| DrawObjCreatorMultiArc | ShapeObject | 반복해서 호 그리기 | |
| DrawObjCreatorMultiCanvas | ShapeObject | 반복해서 캔버스 그리기 | |
| DrawObjCreatorMultiCurve | ShapeObject | 반복해서 곡선 그리기 | |
| DrawObjCreatorMultiEllipse | ShapeObject | 반복해서 원 그리기 | |
| DrawObjCreatorMultiFreeDrawing | ShapeObject | 반복해서 펜 그리기 | |
| DrawObjCreatorMultiLine | ShapeObject | 반복해서 선 그리기 | |
| DrawObjCreatorMultiPolygon | ShapeObject | 반복해서 다각형 그리기 | |
| DrawObjCreatorMultiRectangle | ShapeObject | 반복해서 사각형 그리기 | |
| DrawObjCreatorMultiTextBox | ShapeObject | 반복해서 글상자 그리기 | |
| DrawObjCreatorObject | ShapeObject | 그리기 개체 | |
| DrawObjCreatorPolygon | ShapeObject | 다각형 그리기 | |
| DrawObjCreatorRectangle | ShapeObject | 사각형 그리기 | |
| DrawObjCreatorTextBox | ShapeObject | 글상자 | |
| DrawObjCreatorVertTextBox | ShapeObject | 세로 글상자 만들기 | |
| DrawObjEditDetail | - | 그리기 개체 편집 | |
| DrawObjOpenClosePolygon | - | 다각형 열기/닫기 | |
| DrawObjTemplateLoad | ShapeObject | 그리기 마당에서 불러오기 | |
| DrawObjTemplateSave | - | 그리기 마당에 등록 | |
| DrawShapeObjShadow | ShapeObject | 그리기 개체 그림자 만들기/지우기 | 개체가 선택된 상태에서만 동작 |
| DropCap | DropCap | 문단 첫 글자 장식 | |
| DutmalChars | Dutmal | 덧말 넣기 | |
| EditFieldMemo | - | 메모 내용 편집 | |
| EditParaDown | | 문단 옮기기 | |
| EditParaUp | | 문단 옮기기 | |
| EndnoteEndOfDocument | SecDef | 미주-문서의 끝 | |
| EndnoteEndOfSection | SecDef | 미주-구역의 끝 | |
| EndnoteToFootnote | ExchangeFootnoteEndNote | 모든 미주를 각주로 | |
| EquationCreate | EqEdit | 수식 만들기 | |
| EquationModify | EqEdit | 수식 편집하기 | |
| EquationPropertyDialog | ShapeObject | 수식 개체 속성 고치기 | |
| Erase | - | 지우기 | |
| ExchangeFootnoteEndnote | ExchangeFootnoteEndNote | 각주/미주 변환 | |
| ExecReplace | FindReplace* | 바꾸기(실행) | |
| ExtractImagesFromDoc | SummaryInfo | 삽입 그림을 연결 그림으로 추출 | |
| *FileClose* | - | 문서 닫기 | |
| *FileNew* | - | 새문서 | |
| *FileOpen* | - | 파일 열기 | |
| *FileOpenMRU* | - | 최근 작업 문서 | |
| FilePassword | Password | 문서 암호 | |
| FilePasswordChange | Password | 문서 암호 변경 및 해제 | |
| FilePreview | - | 미리 보기 | |
| *FileQuit* | - | 끝 | |
| FileRWPasswordChange | Password | 문서 열기/쓰기 암호 설정 | |
| FileRWPasswordNew | Password | 문서 열기/쓰기 암호 설정 | |
| *FileSave* | - | 파일 저장 | |
| *FileSaveAs* | - | 다른 이름으로 저장 | |
| FileSaveAsImage | Print | 그림 포맷으로 저장하기 | |
| FileSaveAsImageOption | Print | 그림 포맷으로 저장할 때 옵션 설정하기 | |
| FileSaveOptionDlg | | 저장 옵션 | |
| FileSetSecurity | FileSetSecurity* | 배포용 문서(출판 전용 문서). 대화상자를 띄우지 않고 ParameterSet을 직접 설정하여 생성한다. 예제는 "배포용 문서 만들기" 참고 | |
| FileTemplate | FileOpen | 문서마당 | |
| FillColorShadeDec | - | 면 색 음영 비율 감소 | |
| FillColorShadeInc | - | 면 색 음영 비율 증가 | |
| FindAll | FindReplace* | 모두 찾기 | |
| FindDlg | FindReplace | 찾기 | |
| FindForeBackBookmark | - | 앞뒤로 찾아가기 : 책갈피 | |
| FindForeBackCtrl | - | 앞뒤로 찾아가기 : 조판 부호 | |
| FindForeBackFind | - | 앞뒤로 찾아가기 : 찾기 | |
| FindForeBackLine | - | 앞뒤로 찾아가기 : 줄 | |
| FindForeBackPage | - | 앞뒤로 찾아가기 : 페이지 | |
| FindForeBackSection | - | 앞뒤로 찾아가기 : 구역 | |
