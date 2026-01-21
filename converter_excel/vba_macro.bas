' ============================================================
' 엑셀 -> 한글 데이터 전송 VBA 매크로
' ============================================================
' 사용법:
' 1. 엑셀에서 Alt+F11 → VBA 편집기 열기
' 2. 삽입 → 모듈 → 이 코드 붙여넣기
' 3. 도구 → 참조 → "HWP Type Library" 체크 (없으면 찾아보기로 추가)
' 4. 시트에 버튼 추가 → 매크로 연결
' ============================================================

Option Explicit

' 한글 객체 (전역)
Dim hwp As Object

' ------------------------------------------------------------
' 메인: 엑셀 데이터를 한글로 전송
' ------------------------------------------------------------
Sub SendToHwp()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim wsData As Worksheet
    Dim sheetName As String
    Dim lastRow As Long
    Dim i As Long
    Dim listId As Long
    Dim fieldName As String
    Dim cellValue As String
    Dim successCount As Long
    Dim totalCount As Long

    ' 현재 시트 이름에서 메인 시트 이름 추출
    sheetName = ActiveSheet.Name
    If InStr(sheetName, "_") > 0 Then
        sheetName = Left(sheetName, InStr(sheetName, "_") - 1)
    End If

    ' _cells 시트 찾기
    On Error Resume Next
    Set wsData = ThisWorkbook.Sheets(sheetName & "_cells")
    On Error GoTo ErrorHandler

    If wsData Is Nothing Then
        MsgBox sheetName & "_cells 시트를 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If

    ' 메인 시트
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' 한글 연결
    Set hwp = GetHwpInstance()
    If hwp Is Nothing Then
        MsgBox "한글이 실행 중이지 않습니다.", vbExclamation
        Exit Sub
    End If

    ' _cells 시트에서 list_id, row, col, field_name 읽기
    ' 헤더: list_id(1), row(2), col(3), ..., field_name(마지막-1), field_source(마지막)
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    successCount = 0
    totalCount = 0

    For i = 2 To lastRow  ' 헤더 제외
        listId = wsData.Cells(i, 1).Value  ' list_id

        ' 배경색 확인 (bg_color 열 = 14번째)
        If wsData.Cells(i, 14).Value = "" Then
            ' 배경색 없는 셀만 처리 (편집 가능한 셀)
            Dim excelRow As Long, excelCol As Long
            excelRow = wsData.Cells(i, 2).Value + 1  ' row (0-based -> 1-based)
            excelCol = wsData.Cells(i, 3).Value + 1  ' col (0-based -> 1-based)

            ' 메인 시트에서 값 읽기
            cellValue = CStr(ws.Cells(excelRow, excelCol).Value)

            If listId > 0 Then
                totalCount = totalCount + 1
                ' 한글 셀에 값 설정
                If SetHwpCellValue(listId, cellValue) Then
                    successCount = successCount + 1
                End If
            End If
        End If
    Next i

    MsgBox "전송 완료: " & successCount & "/" & totalCount & "개 셀", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "오류 발생: " & Err.Description, vbCritical
End Sub

' ------------------------------------------------------------
' 한글 인스턴스 가져오기 (ROT에서)
' ------------------------------------------------------------
Function GetHwpInstance() As Object
    On Error Resume Next

    ' 이미 실행 중인 한글에 연결
    Set GetHwpInstance = GetObject(, "HWPFrame.HwpObject")

    If GetHwpInstance Is Nothing Then
        ' 새로 실행
        Set GetHwpInstance = CreateObject("HWPFrame.HwpObject")
        If Not GetHwpInstance Is Nothing Then
            GetHwpInstance.RegisterModule "FilePathCheckDLL", "FilePathCheckerModuleExample"
        End If
    End If
End Function

' ------------------------------------------------------------
' 한글 셀에 값 설정 (list_id 기반)
' ------------------------------------------------------------
Function SetHwpCellValue(listId As Long, value As String) As Boolean
    On Error GoTo ErrorHandler

    ' 셀로 이동
    hwp.SetPos listId, 0, 0

    ' 전체 선택 후 삭제
    hwp.HAction.Run "SelectAll"
    hwp.HAction.Run "Delete"

    ' 새 값 입력
    If Len(value) > 0 Then
        Dim pset As Object
        Set pset = hwp.HParameterSet.HInsertText
        hwp.HAction.GetDefault "InsertText", pset.HSet
        pset.Text = value
        hwp.HAction.Execute "InsertText", pset.HSet
    End If

    SetHwpCellValue = True
    Exit Function

ErrorHandler:
    SetHwpCellValue = False
End Function

' ------------------------------------------------------------
' 한글에서 엑셀로 데이터 가져오기 (역방향)
' ------------------------------------------------------------
Sub GetFromHwp()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim wsData As Worksheet
    Dim sheetName As String
    Dim lastRow As Long
    Dim i As Long
    Dim listId As Long
    Dim cellValue As String
    Dim successCount As Long

    ' 현재 시트 이름에서 메인 시트 이름 추출
    sheetName = ActiveSheet.Name
    If InStr(sheetName, "_") > 0 Then
        sheetName = Left(sheetName, InStr(sheetName, "_") - 1)
    End If

    ' _cells 시트 찾기
    On Error Resume Next
    Set wsData = ThisWorkbook.Sheets(sheetName & "_cells")
    On Error GoTo ErrorHandler

    If wsData Is Nothing Then
        MsgBox sheetName & "_cells 시트를 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If

    ' 메인 시트
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' 한글 연결
    Set hwp = GetHwpInstance()
    If hwp Is Nothing Then
        MsgBox "한글이 실행 중이지 않습니다.", vbExclamation
        Exit Sub
    End If

    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    successCount = 0

    For i = 2 To lastRow
        listId = wsData.Cells(i, 1).Value

        ' 배경색 없는 셀만 처리
        If wsData.Cells(i, 14).Value = "" Then
            Dim excelRow As Long, excelCol As Long
            excelRow = wsData.Cells(i, 2).Value + 1
            excelCol = wsData.Cells(i, 3).Value + 1

            If listId > 0 Then
                ' 한글 셀에서 값 읽기
                cellValue = GetHwpCellValue(listId)

                ' 엑셀에 값 설정
                ws.Cells(excelRow, excelCol).Value = cellValue
                successCount = successCount + 1
            End If
        End If
    Next i

    MsgBox "가져오기 완료: " & successCount & "개 셀", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "오류 발생: " & Err.Description, vbCritical
End Sub

' ------------------------------------------------------------
' 한글 셀에서 값 읽기
' ------------------------------------------------------------
Function GetHwpCellValue(listId As Long) As String
    On Error GoTo ErrorHandler

    hwp.SetPos listId, 0, 0
    hwp.HAction.Run "SelectAll"

    Dim text As String
    text = hwp.GetTextFile("TEXT", "saveblock")

    hwp.HAction.Run "Cancel"

    ' 줄바꿈 제거
    text = Replace(text, vbCrLf, " ")
    text = Replace(text, vbCr, " ")
    text = Replace(text, vbLf, " ")
    text = Trim(text)

    GetHwpCellValue = text
    Exit Function

ErrorHandler:
    GetHwpCellValue = ""
End Function

' ------------------------------------------------------------
' 버튼 생성 헬퍼 (개발자 도구에서 실행)
' ------------------------------------------------------------
Sub CreateButtons()
    Dim ws As Worksheet
    Dim btn As Button

    Set ws = ActiveSheet

    ' 기존 버튼 삭제
    On Error Resume Next
    ws.Buttons("btnSendToHwp").Delete
    ws.Buttons("btnGetFromHwp").Delete
    On Error GoTo 0

    ' 한글로 보내기 버튼
    Set btn = ws.Buttons.Add(10, 10, 100, 25)
    btn.Name = "btnSendToHwp"
    btn.Caption = "한글로 보내기"
    btn.OnAction = "SendToHwp"

    ' 한글에서 가져오기 버튼
    Set btn = ws.Buttons.Add(120, 10, 100, 25)
    btn.Name = "btnGetFromHwp"
    btn.Caption = "한글에서 가져오기"
    btn.OnAction = "GetFromHwp"

    MsgBox "버튼이 생성되었습니다.", vbInformation
End Sub
