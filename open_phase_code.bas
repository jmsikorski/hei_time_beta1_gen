Attribute VB_Name = "open_phase_code"
Private Enum state
    open_phase = 1
    close_phase = 2
    update_phase = 3
End Enum

Private Const pw = ""

Public Sub close_phase_code()
    'on error goto 10
    Dim new_code As Double
    Dim new_desc As String
    Dim rng As Range
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Open Phase Codes")
    Set rng = ws.ListObjects("phase_list").Range(1, 1)
    new_code = get_code(close_phase)
    For Each rng In ws.ListObjects("phase_list").ListColumns(1).DataBodyRange
        If rng.Value = new_code Then
            ws.Unprotect pw
            ws.Rows(rng.Row).EntireRow.Delete
            ws.Protect pw
            Exit Sub
        End If
    Next rng
    MsgBox "Phase Code does not exist", vbExclamation, "ERROR!"
    Exit Sub
10:
    MsgBox "Error: Unable to close Phase Code", vbExclamation, "ERROR!"
End Sub

Public Sub update_phase_code()
    Application.ScreenUpdating = False
    'on error goto 10
    Dim new_code As Double
    Dim new_desc As String
    Dim rng As Range
    Dim ws As Worksheet
    Dim cnt As Integer
    Dim lc_wb As Workbook
    Dim pct As Single
    Dim uCnt As Integer
    pct = 0.03
    ''loadingMenu.updateProgress "Open Phase Codes", pct
    Set lc_wb = Workbooks("Lead Card.xlsx")
    Workbooks.Open lc_wb.path & "\Labor Report.xlsx"
    uCnt = 0
    cnt = 1
    pct = 0.04
    ''loadingMenu.updateProgress "Open Phase Codes", pct
    Do While Workbooks("Labor Report.xlsx").Worksheets(1).Range("C3").Offset(uCnt, 0) <> vbNullString
        uCnt = uCnt + 1
    Loop
    pct = 0.05
    ''loadingMenu.updateProgress "Open Phase Codes", pct
    Set ws = lc_wb.Worksheets("Open Phase Codes")
    ws.Unprotect pw
    ws.Range(ws.ListObjects("phase_list").DataBodyRange(1, 1), ws.ListObjects("phase_list").DataBodyRange(ws.ListObjects("phase_list").ListRows.count - 6, 2)).Delete
    ws.Range(ws.ListObjects("phase_list").DataBodyRange(1, 1), ws.ListObjects("phase_list").DataBodyRange(1, 2)).Clear
    Set rng = ws.ListObjects("phase_list").DataBodyRange(1, 1)
    Application.ScreenUpdating = True
    new_code = 1
    Do While new_code <> 0
1:
        pct = 0.06 + ((cnt / uCnt) * 0.91)
        ''loadingMenu.updateProgress "Open Phase Codes", pct
        new_code = get_code(update_phase, cnt)
        If new_code = -1 Then
            GoTo 20
        ElseIf new_code = -2 Then
            cnt = cnt + 1
            GoTo 1
        ElseIf new_code = 0 Then
            Exit Do
        End If
        For Each rng In ws.ListObjects("phase_list").ListColumns(1).DataBodyRange
            If rng.Value = vbNullString Then
                GoTo 5
            End If
            If rng.Value = new_code Then
                cnt = cnt + 1
                GoTo 1
            End If
        Next rng
5:
        new_desc = get_description(update_phase, cnt)
        If new_desc = vbNullString Then
            GoTo 20
        Else
            If insert_code(new_code, new_desc) = -1 Then
                GoTo 20
            End If
            cnt = cnt + 1
        End If
    Loop
    On Error GoTo 0
    pct = 0.98
    ''loadingMenu.updateProgress "Open Phase Codes", pct
    
    ws.ListObjects("phase_list").ListRows(ws.ListObjects("phase_list").ListRows.count - 5).Delete
    ws.Protect pw
    Workbooks("Labor Report.xlsx").Close False
    pct = 0.99
    ''loadingMenu.updateProgress "Open Phase Codes", pct
    Exit Sub
10:
    Dim ans As Integer
    With Application.FileDialog(msoFileDialogOpen)
        .title = "Select Labor Report"
        .Filters.Add "Excel Files", "*.xls*", 1
        .InitialFileName = ActiveWorkbook.path & "\"
        ans = .Show
        If ans = 0 Then
            Exit Sub
        Else
            Workbooks.Open .SelectedItems(1)
        End If
    End With
    Set mb = Workbooks("Labor Report.xlsx")
    Resume Next
    Exit Sub
20:
    MsgBox "ERROR: Unable to update phase Codes Err:20", vbCritical, "ERROR!"
    On Error GoTo 0
    ws.Protect pw
    Application.ScreenUpdating = True
End Sub

Public Sub open_phase_code()
    On Error GoTo 10
    Dim new_code As Double
    Dim new_desc As String
    Dim rng As Range
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Open Phase Codes")
    ws.Unprotect pw
    new_code = get_code(open_phase)
    For Each rng In ws.ListObjects("phase_list").ListColumns(1).DataBodyRange
        If rng.Value = new_code Then
            MsgBox "Phase code already open!", vbExclamation, "ERROR!"
            Exit Sub
        End If
    Next rng
    If new_code = -1 Then
        GoTo 10
    Else
        new_desc = get_description(open_phase)
        If new_desc = vbNullString Then
            GoTo 10
        Else
            If insert_code(new_code, new_desc) = -1 Then
                GoTo 10
            End If
        End If
    End If
    Set rng = ws.Range("A1").End(xlDown)
    resize_name_range "open_codes", ws, ws.Range("C2"), rng.Offset(0, 2)
    ws.Protect pw
    Exit Sub
10:
    MsgBox "Unable to open phase code", vbExclamation
    On Error GoTo 0
    ws.Protect pw
End Sub

Private Function get_code(state As Integer, Optional cnt As Integer = 1) As Double
    Dim new_code As Double
    Dim ans As Integer
1:
    Select Case state
    Case open_phase
        new_code = InputBox("Enter Phase Code to Open", "Open Phase Code")
        If new_code < 0 Or new_code > 99999 Then
            MsgBox "Invalid Phase Code Entered!", vbCritical, "ERROR!"
            GoTo 1
        Else
            If new_code > 89999 Or new_code < 89000 Then
                ans = MsgBox("Unexpected Phase Code!" & vbNewLine & "Do you want to add " & new_code, vbYesNoCancel)
                If ans = vbYes Then
                    get_code = new_code
                    Exit Function
                ElseIf ans = vbCancel Then
                    get_code = -1
                Else
                    GoTo 1
                End If
            Else
                get_code = new_code
            End If
        End If
    Case close_phase
        new_code = InputBox("Enter Phase Code to Close", "Close Phase Code")
        If new_code < 0 Or new_code > 99999 Then
            MsgBox "Invalid Phase Code Entered!", vbCritical, "ERROR!"
            GoTo 1
        Else
            get_code = new_code
        End If
    Case update_phase
        Dim mb As Workbook
        Dim xlFile As String
        On Error GoTo 30
        Set mb = Workbooks("Labor Report.xlsx")
        On Error GoTo 0
        If mb.Worksheets(1).Range("C2").Offset(cnt, 0).Interior.Color = 255 Then
            get_code = -2
            Exit Function
        End If
        new_code = mb.Worksheets(1).Range("C2").Offset(cnt, 0)
        If new_code < 0 Or new_code > 99999 Then
            MsgBox "Invalid Phase Code Entered!", vbCritical, "ERROR!"
            get_code = -1
        Else
            get_code = new_code
        End If
    Case Else
        get_code = -1
    End Select
    Exit Function
30:
    xlFile = ThisWorkbook.path & "\Labor Report.xlsx"
    Workbooks.Open xlFile
    Set mb = Workbooks("Labor Report.xlsx")
    Resume Next
End Function


Private Function get_description(state As Integer, Optional cnt As Integer = 1) As String
    Dim desc As String
    Dim ans As Integer
1:
    Select Case state
    Case open_phase
        desc = InputBox("Enter Phase Code Description", "Open Phase Code")
        If desc = vbNullString Then
            MsgBox "Description can not be empty!", vbCritical, "ERROR!"
            GoTo 1
        End If
        If Len(desc) > 50 Then
            ans = MsgBox("Description is too long!" & vbNewLine & "Do you want to add it anyway?", vbCritical + vbAbortRetryIgnore)
            If ans = vbIgnore Then
                get_description = desc
                Exit Function
            ElseIf ans = vbCancel Then
                get_description = vbNullString
            ElseIf ans = vbAbort Then
                get_description = vbNullString
            Else
                GoTo 1
            End If
        Else
            get_description = desc
        End If
    Case update_phase
        Dim mb As Workbook
        Dim xlFile As String
        'on error goto 10
        Set mb = Workbooks("Labor Report.xlsx")
        On Error GoTo 0
        desc = mb.Worksheets(1).Range("D2").Offset(cnt, 0)
        get_description = desc
    Case Else
        get_description = vbNullString
    End Select
    Exit Function
10:
    Stop
    xlFile = ThisWorkbook.path & "\Labor Report.xlsx"
    Workbooks.Open xlFile
    Set mb = Workbooks("Labor Report.xlsx")
    Resume Next
End Function


Private Function insert_code(code As Double, desc As String) As Integer
    'on error goto 10
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Set wb = Workbooks("Lead Card.xlsx")
    Set ws = wb.Worksheets("Open Phase Codes")
    For Each rng In ws.ListObjects("phase_list").ListColumns(1).DataBodyRange 'ws.Range("A2", ws.Range("A1").End(xlDown))
        If rng.Value = code Then
            MsgBox "Phase code already open!", vbCritical, "ERROR!"
            GoTo 10
        End If
        If rng.Value > code Then
1:
            With rng
                .Value = code
                .Font.name = "Arial"
                .Font.Bold = False
                .Font.Size = 12
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlThin
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).Weight = xlThin
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeRight).Weight = xlThin
                .Offset(0, 1).Value = desc
                .Offset(0, 1).Font.name = "Arial"
                .Offset(0, 1).Font.Size = 12
                .Offset(0, 1).Font.Bold = False
                .Offset(0, 1).HorizontalAlignment = xlLeft
                .Offset(0, 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Offset(0, 1).Borders(xlEdgeBottom).Weight = xlThin
                .Offset(0, 1).Borders(xlEdgeTop).LineStyle = xlContinuous
                .Offset(0, 1).Borders(xlEdgeTop).Weight = xlThin
                .Offset(0, 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Offset(0, 1).Borders(xlEdgeLeft).Weight = xlThin
                .Offset(0, 1).Borders(xlEdgeRight).LineStyle = xlContinuous
                .Offset(0, 1).Borders(xlEdgeRight).Weight = xlThin
                .Offset(0, 2) = rng.Offset(-1, 2).Formula
                insert_code = 1
            End With
            ws.ListObjects("phase_list").ListRows.Add ws.ListObjects("phase_list").ListRows.count - 4
            Exit Function
        ElseIf rng.Value = vbNullString Then
            GoTo 1
        End If
    Next rng
    Set rng = ws.Range("A1").End(xlDown).Offset(1, 0)
    GoTo 1
10:
    insert_code = -1
    On Error GoTo 0
End Function


Private Function resize_name_range(name As String, ws As Worksheet, c1 As Range, c2 As Range) As Integer
    'on error goto 10
    Dim wb As Workbook
    Dim nr As name
    Dim rng As Range
    Set wb = ThisWorkbook
    Set nr = wb.Names.Item(name)
    Set rng = ws.Range(c1, c2)
    nr.RefersTo = rng
    resize_name_range = 1
    Exit Function
10:
    resize_name_range = -1
    On Error GoTo 0
End Function

