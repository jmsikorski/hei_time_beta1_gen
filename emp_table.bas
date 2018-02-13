Attribute VB_Name = "emp_table"
Private Enum state
    open_phase = 1
    close_phase = 2
    update_phase = 3
End Enum

Private Const pw = ""
    
Public Sub update_emp_table()
    On Error Resume Next
    hiddenApp.Workbooks("Attendance Tracking.xlsx").Close False
    Workbooks("Attendance Tracking.xlsx").Close False
    On Error GoTo 0
    Application.ScreenUpdating = False
    'on error goto 10
    Dim new_emp As Range
    Dim rng As Range
    Dim ws As Worksheet
    Dim cnt As Integer
    Dim pct As Single
    Dim emNum As Integer
    Set hiddenApp = New Excel.Application
    On Error Resume Next
    pct = 0
    loadingMenu.Show
    loadingMenu.updateProgress "Employee Roster", pct
    hiddenApp.Workbooks.Open (timeCard.Getlnkpath(ThisWorkbook.path & "\Data.lnk") & "\Attendance Tracking.xlsx")
    emNum = hiddenApp.Workbooks("Attendance Tracking.xlsx").Worksheets(1).ListObjects("emp_roster").ListRows.count
    On Error GoTo 0
    emNum = emNum + 3
    cnt = 1
    Set ws = ThisWorkbook.Worksheets("ROSTER")
    ws.Unprotect xPass
    ws.Range(ws.ListObjects("emp_roster").DataBodyRange(1, 1), ws.ListObjects("emp_roster").DataBodyRange(ws.ListObjects("emp_roster").ListRows.count - 1, 7)).Clear
    ws.Range(ws.ListObjects("emp_roster").DataBodyRange(1, 1), ws.ListObjects("emp_roster").DataBodyRange(1, 7)).Clear
1:
    pct = (cnt + 3) / emNum
    loadingMenu.updateProgress "Employee Roster", pct
    Set new_emp = get_emp(cnt)
    If new_emp Is Nothing Then
        cnt = cnt + 1
        GoTo 1
    ElseIf new_emp.Cells(1, 2) = vbNullString Then
        GoTo update_done
    End If
    For Each rng In ws.ListObjects("emp_roster").ListColumns(1).DataBodyRange
        If rng.Value = vbNullString Then
            GoTo 5
        End If
        If rng.Value = new_emp.Cells(1, 1) Then
            cnt = cnt + 1
            GoTo 1
        End If
    Next rng
    ws.ListObjects("emp_roster").Resize ws.Range(ws.ListObjects("emp_roster").Range(1, 1), ws.ListObjects("emp_roster").Range(1, ws.ListObjects("emp_roster").ListColumns.count).End(xlDown).Offset(10, 0))
5:
    If insert_emp(new_emp) = -1 Then
        GoTo 20
    End If
    cnt = cnt + 1
    GoTo 1
update_done:
    On Error GoTo 0
    Set rng = ws.Range(ws.ListObjects("emp_roster").Range(1, 1), ws.ListObjects("emp_roster").Range(1, ws.ListObjects("emp_roster").ListColumns.count).End(xlDown))
    ws.ListObjects("emp_roster").Resize rng
    hiddenApp.Workbooks("Attendance Tracking.xlsx").Close
    With ThisWorkbook.Worksheets("ROSTER")
        .Unprotect
        .Range("emp_table_updated") = Now()
        .Protect
    End With
    Application.ScreenUpdating = True
    hiddenApp.Quit
    Set hiddenApp = Nothing
    ws.Protect xPass
    pct = 1
    loadingMenu.updateProgress "Employee Roster", pct
    Unload loadingMenu
    Exit Sub
10:
    Dim ans As Integer
    Dim xFile As String
    With Application.FileDialog(msoFileDialogOpen)
        .title = "Find Attendance Tracking Roster"
        .Filters.Add "Excel Files", "*.xls*", 1
        .InitialFileName = ThisWorkbook.path & "\"
        ans = .Show
        If ans = 0 Then
            Exit Sub
        Else
            Workbooks.Open .SelectedItems(1)
            xFile = .SelectedItems(1)
            SetAttr xFile, vbNormal
        End If
    End With
    Set mb = Workbooks(xFile)
    Resume Next
    Exit Sub
20:
    MsgBox "ERROR: Unable to update roster", vbCritical, "ERROR!"
    On Error GoTo 0
    ws.Protect xPass
    Application.ScreenUpdating = True
End Sub

Private Function get_emp(Optional cnt As Integer = 1) As Range
    Dim new_emp As Range
    Dim ans As Integer
    Dim datFile As String
    datFile = "Attendance Tracking.xlsx"
1:
    Dim mb As Workbook
    Dim xlFile As String
    On Error GoTo book_closed
    Set mb = hiddenApp.Workbooks(datFile)
    On Error GoTo 0
    Dim rng As Range
    Set rng = mb.Worksheets(1).ListObjects("emp_roster").ListRows(cnt).Range
    If rng.Cells(1, 1).Interior.Color = 255 Then
        Set get_emp = Nothing
        Exit Function
    End If
    Set new_emp = mb.Worksheets(1).ListObjects("emp_roster").ListRows(cnt).Range
    Set get_emp = new_emp
    Exit Function
book_closed:
    hiddenApp.Workbooks.Open Getlnkpath(ThisWorkbook.path & "\Data.lnk") & "\" & datFile
    Set mb = hiddenApp.Workbooks(datFile)
    Resume Next
End Function


Private Function insert_emp(emp As Range, Optional c As Integer) As Integer
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("ROSTER")
    Dim r As Integer
    r = 0
    For Each rng In ws.ListObjects("emp_roster").ListColumns(1).DataBodyRange 'ws.Range("A2", ws.Range("A1").End(xlDown))
        If rng = emp.Cells(1, 1) Then
            GoTo 10
        End If
        Do While rng.Offset(r, 0).Value <> vbNullString
            r = r + 1
        Loop
        Set rng = rng.Offset(r, 0)
        With ws.ListObjects("emp_roster")
            ws.Range(.ListRows(r + 1).Range(1, 1), .ListRows(r + 1).Range(1, 7)) = emp.Value
            .ListRows(r + 1).Range(1, 1) = r + 1
            For i = 1 To .DataBodyRange.Columns.count
                With .ListRows(r + 1).Range(1, i)
'                    .Value = emp.Cells(1, i).Value
                    .Font.name = "Helvetica"
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
                End With
            Next i
        End With
        Exit For
5:
    Next rng
    insert_emp = 1
    Exit Function
10:
    insert_emp = -1
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

