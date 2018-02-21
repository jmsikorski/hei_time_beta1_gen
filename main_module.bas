Attribute VB_Name = "main_module"
Public Sub blank()

End Sub

Public Sub showBooks()
    If Environ$("username") = "jsikorski" Then
        On Error Resume Next
        ActiveWorkbook.Unprotect getXPass
        For i = 1 To ThisWorkbook.Sheets.count
            If ThisWorkbook.Worksheets(i).Visible = xlVeryHidden Then
                ThisWorkbook.Worksheets(i).Visible = True
            End If
            If ThisWorkbook.Worksheets(i).Visible = False Then
                ThisWorkbook.Worksheets(i).Visible = True
            End If
        Next i
        ThisWorkbook.Worksheets("KEY").Visible = xlVeryHidden
    Else
        MsgBox ("Sorry this toy is not for you to play with")
    End If
    On Error GoTo 0
End Sub

Public Sub HideBooks()
    If Environ$("username") = "jsikorski" Then
        On Error Resume Next
        ActiveWorkbook.Unprotect xPass
        For i = 1 To ThisWorkbook.Sheets.count
            If ThisWorkbook.Worksheets(i).Visible = True Then
                Debug.Print ThisWorkbook.Worksheets(i).name
                If ThisWorkbook.Worksheets(i).name <> ThisWorkbook.Worksheets("HOME").name Then
                    ThisWorkbook.Worksheets(i).Visible = False
                End If
            End If
        Next i
        For i = 1 To ThisWorkbook.Sheets.count
            Debug.Print ThisWorkbook.Worksheets(i).name
            If ThisWorkbook.Worksheets(i).Visible = False Then
                ThisWorkbook.Worksheets(i).Visible = xlVeryHidden
            End If
        Next i
    Else
        MsgBox ("Sorry this toy is not for you to play with")
    End If
    On Error GoTo 0
End Sub

Sub delSheets()
    Application.DisplayAlerts = False
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Sheets
        ws.Visible = True
        ws.Unprotect
        Debug.Print ws.name
        If MsgBox("Delete " & ws.name & "?", vbYesNo) = vbYes Then
            ws.Delete
        End If
    Next
    Application.DisplayAlerts = True
End Sub

Public Sub showVeryHidden()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Sheets
        ws.Visible = True
    Next
End Sub

Public Sub hideVeryHidden()
    ActiveSheet.Visible = xlVeryHidden
End Sub
