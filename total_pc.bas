Attribute VB_Name = "total_pc"
Private Enum state
    open_phase = 1
    close_phase = 2
    update_phase = 3
End Enum

Private Const pw = ""
Private Function resize_name_range(name As String, ws As Worksheet, c1 As Range, c2 As Range) As Integer
    On Error GoTo 10
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

Public Sub update_file()
Attribute update_file.VB_ProcData.VB_Invoke_Func = "U\n14"
    Application.DisplayAlerts = False
    On Error GoTo 0
    Dim phase_wb As Workbook
    Dim wb As Workbook
    Set wb = Workbooks("Lead Card.xlsx")
    
    Dim ws As Worksheet
    Dim xlFile As String
    Dim rng As Range
'    On Error GoTo update_err
    With wb.Worksheets("ADD NEW PHASE CODE")
        On Error GoTo 0
        .Unprotect pw
        .Range("updated") = Now()
        .Protect pw
    End With
    open_phase_code.update_phase_code
    DisplayAlerts = False
    wb.SaveAs wb.path & "\" & wb.name
    wb.Close
    Exit Sub
update_err:
    MsgBox "Unable to update phase codes UPDATE ERR", vbExclamation, "ERROR"
    wb.Worksheets("Roster").Activate
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
End Sub
