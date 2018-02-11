Attribute VB_Name = "user_form"
Public Sub get_user_list1()
    Dim auth As String
    Dim user As String
    On Error GoTo err_tag
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Dim url As String
    Dim qt As QueryTable
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("USER")
    
    url = "https://github.com/jmsikorski/hei_misc/blob/master/Modules/Time_Card_User.csv"
    Set qt = ws.QueryTables.Add( _
        Connection:="URL;" & url, _
        Destination:=ws.Range("A1"))
     
    With qt
        .RefreshOnFileOpen = False
        .name = "Users"
        .FieldNames = True
        .WebSelectionType = xlAllTables
        .Refresh
    End With
    ws.Range("A1").EntireColumn.Delete
    ws.Range("D1:F1").EntireColumn.Clear
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
err_tag:
    MsgBox "ERROR LOADING LICENSE!", vbCritical + vbOKOnly
    With ws.UsedRange
        .Offset(1, 0).Clear
        .Value = "X"
    End With
    ThisWorkbook.Close False
 End Sub
 
Public Sub get_user_list()
    Application.ScreenUpdating = False
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim xPath As String
    Dim ws As Worksheet
    Dim rng As Range
    Dim fRng As Range
    
    hiddenApp.Workbooks.Open ThisWorkbook.path & "\User.xlsx", Password:="hei3078USER"
    Set ws = hiddenApp.Workbooks("User.xlsx").Worksheets("USER")
    Set rng = ws.UsedRange
    With wb.Worksheets("USER")
        .UsedRange.Offset(1, 0).Clear
        .Range("A2", .Range("A2").Offset(rng.Rows.count - 1, rng.Columns.count - 1)) = rng.Offset(1, 0).Value
        .Range("user_updated") = Now()
    End With
    ws.Parent.Close False
    
    
End Sub

Public Sub extract_users()
    If Environ$("Username") = "jsikorski" Then
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
        showBooks
        Dim wb As Workbook
        Dim dwb As Workbook
        Set wb = ThisWorkbook
        Set dwb = Workbooks.Add
        Debug.Print wb.name
        Debug.Print dwb.name
        wb.Worksheets("USER").Copy after:=dwb.Worksheets(1)
        dwb.Worksheets(1).Delete
        dwb.SaveAs "C:\Users\jsikorski\Documents\GitHub\hei_misc\Modules\Time_Card_User.csv", xlCSV
        dwb.Close
        HideBooks
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
    End If
End Sub


Public Sub export_user_sheet()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim xPath As String
    Dim ws As Worksheet
    Dim rng As Range
    Dim xFile As String
    
    hiddenApp.DisplayAlerts = False
    hiddenApp.Workbooks.Open ThisWorkbook.path & "\User.xlsx", Password:="hei3078USER"
    Set ws = hiddenApp.Workbooks("User.xlsx").Worksheets("USER")
    Set rng = wb.Worksheets("USER").UsedRange
    With ws
        .UsedRange.Clear
        .Range("A1", .Range("A1").Offset(rng.Rows.count - 1, rng.Columns.count - 1)) = rng.Value
        .Range("user_updated") = Now()
    End With
    xFile = ws.Parent.path & "\" & ws.Parent.name
    ws.Parent.SaveAs xFile
    SetAttr xFile, vbHidden
    ws.Parent.Close
    
    
End Sub
