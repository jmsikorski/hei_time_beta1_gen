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
    Dim tEst As Integer
    Dim pct As Single
    pct = 0
    Dim t1 As Date, t2 As Date
    t1 = Now
    tEst = 8
    loadingMenu.Show
    'loadingMenu.updateProgress "User File", pct
    pct = DateDiff("s", t1, Now()) / tEst
    'loadingMenu.updateProgress "User File", pct
    Workbooks.Open Getlnkpath(ThisWorkbook.path & "\Data.lnk") & "\User.xlsx", Password:="hei3078USER"
    pct = DateDiff("s", t1, Now()) / tEst
    'loadingMenu.updateProgress "User File", pct
    Set ws = Workbooks("User.xlsx").Worksheets("USER")
    pct = DateDiff("s", t1, Now()) / tEst
    'loadingMenu.updateProgress "User File", pct
    Set rng = ws.UsedRange
    pct = DateDiff("s", t1, Now()) / tEst
    'loadingMenu.updateProgress "User File", pct
    With wb.Worksheets("USER")
        .UsedRange.Offset(1, 0).Clear
        pct = DateDiff("s", t1, Now()) / tEst
        'loadingMenu.updateProgress "User File", pct
        .Range("A2", .Range("A2").Offset(rng.Rows.count - 1, rng.Columns.count - 1)) = rng.Offset(1, 0).Value
        pct = DateDiff("s", t1, Now()) / tEst
        'loadingMenu.updateProgress "User File", pct
        .Range("user_updated") = Now()
        pct = DateDiff("s", t1, Now()) / tEst
        'loadingMenu.updateProgress "User File", pct
    End With
    ws.Parent.Close False
    pct = DateDiff("s", t1, Now()) / tEst
    'loadingMenu.updateProgress "User File", pct
    t2 = Now
    pct = DateDiff("s", t1, Now()) / tEst
    'loadingMenu.updateProgress "User File", pct
    Debug.Print DateDiff("s", t1, t2)
    pct = 1
    'loadingMenu.updateProgress "User File", pct
    Unload loadingMenu
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
    DisplayAlerts = False
    Workbooks.Open Getlnkpath(ThisWorkbook.path & "\Data.lnk") & "\User.xlsx", Password:="hei3078USER"
    Set ws = Workbooks("User.xlsx").Worksheets("USER")
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
