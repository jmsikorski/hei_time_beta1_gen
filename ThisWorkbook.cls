VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    ThisWorkbook.Save
    For i = 1 To ThisWorkbook.Sheets.count
    ThisWorkbook.Unprotect xPass
        If ThisWorkbook.Worksheets(i).name <> "HOME" Then
            ThisWorkbook.Worksheets(i).Visible = xlVeryHidden
        End If
    Next i
    ThisWorkbook.Protect xPass, True, False
    Application.DisplayAlerts = False
    Application.DisplayFullScreen = False
    Application.WindowState = xlMaximized
    hiddenApp.Quit
    Set hiddenApp = Nothing
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next
    If Environ$("Username") = "jsikorski" Then
        ExportVBA ThisWorkbook.name
    End If
    On Error GoTo 0
    If Environ$("username") <> "jsikorski" Then
        ThisWorkbook.Saved = True
        Cancel = True
    End If
End Sub

Private Sub Workbook_Open()
    main
End Sub

