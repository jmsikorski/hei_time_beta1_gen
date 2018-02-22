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
    Application.Visible = True
    ThisWorkbook.Save
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next
    If Environ$("username") = "jsikorski" Then
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

