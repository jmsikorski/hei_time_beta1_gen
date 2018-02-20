VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} addlead 
   Caption         =   "Select Lead to be Added"
   ClientHeight    =   4440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8280.001
   OleObjectBlob   =   "addlead.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "addlead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addLeadEnter_Click()
    Dim tEmp As Variant
    Dim uNum As Range
    Dim dm As ListObject
    For Each tEmp In ThisWorkbook.Worksheets("ROSTER").ListObjects("emp_roster").ListColumns("LEAD").DataBodyRange
        With Me.ComboBox1
            If .Value = tEmp.Offset(0, -3) & " " & tEmp.Offset(0, -4) Then
                ThisWorkbook.Worksheets("ROSTER").Unprotect xPass
                tEmp.Value = "YES"
                ThisWorkbook.Worksheets("ROSTER").Protect xPass
                Exit For
            End If
        End With
    Next
    Unload Me
    Unload lMenu
    Set lMenu = New pjSuperPkt
    lMenu.Show

End Sub

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
    Dim tEmp As Variant
    Dim uNum As Range
    Dim dm As ListObject
    For Each tEmp In ThisWorkbook.Worksheets("ROSTER").ListObjects("emp_roster").ListColumns("LEAD").DataBodyRange
        With Me.ComboBox1
            If tEmp = "NO" Then
                .AddItem tEmp.Offset(0, -3) & " " & tEmp.Offset(0, -4)
                .list(.ListCount - 1, 1) = tEmp.Offset(0, 1).Value
            End If
        End With
    Next
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
       xcl_Click
    End If
End Sub

Private Sub xcl_Click()
    Unload Me
    If addlead.Visible = True Then
        addlead.Hide
    End If
    lMenu.Show
End Sub
