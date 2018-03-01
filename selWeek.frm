VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} selWeek 
   Caption         =   "Select Week"
   ClientHeight    =   4440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8280.001
   OleObjectBlob   =   "selWeek.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "selWeek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub setWeek_Click()
    week = Me.ComboBox1.Value
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
    Dim thisWeek As Date
    Dim nextWeek As Date
    thisWeek = calcWeek(Date)
    nextWeek = calcWeek(Date + 7)
    With Me.ComboBox1
        .AddItem thisWeek
        .AddItem nextWeek
'            .AddItem tEmp.Offset(0, -3) & " " & tEmp.Offset(0, -4)
'            .list(.ListCount - 1, 1) = tEmp.Offset(0, 1).Value
    End With
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
    mMenu.Show
End Sub
