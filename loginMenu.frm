VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} loginMenu 
   Caption         =   "Select Job Number"
   ClientHeight    =   6390
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8280.001
   OleObjectBlob   =   "loginMenu.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "loginMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub loginButton_Click()
    Me.Hide
End Sub

Public Sub mCancel_Click()
    Application.DisplayAlerts = False
    Dim unLockIn As String
    Dim ans As Integer, attempt As Integer
    Dim correct As Boolean
    correct = False
    attempt = 1
    ans = MsgBox("This file is locked" & vbNewLine & "Are you sure you want to quit?", 4147, "EXIT")
    If ans = 6 Then
        Application.DisplayAlerts = False
        Unload Me
        ThisWorkbook.Close
    ElseIf ans = 2 Then
        If Environ$("username") = "jsikorski" Then
            On Error Resume Next
            If loginMenu.Visible = True Then
                loginMenu.Hide
            End If
            If mMenu.Visible = False Then
                mMenu.Hide
            End If
            If sMenu.Visible = True Then
                sMenu.Hide
            End If
            If Application.Visible = False Then
                Application.Visible = True
            End If
            End
        End If
        Do While correct = False And attempt > 0
            unLockIn = InputBox("This file is locked for editing" & vbNewLine & "Please enter the unlock password:", "UNLOCK FILE ATTEMPT " & attempt & "/3")
            If unLockIn = "" Then
                attempt = attempt + 1
            ElseIf unLockIn = "jms7481" Then
                On Error Resume Next
                If loginMenu.Visible = True Then
                    loginMenu.Hide
                End If
                If mMenu.Visible = True Then
                    mMenu.Hide
                End If
                If sMenu.Visible = True Then
                    sMenu.Hide
                End If
                If Application.Visible = False Then
                    Application.Visible = True
                End If
                On Error GoTo 0
                End
            Else
                attempt = attempt + 1
            End If
            If attempt = 4 Then
                MsgBox "You have made 3 failed attempts!", 16, "FAILED UNLOCK"
                Application.DisplayAlerts = False
                Unload Me
                If Application.Visible = False Then
                    Application.Visible = True
                End If
                ThisWorkbook.Close
            End If
        Loop
    End If
End Sub


Private Sub pw_reset_Click()
    MsgBox ("This feature is not implemented yet")
End Sub

Private Sub reqUser_Click()
    Me.Hide
    reqMenu.Show
End Sub

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .TextBox2.Value = timeCard.user
        .TextBox1.SetFocus
        .Caption = "LOGIN"
    End With

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        mCancel_Click
    End If
End Sub
