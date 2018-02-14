VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} loadingMenu 
   Caption         =   "Working..."
   ClientHeight    =   2130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "loadingMenu.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "loadingMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dots As Integer
Dim done As Boolean
Dim task As String

Public Sub go()
    Dim i As Integer
    Dim pct As Single
    Dim steps As Integer
    Dim waittime As Date
    waittime = Now + TimeValue("00:00:01")
    steps = 20
    i = 0
    Do While Not done
        If Application.wait(waittime) Then
            waittime = Now + TimeValue("00:00:01")
            i = i + 1
            If i > steps Then i = 1
            pct = i / steps
            loadingMenu.updateProgress task, pct
        End If
    Loop
    Unload Me
End Sub

Public Sub updateTask(t As String)
    task = t
End Sub

Public Sub updateProgress(t As String, Optional pct As Single)
    If t = vbNullString Then t = task
    Me.Label1.Caption = t & vbNewLine & "This might take a few moments."
    Me.Caption = "Working"
    Me.Caption = Me.Caption & Left("..........", dots)
    If dots > 10 Then
        dots = 1
    Else
        dots = dots + 1
    End If
    Me.ProgressLabel.Width = (Me.ProgressFrame.Width - 12) * pct
    DoEvents
End Sub

Private Sub UserForm_Initialize()
    Me.Label1.Caption = "Loading " & vbNewLine & "This might take a few moments."
    Me.ProgressLabel.Width = 0
    Me.ProgressFrame.Caption = vbNullString
    dots = 0
    task = vbNullString
    Me.go
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
retry:
        Dim ans As Integer
        ans = MsgBox("The application is still running!" & vbNewLine & _
        "Exiting before completion is not allowed.", vbCritical + vbAbortRetryIgnore, "STOP!")
        If ans = vbAbort Then
            Cancel = True
        ElseIf ans = vbIgnore Then
            Unload Me
        Else
            GoTo retry
        End If
    End If
End Sub
