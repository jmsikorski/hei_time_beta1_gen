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
Private dots As Integer
Public done As Boolean
Public task As String
Private i As Integer
Const steps = 20

Public Sub stopLoading()
    done = True
End Sub

Public Sub updateTask(t As String)
    task = t
End Sub

Public Sub updateProgress()
    Dim pct As Single
    pct = i / steps
    i = i + 1
    t = task
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
    i = 0
    done = False
    task = "Loading"
    continueLoading
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
retry:
        Dim ans As Integer
        ans = MsgBox("The application is still running!" & vbNewLine & _
        "Exiting before completion is dangerous.", vbCritical + vbAbortRetryIgnore, "STOP!")
        If ans = vbAbort Then
            Cancel = True
        ElseIf ans = vbIgnore Then
            Unload Me
        Else
            GoTo retry
        End If
    End If
End Sub
