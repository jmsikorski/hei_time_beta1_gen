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

Public Sub updateProgress(task As String, pct As Single)
    Me.Label1.Caption = "Loading " & task & vbNewLine & "This might take a few moments."
    Me.ProgressLabel.Width = (Me.ProgressFrame.Width - 12) * pct
    Me.ProgressFrame.Caption = Format(pct, "0%")
    DoEvents
End Sub

Private Sub UserForm_Initialize()
    Me.Label1.Caption = "Loading " & vbNewLine & "This might take a few moments."
    Me.ProgressLabel.Width = 0
    Me.ProgressFrame.Caption = Format(0, "0%")
End Sub
