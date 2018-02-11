VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} pjSuperMenu 
   Caption         =   "Superintendent Menu"
   ClientHeight    =   3255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5625
   OleObjectBlob   =   "pjSuperMenu.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "pjSuperMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

















Private Sub smBuild_Click()
    Dim we As String
    Dim xlFile As String
    Dim killFile As String
    Dim xStrPath As String
    we = Format(week, "mm.dd.yy")
    xlFile = jobPath & "\" & jobNum & "\Week_" & we & "\TimePackets\" & jobNum & "_Week_" & we & ".xlsx"
    If testFileExist(xlFile) > 0 Then
        On Error Resume Next
        Dim ans As Integer
        ans = MsgBox("The packet already exists, Are you sure you want to overwrite it?", vbYesNo + vbQuestion)
        If ans = vbYes Then
            xStrPath = jobPath & "\" & jobNum & "\Week_" & we & "\TimeSheets\"
            Kill xlFile
            xStrPath = jobPath & "\" & jobNum & "\Week_" & we & "\TimeSheets\"
            killFile = Dir(xStrPath & "\*.xlsx")
            Do While killFile <> ""
                Kill xStrPath & killFile
                killFile = Dir
            Loop
        Else
            Exit Sub
        End If
        On Error GoTo 0
    End If
    sMenu.Hide
    Set lMenu = New pjSuperPkt
    lMenu.Show
End Sub

Private Sub smEdit_Click()
    If timeCard.loadRoster = -1 Then GoTo 10
    sMenu.Hide
    Set lMenu = New pjSuperPkt
    lMenu.Show
    Exit Sub
10:
    MsgBox ("Unable to Edit Packet - The file does not exist")
End Sub

Public Sub smExit_Click()
    sMenu.Hide
    mMenu.Show
End Sub

Private Sub smSubmit_Click()
    timeCard.genTimeCard
    
    timeCard.updatePacket
    MsgBox "Time Cards Complete"
    Unload Me
    mMenu.Show
End Sub

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Label2.Caption = job & vbNewLine & Format(week, "mm-dd-yy")
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Me.smExit_Click
    End If
End Sub

