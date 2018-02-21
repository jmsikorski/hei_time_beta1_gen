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
    Dim lastWE As String
    Dim ans As Integer
    Dim FSO As FileSystemObject
    Dim testPacket As Boolean
    Set FSO = New FileSystemObject
    lastWE = Format(week - 7, "mm.dd.yy")
    we = Format(week, "mm.dd.yy")
    xlFile = jobNum & "\Week_" & we & "\TimePackets\" & jobNum & "_Week_" & we & ".xlsx"
    lwXLFile = jobPath & jobNum & "\Week_" & lastWE & "\TimePackets\" & jobNum & "_Week_" & lastWE & ".xlsx"

    If publish = vbYes Then
        If testFileExist(sharePointPath & xlFile) > 0 Or testFileExist(jobPath & xlFile) > 0 Then
            testPacket = True
        Else
            testPacket = False
        End If
    Else
        If testFileExist(jobPath & xlFile) > 0 Then
            testPacket = True
        Else
            testPacket = False
        End If
    End If
    If testPacket Then
        On Error Resume Next
        ans = MsgBox("The packet already exists, Are you sure you want to overwrite it?", vbYesNo + vbQuestion)
        If ans = vbYes Then
            Kill jobPath & xlFile
            Kill sharePointPath & xlFile
            xStrPath = jobPath & jobNum & "\Week_" & we & "\TimeSheets\"
            killFile = Dir(xStrPath & "\*.xlsx")
            Do While killFile <> ""
                Kill xStrPath & killFile
                killFile = Dir
            Loop
            xStrPath = sharePointPath & jobNum & "\Week_" & we & "\TimeSheets\"
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
    If FSO.FileExists(lwXLFile) Then
        ans = MsgBox("Copy from last week?", vbYesNoCancel + vbQuestion, "COPY?")
        If ans = vbYes Then
            On Error Resume Next
            MkDir jobPath & jobNum & "\Week_" & we
            MkDir jobPath & jobNum & "\Week_" & we & "\TimePackets"
            FSO.CopyFile lwXLFile, jobPath & xlFile, True
            FSO.CopyFile lwXLFile, sharePointPath & xlFile, True
            On Error GoTo 0
            smEdit_Click
        ElseIf ans = vbCancel Then
            GoTo clean_up
        End If
    End If
clean_up:
    Set FSO = Nothing
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
    Dim st As Date
    Dim fn As Date
    st = Now
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Set lApp = New Excel.Application
    Unload Me
    lApp.Workbooks.Open ThisWorkbook.path & "\loadingtimer.xlsm"
    lApp.Run "'loadingtimer.xlsm'!main"
    If loadRoster = -1 Then
        Stop
    End If
    If timeCard.loadShifts = -1 Then
        Stop
    End If
    timeCard.genTimeCard
    timeCard.updatePacket
    lApp.Run "'loadingtimer.xlsm'!stopLoading"
    lApp.Quit
    Set lApp = Nothing
    Application.Visible = True
    fn = Now
    MsgBox "Time to complete: " & Format(fn - st, "h:mm:ss")
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    mMenu.Show
End Sub

Private Sub UserForm_Initialize()
    publish = MsgBox("Upload to Sharepoint?", vbQuestion + vbYesNo, "UPLOAD?")
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

