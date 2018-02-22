VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} mainMenu 
   Caption         =   "Select Job Number"
   ClientHeight    =   4440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8280.001
   OleObjectBlob   =   "mainMenu.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "mainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()
    On Error GoTo 1
    job = ComboBox1.Value
    Dim tEmp() As String
    tEmp = Split(job, " - ")
    jobNum = tEmp(0)
    jobName = tEmp(1)
1:
End Sub
Public Sub mCancel_Click()
    Unload Me
End Sub

Private Sub pjCoordinator_Click()
    MsgBox ("This feature is not implemented yet")
End Sub

Private Sub pjSuper_Click()
    selWeek.Show
    jobPath = ThisWorkbook.path & "\Data\"
    sharePointPath = "C:\Users\" & Environ$("username") & "\Helix Electric Inc\TeslaTimeCard - Documents\Time Card Files\Data\"
    getUpdatedFiles jobPath, sharePointPath, jobNum & "\Week_" & Format(calcWeek(Date), "mm.dd.yy")
    If TypeName(mMenu) <> "mainMenu" Then
        job = "ERROR"
    Else
        If job = vbNullString Then
            MsgBox ("You must enter a job number")
            Exit Sub
        End If
    End If
    mMenu.Hide
    Set sMenu = New pjSuperMenu
    sMenu.Show
End Sub

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
    Dim cJob As Range
    Dim uNum As Range
    timeCard.a
    For Each cJob In ThisWorkbook.Worksheets("JOBS").Range("jobList")
        With Me.ComboBox1
        For Each uNum In ThisWorkbook.Worksheets("USER").Range("A2", ThisWorkbook.Worksheets("USER").Range("A2").End(xlDown))
            If uNum.Value = user Then
                If uNum.Offset(0, cJob.Row + 2) = True Then
                    .AddItem cJob.Value
                    .List(.ListCount - 1, 1) = cJob.Offset(0, 1).Value
                End If
            End If
        Next
      End With
    Next
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
       mCancel_Click
    End If
    timeCard.main True
End Sub

