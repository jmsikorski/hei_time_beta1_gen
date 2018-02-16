VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} reqMenu 
   Caption         =   "Select Job Number"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8280.001
   OleObjectBlob   =   "reqMenu.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "reqMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mCancel_Click()
    mainMenu.mCancel_Click
End Sub

Private Sub reqSubmit_Click()
    Dim xSht As Worksheet
    Dim xOutlookObj As Object
    Dim xEmailObj As Object
    Dim send_to As String
    On Error GoTo 0
    Set xOutlookObj = CreateObject("Outlook.Application")
    Set xEmailObj = xOutlookObj.CreateItem(olMailItem)
    Dim name As String
    Dim user As String
    Dim pw As String
    Dim mgr As String
    Dim rng As Range
    Set xSht = ThisWorkbook.Worksheets("USER")
    name = Me.TextBox1 & " " & Me.TextBox2
    user = Left(Me.TextBox1, 1) & Me.TextBox2
    Dim uRow As Integer
    uRow = find_user(user)
    If uRow > 0 Then
        If xSht.Range("A" & uRow).Offset(0, 2).Value = "NO" Then
            MsgBox "Username is already pending authorization.", vbInformation + vbOKOnly
            GoTo reqSubmit_sub_end
        ElseIf xSht.Range("A" & uRow).Offset(0, 2).Value = "YES" Then
            Dim ans As Integer
10:
            ans = MsgBox("Username already exists", vbCritical + vbRetryCancel, "INVALID USERNAME")
            If ans = vbRetry Then
                Exit Sub
            ElseIf ans = vbCancel Then
                GoTo reqSubmit_sub_end
            Else
                GoTo 10
            End If
        End If
    End If
    pw = encryptPassword(Me.TextBox4)
    mgr = Me.TextBox3
    Set rng = xSht.Range("A" & xSht.UsedRange.Rows.count + 1)
    rng.Value = LCase(user)
    rng.Offset(0, 1).Value = pw
    rng.Offset(0, 2).Value = "NO"
    With xEmailObj
        .To = "jsikorski@helixelectric.com"
        .Subject = "Time Card User Request"
        .Body = name & vbNewLine & user & vbNewLine & mgr
        .display ' REMOVE AFTER BETA
'        .Send
    End With
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    user_form.export_user_sheet
    MsgBox "Thank you for your request." & vbNewLine & "You will recieve an email when your account is activated.", vbInformation + vbOKOnly, "SUBMITTED"
reqSubmit_sub_end:
    Unload Me
    ThisWorkbook.Close True
    Stop
End Sub

Private Function find_user(user As String) As Integer
    Dim rng As Range
    Dim trng As Range
    Set rng = ThisWorkbook.Worksheets("USER").Range("B1", ThisWorkbook.Worksheets("USER").Range("B1").End(xlDown))
    For Each trng In rng
        If trng.Value = user Then
            find_user = trng.Row
        End If
    Next
    find_user = -1
End Function

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .TextBox1.SetFocus
        .Caption = "REGISTER"
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        loginMenu.mCancel_Click
    End If
End Sub

Private Function encryptPassword(pw As String) As String
    Dim pwi() As Long
    Dim tEst() As String
    Dim epw As String
    epw = vbnullStrig
    ReDim tEst(Len(pw))
    ReDim pwi(Len(pw))
    Dim x As Integer
    x = 1
    For i = 0 To Len(pw) - 1
        tEst(i) = Left(pw, 1)
        pwi(i) = Asc(tEst(i))
        pw = Right(pw, Len(pw) - 1)
        pwi(i) = pwi(i) Xor ThisWorkbook.Worksheets("KEY").Range("A" & x).Value
        'If pwi(i) = 0 Then pwi(i) = 1
        epw = epw & Chr(pwi(i))
        x = x + 1
    Next i
    encryptPassword = epw
End Function
