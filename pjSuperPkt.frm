VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} pjSuperPkt 
   Caption         =   "Add Leads"
   ClientHeight    =   11250
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10365
   OleObjectBlob   =   "pjSuperPkt.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "pjSuperPkt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






















Private Sub spAdd_Click()
    Set aLead = New addlead
    aLead.Show
    
End Sub

Private Sub spDone_Click()
    Dim tLead As Employee
    Dim tmpRoster() As Employee
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("ROSTER")
    Set tLead = New Employee
    Dim lBox As Integer
    Dim tlist As Object
    lBox = Me.Controls.count - 5
    Dim tmp As Range
    Dim lIndex As Integer
    Dim leadNum As Integer
    leadNum = UBound(weekRoster)
    lIndex = 0
    For i = 1 To lBox
        Set tlist = Me.Controls.Item("empList" & i)
        For x = 0 To tlist.ListCount - 1
            If tlist.Selected(x) Then
            lIndex = lIndex + 1
            End If
        Next x
    Next i
    If lIndex = 0 Then
        MsgBox "You must Select a Lead!", vbExclamation + vbOKOnly
        Exit Sub
    End If
    If UBound(menuList) = 0 And isSave <> 1 Then 'isSave < 0 Then
        ReDim weekRoster(lIndex - 1, eCount)
    Else
        ReDim tmpRoster(lIndex - 1, eCount)
    End If
    lIndex = 0
    For i = 1 To lBox
        Set tlist = Me.Controls.Item("empList" & i)
        For x = 0 To tlist.ListCount - 1
            If tlist.Selected(x) Then
                leadRoster(i - 1, x).eLead = 0
                If UBound(menuList) = 0 And isSave <> 1 Then 'isSave < 0 Then
                    Set weekRoster(lIndex, 0) = leadRoster(i - 1, x)
                Else
                    Set tmpRoster(lIndex, 0) = leadRoster(i - 1, x)
                End If
                lIndex = lIndex + 1
            End If
        Next x
    Next i
    lNum = 1
    Dim ldn As Integer
    ldn = 0
    ReDim menuList(lIndex)
    
    If UBound(weekRoster) <> lIndex - 1 Then
        If UBound(weekRoster) < lIndex - 1 Then
            For i = 0 To UBound(tmpRoster)
                If (tmpRoster(i, 0).getFullname = weekRoster(ldn, 0).getFullname) Then
                    For x = 0 To eCount
                        Set tmpRoster(i, x) = weekRoster(ldn, x)
                    Next x
                    If ldn = UBound(weekRoster) Then Exit For
                    ldn = ldn + 1
                End If
            Next i
        Else
            For i = 0 To UBound(weekRoster)
                If (tmpRoster(ldn, 0).getFullname = weekRoster(i, 0).getFullname) Then
                    For x = 0 To eCount
                        Set tmpRoster(ldn, x) = weekRoster(i, x)
                    Next x
                    If ldn = UBound(tmpRoster) Then Exit For
                    ldn = ldn + 1
                End If
            Next i
        End If
        resizeRoster lIndex - 1, eCount
        For i = 0 To lIndex - 1
            For e = 0 To eCount
                Set weekRoster(i, e) = tmpRoster(i, e)
            Next e
'            If ldn <> lIndex - 1 Then
'                If weekRoster(i, 0) Is Nothing Then
'                    For x = 0 To eCount
'                        Set weekRoster(i, 0) = tmpRoster(i, 0)
'                    Next x
'                End If
'                If (tmpRoster(i, 0).getFullname = weekRoster(ldn, 0).getFullname) Then
'                    For x = 0 To eCount
'                        Set weekRoster(i, x) = weekRoster(ldn, x)
'                    Next x
'                    ldn = ldn + 1
'                Else
'                    For x = 0 To eCount
'                        Set weekRoster(i, 0) = tmpRoster(i, 0)
'                    Next x
'                End If
'            Else
'                For x = 0 To eCount
'                    Set weekRoster(i, 0) = tmpRoster(i, 0)
'                Next x
'            End If
        Next i
        resizeRoster lIndex - 1, eCount
    End If
    For i = 0 To UBound(weekRoster)
        With menuList(i)
            addMenu (4) 'mType.pjSuperPktEmp)
            menuList(i).setSheet (i)
        End With
    Next i
    Me.Hide
    menuList(0).Show
End Sub

Private Sub UserForm_Initialize()
    Dim tLead As String
    Dim ws As Worksheet
    Dim tmp As Range
    Set ws = ThisWorkbook.Worksheets("ROSTER")
    Dim cnt As Integer
    Dim lBoxHt As Integer
    lBoxHt = 0
    lCnt = 0
    cnt = 0
    For Each tmp In ws.Range("E2", ws.Range("E2").End(xlDown))
        Debug.Print tmp.Offset(0, -1) & " " & tmp.Offset(0, -2)
        If tmp.Offset(0, 2).Value = "YES" Then
            lCnt = lCnt + 1
        End If
        cnt = cnt + 1
    Next tmp
    Dim numBox As Integer
    numBox = lCnt \ 24 + 1
    If numBox < 2 Then
        numBox = 2
    End If
    lBoxHt = 1 + (lCnt \ numBox)
    ReDim leadRoster(numBox - 1, lBoxHt - 1)
    For i = 1 To numBox
        Dim eBox As Control
        Set eBox = Me.Controls.Add("Forms.ListBox.1", "empList" & i)
        eBox.Visible = True
        eBox.Top = 84
        If lBoxHt < 12 Then
            eBox.Height = 198
        Else
            eBox.Height = lBoxHt * 16.5
        End If
        eBox.SpecialEffect = 0
        eBox.ListStyle = 1
        eBox.MultiSelect = 1
    Next i
    Dim name As String
    Dim maxLen As Integer
    Dim wide As Integer
    Dim eBoxIndex As Integer
    Dim eBoxCol As Integer
    eBoxCol = 1
    eBoxIndex = 0
    wide = 0
    maxLen = 0
    For i = 0 To cnt
        Dim tBox As Control
        If ws.Range("G" & i + 2).Value = "YES" Then
            Dim tEmp As Employee
            Set tEmp = New Employee
            leg = tEmp.newEmployee(i)
            name = tEmp.getFName & " " & tEmp.getLName
            If maxLen < Len(name) Then
                maxLen = Len(name)
            End If
            eBoxIndex = eBoxIndex + 1
            If eBoxIndex = lBoxHt Then
                If (lCnt Mod numBox) < eBoxCol Then
                    eBoxCol = eBoxCol + 1
                    eBoxIndex = 1
                End If
            ElseIf eBoxIndex > lBoxHt Then
                eBoxCol = eBoxCol + 1
                eBoxIndex = 1
            End If
            Set tBox = Me.Controls.Item("empList" & eBoxCol)
            tBox.AddItem name
            For tl = 0 To UBound(weekRoster)
                Dim tempLead As Employee
                Set tempLead = weekRoster(tl, 0)
                If tempLead Is Nothing Then
                Else
                    If tEmp.getNum = tempLead.getNum Then
                        tBox.Selected(tBox.ListCount - 1) = True
                    End If
                End If
            Next tl
            Set leadRoster(eBoxCol - 1, eBoxIndex - 1) = tEmp
        End If
    Next i
    wide = maxLen * 10
    With Me
        .Height = .Controls("empList1").Height + .L1.Height + .Label2.Height + .spAdd.Height + 78
        .Width = wide * numBox + 18
        .Label2.Caption = job & vbNewLine & "Week Ending: " & Format(week, "mm-dd-yy")
        .Label2.Left = 6
        .Label2.Width = wide * numBox
        .L1.Caption = "Select Leads"
        .L1.Left = 6
        .L1.Width = wide * numBox
        .spAdd.Left = (.Width - 272) / 3
        .spAdd.Top = Controls("empList1").Top + Controls("empList1").Height
        .spDone.Left = (.Width - 272) / 3 * 2 + 130
        .spDone.Top = Controls("empList1").Top + Controls("empList1").Height
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
    For i = 1 To numBox
        With Me.Controls.Item("empList" & i)
            .Left = 6 + (i - 1) * wide
            .Width = wide * (i + 1)
        End With
    Next i
    GoTo 20
10
    tEmp.emnum = -1
    Resume Next
20
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Me.Hide
        mMenu.Show
    End If
End Sub



