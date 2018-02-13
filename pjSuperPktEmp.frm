VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} pjSuperPktEmp 
   Caption         =   "Add Employees"
   ClientHeight    =   7860
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5625
   OleObjectBlob   =   "pjSuperPktEmp.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "pjSuperPktEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






















Private Sub nLead_Click()
    Dim thisMenu As String
    For i = 1 To UBound(menuList)
        thisMenu = "Lead " & i & "/" & UBound(menuList) & " " & weekRoster(i - 1, 0).getFName & " " & weekRoster(i - 1, 0).getLName
        If thisMenu = Me.E1.Caption Then
            Exit For
        End If
    Next i
    If i = UBound(menuList) Then
        MsgBox ("No more leads")
    Else
        loadRoster i - 1
        Me.Hide
        menuList(i).Show
    End If
End Sub

Private Sub pLead_Click()
    Dim thisMenu As String
    For i = 1 To UBound(menuList)
        thisMenu = "Lead " & i & "/" & UBound(menuList) & " " & weekRoster(i - 1, 0).getFName & " " & weekRoster(i - 1, 0).getLName
        If thisMenu = Me.E1.Caption Then
            Exit For
        End If
    Next i
    If i - 2 < 0 Then
        MsgBox ("You are at the first lead")
    Else
        loadRoster i - 1
        Me.Hide
        menuList(i - 2).Show
    End If
End Sub

Private Sub spAdd_Click()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("ROSTER")
    Dim lBox As Integer
    Dim tlist As Object
    lBox = Me.Controls.count - 6
    Dim tmp As Range
    Dim lIndex As Integer, mSize As Integer, cnt As Integer
    cnt = 0
'    mSize = 0
    For ld = 0 To UBound(menuList) - 1
        For i = 1 To lBox
            Set tlist = menuList(ld).Controls.Item("empList" & i)
            For x = 0 To tlist.ListCount - 1
                If tlist.Selected(x) Then
                cnt = cnt + 1
                End If
            Next x
        Next i
'        If mSize < cnt Then mSize = cnt
'        If mSize > 0 Then
'        resizeRoster UBound(menuList) - 1, eCount
        lIndex = 1
        For i = 1 To lBox
            Set tlist = menuList(ld).Controls.Item("empList" & i)
            For x = 0 To tlist.ListCount - 1
                If tlist.Selected(x) Then
                    empRoster(i - 1, x).eLead = ld
                    Set weekRoster(ld, lIndex) = empRoster(i - 1, x)
                    lIndex = lIndex + 1
                    If lIndex > eCount Then
                        MsgBox ("ERROR, Lead can only have " & eCount & " workers!")
                        Exit Sub
                    End If
                End If
            Next x
        Next i
    Next ld
'    savePacket
    Me.Hide
    For i = 0 To UBound(menuList) - 1
        loadRoster i
        Unload menuList(i)
    Next i
    lMenu.Show
End Sub

Private Sub loadRoster(ld)
    Dim lBox As Integer
    Dim tlist As Object
    lBox = Me.Controls.count - 6
    Dim tmp As Range
    Dim lIndex As Integer, mSize As Integer, cnt As Integer
    cnt = 0
    For i = 1 To lBox
        Set tlist = menuList(ld).Controls.Item("empList" & i)
        For x = 0 To tlist.ListCount - 1
            If tlist.Selected(x) Then
            cnt = cnt + 1
            End If
        Next x
    Next i
'    resizeRoster UBound(menuList) - 1, eCount
    lIndex = 1
    For i = 1 To lBox
        Set tlist = menuList(ld).Controls.Item("empList" & i)
        For x = 0 To tlist.ListCount - 1
            If tlist.Selected(x) Then
                empRoster(i - 1, x).eLead = ld
                Set weekRoster(ld, lIndex) = empRoster(i - 1, x)
                lIndex = lIndex + 1
                If lIndex > eCount Then
                    MsgBox ("ERROR, Lead can only have " & eCount & " workers!")
                    Exit Sub
                End If
            End If
        Next x
    Next i
End Sub

Private Sub spDone_Click()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Unload sMenu
    Unload lMenu
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("ROSTER")
'    Stop
'    loadRoster
    For i = 0 To UBound(menuList) - 1
        loadRoster i
        Unload menuList(i)
    Next i
    savePacket
    genLeadSheets
    mMenu.Show
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Public Sub setSheet(menuNum As Integer)
    Dim ws As Worksheet
    Dim tmp As Range
    Set ws = ThisWorkbook.Worksheets("ROSTER")
    Dim cnt As Integer
    Dim lBoxHt As Integer
    lBoxHt = 0
    cnt = 0
    For Each tmp In ws.Range("E2", ws.Range("E2").End(xlDown))
        cnt = cnt + 1
    Next tmp
    Dim numBox As Integer
    numBox = (cnt - UBound(menuList)) \ 30 + 1
    If numBox < 2 Then
        numBox = 2
    End If
    lBoxHt = 1 + ((cnt - UBound(menuList)) \ numBox)
    ReDim empRoster(numBox - 1, lBoxHt - 1)
    For i = 1 To numBox
        Dim eBox As Control
        Set eBox = Me.Controls.Add("Forms.ListBox.1", "empList" & i)
        eBox.Visible = True
        eBox.Top = 84
        If lBoxHt < 12 Then
            eBox.Height = 198
        Else
            eBox.Height = lBoxHt * 17
        End If
        eBox.SpecialEffect = 0
        eBox.ListStyle = 1
        eBox.MultiSelect = 1
    Next i
    Dim eName As String
    Dim maxLen As Integer
    Dim wide As Integer
    Dim eBoxIndex As Integer
    Dim eBoxCol As Integer
    eBoxCol = 1
    eBoxIndex = 0
    wide = 0
    maxLen = 0
    For i = 0 To cnt - 1
50
        Dim tBox As Control
        Dim tEmp
        Set tEmp = New Employee
        leg = tEmp.newEmployee(i)
    
        For p = 0 To UBound(weekRoster)
            If weekRoster(p, 0) Is Nothing Then
            Else
                If tEmp.getNum = weekRoster(p, 0).getNum Then
                    i = i + 1
                    GoTo 50
                End If
            End If
        Next p
    
        eName = tEmp.getFullname
        If maxLen < Len(eName) Then
            maxLen = Len(eName)
        End If
        eBoxIndex = eBoxIndex + 1
        If eBoxIndex = lBoxHt Then
            If (cnt - UBound(menuList) Mod numBox) < eBoxCol Then
                eBoxCol = eBoxCol + 1
                eBoxIndex = 1
            End If
        ElseIf eBoxIndex > lBoxHt Then
            eBoxCol = eBoxCol + 1
            eBoxIndex = 1
        End If
        Set empRoster(eBoxCol - 1, eBoxIndex - 1) = tEmp
        Set tBox = Me.Controls.Item("empList" & eBoxCol)
        tBox.AddItem eName
        For te = 1 To UBound(weekRoster, 2)
            Dim tempEmp As Employee
            Set tempEmp = weekRoster(menuNum, te)
            If tempEmp Is Nothing Then
            Else
                If tEmp.getNum = tempEmp.getNum Then
                    tBox.Selected(tBox.ListCount - 1) = True
                End If
            End If
        Next te
    Next i
    wide = maxLen * 10
    With Me
        .Height = .Controls("empList1").Height + .E1.Height + .Label2.Height + .spAdd.Height + 78
        .Width = wide * numBox + 18
        .Label2.Caption = job & vbNewLine & "Week Ending: " & Format(week, "mm-dd-yy")
        .Label2.Left = 6
        .Label2.Width = wide * numBox
        .E1.Caption = "Lead " & lNum & "/" & UBound(menuList) & " "
        .E1.Caption = .E1.Caption & weekRoster(lNum - 1, 0).getFName & " " & weekRoster(lNum - 1, 0).getLName
        .E1.Left = 6
        .E1.Top = pjSuperPkt.L1.Top
        .E1.Width = wide * numBox
        .spAdd.Left = (.Width - 532) / 5
        .spAdd.Top = .Controls("empList1").Top + .Controls("empList1").Height
        .spDone.Left = (.Width - 532) / 5 * 2 + 130
        .spDone.Top = .Controls("empList1").Top + .Controls("empList1").Height
        .pLead.Left = (.Width - 532) / 5 * 3 + 260
        .pLead.Top = .Controls("empList1").Top + .Controls("empList1").Height
        .nLead.Left = (.Width - 532) / 5 * 4 + 390
        .nLead.Top = .Controls("empList1").Top + .Controls("empList1").Height
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        lNum = lNum + 1
    End With
    For i = 1 To numBox
        With Me.Controls.Item("empList" & i)
            .Left = 6 + (i - 1) * wide
            .Width = wide * (i + 1)
        End With
    Next i
    Exit Sub
10
    tEmp.emnum = -1
    Resume Next
20
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Me.Hide
        lMenu.Show
    End If
End Sub
