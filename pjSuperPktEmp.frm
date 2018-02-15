VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} pjSuperPktEmp 
   Caption         =   "Add Employees"
   ClientHeight    =   9105.001
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
    lBox = Me.Controls.count - 8
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
    Application.EnableEvents = False
    Set lApp = New Excel.Application
    lApp.Workbooks.Open ThisWorkbook.path & "\loadingtimer.xlsm"
    lApp.Run "'loadingtimer.xlsm'!main"
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
    lApp.Run "'loadingtimer.xlsm'!stopLoading"
    lApp.Quit
    Set lApp = Nothing
    mMenu.Show
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
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
        Set eBox = Me.Controls("EmpFrame").Controls.Add("Forms.ListBox.1", "empList" & i)
        eBox.Visible = True
        eBox.Top = 6
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
    Dim bCnt As Integer
    bCnt = 5
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
    Dim bNames() As String
    ReDim bNames(bCnt)
    bNames = Split("spAdd,spDone,pLead,nLead,updateGoals", ",")
    bCnt = bCnt + 1
    Dim buff As Integer
    Dim space As Double
    Dim bWide As Double
    Dim bRows As Integer
    bWide = Me.spAdd.Width
    With Me.Controls("EmpFrame")
        If (wide * numBox) + 36 > Application.Width * 0.95 Then
            .Width = Application.Width * 0.95
            .ScrollBars = fmScrollBarsHorizontal
            .ScrollWidth = wide * numBox + 24
        Else
            .Width = wide * numBox + 24
        End If
        Dim header As Integer
        Dim Footer As Integer
        bRows = Application.WorksheetFunction.RoundUp((bCnt * Me.spAdd.Width) / (Me.Controls("EmpFrame").Width + 36), 0)
        head = 24 + Me.E1.Height + Me.Label2.Height
        foot = bRows * Me.spAdd.Height + 78
        If (.Controls("empList1").Height + head + foot) > Application.Height * 0.95 Then
            .Height = Application.Height * 0.95
            .Height = .Height - head - foot
            If .ScrollBars = fmScrollBarsHorizontal Then
                .ScrollBars = fmScrollBarsBoth
            Else
                .ScrollBars = fmScrollBarsVertical
            End If
            .ScrollHeight = .Controls("empList1").Height + 12
        Else
            .Height = .Controls("empList1").Height + 12
        End If
    End With
    With Me
        .Width = .Controls("EmpFrame").Width + 36
        bRows = Application.WorksheetFunction.RoundUp((bCnt * .spAdd.Width) / Me.Width, 0)
        bCnt = Application.WorksheetFunction.RoundUp((bCnt) / bRows, 0)
        .Height = .Controls("EmpFrame").Height + head + foot
        .Width = .Controls("EmpFrame").Width + 36
        buff = (Me.Width - (bCnt * bWide)) / bCnt
        space = (Me.Width - buff - ((bCnt - 1) * bWide))
        space = (space / bCnt)
        .Label2.Caption = job & vbNewLine & "Week Ending: " & Format(week, "mm-dd-yy")
        .Label2.Left = .Controls("EmpFrame").Left
        .Label2.Width = .Controls("EmpFrame").Width
        .E1.Caption = "Lead " & lNum & "/" & UBound(menuList) & " "
        .E1.Caption = .E1.Caption & weekRoster(lNum - 1, 0).getFName & " " & weekRoster(lNum - 1, 0).getLName
        .E1.Left = .Controls("EmpFrame").Left
        .E1.Top = lMenu.L1.Top
        .E1.Width = .Controls("EmpFrame").Width
        i = 0
        For r = 0 To bRows - 1
            For c = 0 To UBound(bNames) / bRows
            .Controls(bNames(i)).Left = buff + (c * (bWide + space)) + 24
            Debug.Print .Controls(bNames(i)).Left
            .Controls(bNames(i)).Top = .Controls("EmpFrame").Top + .Controls("EmpFrame").Height + (r * .spAdd.Height) + 12
            i = i + 1
            If i > UBound(bNames) Then Exit For
            Next
        Next
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        lNum = lNum + 1
    End With
    For i = 1 To numBox
        With Me.Controls("EmpFrame").Controls.Item("empList" & i)
            .Left = 6 + (i - 1) * wide
            .Width = wide * (i + 1)
        End With
    Next i
    Exit Sub
10
    tEmp.emNum = -1
    Resume Next
End Sub


Private Sub updateGoals_Click()
    Workbooks.Open ThisWorkbook.path & "\UnitGoals.xlsx"
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Integer
    Set wb = Workbooks("UnitGoals.xlsx")
    For i = 1 To wb.Sheets.count
        If wb.Worksheets(i).Visible = xlVeryHidden Then
            wb.Worksheets(i).Visible = True
        End If
    Next
    Dim tmp() As String
    Dim leadName As String
    Dim rng As Range
    tmp() = Split(Me.E1.Caption, " ")
    leadName = tmp(UBound(tmp))
    On Error Resume Next
    Set ws = wb.Worksheets(leadName)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        wb.Worksheets("MASTER").Copy before:=wb.Worksheets(1)
        Set ws = wb.Worksheets(1)
        ws.name = leadName
        With ws.ListObjects(1)
            ws.Unprotect
            .name = leadName & "_goals"
            ws.Protect
        End With
    End If
    For i = 1 To wb.Sheets.count
        With wb.Worksheets(i)
        If .name <> leadName Then
            .Visible = xlVeryHidden
        End If
        End With
    Next i
    Visible = True
    WindowState = xlMaximized
    Do While done = False
        On Error GoTo wb_closed
        Set wb = Workbooks("UnitGoals.xlsx")
        done = False
    Loop
wb_closed:
    Err.Clear
    Visible = False
    Set wb = Nothing
    Set ws = Nothing
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Me.Hide
        lMenu.Show
    End If
End Sub
