Attribute VB_Name = "test_code"
Public Sub t12()
    Dim lApp As Excel.Application
    Set lApp = New Excel.Application
    lApp.Workbooks.Open ThisWorkbook.path & "\loadingtimer.xlsm"
    lApp.Run "'loadingtimer.xlsm'!main"
    lApp.Run "'loadingtimer.xlsm'!update", "Task1"
    MsgBox "Does it keep going?"
    lApp.Run "'loadingtimer.xlsm'!update", "Task2"
    Application.Wait Now + TimeValue("00:00:10")
    lApp.Run "'loadingtimer.xlsm'!stopLoading"
    lApp.Quit
    Set lApp = Nothing
End Sub

Sub copyDailyJobDescription()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Dim done As Boolean
    Dim emNum As Integer
    Dim i As Integer
    Dim dataWs As Worksheet
    Set wb = ThisWorkbook
    Set dataWs = Workbooks("461702 WE 02.18.18 .Alex Brauning.revisions.xlsx").Worksheets("DAILY JOB REPORT")
    Set ws = wb.Worksheets("LEAD")
    Dim datarng As Range
    Dim empCnt As Integer
    Dim dCnt As Integer
    empCnt = 0
'    emNum = ActiveCell.Value
    For dCnt = 1 To 7
        For Each rng In ws.ListObjects(dCnt).ListColumns("DAILY JOB DESCRIPTION").DataBodyRange
            Set datarng = dataWs.Range("A8").Offset(0, (dCnt - 1) * 9)
            Do
                If rng.Offset(0, -5).Value = vbNullString Then Exit For
                If rng.Offset(0, -5).Value = datarng.Value Then
                    rng = datarng.Offset(0, 5)
                    Exit Do
                Else
                    Set datarng = datarng.Offset(1, 0)
                End If
            Loop
        Next
    Next
End Sub

Public Sub t13()
    loadingMenu.stopLoading
End Sub

Public Sub t14() ' test timeCard.makeWeekPath
    Dim wb As Workbook
    jobNum = "46XXXX"
    jobName = "TEST"
    Dim we As String
    we = Format(calcWeek(Now), "mm.dd.yy")
    jobPath = ThisWorkbook.path & "\Data\"
    sharePointPath = "C:\Users\jsikorski\Helix Electric Inc\TeslaTimeCard - Documents\Time Card Files\Data\"
    timeCard.makeWeekPath (we)
        
End Sub

Public Sub t15() ' test timeCard.updatedFile
    Dim result As Boolean
    Dim fa As String
    Dim fb As String
    fa = "C:\Users\jsikorski\Helix Electric Inc\TeslaTimeCard - Documents\Time Card Files\Data\461625\Week_02.18.18\TimeSheets\Hood_Week_02.18.18.xlsx"
    fb = "C:\Users\jsikorski\AppData\Roaming\HelixTimeCard\Data\461625\Week_02.18.18\TimeSheets\Hood_Week_02.18.18.xlsx"
    result = timeCard.updatedFile(fa, fb)
    If result Then
        Debug.Print "updated"
    Else
        Debug.Print "not updated"
    End If
End Sub

Public Sub t16()
    Dim result As Boolean
    Dim fa As String
    Dim fb As String
    Dim fol As String
    jobPath = "C:\Users\jsikorski\AppData\Roaming\HelixTimeCard\Data\"
    sharePointPath = "C:\Users\jsikorski\Helix Electric Inc\TeslaTimeCard - Documents\Time Card Files\Data\"
    fol = "461625\Week_02.18.18"
    fa = "C:\Users\jsikorski\Helix Electric Inc\TeslaTimeCard - Documents\Time Card Files\Data\461625\Week_02.18.18\TimeSheets\Cox_Week_02.18.18.xlsx"
    fb = "C:\Users\jsikorski\AppData\Roaming\HelixTimeCard\Data\461625\Week_02.18.18\TimeSheets\Cox_Week_02.18.18.xlsx"
    result = timeCard.updatedFile(fa, fb)
    If result Then
        Debug.Print "1: updated"
    Else
        Debug.Print "1: not updated"
    End If
    timeCard.getUpdatedFiles jobPath, sharePointPath, fol
    result = timeCard.updatedFile(fa, fb)
    If result Then
        Debug.Print "2: updated"
    Else
        Debug.Print "2: not updated"
    End If
    
End Sub

Public Sub t17() 'Test Set week userform
    selWeek.Show
    Stop
End Sub
Private Sub t8() 'Test of loadTestRoster function
    Dim l As Integer
    Dim e As Integer
    l = 3
    e = 5
    If loadTestRoster(l, e) = -1 Then
        Debug.Print "FAIL"
    Else
        print_roster
        Debug.Print "SUCCESS " & l * e & " Records Loadaed!"
    End If

End Sub
'
'Private Sub t9() 'Test of saveWeekRoster function
'    If loadTestRoster(3, 5) <> 1 Then Debug.Print "saveWeekRoster TEST FAIL #1"
'    If saveWeekRoster <> 1 Then Debug.Print "saveWeekRoster TEST FAIL #2" Else Debug.Print "saveWeekRoster TEST SUCCESS!"
'End Sub
'
'Private Sub t10() 'Test of loadRoster function
'    ReDim weekRoster(0, 0)
'    print_roster
'    jobNum = 461702
'    If loadRoster(Format(calcWeek(Now), "mm.dd.yy") & ".csv") <> 1 Then
'        Debug.Print "loadRoster TEST FAIL"
'    Else
'        print_roster
'        Debug.Print "loadRoster TEST SUCCESS"
'    End If
'End Sub
'
'Private Sub t11()
'    t10
'    If loadShifts = -1 Then Debug.Print "LOAD SHIFTS FAIL"
'End Sub

Private Function loadTestRoster(leads As Integer, emps As Integer) As Integer
    On Error GoTo load_err
    ReDim weekRoster(leads - 1, emps - 1)
    Dim cnt As Integer
    cnt = 1
    For i = 0 To leads - 1
        For x = 0 To emps - 1
            Dim tEmp As Employee
            Set tEmp = New Employee
            If tEmp.newEmployee(cnt) = -1 Then GoTo load_err
            cnt = cnt + 1
            Set weekRoster(i, x) = tEmp
        Next
    Next
    On Error GoTo 0
    loadTestRoster = 1
    Exit Function
load_err:
    loadTestRoster = -1
    On Error GoTo 0
End Function

Public Function print_roster() As Integer
    On Error GoTo print_err
    Dim cnt As Integer
    Dim l As Integer
    Dim e As Integer
    l = UBound(weekRoster)
    e = UBound(weekRoster, 2)
    cnt = 1
    For i = 0 To l
        For x = 0 To e
            If weekRoster(i, x) Is Nothing Then
                Exit For
            Else
                Debug.Print cnt & ": Lead(" & i & ") " & weekRoster(i, x).getFullname
                Debug.Print "Shift Count: " & weekRoster(i, x).getShifts.count
                cnt = cnt + 1
            End If
        Next
    Next
    print_roster = 1
    Debug.Print cnt - 1 & " records printed"
    On Error GoTo 0
    Exit Function
print_err:
    Debug.Print "ERROR PRINTING (" & i & "," & x & ")"
    print_roster = -1
    On Error GoTo 0
End Function
