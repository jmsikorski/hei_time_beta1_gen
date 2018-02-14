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

Public Sub t13()
    loadingMenu.stopLoading
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

Private Function print_roster() As Integer
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
                Debug.Print cnt & ": " & weekRoster(i, x).getFullname
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
