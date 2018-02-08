Attribute VB_Name = "encryption"
Private Function encryptPassword(pw As String) As String
    Dim pwi As Long
    Dim test As String
    Dim epw As String
    Dim key As Long
    epw = vbnullStrig
    For i = 0 To Len(pw) - 1
        test = Left(pw, 1)
        pwi = Asc(test)
        pw = Right(pw, Len(pw) - 1)
        key = ThisWorkbook.Worksheets("KEY").Range("A" & i + 1).Value
        If key = pwi Then key = key + 128
        pwi = pwi Xor key
        If pwi = key + 128 Then
            pwi = key
        End If
        epw = epw & Chr(pwi)
    Next i
    encryptPassword = epw
End Function

Public Function testPW(pw As String, tpw As String) As Boolean
    If encryptPassword(tpw) = pw Then
        testPW = True
    Else
        testPW = False
    End If
End Function

