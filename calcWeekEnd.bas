Attribute VB_Name = "calcWeekEnd"
Function calcWeek(d As Date) As Date
    Dim weekEnd As Date
    Dim day As Integer
    day = (d Mod 7) + 1
    Select Case day
        Case 1
            weekEnd = DateAdd("d", 1, d)
        Case 2
            weekEnd = Now
        Case 3
            weekEnd = DateAdd("d", 6, d)
        Case 4
            weekEnd = DateAdd("d", 5, d)
        Case 5
            weekEnd = DateAdd("d", 4, d)
        Case 6
            weekEnd = DateAdd("d", 3, d)
        Case 7
            weekEnd = DateAdd("d", 2, d)
    
    End Select
        
    calcWeek = weekEnd
End Function

