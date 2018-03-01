Attribute VB_Name = "getJobs"
'getJobs Uses the R:\Data\Jobfiles to get all folders begingin with 46XXXX
'Returns list of folder names
Public lst As Range

Function genJobsList() As Collection
    Application.DisplayAlerts = False
    Dim xPath As String
    Dim xFile As String
    Set List = New Collection
    xPath = "R:\Data\Jobfiles\"
    Dim tEmp() As String
    Dim i As Integer
    Dim t As Double
    i = 1
    xFile = Dir(xPath, vbDirectory)
    Do While xFile <> ""
        tEmp = Split(xFile, "-")
        Dim v As Variant
        For Each v In tEmp
            If v = " Shortcut.lnk" Then
                tEmp(0) = "x"
            End If
        Next v
        If IsNumeric(tEmp(0)) Then
            t = CDbl(tEmp(0))
        End If
        If t > 459999 And t < 470000 Then
            List.Add (xFile)
            t = 0
        End If
        xFile = Dir
    Loop
    Set genJobsList = List
    Application.DisplayAlerts = True
End Function

'printList prints the Values of each item in a list in the cell
'referenced by c column 1 row 1

Sub printList(l As Collection, c As Range)
    Dim actSheet As Worksheet
    Set actSheet = ActiveSheet
    Dim v As Variant
    Dim i As Integer
    i = 0
    Stop
    With c.Cells(1, 1)
        For Each v In l
            .Offset(i, 0).Value = v
            i = i + 1
        Next v
    End With
    Set lst = ThisWorkbook.Worksheets("JOBS").UsedRange
    lst.name = "jobList"
    lst.Sort key1:=Worksheets("JOBS").Range("A1"), order1:=xlAscending
    actSheet.Activate
End Sub

'sheetExists checks to see if a sheet is in the current Workbook
Function sheetExists(sheetToFind As String) As Boolean
    sheetExists = False
    For Each Sheet In Worksheets
        If sheetToFind = Sheet.name Then
            sheetExists = True
            Exit Function
        End If
    Next Sheet
End Function

Private Sub get_it()
    printList genJobsList, ThisWorkbook.Worksheets("JOBS").Range("A2")
End Sub
