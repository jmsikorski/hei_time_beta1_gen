VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Fname As String
Private lName As String
Private num As Integer
Private eClass As String
Private ePerDiem As Boolean
Private shifts As Collection
Private lead As Integer
Private tPerDiem(7) As Double

Public Property Let eLead(i As Integer)
    lead = i
End Property

Public Property Let efName(s As String)
    Fname = s
End Property

Public Property Let elName(s As String)
    lName = s
End Property

Public Property Let emNum(s As Integer)
    num = s
End Property

Public Property Let emClass(s As String)
    eClass = s
End Property

Public Sub addShift(s As shift)
    shifts.Add s
End Sub

Private Sub calcPerDiem()
    Dim s As shift
    Dim n As Double
    For Each s In shifts
        n = s.getPhase
        If n <> 88081 And n <> 88086 And n <> -1 And n <> 16950 Then
            tPerDiem(s.getDay) = tPerDiem(s.getDay) + s.getHrs
        End If
    Next s
    For i = 0 To 6
        If tPerDiem(i) >= 6 Then
            tPerDiem(i) = 100
        Else
            tPerDiem(i) = 0
        End If
    Next i
End Sub

Public Property Get getCalcPerDiem() As Double
    Dim amt As Double
    amt = 0
    If ePerDiem Then
        calcPerDiem
        For i = 0 To 6
            amt = amt + tPerDiem(i)
        Next i
        getCalcPerDiem = amt
    Else
        getCalcPerDiem = amt
    End If
End Property

Public Property Get getLead() As Integer
    getLead = lead
End Property

Public Property Get getFName() As String
    getFName = Fname
End Property

Public Property Get getLName() As String
    getLName = lName
End Property

Public Property Get getNum() As Integer
    getNum = num
End Property

Public Property Get getClass() As String
    getClass = eClass
End Property

Public Property Get getShifts() As Collection
    Set getShifts = shifts
End Property

Public Property Let emPerDiem(pd As Boolean)
    ePerDiem = pd
End Property

Public Property Get getPerDiem() As Boolean
    getPerDiem = ePerDiem
End Property

Private Sub Class_Initialize()
    Fname = "NONE"
    lName = "NONE"
    num = -1
    eClass = "NONE"
    ePerDiem = False
    For i = 0 To 6
        tPerDiem(i) = 0
    Next i
    lead = -1
    Set shifts = New Collection
End Sub

Public Property Get getFullname() As String
    getFullname = Fname & " " & lName
End Property

Public Property Get newEmployee(ByVal i As Integer) As Integer
    On Error GoTo 10
    Dim rg As Range
    Set rg = ThisWorkbook.Worksheets("ROSTER").Range("A" & i + 2)
        With Me
            .efName = rg.Offset(0, 3).Value
            .elName = rg.Offset(0, 2).Value
            .emNum = rg.Offset(0, 4).Value
            .emClass = rg.Offset(0, 1).Value
            If rg.Offset(0, 5).Value = "YES" Then
                .emPerDiem = True
            Else
                .emPerDiem = False
            End If
        End With
    On Error GoTo 0
    newEmployee = 1
    Exit Property
10
    On Error GoTo 0
    newEmployee = -1
End Property

