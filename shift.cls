VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private d As Double
Private phase As Double
Private phase_desc As String
Private units As Double
Private hrs As Double
Private daily_job_desc As String

Private Sub Class_Initialize()
    d = 0
    phase = 0
    phase_desc = vbNullString
    units = 0
    hrs = 0
    daily_job_desc = vbNullString
End Sub

Public Property Let setDayDesc(desc As String)
    daily_job_desc = desc
End Property

Public Property Get getDayDesc() As String
    getDayDesc = daily_job_desc
End Property

Public Property Let setDay(dy As Double)
    d = dy
End Property

Public Property Get getDay() As Double
    getDay = d
End Property

Public Property Let setPhaseDesc(desc As String)
    phase_desc = desc
End Property

Public Property Get getPhaseDesc() As String
    getPhaseDesc = phase_desc
End Property

Public Property Let setUnits(u As Double)
    units = u
End Property

Public Property Get getUnits() As Double
    getUnits = units
End Property

Public Property Let setPhase(p As Double)
    phase = p
End Property

Public Property Get getPhase() As Double
    getPhase = phase
End Property

Public Property Let setHrs(h As Double)
    hrs = h
End Property

Public Property Get getHrs() As Double
    getHrs = hrs
End Property
