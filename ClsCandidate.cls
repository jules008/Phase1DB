VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsCandidate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 24 Aug 16
'===============================================================
' Methods
'---------------------------------------------------------------
' LoadDB - Load Class from DB
' NewDB - Creates a new Class in the DB
' UpdateDB - Update Class to DB
' DeleteDB - Delete Class from DB
' GetParentFromPtr - Private routine to copy memory address of parent class
'===============================================================
Option Explicit
Private pCrewNo As String
Private pName As String
Private pStationNo As String
Private pWCS As Supervisor
Private pDC As Supervisor
Private pDDC1 As Supervisor
Private pDDC2 As Supervisor
Private pDivision As String
Private pStatus As String
Private pAssessments As ClsAssessments
Private pDevelopmentPlans As ClsDevelopmentPlans
Private pDailyLogs As ClsDailyLogs
Private pParent As Long

'-----------------------------------------------------------------
Public Property Get CrewNo() As String
    CrewNo = pCrewNo
End Property
Public Property Let CrewNo(Value As String)
    pCrewNo = Value
End Property

'-----------------------------------------------------------------
Public Property Get Name() As String
    Name = pName
End Property
Public Property Let Name(Value As String)
    pName = Value
End Property

'-----------------------------------------------------------------
Public Property Get StationNo() As String
    StationNo = pStationNo
End Property

Public Property Let StationNo(ByVal vNewValue As String)
    pStationNo = vNewValue
End Property

'-----------------------------------------------------------------
Public Property Get WCS() As Supervisor
    Dim RngTarget As Range
    Dim RngLookUp As Range
    
    Set RngLookUp = ShtLists.Range("Stations")
    
    Set RngTarget = RngLookUp.Find(pStationNo)
    
    With pWCS
        .UserName = RngTarget.Offset(0, 1)
        .Role = "WCS"
    End With
    
    WCS = pWCS
End Property

'-----------------------------------------------------------------
Public Property Get DC() As Supervisor
    Dim RngTarget As Range
    Dim RngLookUp As Range
    
    Set RngLookUp = ShtLists.Range("Divisions")
    
    Set RngTarget = RngLookUp.Find(pDivision)
    
    With pDC
        .UserName = RngTarget.Offset(0, 1)
        .Role = "DC"
    End With
    
    DC = pDC
End Property

'-----------------------------------------------------------------
Public Property Get DDC1() As Supervisor
    Dim RngTarget As Range
    Dim RngLookUp As Range
    
    Set RngLookUp = ShtLists.Range("Divisions")
    
    Set RngTarget = RngLookUp.Find(pDivision)
    
    With pDDC1
        .UserName = RngTarget.Offset(0, 2)
        .Role = "DDC1"
    End With
    
    DDC1 = pDDC1
End Property

'-----------------------------------------------------------------
Public Property Get DDC2() As Supervisor
    Dim RngTarget As Range
    Dim RngLookUp As Range
    
    Set RngLookUp = ShtLists.Range("Divisions")
    
    Set RngTarget = RngLookUp.Find(pDivision)
    
    With pDDC2
        .UserName = RngTarget.Offset(0, 3)
        .Role = "DDC2"
    End With
    
    DDC2 = pDDC2
End Property

'-----------------------------------------------------------------
Public Property Get Division() As String
    Division = pDivision
End Property

Public Property Let Division(ByVal vNewValue As String)
    pDivision = vNewValue
End Property

'-----------------------------------------------------------------
Public Property Get Status() As String
    Status = pStatus
End Property

Public Property Let Status(ByVal vNewValue As String)
    pStatus = vNewValue
End Property

'-----------------------------------------------------------------
Public Property Get Assessments() As ClsAssessments
    Set Assessments = pAssessments
End Property

Public Property Let Assessments(ByVal vNewValue As ClsAssessments)
    Set pAssessments = vNewValue
End Property

'-----------------------------------------------------------------
Public Property Get Dailylogs() As ClsDailyLogs
    Set Dailylogs = pDailyLogs
End Property

Public Property Let Dailylogs(ByVal vNewValue As ClsDailyLogs)
    Set pDailyLogs = vNewValue
End Property

'-----------------------------------------------------------------
Public Property Get DevelopmentPlans() As ClsDevelopmentPlans
    Set DevelopmentPlans = pDevelopmentPlans
End Property

Public Property Let DevelopmentPlans(ByVal vNewValue As ClsDevelopmentPlans)
    Set pDevelopmentPlans = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Parent() As ClsCourse
    If pParent <> 0 Then
        Set Parent = GetParentFromPtr(pParent)
    End If
End Property

Friend Function SetParent(ByVal Ptr As Long) As Boolean
    pParent = Ptr
End Function

'===============================================================
'Method LoadDB
'Load Class from DB
'---------------------------------------------------------------
Public Sub LoadDB()
    Dim RstCandidate As Recordset
    Dim StrCrewNo As String
        
    StrCrewNo = "'" & pCrewNo & "'"
    
    Set RstCandidate = ModDatabase.SQLQuery("SELECT * FROM candidate WHERE crewno = " & StrCrewNo & " AND deleted IS NULL")
    
    With RstCandidate
        If .RecordCount <> 0 Then
            If Not IsNull(!Name) Then pName = !Name
            If Not IsNull(!CrewNo) Then pCrewNo = !CrewNo
            If Not IsNull(!Status) Then pStatus = !Status
            If Not IsNull(!StationNo) Then pStationNo = !StationNo
            If Not IsNull(!Division) Then pDivision = !Division
        
            pAssessments.GetCollection
            pDailyLogs.GetCollection
            pDevelopmentPlans.GetCollection
        End If
        Set RstCandidate = Nothing
    End With
End Sub

'===============================================================
'Method NewDB
' Creates a new Class in the DB
'---------------------------------------------------------------
Public Sub NewDB()
    Dim RstTrgSession As Recordset
    Dim StrCrewNo As String
    
    StrCrewNo = "'" & pCrewNo & "'"
    
    Set RstTrgSession = ModDatabase.SQLQuery("SELECT * FROM Candidate WHERE CrewNo = " & StrCrewNo)
        
    With RstTrgSession
        If .RecordCount <> 0 Then
            MsgBox "Candidate already has a record in the Database.  This may need undeleting or removing from the database before a new record can be created"
        Else
            .AddNew
            !CrewNo = pCrewNo
            .Update
        End If
    End With
    
    Set RstTrgSession = Nothing
End Sub
'===============================================================
' Method UpdateDB
' Update Candidate to database
'---------------------------------------------------------------
Public Function UpdateDB() As Boolean
    Dim RstCandidate As Recordset
    Dim StrCrewNo As String
    
    StrCrewNo = "'" & pCrewNo & "'"
        
    Set RstCandidate = ModDatabase.SQLQuery("SELECT * FROM candidate WHERE crewno = " & StrCrewNo & " AND deleted IS NULL")

    With RstCandidate
        If .RecordCount = 0 Then
            UpdateDB = False
        Else
            .Edit
            !CrewNo = pCrewNo
            !CourseNo = Parent.CourseNo
            !Division = pDivision
            !Name = pName
            !StationNo = pStationNo
            !Status = pStatus
            .Update
            UpdateDB = True
        End If
    End With
    
    Set RstCandidate = Nothing
End Function

'===============================================================
'Method DeleteDB
' Delete Class from DB
'---------------------------------------------------------------
Public Sub DeleteDB()
    Dim RstCandidate As Recordset
    Dim StrCrewNo As String
    
    StrCrewNo = "'" & pCrewNo & "'"
    
    Set RstCandidate = ModDatabase.SQLQuery("SELECT * FROM candidate WHERE crewno = " & StrCrewNo & " AND deleted IS NULL")

    With RstCandidate
        If .RecordCount <> 0 Then
            .Delete
            pAssessments.DeleteCollection
            pDailyLogs.DeleteCollection
            pDevelopmentPlans.DeleteCollection
        End If
    End With
    Set RstCandidate = Nothing
End Sub

'===============================================================
'Method GetParentFromPtr
' Private routine to copy memory address of parent class
'---------------------------------------------------------------
Private Function GetParentFromPtr(ByVal Ptr As Long) As ClsCourse
    Dim tmp As ClsCourse
    
    CopyMemory tmp, Ptr, 4
    Set GetParentFromPtr = tmp
    CopyMemory tmp, 0&, 4
End Function

'=========================================================================
Private Sub Class_Initialize()
    Set pAssessments = New ClsAssessments
    Set pDevelopmentPlans = New ClsDevelopmentPlans
    Set pDailyLogs = New ClsDailyLogs
    
    pAssessments.SetParent ObjPtr(Me)
    pDevelopmentPlans.SetParent ObjPtr(Me)
    pDailyLogs.SetParent ObjPtr(Me)

End Sub

Private Sub Class_Terminate()
    pAssessments.SetParent 0
    pDevelopmentPlans.SetParent 0
    pDailyLogs.SetParent 0
    
    Set pAssessments = Nothing
    Set pDevelopmentPlans = Nothing
    Set pDailyLogs = Nothing
End Sub
'=========================================================================


