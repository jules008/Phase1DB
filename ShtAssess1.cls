VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtAssess1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 03 Nov 16
'===============================================================
Option Explicit

Private Const StrMODULE As String = "ShtAssess"

Dim Course As ClsCourse
Private Sub BtnAddAssess_Click()
    
    Const StrPROCEDURE As String = "BtnAddAssess_Click()"

    On Error GoTo ErrorHandler
    
    If Courses Is Nothing Then
        If Not Initialise Then Err.Raise HANDLED_ERROR
    End If
    
    If Course Is Nothing Then
        If Not SetActiveCourse Then Err.Raise HANDLED_ERROR
    End If
    
    If Not FrmAssessment.ShowForm(Course) Then Err.Raise HANDLED_ERROR
        
Exit Sub

ErrorExit:

Exit Sub

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

Private Sub Worksheet_Activate()
    Const StrPROCEDURE As String = "Worksheet_Activate()"
    
    On Error GoTo ErrorHandler
    
    If Not WorksheetActivate Then Err.Raise HANDLED_ERROR
Exit Sub

ErrorExit:

Exit Sub

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

Public Function WorksheetActivate() As Boolean
    Const StrPROCEDURE As String = "WorksheetActivate()"
    
    Dim CourseNo As String
       
    On Error GoTo ErrorHandler
    
'    Me.Activate
    If Not SetActiveCourse Then Err.Raise HANDLED_ERROR
    If Not PopulateSheet Then Err.Raise HANDLED_ERROR
        
    WorksheetActivate = True
    
Exit Function

ErrorExit:
    WorksheetActivate = False
    
Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

Public Function SetActiveCourse() As Boolean
    
    Const StrPROCEDURE As String = "SetActiveCourse()"
    
    On Error GoTo ErrorHandler
    
    If Courses Is Nothing Then
        If Not Initialise Then Err.Raise HANDLED_ERROR
    End If
    
    If Courses.Count <> 0 Then
        If [CourseNo] <> "" Then
            Set Course = Courses.FindItem(CStr([CourseNo]))
        Else
            Set Course = Courses.FindItem(1)
        End If
    Else
        ClearSheet
    End If

    SetActiveCourse = True

Exit Function

ErrorExit:
    SetActiveCourse = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

Public Function PopulateSheet() As Boolean
    Const StrPROCEDURE As String = "PopulateSheet()"
    
    Dim TargetRange As Range
    Dim TargetRange2 As Range
    Dim i As Integer
    Dim x As Integer
    Dim Candidate As ClsCandidate
    Dim Assessment As ClsAssessment
    
    On Error GoTo ErrorHandler
    
    Set TargetRange = Range("B7")
    
    Application.ScreenUpdating = False
    
    ClearSheet
    
    If Courses Is Nothing Then
        If Not Initialise Then Err.Raise HANDLED_ERROR
    End If
    
    If Course Is Nothing Then
        If Not SetActiveCourse Then Err.Raise HANDLED_ERROR
    End If
    
    If Courses.Count <> 0 Then
            
        'populate candidates
        With Course.Candidates
            For i = 1 To .Count
                Me.Unprotect
                Set Candidate = .FindItem(i)
                TargetRange.Offset(i - 1, 0) = Candidate.crewno
                TargetRange.Offset(i - 1, 1) = Candidate.Name
                
                With Candidate.Assessments
                    For x = 1 To .Count
                        Set Assessment = .FindItem(x)
                        Set TargetRange2 = GetAssessmentRange(Candidate.crewno, Assessment.Module.DayNo, Assessment.AssessType, Assessment.Attempt)
                        
                        If TargetRange2 Is Nothing Then Err.Raise HANDLED_ERROR
                        
                        TargetRange2.Select
                        TargetRange2.Value = Assessment.Score
                    Next
                End With
                TargetRange.Offset(i - 1, 87) = Candidate.Assessments.AverageScore
                TargetRange.Offset(i - 1, 88) = Candidate.DevelopmentPlans.Count
                Me.Protect
           Next
        End With
    
        'sort by average score
        With Range("B7:CL35")
            Me.Unprotect
            .Sort key1:=.Range("Cj9"), order1:=xlDescending
            Me.Protect
        End With
    End If
    
    Application.ScreenUpdating = True
    Set Candidate = Nothing
    Set Assessment = Nothing
    Set TargetRange = Nothing
    Set TargetRange2 = Nothing
    PopulateSheet = True
    
Exit Function

ErrorExit:
    Me.Protect
    Application.ScreenUpdating = True
    Set Candidate = Nothing
    Set Assessment = Nothing
    Set TargetRange = Nothing
    Set TargetRange2 = Nothing
    PopulateSheet = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

Public Sub ClearSheet()
    
    On Error Resume Next
    Me.Unprotect
    Range("B7:CL35").ClearContents
    Me.Protect
End Sub


Private Sub BtnEditAssess_Click()
    Const StrPROCEDURE As String = "BtnEditAssess_Click()"

    On Error GoTo ErrorHandler

    If Courses Is Nothing Then
        If Not Initialise Then Err.Raise HANDLED_ERROR
    End If
    
    If Course Is Nothing Then
        If Not SetActiveCourse Then Err.Raise HANDLED_ERROR
    End If
    
    If Not FrmEditAssessment.ShowForm(Course) Then Err.Raise HANDLED_ERROR
Exit Sub

ErrorExit:

Exit Sub

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub
Public Function GetAssessmentRange(crewno As String, DayNo As Integer, AssessType As String, Attempt As Integer) As Range
    
    Const StrPROCEDURE As String = "GetAssessmentRange()"

    On Error GoTo ErrorHandler

    Dim RowNo As Integer
    Dim ColNo As Integer
    Dim ColNoOffset As Integer
    Dim RngCrewNoLookUp As Range
    Dim RngDayNoLookUp As Range
    Dim RngTarget As Range
    Dim RngTypeLookUp As Range
    
    Set RngCrewNoLookUp = Range("B7:B35")
    Set RngDayNoLookUp = Range("D3:CJ3")

    'get rowno from crewno
    Set RngTarget = RngCrewNoLookUp.Find(crewno)
    RowNo = RngTarget.Row
    
    'get colno from dayno
    Set RngTarget = RngDayNoLookUp.Find(DayNo)
    RngTarget.Select
    Debug.Print RngTarget.Address
    
    ColNo = RngTarget.Column
    ColNoOffset = ColNo + RngTarget.MergeArea.Cells.Count - 1
    
    'set type look up range to search from the current rowno
    Set RngTypeLookUp = Range(Cells(6, ColNo - 1), Cells(6, ColNoOffset))
    
    Set RngTarget = Nothing
    
    RngTypeLookUp.Select
    Set RngTarget = RngTypeLookUp.Find(What:=AssessType, _
                            LookIn:=xlValues, _
                            LookAt:=xlPart, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)
    
    ColNo = RngTarget.Column
    RngTarget.Select
    
    'offset pointer by number of attempts
    ColNo = ColNo + Attempt - 1
    Set RngTarget = Cells(RowNo, ColNo)
    Set GetAssessmentRange = RngTarget
    
    Set RngCrewNoLookUp = Nothing
    Set RngDayNoLookUp = Nothing
    Set RngTarget = Nothing
    Set RngTypeLookUp = Nothing

Exit Function

ErrorExit:
    Set RngCrewNoLookUp = Nothing
    Set RngDayNoLookUp = Nothing
    Set RngTarget = Nothing
    Set RngTypeLookUp = Nothing
    
    Set GetAssessmentRange = Nothing

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
Private Sub Worksheet_Deactivate()
    Set Course = Nothing
End Sub
