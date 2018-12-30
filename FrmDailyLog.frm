VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmDailyLog 
   Caption         =   "Candidate Assessment"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16230
   OleObjectBlob   =   "FrmDailyLog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmDailyLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===============================================================
' v0,0 - Initial version
' v0,1 - Changes to guidance forms
' v0,2 - WT2019 Version
'---------------------------------------------------------------
' Date - 30 Dec 18
'===============================================================
' Methods
'===============================================================
Option Explicit
Private Const StrMODULE As String = "FrmDailyLog"

Private Candidate As ClsCandidate
Private DailyLog As ClsDailyLog
Private FormChanged As Boolean

' ===============================================================
' ShowForm
' Shows Daily Log Form for selected person and day
' ---------------------------------------------------------------
Public Function ShowForm(Optional LocalDailyLog As ClsDailyLog, Optional LocalCandidate) As Boolean

    Const StrPROCEDURE As String = "ShowForm()"

    On Error GoTo ErrorHandler

    If LocalDailyLog Is Nothing Then
        Set DailyLog = New ClsDailyLog
        
        If Not LocalCandidate Is Nothing Then
            Set Candidate = LocalCandidate
        End If
        
        Candidate.Dailylogs.AddItem DailyLog
    Else
        Set DailyLog = LocalDailyLog
        Set Candidate = DailyLog.Parent
    End If
    
    If Not PopulateForm Then Err.Raise HANDLED_ERROR
    
    FormChanged = False
    Show
    ShowForm = True

Exit Function

ErrorExit:
    ShowForm = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' PopulateForm
' Populates Daily Log form
' ---------------------------------------------------------------
Private Function PopulateForm() As Boolean
    
    Const StrPROCEDURE As String = "PopulateForm()"

    On Error GoTo ErrorHandler

    With DailyLog
        If .DLDate = #12:00:00 AM# Then TxtAssessDate = VBA.Format(Now, "dd/mm/yy") Else TxtAssessDate = .DLDate
        TxtComments1 = .Comments1
        TxtComments2 = .Comments2
        TxtComments3 = .Comments3
        TxtComments4 = .Comments4
        TxtCourseNo = Candidate.Parent.CourseNo
        TxtCrewNo = Candidate.CrewNo
        TxtDayNo = DailyLog.Module.DayNo
        TxtDivision = Candidate.Division
        TxtMiscComments = .CommentsMisc
        TxtModuleDesc = DailyLog.Module.Module
        TxtModuleNo = DailyLog.Module.ModuleNo
        TxtName = Candidate.Name
        TxtOverallGrade = .OverallGrade
        TxtStartDate = Candidate.Parent.StartDate
        TxtStationNo = Candidate.StationNo
        CmoAssessors = .Assessor
        Cmo1 = .Score1
        Cmo2 = .Score2
        Cmo3 = .Score3
        Cmo4 = .Score4
    End With

    PopulateForm = True

Exit Function

ErrorExit:

    PopulateForm = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' ValidateData
' Validates input data before saving
' ---------------------------------------------------------------
Private Function ValidateData() As EnumFormValidation
    Const StrPROCEDURE As String = "ValidateData()"

    On Error GoTo ErrorHandler

    If Me.CmoAssessors = "" Then
        MsgBox "Please enter an assessor name"
        ValidateData = ValidationError
        Exit Function
    End If
    
    If Me.Cmo1 = "0" Then
        MsgBox "Please enter a score for Attitude"
        ValidateData = ValidationError
        Exit Function
    End If

    If Me.Cmo2 = "0" Then
        MsgBox "Please enter a score for Practical Ability"
        ValidateData = ValidationError
        Exit Function
    End If

    If Me.Cmo3 = "0" Then
        MsgBox "Please enter a score for Knowledge"
        ValidateData = ValidationError
        Exit Function
    End If

    If Me.Cmo4 = "0" Then
        MsgBox "Please enter a score for Safety"
        ValidateData = ValidationError
        Exit Function
    End If

    ValidateData = FormOK

Exit Function

ErrorExit:

    ValidateData = FunctionalError

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' ResetForm
' Resets form
' ---------------------------------------------------------------
Private Function ResetForm() As Boolean
    Const StrPROCEDURE As String = "ResetForm()"

    On Error GoTo ErrorHandler

    FormChanged = False
    
    Me.TxtAssessDate = ""
    Me.TxtComments1 = ""
    Me.TxtComments2 = ""
    Me.TxtComments3 = ""
    Me.TxtComments4 = ""
    Me.TxtCourseNo = ""
    Me.TxtCrewNo = ""
    Me.TxtDayNo = ""
    Me.TxtDivision = ""
    Me.TxtMiscComments = ""
    Me.TxtModuleDesc = ""
    Me.TxtModuleNo = ""
    Me.TxtName = ""
    Me.TxtOverallGrade = ""
    Me.TxtStartDate = ""
    Me.TxtStationNo = ""
    Me.Cmo1 = ""
    Me.Cmo2 = ""
    Me.Cmo3 = ""
    Me.Cmo4 = ""
    Me.CmoAssessors = ""

    ResetForm = True

Exit Function

ErrorExit:

    ResetForm = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' CalcGrade
' calculates score based on input value.  returns 99 on error
' ---------------------------------------------------------------
Private Function CalcGrade(Score_1 As Integer, Score_2 As Integer, Score_3 As Integer, Score_4 As Integer) As Single
    
    Const StrPROCEDURE As String = "CalcGrade()"
    
    On Error GoTo ErrorHandler
    
    Dim OverAchieve As Integer
    Dim Acheived As Integer
    Dim TotalScore As Single
    Dim UnderAchieve As Integer
    Dim i As Integer
    Dim Scores(1 To 4) As Integer
    
    'add scores to array
    Scores(1) = Score_1
    Scores(2) = Score_2
    Scores(3) = Score_3
    Scores(4) = Score_4
    
    For i = 1 To 4
        If Scores(i) = 1 Then OverAchieve = OverAchieve + 1
        If Scores(i) = 2 Then Acheived = Acheived + 1
        If Scores(i) = 3 Then UnderAchieve = UnderAchieve + 1
        If Scores(i) = 4 Then UnderAchieve = UnderAchieve + 1
        If Scores(i) = 5 Then UnderAchieve = UnderAchieve + 1
        TotalScore = TotalScore + Scores(i)
    Next
    
    If UnderAchieve = 1 Then CalcGrade = 3
    If UnderAchieve > 1 And UnderAchieve < 4 Then CalcGrade = 4
    If UnderAchieve = 4 Then CalcGrade = 5
    
    'if all tests missed, set to 2
    If CalcGrade = 0 Then CalcGrade = TotalScore / 4

Exit Function

ErrorExit:
    
    CalcGrade = 99

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' BtnDelete_Click
' Event process for delete button
' ---------------------------------------------------------------
Private Sub BtnDelete_Click()
    Dim ErrNo As Integer
    Dim Response As Integer

    Const StrPROCEDURE As String = "BtnDelete_Click()"

    On Error GoTo ErrorHandler

Restart:

    If Courses Is Nothing Then Err.Raise SYSTEM_RESTART

    
    Response = MsgBox("Are you sure you want to delete the Daily Log?", vbYesNo)
    
    If Response = 6 Then
        DailyLog.DeleteDB
        Candidate.Dailylogs.RemoveItem CStr(DailyLog.Module.DayNo)
        Set DailyLog = Nothing
        
        If Not ResetForm Then Err.Raise HANDLED_ERROR
        
        MsgBox "Daily Log has been deleted"
        
        Unload Me
        
    End If
    
GracefulExit:

Exit Sub

ErrorExit:

Exit Sub

ErrorHandler:
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        ErrNo = Err.Number
        CustomErrorHandler (Err.Number)
        If ErrNo = SYSTEM_RESTART Then Resume Restart Else Resume GracefulExit
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' BtnUpdate_Click
' Event process for Update button
' ---------------------------------------------------------------
Private Sub BtnUpdate_Click()
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "BtnUpdate_Click()"

    On Error GoTo ErrorHandler

Restart:

    If Courses Is Nothing Then Err.Raise SYSTEM_RESTART

    
    If ValidateData = FunctionalError Then Err.Raise HANDLED_ERROR
    
    If ValidateData = FormOK Then
        
        TxtOverallGrade = CalcGrade(Cmo1, Cmo2, Cmo3, Cmo4)
        
        If TxtOverallGrade = 99 Then Err.Raise HANDLED_ERROR
        
        With DailyLog
            .DLDate = TxtAssessDate
            .Assessor = CmoAssessors
            .Score1 = Me.Cmo1
            .Score2 = Me.Cmo2
            .Score3 = Me.Cmo3
            .Score4 = Me.Cmo4
            .Comments1 = Me.TxtComments1
            .Comments2 = Me.TxtComments2
            .Comments3 = Me.TxtComments3
            .Comments4 = Me.TxtComments4
            .CommentsMisc = Me.TxtMiscComments
            .OverallGrade = TxtOverallGrade
            .UpdateDB
        End With
        
        Selection = CInt(TxtOverallGrade)
        
        Hide
        
        If Not FrmGrade.ShowForm(DailyLog) Then Err.Raise HANDLED_ERROR
        
    End If

GracefulExit:

Exit Sub
ErrorExit:

Exit Sub

ErrorHandler:
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        ErrNo = Err.Number
        CustomErrorHandler (Err.Number)
        If ErrNo = SYSTEM_RESTART Then Resume Restart Else Resume GracefulExit
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' BtnGuidanced1_Click
' Grade guidance
' ---------------------------------------------------------------
Private Sub BtnGuidanced1_Click()
    FrmGuidance_1_.Show
End Sub

' ===============================================================
' BtnGuidanced2_Click
' Grade guidance
' ---------------------------------------------------------------
Private Sub BtnGuidanced2_Click()
    FrmGuidance_2_.Show
End Sub

' ===============================================================
' BtnGuidanced3_Click
' Grade guidance
' ---------------------------------------------------------------
Private Sub BtnGuidanced3_Click()
    FrmGuidance_3_.Show
End Sub

' ===============================================================
' BtnGuidanced4_Click
' Grade guidance
' ---------------------------------------------------------------
Private Sub BtnGuidanced4_Click()
  FrmGuidance_4_.Show
End Sub

' ===============================================================
' BtnSpellChk_Click
' Spell checks form
' ---------------------------------------------------------------
Private Sub BtnSpellChk_Click()
    On Error Resume Next
    
    Dim Cntrls As Collection
    Dim i As Integer
    Dim Cntrl As Control
    
    Set Cntrls = New Collection
    
    For i = 0 To Me.Controls.Count - 1
        Cntrls.Add Controls(i)
    Next
    ModLibrary.SpellCheck Cntrls
    
    Set Cntrls = Nothing
End Sub

' ===============================================================
' BtnXtraTrng_Click
' Event process to show extra training form
' ---------------------------------------------------------------
Private Sub BtnXtraTrng_Click()
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "BtnXtraTrng_Click()"

    On Error GoTo ErrorHandler

Restart:

    If Courses Is Nothing Then Err.Raise SYSTEM_RESTART


    If Not FrmExtraTraining.ShowForm(DailyLog) Then Err.Raise HANDLED_ERROR

GracefulExit:

Exit Sub

ErrorExit:

Exit Sub

ErrorHandler:
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        ErrNo = Err.Number
        CustomErrorHandler (Err.Number)
        If ErrNo = SYSTEM_RESTART Then Resume Restart Else Resume GracefulExit
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' CmdClose_Click
' Event process for Close Button
' ---------------------------------------------------------------
Private Sub CmdClose_Click()
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "CmdClose_Click()"

    On Error GoTo ErrorHandler

Restart:
    
    If Courses Is Nothing Then Err.Raise SYSTEM_RESTART

    Dim Response As Integer
    
    If FormChanged = True Then
        Response = MsgBox("The form has been changed, would you like to save these changes?", vbYesNo)
        
        If Response = 6 Then BtnUpdate_Click
        FormChanged = False
    End If
    
    Candidate.Dailylogs.CleanUp

    Unload Me

GracefulExit:

Exit Sub
ErrorExit:

Exit Sub

ErrorHandler:
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        ErrNo = Err.Number
        CustomErrorHandler (Err.Number)
        If ErrNo = SYSTEM_RESTART Then Resume Restart Else Resume GracefulExit
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' UserForm_Activate
' Form activate event
' ---------------------------------------------------------------
Private Sub UserForm_Activate()
    On Error Resume Next
    FormActivate
End Sub

' ===============================================================
' UserForm_Initialize
' Form initialise event
' ---------------------------------------------------------------
Private Sub UserForm_Initialize()
    On Error Resume Next
    FormInitialise
End Sub

' ===============================================================
' UserForm_Terminate
' Form terminate event
' ---------------------------------------------------------------
Private Sub UserForm_Terminate()
    FormTerminate
End Sub

' ===============================================================
' FormActivate
' Form activate processing
' ---------------------------------------------------------------
Private Sub FormActivate()
    Dim ErrNo As Integer
    Dim RstUsers As Recordset
    
    Const StrPROCEDURE As String = "FormActivate()"

    On Error GoTo ErrorHandler

Restart:

    If Courses Is Nothing Then Err.Raise SYSTEM_RESTART
    
    Set RstUsers = GetAccessList
    
    If DailyLog.Module.Assessment Then
        LblAssess = "Assessment"
    Else
        LblAssess = "Daily Log"
    End If

    Me.CmoAssessors.Clear
    
    With RstUsers
        Do
        Me.CmoAssessors.AddItem !UserName
        .MoveNext
        Loop While Not .EOF
    End With

GracefulExit:
    
    Set RstUsers = Nothing
Exit Sub

ErrorExit:
    Set RstUsers = Nothing

Exit Sub

ErrorHandler:
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        ErrNo = Err.Number
        CustomErrorHandler (Err.Number)
        If ErrNo = SYSTEM_RESTART Then Resume Restart Else Resume GracefulExit
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' FormInitialise
' Event process for form initialise
' ---------------------------------------------------------------
Private Sub FormInitialise()
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "FormInitialise()"

    On Error GoTo ErrorHandler

Restart:

    If Courses Is Nothing Then Err.Raise SYSTEM_RESTART
    
    Set Candidate = New ClsCandidate
    Set DailyLog = New ClsDailyLog
    
    Cmo1.AddItem "0"
    Cmo2.AddItem "0"
    Cmo3.AddItem "0"
    Cmo4.AddItem "0"
    
    Cmo1.AddItem "1"
    Cmo2.AddItem "1"
    Cmo3.AddItem "1"
    Cmo4.AddItem "1"
    
    Cmo1.AddItem "2"
    Cmo2.AddItem "2"
    Cmo3.AddItem "2"
    Cmo4.AddItem "2"
    
    Cmo1.AddItem "3"
    Cmo2.AddItem "3"
    Cmo3.AddItem "3"
    Cmo4.AddItem "3"
    
    Cmo1.AddItem "4"
    Cmo2.AddItem "4"
    Cmo3.AddItem "4"
    Cmo4.AddItem "4"
    
    Cmo1.AddItem "5"
    Cmo2.AddItem "5"
    Cmo3.AddItem "5"
    Cmo4.AddItem "5"
   
Exit Sub

GracefulExit:

Exit Sub

ErrorExit:

Exit Sub

ErrorHandler:
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        ErrNo = Err.Number
        CustomErrorHandler (Err.Number)
        If ErrNo = SYSTEM_RESTART Then Resume Restart Else Resume GracefulExit
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' FormTerminate
' Event process for form terminate
' ---------------------------------------------------------------
Public Sub FormTerminate()
    On Error Resume Next
    
    Candidate.Dailylogs.CleanUp
    Set Candidate = Nothing
    Set DailyLog = Nothing

End Sub
