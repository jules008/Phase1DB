VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmDailyLog 
   Caption         =   "Candidate Assessment"
   ClientHeight    =   11685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11445
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
'---------------------------------------------------------------
' Date - 29 Dec 18
'===============================================================
' Methods
'---------------------------------------------------------------
' ShowForm
' PopulateForm
' BtnCancel_Click
' BtnDelete_Click
' BtnOk_Click
' UserForm_Activate
' ValidateData
' ResetForm
' UpdateClass
' UserForm_Initialize
' UserForm_Terminate
'===============================================================
Option Explicit
Private Const StrMODULE As String = "FrmDailyLog"

Private Candidate As ClsCandidate
Private DailyLog As ClsDailyLog
Private FormChanged As Boolean

' Routines =====================================================
'---------------------------------------------------------------
Public Function ShowForm(Optional LocalDailyLog As ClsDailyLog, Optional LocalCandidate) As Boolean

    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
    
    ResetForm
    
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
    FormTerminate
    Terminate
Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
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
    FormTerminate
    Terminate
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

Private Function ValidateData() As Boolean
    
    On Error Resume Next
    
    If Me.CmoAssessors = "" Then
        MsgBox "Please enter an assessor name"
        ValidateData = False
        Exit Function
    End If
    
    If Me.Cmo1 = "0" Then
        MsgBox "Please enter a score for Attitude"
        ValidateData = False
        Exit Function
    End If

    If Me.Cmo2 = "0" Then
        MsgBox "Please enter a score for Practical Ability"
        ValidateData = False
        Exit Function
    End If

    If Me.Cmo3 = "0" Then
        MsgBox "Please enter a score for Knowledge"
        ValidateData = False
        Exit Function
    End If

    If Me.Cmo4 = "0" Then
        MsgBox "Please enter a score for Safety"
        ValidateData = False
        Exit Function
    End If

    ValidateData = True
End Function

Private Sub ResetForm()
    
    On Error Resume Next
    
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
End Sub

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
    FormTerminate
    Terminate
    CalcGrade = 0

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
Private Function UpdateClass() As Boolean
   Const StrPROCEDURE As String = "UpdateClass()"

   On Error GoTo ErrorHandler

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
    End With

    UpdateClass = True

Exit Function

ErrorExit:
    FormTerminate
    Terminate
    UpdateClass = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

Private Sub BtnDelete_Click()
    
   Const StrPROCEDURE As String = "BtnDelete_Click()"
    
   On Error GoTo ErrorHandler
    
    Dim Response As Integer
    
    Response = MsgBox("Are you sure you want to delete the Daily Log?", vbYesNo)
    
    If Response = 6 Then
        DailyLog.DeleteDB
        Candidate.Dailylogs.RemoveItem CStr(DailyLog.Module.DayNo)
        ResetForm
        
        Set DailyLog = Nothing
        
        MsgBox "Daily Log has been deleted"
        Hide
    End If
Exit Sub

ErrorExit:
    FormTerminate
    Terminate
    Set DailyLog = Nothing


Exit Sub

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' Events =====================================================
'-------------------------------------------------------------
Private Sub BtnUpdate_Click()
   
   Const StrPROCEDURE As String = "BtnUpdate_Click()"
    
    Dim Success As Boolean
            
   On Error GoTo ErrorHandler
            
    If ValidateData Then
        
        'calculate grade
        TxtOverallGrade = CalcGrade(Cmo1, Cmo2, Cmo3, Cmo4)
        If TxtOverallGrade = 0 Then Err.Raise HANDLED_ERROR
        
        If Not UpdateClass Then Err.Raise HANDLED_ERROR
        
        With DailyLog
            Success = .UpdateDB
            
            If Success = False Then
                .NewDB
                .UpdateDB
            End If
        End With
        Selection = CInt(TxtOverallGrade)
        Me.Hide
        
        If Not FrmGrade.ShowForm(DailyLog) Then Err.Raise HANDLED_ERROR
        FormTerminate
        
    End If

Exit Sub
ErrorExit:
    FormTerminate
    Terminate

Exit Sub

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

Private Sub BtnGuidanced1_Click()
    FrmGuidance_1_.Show
End Sub

Private Sub BtnGuidanced2_Click()
    FrmGuidance_2_.Show
End Sub

Private Sub BtnGuidanced3_Click()
    FrmGuidance_3_.Show
End Sub

Private Sub BtnGuidanced4_Click()
  FrmGuidance_4_.Show
End Sub

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

Private Sub BtnXtraTrng_Click()
    
    Const StrPROCEDURE As String = "BtnXtraTrng_Click()"

    On Error GoTo ErrorHandler
    
    If Not FrmExtraTraining.ShowForm(DailyLog) Then Err.Raise HANDLED_ERROR

Exit Sub

ErrorExit:
    FormTerminate
    Terminate

Exit Sub

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub
Private Sub CmdClose_Click()
   
   Const StrPROCEDURE As String = "CmdClose_Click()"
    
   On Error GoTo ErrorHandler
    
    Dim Response As Integer
    
    If FormChanged = True Then
        Response = MsgBox("The form has been changed, would you like to save these changes?", vbYesNo)
        
        If Response = 6 Then BtnUpdate_Click
        FormChanged = False
    End If
    Candidate.Dailylogs.CleanUp
    FormTerminate

    Me.Hide
    
Exit Sub
ErrorExit:
    FormTerminate
    Terminate
    FormChanged = False
    Me.Hide

Exit Sub

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

Private Sub Cmo1_Change()
    FormChanged = True
End Sub

Private Sub Cmo2_Change()
    FormChanged = True
End Sub

Private Sub Cmo3_Change()
    FormChanged = True
End Sub

Private Sub Cmo4_Change()
    FormChanged = True
End Sub

Private Sub CmoAssessors_Change()
    FormChanged = True
End Sub

Private Sub TxtAssessDate_Click()
    FormChanged = True
End Sub

Private Sub TxtComments1_Change()
    FormChanged = True
End Sub

Private Sub TxtComments2_Change()
    FormChanged = True
End Sub

Private Sub TxtComments3_Change()
    FormChanged = True
End Sub

Private Sub TxtComments4_Change()
    FormChanged = True
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    FormActivate
End Sub

Private Sub UserForm_Initialize()
    On Error Resume Next
    FormInitialise
End Sub

Private Sub UserForm_Terminate()
    FormTerminate
End Sub

Public Sub FormActivate()
   
   Const StrPROCEDURE As String = "FormActivate()"
    
    Dim RstUsers As Recordset
    
    On Error GoTo ErrorHandler
    
    Set RstUsers = GetAccessList
    
    Select Case DailyLog.Module.DayNo
        
        Case 3, 11, 17, 20, 9, 27, 28, 29
            LblAssess = "Assessment"
        
        Case Else
            LblAssess = "Daily Log"
    End Select

    Me.CmoAssessors.Clear
    
    With RstUsers
        Do
        Me.CmoAssessors.AddItem !UserName
        .MoveNext
        Loop While Not .EOF
    End With
    Set RstUsers = Nothing
Exit Sub

ErrorExit:
    Set RstUsers = Nothing
    FormTerminate
    Terminate

Exit Sub

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub


Public Sub FormInitialise()
    
    Const StrPROCEDURE As String = "FormInitialise()"
    
    Dim cell As Range
    Dim screenheight As Integer
    
    On Error GoTo ErrorHandler
    
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
   
    'set form height dependant on screen size
    screenheight = GetScreenHeight
      
    If screenheight < 900 Then
        Me.Height = screenheight - 300
        Me.ScrollHeight = 605
    End If
    
Exit Sub

ErrorExit:
    FormTerminate
    Terminate

Exit Sub

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub


Public Sub FormTerminate()
    On Error Resume Next
    
    Candidate.Dailylogs.CleanUp
    Set Candidate = Nothing
    Set DailyLog = Nothing

End Sub
