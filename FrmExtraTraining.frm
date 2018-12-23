VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmExtraTraining 
   Caption         =   "Action Plan"
   ClientHeight    =   9675
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11955
   OleObjectBlob   =   "FrmExtraTraining.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmExtraTraining"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'===============================================================
' v0,0 - Initial version
' v0,1 - WT2018 Version
'---------------------------------------------------------------
' Date - 23 Dec 18
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmExtraTraining"

Private Candidate As ClsCandidate
Private DailyLog As ClsDailyLog
Private ActiveSession As ClsXTrainingSession

' ===============================================================
' ShowForm
' Shows Extra Training Form
' ---------------------------------------------------------------
Public Function ShowForm(Optional ExistDailyLog As ClsDailyLog) As Boolean
    
    Const StrPROCEDURE As String = "ShowForm()"

    On Error GoTo ErrorHandler

    If Not ResetForm Then Err.Raise HANDLED_ERROR
    
    If ExistDailyLog Is Nothing Then
        Set DailyLog = New ClsDailyLog
        Show
    Else
        Set Candidate = ExistDailyLog.Parent
        Set DailyLog = ExistDailyLog
        
        If Not PopulateForm Then Err.Raise HANDLED_ERROR
        
        Show
    End If
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
' BtnClose_Click
' Closes form
' ---------------------------------------------------------------
Private Sub BtnClose_Click()
    On Error Resume Next
    Unload Me
End Sub

' ===============================================================
' BtnDelete_Click
' Event process for the Delete button
' ---------------------------------------------------------------
Private Sub BtnDelete_Click()
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "BtnDelete_Click()"

    On Error GoTo ErrorHandler

Restart:

    If Course Is Nothing Then Err.Raise SYSTEM_RESTART
    
    DailyLog.XtrainingSessions.RemoveItem (CStr(ActiveSession.ExtraTrainingNo))
    ActiveSession.DeleteDB
    LstTrainingList.ListIndex = -1
    
    If Not PopulateForm Then Err.Raise HANDLED_ERROR

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
' BtnNew_Click
' Event process for New Button
' ---------------------------------------------------------------
Private Sub BtnNew_Click()
    Dim NewSession As ClsXTrainingSession
    Dim ExtraTrainingNo As Integer
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "BtnNew_Click()"

    On Error GoTo ErrorHandler

Restart:

    If Course Is Nothing Then Err.Raise SYSTEM_RESTART
    
    Set NewSession = New ClsXTrainingSession
    
    NewSession.ExtraTrainingNo = NewSession.NewDB
    DailyLog.XtrainingSessions.AddItem NewSession
    
    If Not PopulateForm Then Err.Raise HANDLED_ERROR
    
    With LstTrainingList
        .Selected(.ListCount - 1) = True
    End With

GracefulExit:
    
    Set NewSession = Nothing
Exit Sub

ErrorExit:
    Set NewSession = Nothing

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
' BtnSpellChk_Click
' Event process for spell checker
' ---------------------------------------------------------------
Private Sub BtnSpellChk_Click()
    Dim Cntrls As Collection
    Dim i As Integer
    Dim Cntrl As Control
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "BtnSpellChk_Click()"

    On Error GoTo ErrorHandler

Restart:

    If Course Is Nothing Then Err.Raise SYSTEM_RESTART
    
    Set Cntrls = New Collection
    
    For i = 0 To Me.Controls.Count - 1
        Cntrls.Add Controls(i)
    Next
    
    ModLibrary.SpellCheck Cntrls
    
    Set Cntrls = Nothing

GracefulExit:

Exit Sub

ErrorExit:

    Set Cntrls = Nothing

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
' Event Process for Update button
' ---------------------------------------------------------------
Private Sub BtnUpdate_Click()
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "BtnUpdate_Click()"

    On Error GoTo ErrorHandler

Restart:

    If Course Is Nothing Then Err.Raise SYSTEM_RESTART
    
    Select Case ValidateData
        
        Case Is = ForkOK
            With ActiveSession
            
                If TxtTrainingDate <> "" Then .TrainingDate = TxtTrainingDate
                
                .TrainingDetails = TxtTrainingDetails
                .TrainingResults = TxtTrainingResults
                
                If OptTrngAccpted.Value = True Then
                    .TrainingTaken = True
                Else
                    .TrainingTaken = False
                End If
                
                ActiveSession.UpdateDB
                MsgBox "Training Session Updated"
                
                If Not PopulateForm Then Err.Raise HANDLED_ERROR
                
            End With
        
        Case Is = FunctionalError
            Err.Raise HANDLED_ERROR, , "Functional Error in Validation"
        
        Case Is = ValidationError
            GoTo GracefulExit
            
    End Select

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
' LstTrainingList_Click
' Event Process for Training List Click
' ---------------------------------------------------------------
Private Sub LstTrainingList_Click()
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "LstTrainingList_Click()"

    On Error GoTo ErrorHandler

Restart:
    
    If Course Is Nothing Then Err.Raise SYSTEM_RESTART
    If Not SetActiveSession Then Err.Raise HANDLED_ERROR
    If Not PopulateTrainingDetails Then Err.Raise HANDLED_ERROR

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
' UserForm_Initialize
' Form Start up process
' ---------------------------------------------------------------
Private Sub UserForm_Initialize()
    On Error Resume Next
    
    Set Candidate = New ClsCandidate
End Sub

' ===============================================================
' UserForm_Terminate
' Form Close down process
' ---------------------------------------------------------------
Private Sub UserForm_Terminate()
    On Error Resume Next
    
    DailyLog.XtrainingSessions.CleanUp
    Set Candidate = Nothing
    Set DailyLog = Nothing
    Set ActiveSession = Nothing
End Sub

' ===============================================================
' PopulateForm
' Polulates data in the form
' ---------------------------------------------------------------
Private Function PopulateForm() As Boolean
    Dim i As Integer
    Dim ListIndex As Integer
    Dim TrainingDate As String
    Dim TrainingTaken As String
    Dim TrainingSession As ClsXTrainingSession
    
    Const StrPROCEDURE As String = "PopulateForm()"
    On Error GoTo ErrorHandler

    Set TrainingSession = New ClsXTrainingSession
    
    'candidate details
    With Candidate
        TxtCourseNo = .Parent.CourseNo
        TxtCrewNo = .CrewNo
        TxtName = .Name
        TxtModuleNo = DailyLog.Module.ModuleNo
        TxtDayNo = DailyLog.Module.DayNo
        TxtModuleDesc = DailyLog.Module.Module
    End With
    
    'list heading
    With LstHeadings
        .Clear
        .AddItem
        .List(0, 0) = ""
        .List(0, 1) = "No"
        .List(0, 2) = "Date"
        .List(0, 3) = "Training Taken"
    End With
    
    'list contents
    With LstTrainingList
        ListIndex = .ListIndex
            
        .Clear
        If DailyLog.XtrainingSessions.Count = 0 Then
        
            If Not DisableEntryForm Then Err.Raise HANDLED_ERROR
        Else
            
            For i = 1 To DailyLog.XtrainingSessions.Count
                            
                Set TrainingSession = DailyLog.XtrainingSessions.FindItem(i)
                
                If TrainingSession.TrainingTaken = True Then
                    TrainingTaken = "Yes"
                    TrainingDate = Format(TrainingSession.TrainingDate, "dd/mm/yy")
                Else
                    TrainingTaken = "No"
                    TrainingDate = ""
                End If
                
                .AddItem
                .List(i - 1, 0) = TrainingSession.ExtraTrainingNo
                .List(i - 1, 1) = i
                .List(i - 1, 2) = TrainingDate
                .List(i - 1, 3) = TrainingTaken
            Next
            

        End If
        If ListIndex <> -1 Then .Selected(ListIndex) = True
    End With
    Set TrainingSession = Nothing
    PopulateForm = True

Exit Function

ErrorExit:

    Set TrainingSession = Nothing
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
' PopulateTrainingDetails
' Populates training details on form
' ---------------------------------------------------------------
Private Function PopulateTrainingDetails() As Boolean
    Const StrPROCEDURE As String = "PopulateTrainingDetails()"

    On Error GoTo ErrorHandler

    With ActiveSession
    
        If .TrainingDate <> 0 Then TxtTrainingDate = Format(.TrainingDate, "dd/mm/yy")
        
        TxtTrainingDetails = .TrainingDetails
        TxtTrainingResults = .TrainingResults
        
        If .TrainingTaken = True Then
            OptTrngAccpted.Value = True
        Else
            OptTrngRejected.Value = True
        End If
    
    End With
    
    If Not ActiveSession Is Nothing Then
        EnableFormEntry
    Else
        DisableEntryForm
    End If

    PopulateTrainingDetails = True

Exit Function

ErrorExit:

    PopulateTrainingDetails = False

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
' Validates form data before updating DB
' ---------------------------------------------------------------
Private Function ValidateData() As EnumFormValidation
    Const StrPROCEDURE As String = "ValidateData()"

    On Error GoTo ErrorHandler

    If Me.TxtTrainingDetails = "" Then
        MsgBox "Please enter details of the training offered"
        ValidateData = ValidationError
        Exit Function
    End If
    
    If TxtTrainingDate <> "" And Not IsDate(Me.TxtTrainingDate) Then
        MsgBox "Please enter a valid date"
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
' SetActiveSession
' Sets the active x training session when it is selected
' ---------------------------------------------------------------
Private Function SetActiveSession() As Boolean
    Dim ListSel As Integer
    Dim Index As Integer
    
    Const StrPROCEDURE As String = "SetActiveSession()"

    On Error GoTo ErrorHandler
    
    If ActiveSession Is Nothing Then Set ActiveSession = New ClsXTrainingSession
    
    ListSel = LstTrainingList.ListIndex
    
    If ListSel <> -1 Then
        Index = LstTrainingList.List(ListSel, 0)
        Set ActiveSession = DailyLog.XtrainingSessions.FindItem(CStr(Index))
    End If
    SetActiveSession = True

Exit Function

ErrorExit:

    SetActiveSession = False

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

    TxtCourseNo = ""
    TxtCrewNo = ""
    TxtDayNo = ""
    TxtModuleDesc = ""
    TxtModuleNo = ""
    TxtName = ""
    TxtTrainingDate = ""
    TxtTrainingDetails = ""

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
' DisableEntryForm
' disables the form to prevent input
' ---------------------------------------------------------------
Private Function DisableEntryForm() As Boolean
    Const StrPROCEDURE As String = "DisableEntryForm()"

    On Error GoTo ErrorHandler

    With TxtTrainingDetails
        .Enabled = False
        .Value = ""
        .BackColor = RGB(211, 213, 212)
    End With
    
    With TxtTrainingDate
        .Enabled = False
        .Value = ""
        .BackColor = RGB(211, 213, 212)
    End With
    
    With TxtTrainingResults
        .Enabled = False
        .Value = ""
        .BackColor = RGB(211, 213, 212)
    End With
    
    OptTrngAccpted.Enabled = False
    OptTrngRejected.Enabled = False
    BtnDelete.Enabled = False
    BtnUpdate.Enabled = False
    
    DisableEntryForm = True

Exit Function

ErrorExit:

    DisableEntryForm = False

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
' EnableFormEntry
' Enables form to allow data entry
' ---------------------------------------------------------------
Private Function EnableFormEntry() As Boolean
    Const StrPROCEDURE As String = "EnableFormEntry()"

    On Error GoTo ErrorHandler

    With TxtTrainingDetails
        .Enabled = True
        .BackColor = RGB(255, 255, 255)
    End With
    
    With TxtTrainingDate
        .Enabled = True
        .BackColor = RGB(255, 255, 255)
    End With
    
    With TxtTrainingResults
        .Enabled = True
        .BackColor = RGB(255, 255, 255)
    End With
    
    OptTrngAccpted.Enabled = True
    OptTrngRejected.Enabled = True
    BtnDelete.Enabled = True
    BtnUpdate.Enabled = True

    EnableFormEntry = True

Exit Function

ErrorExit:
    
    EnableFormEntry = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
