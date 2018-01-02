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
'---------------------------------------------------------------
' Date - 14 Sep 16
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmExtraTraining"

Private Candidate As ClsCandidate
Private DailyLog As ClsDailyLog
Private ActiveSession As ClsXTrainingSession

Public Function ShowForm(Optional ExistDailyLog As ClsDailyLog, Optional ShowAll As Boolean) As Boolean
    
   Const StrPROCEDURE As String = "ShowForm()"
   
   On Error GoTo ErrorHandler
   
    ResetForm
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
    FormTerminate
    Terminate
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

Private Sub BtnClose_Click()
    On Error Resume Next
    
    FormTerminate
    Me.Hide
End Sub

Private Sub BtnDatePicker_Click()
    On Error Resume Next
    
    FrmDatePicker.Show
    
    Me.TxtTrainingDate = FrmDatePicker.Tag

End Sub

Private Sub BtnDelete_Click()
    Const StrPROCEDURE As String = "BtnDelete_Click()"
    
    On Error GoTo ErrorHandler
    
    If Not SetActiveSession Then Err.Raise HANDLED_ERROR
    
    DailyLog.XtrainingSessions.RemoveItem (CStr(ActiveSession.ExtraTrainingNo))
    ActiveSession.DeleteDB
    LstTrainingList.ListIndex = -1
    
    If Not PopulateForm Then Err.Raise HANDLED_ERROR
    
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


Private Sub BtnNew_Click()
    Const StrPROCEDURE As String = "BtnNew_Click()"
    
    On Error GoTo ErrorHandler
    
    Dim NewSession As ClsXTrainingSession
    Dim ExtraTrainingNo As Integer
    
    Set NewSession = New ClsXTrainingSession
    
    NewSession.ExtraTrainingNo = NewSession.NewDB
    DailyLog.XtrainingSessions.AddItem NewSession
    
    If Not PopulateForm Then Err.Raise HANDLED_ERROR
    
    With LstTrainingList
        .Selected(.ListCount - 1) = True
    End With
    Set NewSession = Nothing
Exit Sub

ErrorExit:
    Set NewSession = Nothing
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
Private Sub BtnSpellChk_Click()
    On Error Resume Next
    
    'to be added to selection button
    Dim Cntrls As Collection
    Dim i As Integer
    Dim Cntrl As Control
    
    Set Cntrls = New Collection
    
    For i = 0 To Me.Controls.Count - 1
        Cntrls.Add Controls(i)
    Next
    Library.SpellCheck Cntrls
    Set Cntrls = Nothing
End Sub
Private Sub BtnUpdate_Click()
    Const StrPROCEDURE As String = "BtnUpdate_Click()"

    On Error GoTo ErrorHandler
    
    If ValidateData = True Then
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
Private Sub LstTrainingList_Click()
    Const StrPROCEDURE As String = "LstTrainingList_Click()"
    
    On Error GoTo ErrorHandler
    
    If Not SetActiveSession Then Err.Raise HANDLED_ERROR
    If Not PopulateTrainingDetails Then Err.Raise HANDLED_ERROR
    
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
Private Sub UserForm_Initialize()
    On Error Resume Next
    
    Set Candidate = New ClsCandidate
End Sub

Private Sub UserForm_Terminate()
    On Error Resume Next
    FormTerminate
End Sub

Private Function PopulateForm(ShowAll As Boolean) As Boolean
    Const StrPROCEDURE As String = "PopulateForm()"
    
    Dim i As Integer
    Dim ListIndex As Integer
    Dim TrainingDate As String
    Dim TrainingTaken As String
    Dim TrainingSession As ClsXTrainingSession
    
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
    
    Set TrainingSession = Nothing
    PopulateForm = True

Exit Function

ErrorExit:
    FormTerminate
    Terminate
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
    FormTerminate
    Terminate
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

Private Function ValidateData() As Boolean
    On Error Resume Next
    
    If Me.TxtTrainingDetails = "" Then
        MsgBox "Please enter details of the training offered"
        ValidateData = False
        Exit Function
    End If
    
    ValidateData = True
End Function

Private Function SetActiveSession() As Boolean
    Const StrPROCEDURE As String = "SetActiveSession()"
    
    On Error GoTo ErrorHandler
    
    Dim ListSel As Integer
    Dim Index As Integer
    
    If ActiveSession Is Nothing Then Set ActiveSession = New ClsXTrainingSession
    
    ListSel = LstTrainingList.ListIndex
    
    If ListSel <> -1 Then
        Index = LstTrainingList.List(ListSel, 0)
        Set ActiveSession = DailyLog.XtrainingSessions.FindItem(CStr(Index))
    End If
    SetActiveSession = True

Exit Function

ErrorExit:
    FormTerminate
    Terminate
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

Private Sub ResetForm()
    On Error Resume Next
    
    TxtCourseNo = ""
    TxtCrewNo = ""
    TxtDayNo = ""
    TxtModuleDesc = ""
    TxtModuleNo = ""
    TxtName = ""
    TxtTrainingDate = ""
    TxtTrainingDetails = ""

End Sub


Public Sub FormTerminate()
    On Error Resume Next
    
    DailyLog.XtrainingSessions.CleanUp
    Set Candidate = Nothing
    Set DailyLog = Nothing
    Set ActiveSession = Nothing
End Sub

Public Sub DisableEntryForm()
    On Error Resume Next
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

End Sub

Public Sub EnableFormEntry()
    On Error Resume Next
    
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

End Sub

Public Function PopulateList(ShowAll As Boolean) As Boolean

    'list contents
    With LstTrainingList
        ListIndex = .ListIndex
            
        .Clear
        If DailyLog.XtrainingSessions.Count = 0 Then
        
            DisableEntryForm
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

End Function
