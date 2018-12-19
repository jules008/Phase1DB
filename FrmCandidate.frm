VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmCandidate 
   Caption         =   "Candidate"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11565
   OleObjectBlob   =   "FrmCandidate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmCandidate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===============================================================
' v0,0 - Initial version
' v0,1 - WT2018 Version
'---------------------------------------------------------------
' Date - 19 Dec 18
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmCandidate"

Private Candidate As ClsCandidate
Private FormChanged As Boolean

' ===============================================================
' ShowForm
' Shows Candidate form and passes Candidate object if available
' ---------------------------------------------------------------
Public Function ShowForm(Optional ExistCandidate As ClsCandidate) As Boolean
    Const StrPROCEDURE As String = "ShowForm()"

    On Error GoTo ErrorHandler

    If Not ResetForm Then Err.Raise HANDLED_ERROR
    
    If ExistCandidate Is Nothing Then
        Set Candidate = New ClsCandidate
        TxtCrewNo.Enabled = True
    Else
        Set Candidate = ExistCandidate
        TxtCrewNo.Enabled = False
        If Not PopulateForm Then Err.Raise HANDLED_ERROR
    End If
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
' BtnClose_Click
' Event for Close Button press
' ---------------------------------------------------------------
Private Sub BtnClose_Click()
    Dim ErrNo As Integer
    Dim Response As Integer

    Const StrPROCEDURE As String = "BtnClose_Click()"

    On Error GoTo ErrorHandler

Restart:

    If Course Is Nothing Then Err.Raise SYSTEM_RESTART

    If FormChanged = True Then
        Response = MsgBox("The form has been changed, would you like to save these changes?", vbYesNo)
        
        If Response = 6 Then BtnUpdate_Click
        FormChanged = False
    End If
    
    Course.Candidates.CleanUp
    Hide

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
' Event process for New Candidate
' ---------------------------------------------------------------
Private Sub BtnNew_Click()
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "BtnNew_Click()"

    On Error GoTo ErrorHandler

Restart:

    If Course Is Nothing Then Err.Raise SYSTEM_RESTART

    If Not ResetForm Then Err.Raise HANDLED_ERROR
    
    Set Candidate = New ClsCandidate
    Course.Candidates.AddItem Candidate
    TxtCrewNo.Enabled = True

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
' PopulateForm
' Populates form with candidate details
' ---------------------------------------------------------------
Private Function PopulateForm() As Boolean
    
    Const StrPROCEDURE As String = "PopulateForm()"

    On Error GoTo ErrorHandler

    With Candidate
        CmoCourseNo = .Parent.CourseNo
        TxtCrewNo = .CrewNo
        CmoDivision = .Division
        TxtName = .Name
        CmoStationNo = .StationNo
        CmoStatus = .Status
    End With
    
    With Candidate.DevelopmentPlans
        TxtDPsClosed = .NoClosed
        TxtDPsOpen = .NoOpen
        TxtDPsOverdue = .NoOverDue
        TxtDPsTotal = .NoClosed + .NoOpen
    End With
    
    With Candidate.Dailylogs
        TxtETOffered = .ETOffered
        TxtETRefused = .ETRefused
        TxtETTaken = .ETTaken
        TxtETTotal = .ETOffered
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
' Validate input data before updating
' ---------------------------------------------------------------
Private Function ValidateData() As Boolean
    Const StrPROCEDURE As String = "ValidateData()"

    On Error GoTo ErrorHandler

    If Me.TxtName = "" Then
        MsgBox "Please enter a candidate name"
        ValidateData = False
        Exit Function
    End If
    
    If Me.TxtCrewNo = "" Then
        MsgBox "Please enter a Crew No"
        ValidateData = False
        Exit Function
    End If

    If Not IsNumeric(Me.TxtCrewNo) Then
        MsgBox "Please enter only numeric characters for crew no"
        ValidateData = False
        Exit Function
    End If
    
    If Len(Me.TxtCrewNo) > 4 Then
        MsgBox "Please check the crew no"
        ValidateData = False
        Exit Function
    End If
    
    If Me.CmoDivision = "" Then
        MsgBox "Please enter the Division"
        ValidateData = False
        Exit Function
    End If

    If Me.CmoStationNo = "" Then
        MsgBox "Please enter a Station"
        ValidateData = False
        Exit Function
    End If

    If Me.CmoCourseNo = "" Then
        MsgBox "Please enter a Course"
        ValidateData = False
        Exit Function
    End If

    If Me.CmoStatus = "" Then
        MsgBox "Please enter a Status"
        ValidateData = False
        Exit Function
    End If

    ValidateData = True

Exit Function

ErrorExit:

    ValidateData = False

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
' Resets candidate form
' ---------------------------------------------------------------
Private Function ResetForm() As Boolean
    Const StrPROCEDURE As String = "ResetForm()"

    On Error GoTo ErrorHandler

    TxtCrewNo = ""
    TxtName = ""
    CmoStatus = ""
    CmoStatus.Value = ""

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
' BtnUpdate_Click
' Event for pressing update button
' ---------------------------------------------------------------
Private Sub BtnUpdate_Click()
    Dim ErrNo As Integer
    Dim Success As Boolean
    
    Const StrPROCEDURE As String = "BtnUpdate_Click()"

    On Error GoTo ErrorHandler

Restart:

    If Course Is Nothing Then Err.Raise HANDLED_ERROR
    
    If ValidateData Then
    
        With Candidate
            .CrewNo = TxtCrewNo
            .Division = CmoDivision
            .Name = TxtName
            .StationNo = CmoStationNo
            .Status = CmoStatus
            .UpdateDB
        End With

        Me.Hide
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

Private Sub TxtCrewNo_Change()
    FormChanged = True
End Sub

Private Sub CmoDivision_Change()
    FormChanged = True
End Sub

Private Sub TxtName_Change()
    FormChanged = True
End Sub

Private Sub CmoStationNo_Change()
    FormChanged = True
End Sub

Private Sub CmoStatus_Change()
    FormChanged = True
End Sub

Private Sub UserForm_Initialize()
    On Error Resume Next
    FormInitialise
End Sub

' ===============================================================
' BtnDelete_Click
' Event for delete button
' ---------------------------------------------------------------
Private Sub BtnDelete_Click()
    Dim ErrNo As Integer
    Dim Response As Integer

    Const StrPROCEDURE As String = "BtnDelete_Click()"
    On Error GoTo ErrorHandler

Restart:
    
    If Course Is Nothing Then Err.Raise HANDLED_ERROR

    Response = MsgBox("Are you sure you want to mark the candidate as deleted?", vbYesNo)
    
    If Response = 6 Then
        Candidate.Parent.Candidates.RemoveItem (Candidate.CrewNo)
        Candidate.DeleteDB
        Set Candidate = Nothing
    End If
    
    If Not ResetForm Then Err.Raise HANDLED_ERROR
    
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

Private Sub UserForm_Terminate()
    On Error Resume Next
    FormTerminate
End Sub

' ===============================================================
' FormInitialise
' Initialises candidate form
' ---------------------------------------------------------------
Private Function FormInitialise() As Boolean
    Dim cell As Range
    
    Const StrPROCEDURE As String = "FormInitialise()"
    
    On Error GoTo ErrorHandler

    For Each cell In ShtLists.Range("Status")
        Me.CmoStatus.AddItem cell
    Next
    
    With CmoDivision
        .Clear
        .AddItem "Not Applicable"
        .Value = "Not Applicable"
        .Enabled = False
    End With

    With CmoStationNo
        .Clear
        .AddItem "Not Applicable"
        .Value = "Not Applicable"
        .Enabled = False
    End With

    With CmoCourseNo
        .Clear
        .AddItem "WT2018"
        .Value = "WT2018"
        .Enabled = False
    End With

    FormInitialise = True

Exit Function

ErrorExit:

    FormInitialise = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If

End Function

Public Sub FormTerminate()
    On Error Resume Next
    Set Candidate = Nothing
End Sub
