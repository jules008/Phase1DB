VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmCandidate 
   Caption         =   "Candidate"
   ClientHeight    =   10515
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11670
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
'---------------------------------------------------------------
' Date - 24 Aug 16
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmCandidate"

Private Course As ClsCourse
Private Candidate As ClsCandidate
Private FormChanged As Boolean

Public Function ShowForm(Optional ExistCandidate As ClsCandidate) As Boolean
    
    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
    
    ResetForm
    
    If ExistCandidate Is Nothing Then
        Set Candidate = New ClsCandidate
        Set Course = New ClsCourse
        TxtCrewNo.Enabled = True
    Else
        Set Candidate = ExistCandidate
        Set Course = Candidate.Parent
        TxtCrewNo.Enabled = False
        If Not PopulateForm Then Err.Raise HANDLED_ERROR
    End If
    FormChanged = False
    Show

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
    
    Dim Response As Integer
    
    If FormChanged = True Then
        Response = MsgBox("The form has been changed, would you like to save these changes?", vbYesNo)
        
        If Response = 6 Then BtnUpdate_Click
        FormChanged = False
    End If
    Course.Candidates.CleanUp
    FormTerminate
    Me.Hide
End Sub

Private Sub BtnEmailWCS_Click()
    Const StrPROCEDURE As String = "BtnEmailWCS_Click()"

    On Error GoTo ErrorHandler
    
    With MailSystem
        .MailItem.To = TxtWCS
        .MailItem.Subject = TxtCrewNo & " " & TxtName
        .ReturnMail.CrewNo = TxtCrewNo
        .DisplayEmail
        
    End With
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
    
    ResetForm
    Set Candidate = New ClsCandidate
    Course.Candidates.AddItem Candidate
    TxtCrewNo.Enabled = True
      
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

Private Function PopulateForm() As Boolean
    
    Const StrPROCEDURE As String = "PopulateForm()"
    
    Dim LocCourse As ClsCourse
    Dim cell As Range
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    With Candidate
        TxtCourseNo = .Parent.CourseNo
        TxtCrewNo = .CrewNo
        TxtDivision = .Division
        TxtName = .Name
        TxtStationNo = .StationNo
        TxtStatus = .Status
        TxtWCS = .WCS.UserName
        TxtDC = .DC.UserName
        TxtDDC1 = .DDC1.UserName
        TxtDDC2 = .DDC2.UserName
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
    
    With LstHeadings
        .Clear
        .AddItem
        .List(0, 0) = "From"
        .List(0, 1) = "To"
        .List(0, 2) = "Subject"
        .List(0, 3) = "Date"
    End With
    
    PopulateForm = True
    
    Set LocCourse = Nothing
    
Exit Function

ErrorExit:
    
    FormTerminate
    Terminate
    Set LocCourse = Nothing
    
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
    
    If Me.TxtDivision = "" Then
        MsgBox "Please enter the Division"
        ValidateData = False
        Exit Function
    End If

    If Me.TxtStationNo = "" Then
        MsgBox "Please enter a Station"
        ValidateData = False
        Exit Function
    End If

    If Me.TxtCourseNo = "" Then
        MsgBox "Please enter a Course"
        ValidateData = False
        Exit Function
    End If

    If Me.TxtStatus = "" Then
        MsgBox "Please enter a Status"
        ValidateData = False
        Exit Function
    End If
    
    ValidateData = True
End Function

Private Sub ResetForm()

    On Error Resume Next
    
    FormChanged = False
    
    Me.TxtCourseNo = ""
    Me.TxtCrewNo = ""
    Me.TxtDivision = ""
    Me.TxtName = ""
    Me.TxtStationNo = ""
    Me.TxtStatus = ""
    Me.TxtCourseNo.Value = ""
    Me.TxtDivision.Value = ""
    Me.TxtStationNo.Value = ""
    Me.TxtStatus.Value = ""

End Sub

Private Function UpdateClass() As Boolean
    Const StrPROCEDURE As String = "UpdateClass()"
    
    On Error GoTo ErrorHandler
    
    With Candidate
        .CrewNo = TxtCrewNo
        .Division = TxtDivision
        .Name = TxtName
        .StationNo = TxtStationNo
        .Status = TxtStatus
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

Private Sub BtnUpdate_Click()
    Const StrPROCEDURE As String = "BtnOk_Click()"
    
    Dim Success As Boolean
    
    On Error GoTo ErrorHandler
    
    If ValidateData Then
        If Not UpdateClass Then Err.Raise HANDLED_ERROR
        
        With Candidate
            Success = .UpdateDB
            
            If Success = False Then
                .NewDB
                .UpdateDB
            End If
        End With
        Me.Hide
    End If
Exit Sub
    
ErrorExit:
    FormTerminate
    Terminate
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

Private Sub TxtCourseNo_AfterUpdate()
    
    Const StrPROCEDURE As String = "TxtCourseNo_AfterUpdate()"
    
    Dim NewCourse As ClsCourse
    
    On Error GoTo ErrorHandler
    
    FormChanged = True
    If Course.CourseNo = "" Then
        Set Course = Courses.FindItem(TxtCourseNo)
        Candidate.CrewNo = TxtCrewNo
        Course.Candidates.AddItem Candidate
    Else
    
        If TxtCourseNo.Value <> Course.CourseNo Then
            Set NewCourse = Courses.FindItem(TxtCourseNo.Value)
            
            Course.Candidates.RemoveItem Candidate.CrewNo
            
            NewCourse.Candidates.AddItem Candidate
            
            Set NewCourse = Nothing
        End If
    End If
    Set NewCourse = Nothing
Exit Sub

ErrorExit:
    FormTerminate
    Terminate
    Set NewCourse = Nothing

Exit Sub

ErrorHandler:
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

Private Sub TxtDivision_Change()
    FormChanged = True
End Sub

Private Sub TxtName_Change()
    FormChanged = True
End Sub

Private Sub TxtStationNo_Change()
    FormChanged = True
End Sub

Private Sub TxtStatus_Change()
    FormChanged = True
End Sub

Private Sub UserForm_Initialize()
    On Error Resume Next
    FormInitialise
End Sub

Private Sub BtnDelete_Click()
    
    Const StrPROCEDURE As String = "BtnDelete_Click()"
    
    Dim Response As Integer
    
    On Error GoTo ErrorHandler
    
    Response = MsgBox("Are you sure you want to mark the candidate as deleted?", vbYesNo)
    
    If Response = 6 Then
        Candidate.Parent.Candidates.RemoveItem (Candidate.CrewNo)
        Candidate.DeleteDB
        Set Candidate = Nothing
    End If
    ResetForm
    FormChanged = False
Exit Sub

ErrorExit:
    FormTerminate
    Terminate
    Hide
Exit Sub

ErrorHandler:
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

Public Function FormInitialise() As Boolean
    Const StrPROCEDURE As String = "FormInitialise()"
    
    Dim cell As Range
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    If Courses Is Nothing Then
        If Not Initialise Then Err.Raise HANDLED_ERROR
    End If
    
    For Each cell In ShtLists.Range("F1:F38")
        Me.TxtStationNo.AddItem cell
    Next

    For Each cell In ShtLists.Range("A1:A3")
        Me.TxtDivision.AddItem cell
    Next

    For Each cell In ShtLists.Range("Status")
        Me.TxtStatus.AddItem cell
    Next
    
    TxtCourseNo.Clear
    With Courses
        For i = 1 To .Count
            Set Course = Courses.FindItem(i)
            TxtCourseNo.AddItem Course.CourseNo
        Next
    End With

    FormInitialise = True

Exit Function

ErrorExit:
    FormTerminate
    Terminate
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
    Set Course = Nothing
    Set Candidate = Nothing
End Sub
