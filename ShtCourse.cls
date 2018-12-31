VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtCourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================
' v0,0 - Initial version
' v1,0 - WT2019 Version
'---------------------------------------------------------------
' Date - 30 Dec 18
'===============================================================
Option Explicit

Private Const StrMODULE As String = "ShtCourse"

Dim Course As ClsCourse

' ===============================================================
' CmoCourseNo_Change
' Processing when course no is changed
' ---------------------------------------------------------------
Private Sub CmoCourseNo_Change()
    
    Const StrPROCEDURE As String = "CmoCourseNo_Change()"
    
    Dim Response As Integer
    
    On Error GoTo ErrorHandler
    
    If Courses Is Nothing Then
        If Not Initialise Then Err.Raise HANDLED_ERROR
    End If
    
    If Courses.Count <> 0 Then
        If CmoCourseNo <> "" Then
            If CourseAccessCheck(CmoCourseNo.Value) Or IsAdmin Then
                If Not SetActiveCourse Then Err.Raise HANDLED_ERROR
                If Not PopulateSheet Then Err.Raise HANDLED_ERROR
            Else
                Response = MsgBox("Sorry, access is denied to this course, would you like to request access?", 20)
                If Response = 6 Then
                        With MailSystem
                            .MailItem.To = "Julian Turner"
                            .MailItem.Subject = "Access Request - " & APP_NAME
                            .MailItem.Body = Application.UserName & " has requested access to " & APP_NAME
                            .SendEmail
                        End With

                End If
                
                If Course Is Nothing Then
                    If Not SetActiveCourse Then Err.Raise HANDLED_ERROR
                End If
                
                CmoCourseNo = Course.CourseNo
            End If
        End If
    Else
        ClearSheet
    End If
    
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

' ===============================================================
' WorksheetActivate
' Event processing for worksheet activation
' ---------------------------------------------------------------
Public Function WorksheetActivate()
    Const StrPROCEDURE As String = "WorksheetActivate()"
    
    Dim CourseNo As String
       
    On Error GoTo ErrorHandler
        
    If Courses Is Nothing Then
        If Not Initialise Then Err.Raise HANDLED_ERROR
    End If
    
    If Not RefreshCourses Then Err.Raise HANDLED_ERROR
    
    If Courses.Count <> 0 Then
        If Not SetActiveCourse Then Err.Raise HANDLED_ERROR
        
        If CourseAccessCheck(Course.CourseNo) Or IsAdmin Then
            If Not PopulateSheet Then Err.Raise HANDLED_ERROR
        End If
    Else
        
        Me.Unprotect
        
        ClearSheet
        CmoCourseNo = ""
        Range("CourseNo") = ""
        Set Course = Nothing
        
        Me.Protect
    End If
    
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

' ===============================================================
' PopulateSheet
' Populates Daily Log scores into sheet
' ---------------------------------------------------------------
Public Function PopulateSheet() As Boolean
    Dim TargetRange As Range
    Dim i As Integer
    Dim x As Integer
    Dim Candidate As ClsCandidate
    Dim DailyLog As ClsDailyLog
    
    Const StrPROCEDURE As String = "PopulateSheet()"
    
    On Error GoTo ErrorHandler
    
    Set TargetRange = Range("B6")
    
    Application.ScreenUpdating = False
    
    If Not ClearSheet Then Err.Raise HANDLED_ERROR
    
    If Course Is Nothing Then Err.Raise HANDLED_ERROR, , "No Course Selected"
    
    Unprotect

    If Course Is Nothing Then
        If Not SetActiveCourse Then Err.Raise HANDLED_ERROR
    End If
    
    If Courses.Count <> 0 Then
        With Course
            Me.Unprotect
            [CourseDirector] = "Course Director - " & Course.CourseDirector
            [CourseStat] = "Status - " & Course.Status
            [CourseNo] = .CourseNo
        End With
        
        With Course.Candidates
            For i = 1 To .Count
                Set Candidate = .FindItem(i)
                TargetRange.Offset(i - 1, 0) = Candidate.CrewNo
                TargetRange.Offset(i - 1, 1) = Candidate.Name
                
                With Candidate.Dailylogs
                    For x = 1 To .Count
                        Set DailyLog = .FindItem(x)
                        TargetRange.Offset(i - 1, DailyLog.Module.DayNo + 1) = DailyLog.OverallGrade
                    Next
                End With
                
            Next
        End With
        Me.Protect
        Set Candidate = Nothing
        Set DailyLog = Nothing
    End If
    
    Application.ScreenUpdating = True
    Protect

    PopulateSheet = True

Exit Function

ErrorExit:
    
    Application.ScreenUpdating = True
    Me.Protect
    Set Candidate = Nothing
    Set DailyLog = Nothing
    Set TargetRange = Nothing
   
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

' ===============================================================
' CmoCourseNo_Click
' Event process for clicking Course no
' ---------------------------------------------------------------
Private Sub CmoCourseNo_Click()
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "CmoCourseNo_Click()"

    On Error GoTo ErrorHandler

Restart:

    If Courses Is Nothing Then Err.Raise SYSTEM_RESTART

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
' ClearSheet
' Clears all Daily Log scores from sheet
' ---------------------------------------------------------------
Private Function ClearSheet() As Boolean
    Const StrPROCEDURE As String = "ClearSheet()"

    On Error GoTo ErrorHandler

    [DL_Clear_Area].ClearContents

    ClearSheet = True
    
Exit Function
ErrorExit:

    ClearSheet = False

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
' Worksheet_Deactivate
' Deactivate worksheet event
' ---------------------------------------------------------------
Private Sub Worksheet_Deactivate()
    On Error Resume Next
    Set Course = Nothing
End Sub

' ===============================================================
' RefreshCourses
' Refreshes course list
' ---------------------------------------------------------------
Public Function RefreshCourses() As Boolean
    Const StrPROCEDURE As String = "RefreshCourses()"
    
    Dim i As Integer
    Dim LocCourse As ClsCourse
    
    On Error GoTo ErrorHandler
    
    With CmoCourseNo
        
        .Clear
        
        For i = 1 To Courses.Count
        
            Set LocCourse = Courses.FindItem(i)
            .AddItem LocCourse.CourseNo
            
            Debug.Print LocCourse.CourseNo & " - " & [CourseNo]
            
        Next
            
    End With

    RefreshCourses = True

Exit Function

ErrorExit:
    Set LocCourse = Nothing
    RefreshCourses = False

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
' ShowCourseFrm
' Shows the dialog box for inputting coruse details
' ---------------------------------------------------------------
Private Sub ShowCourseFrm()
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "ShowCourseFrm()"

    On Error GoTo ErrorHandler

Restart:

    If Courses Is Nothing Then Err.Raise SYSTEM_RESTART
       
    If Not FrmCourse.ShowForm(Course) Then Err.Raise HANDLED_ERROR

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
' ShowCandidateFrm
' Shows Candidate form when button is pressed
' ---------------------------------------------------------------
Private Sub ShowCandidateFrm()
    Dim Candidate As ClsCandidate
    Dim CrewNo As String
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "ShowCandidateFrm()"

    On Error GoTo ErrorHandler

Restart:

    If Courses Is Nothing Then Err.Raise SYSTEM_RESTART

    Set Candidate = New ClsCandidate
    
    CrewNo = GetCrewNo
    
    If CrewNo <> "" Then
        Set Candidate = Course.Candidates.FindItem(CrewNo)
        If Not FrmCandidate.ShowForm(Candidate) Then Err.Raise HANDLED_ERROR
    Else
        If Not FrmCandidate.ShowForm Then Err.Raise HANDLED_ERROR
    End If
    
    If Not PopulateSheet Then Err.Raise HANDLED_ERROR
    
    Set Candidate = Nothing

GracefulExit:

Exit Sub

ErrorExit:
    Set Candidate = Nothing

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
' ShowDailyLogFrm
' Event for Daily Log Button to show form
' ---------------------------------------------------------------
Private Sub ShowDailyLogFrm()
    Dim CrewNo As String
    Dim DayNo As String
    Dim Candidate As ClsCandidate
    Dim DailyLog As ClsDailyLog
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "ShowDailyLogFrm()"
    On Error GoTo ErrorHandler

Restart:
    
    If Courses Is Nothing Then Err.Raise SYSTEM_RESTART

    CrewNo = GetCrewNo
    DayNo = GetDayNo
       
    If CrewNo <> "" And DayNo <> "" Then
        Set Candidate = Course.Candidates.FindItem(CrewNo)
        Set DailyLog = Candidate.Dailylogs.FindItem(DayNo)
        
        If DailyLog Is Nothing Then
            Set DailyLog = New ClsDailyLog
            DailyLog.Module.DayNo = DayNo
            DailyLog.Module.LoadDB
            Candidate.Dailylogs.AddItem DailyLog
        End If
        
        If Not FrmDailyLog.ShowForm(LocalDailyLog:=DailyLog) Then Err.Raise HANDLED_ERROR
        
        If Not PopulateSheet Then Err.Raise HANDLED_ERROR
    End If
    
GracefulExit:
        
        Set Candidate = Nothing
        Set DailyLog = Nothing

Exit Sub

ErrorExit:

        Set Candidate = Nothing
        Set DailyLog = Nothing

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

Public Sub ShowDevelopmentPlanFrm()

    Const StrPROCEDURE As String = "ShowDevelopmentPlanFrm()"
    
    Dim CrewNo As String
    Dim Candidate As ClsCandidate
    Dim DevelopmentPlan As ClsDevelopmentPlan
    
    On Error GoTo ErrorHandler
    
    CrewNo = CStr(Format(GetCrewNo, "0000"))
    
    If Course Is Nothing Then
        If Not Initialise Then Err.Raise HANDLED_ERROR
        If Not SetActiveCourse Then Err.Raise HANDLED_ERROR
    End If
    
    If CrewNo <> "" Then
        Set Candidate = Course.Candidates.FindItem(CrewNo)
    
        If Not FrmDevelopmentPlanList.ShowForm(Candidate) Then Err.Raise HANDLED_ERROR
    End If

    Set Candidate = Nothing
    Set DevelopmentPlan = Nothing
Exit Sub
    
ErrorExit:
    Set Candidate = Nothing
    Set DevelopmentPlan = Nothing

Exit Sub

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' GetCrewNo
' returns the crew no of the selected cell
' ---------------------------------------------------------------
Public Function GetCrewNo() As String
    Dim CrewNo As String
    
    On Error Resume Next
    
    CrewNo = Cells(ActiveCell.Row, 2).Value
    If CrewNo <> "" Then
        GetCrewNo = Format(CrewNo, "0000")
    End If
End Function

' ===============================================================
' GetDayNo
' Returns the course day no for the selected cell
' ---------------------------------------------------------------
Public Function GetDayNo() As String
    Dim DayNo As String
    
    On Error Resume Next
        
    DayNo = Cells(3, ActiveCell.Column).Value
    If DayNo <> "" And IsNumeric(DayNo) Then
        GetDayNo = DayNo
    Else
        MsgBox "Module not selected"
    End If
End Function

' ===============================================================
' SetActiveCourse
' Sets the active course
' ---------------------------------------------------------------
Public Function SetActiveCourse() As Boolean
    
    Const StrPROCEDURE As String = "SetActiveCourse()"
    
    On Error GoTo ErrorHandler

    If Courses Is Nothing Then
        If Not Initialise Then Err.Raise HANDLED_ERROR
    End If

    If Courses.Count <> 0 Then
        If CmoCourseNo <> "" Then
            Set Course = Courses.FindItem(CmoCourseNo.Value)
        Else
            Set Course = Courses.FindItem([CourseNo])
        End If
        
        If Course Is Nothing Then
            Set Course = Courses.FindItem(1)
        End If
        
        CmoCourseNo = Course.CourseNo
        Me.Unprotect
        [CourseNo] = CmoCourseNo
        Me.Protect
    Else
        Set Course = Nothing
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

Public Sub SendSupportMail()
    
    Const StrPROCEDURE As String = "SendSupportMail()"

    Dim sPath As String
    
    On Error GoTo ErrorHandler

    If MailSystem Is Nothing Then
        
        If Not Initialise Then Err.Raise HANDLED_ERROR
        
    End If
    
    sPath = ThisWorkbook.Path
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    
    With MailSystem
        .MailItem.To = "Julian Turner"
        .MailItem.Subject = "Support Mail - " & APP_NAME
        .MailItem.Attachments.Add sPath & FILE_ERROR_LOG
        .MailItem.Body = "Please include any feedback, suggestions or problems " _
                           & "that you would like to share with the developer"
        .DisplayEmail
    End With
    
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
Public Sub ShowAdminForm()
    Const StrPROCEDURE As String = "ShowAdminForm()"

    On Error GoTo ErrorHandler

    If IsAdmin Then
        If Not FrmAdminUserList.ShowForm Then Err.Raise HANDLED_ERROR
    Else
        MsgBox "Access Denied - System Admin use only", vbCritical, Title:=APP_NAME
        
    End If
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

' ===============================================================
' RunSupervisionFile
' Generates Supervision file for selected candidate
' ---------------------------------------------------------------
Public Sub RunSupervisionFile()
    Dim Candidate As ClsCandidate
    Dim CrewNo As String
    Dim RngName As Range
    Dim Response As Integer
    
    Const StrPROCEDURE As String = "RunSupervisionFile()"

    On Error GoTo ErrorHandler

    Response = MsgBox("Do you want to run all Supervision Files?  If No is pressed, the selected File will be processed", vbYesNo + vbQuestion, APP_NAME)
   
    If Course Is Nothing Then
        If Not Initialise Then Err.Raise HANDLED_ERROR
        If Not SetActiveCourse Then Err.Raise HANDLED_ERROR
    End If
    
    If Response = 6 Then
    
        Set RngName = ShtCourse.Range("C6")
        
        Do While RngName.Value <> ""
        
            ShtCourse.Activate
            RngName.Activate
                    
            CrewNo = GetCrewNo
            
            If CrewNo <> "" Then Set Candidate = Course.Candidates.FindItem(CrewNo)
        
            If Not SupervisionFile.BtnSupervisionFile(Candidate) Then Err.Raise HANDLED_ERROR
            Set RngName = RngName.Offset(1, 0)
        Loop
    Else
        ShtCourse.Activate
               
        CrewNo = GetCrewNo
               
        If CrewNo <> "" Then Set Candidate = Course.Candidates.FindItem(CrewNo)
    
        If Not SupervisionFile.BtnSupervisionFile(Candidate) Then Err.Raise HANDLED_ERROR
    End If
    
    Set Candidate = Nothing
    Set RngName = Nothing
Exit Sub

ErrorExit:
    
    Set Candidate = Nothing
    Set RngName = Nothing
'    ***CleanUpCode***

Exit Sub

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' GetModules
' Displays module headings on the page
' ---------------------------------------------------------------
Public Function GetModules() As Boolean
    Dim i As Integer
    Dim Module As ClsModule
    Dim RngModules As Range
    Dim AryModules() As String
    
    Const StrPROCEDURE As String = "GetModules()"

    On Error GoTo ErrorHandler

    ReDim AryModules(1 To Modules.Count)
    Set Module = New ClsModule
    
    For Each Module In Modules
        With Module
            AryModules(.ModuleNo) = .Module
        End With
    Next
    
    Set RngModules = Range("C6")
    Set RngModules = RngModules.Resize(1, Modules.Count)
    Unprotect
    RngModules = AryModules
    Protect

    GetModules = True
    Set Module = Nothing
    Set RngModules = Nothing
Exit Function

ErrorExit:

    '***CleanUpCode***
    Set Module = Nothing
    Set RngModules = Nothing
    GetModules = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function





