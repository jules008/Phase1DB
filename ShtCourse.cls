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
'---------------------------------------------------------------
' Date - 10 Oct 16
'===============================================================
Option Explicit

Private Const StrMODULE As String = "ShtCourse"

Dim Course As ClsCourse

Private Sub BtnConnectDB_Click()

    Const StrPROCEDURE As String = "BtnConnectDB_Click()"
    
    On Error GoTo ErrorHandler
    
    If Not ModDatabase.SelectDB Then Err.Raise HANDLED_ERROR
    
Exit Sub

ErrorExit:
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


Private Sub BtnReset_Click()
    Const StrPROCEDURE As String = "BtnReset_Click()"
    
    On Error GoTo ErrorHandler
    
    If Not Initialise Then Err.Raise HANDLED_ERROR
    If Not WorksheetActivate Then Err.Raise HANDLED_ERROR

Exit Sub

ErrorExit:
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

Private Sub CmoCourseNo_Change()
    
    Const StrPROCEDURE As String = "CmoCourseNo_Change()"
    
    Dim Response As Integer
    
    On Error GoTo ErrorHandler
    
    'first make sure there are courses, then make sure the dropdown contains a value
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

Public Function PopulateSheet() As Boolean
    Const StrPROCEDURE As String = "PopulateSheet()"
    
    Dim TargetRange As Range
    Dim i As Integer
    Dim x As Integer
    Dim Candidate As ClsCandidate
    Dim DailyLog As ClsDailyLog
    
    On Error GoTo ErrorHandler
    
    Set TargetRange = Range("B6")
    
    
    ClearSheet
    
    If Courses Is Nothing Then
        If Not Initialise Then Err.Raise HANDLED_ERROR
    End If
    
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

    PopulateSheet = True

Exit Function

ErrorExit:
    Me.Protect
    Set Candidate = Nothing
    Set DailyLog = Nothing
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

Private Sub CmoCourseNo_GotFocus()
    Const StrPROCEDURE As String = "CmoCourseNo_GotFocus()"
    
    On Error GoTo ErrorHandler
    
    If Courses Is Nothing Then
        If Not Initialise Then Err.Raise HANDLED_ERROR
    End If
    
    If CmoCourseNo.ListCount = 0 Then
        If Not RefreshCourses Then Err.Raise HANDLED_ERROR
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
Private Sub Worksheet_Deactivate()
    On Error Resume Next
    Set Course = Nothing
End Sub

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

Public Sub ClearSheet()
    
    On Error Resume Next
    
    Range("B6:AI32").ClearContents
End Sub

Public Sub ShowCourseFrm()
    
    Const StrPROCEDURE As String = "ShowCourseFrm()"
    
    On Error GoTo ErrorHandler
    
    If Course Is Nothing Then
        If Not Initialise Then Err.Raise HANDLED_ERROR
        If Not SetActiveCourse Then Err.Raise HANDLED_ERROR
    End If
    
    If Not FrmCourse.ShowForm(Course) Then Err.Raise HANDLED_ERROR
    If Not RefreshCourses Then Err.Raise HANDLED_ERROR
    If Not SetActiveCourse Then Err.Raise HANDLED_ERROR
    If Not PopulateSheet Then Err.Raise HANDLED_ERROR

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

Public Sub ShowCandidateFrm()
    
    Const StrPROCEDURE As String = "ShowCandidateFrm()"
    
    Dim Candidate As ClsCandidate
    Dim CrewNo As String
    
    On Error GoTo ErrorHandler
    
    Set Candidate = New ClsCandidate
    
    CrewNo = GetCrewNo
    
    If Course Is Nothing Then
        If Not SetActiveCourse Then Err.Raise HANDLED_ERROR
    End If
    
    If CrewNo <> "" Then
        Set Candidate = Course.Candidates.FindItem(CrewNo)
        If Not FrmCandidate.ShowForm(Candidate) Then Err.Raise HANDLED_ERROR
    Else
        If Not FrmCandidate.ShowForm Then Err.Raise HANDLED_ERROR
    End If
    
    If Not PopulateSheet Then Err.Raise HANDLED_ERROR
    
    Set Candidate = Nothing

Exit Sub
    
ErrorExit:
    Set Candidate = Nothing

Exit Sub

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub


Public Sub ShowDailyLogFrm()
    
    Const StrPROCEDURE As String = "ShowDailyLogFrm()"
    
    Dim CrewNo As String
    Dim DayNo As String
    Dim Candidate As ClsCandidate
    Dim DailyLog As ClsDailyLog
    
    On Error GoTo ErrorHandler
    
    CrewNo = GetCrewNo
    DayNo = GetDayNo
    
    If Course Is Nothing Then
        If Not Initialise Then Err.Raise HANDLED_ERROR
        If Not SetActiveCourse Then Err.Raise HANDLED_ERROR
    End If
    
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
        
        Set Candidate = Nothing
        Set DailyLog = Nothing
    End If

Exit Sub

ErrorExit:
    Set Candidate = Nothing
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

Public Function GetCrewNo() As String
    Dim CrewNo As String
    
    On Error Resume Next
    
    CrewNo = Cells(ActiveCell.Row, 2).Value
    If CrewNo <> "" Then
        GetCrewNo = Format(CrewNo, "0000")
    End If
End Function

Public Function GetDayNo() As String
    Dim DayNo As String
    
    On Error Resume Next
        
    DayNo = Cells(3, ActiveCell.Column).Value
    If DayNo <> "" Then
        GetDayNo = DayNo
    Else
        MsgBox "Module not selected"
    End If
End Function

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
    
    Const StrPROCEDURE As String = "RunSupervisionFile()"

    On Error GoTo ErrorHandler

    Set RngName = ShtCourse.Range("C6")
    
    Do While RngName.Value <> ""
    
        ShtCourse.Activate
        RngName.Activate
                
        CrewNo = GetCrewNo
        
        If Course Is Nothing Then
            If Not Initialise Then Err.Raise HANDLED_ERROR
            If Not SetActiveCourse Then Err.Raise HANDLED_ERROR
        End If
        
        If CrewNo <> "" Then Set Candidate = Course.Candidates.FindItem(CrewNo)
    
        If Not SupervisionFile.BtnSupervisionFile(Candidate) Then Err.Raise HANDLED_ERROR
        Set RngName = RngName.Offset(1, 0)
    Loop
    
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

