Attribute VB_Name = "SupervisionFile"
Option Explicit

' ===============================================================
' BtnSupervisionFile
' Runs supervision file
' ---------------------------------------------------------------
Public Function BtnSupervisionFile(Candidate As ClsCandidate) As Boolean
    Const StrPROCEDURE As String = "BtnSupervisionFile()"

    On Error GoTo ErrorHandler

    Dim DlgOpen As FileDialog
    Dim DailyLog As ClsDailyLog
    Dim DevelopmentPlan As ClsDevelopmentPlan
    Dim DevelopmentPlans As ClsDevelopmentPlans
    Dim FSO As FileSystemObject
    Dim Module As ClsModule
    Dim ModuleNo As String
    Dim RstModule As Recordset
    Dim RstDevelopmentPlan  As Recordset
    Dim StrCrewNo As String
    Dim WshtDP As Worksheet
    Dim FilePath As String
    Dim NoDPs As Integer
    Dim Course As ClsCourse
    Dim crewno As String
    Dim Response As Integer
    Dim i As Integer
    

    Set Module = New ClsModule
    Set Candidate = New ClsCandidate
    Set Course = New ClsCourse
    Set DailyLog = New ClsDailyLog
    Set FSO = New FileSystemObject
    Set DevelopmentPlans = New ClsDevelopmentPlans
    Set DevelopmentPlan = New ClsDevelopmentPlan
    
    Response = MsgBox("Are you sure you want to create a Supervision File?", vbYesNo, Title:="Print Supervision File")
    
    If Response = 6 Then
        
        'get destination folder
        Set DlgOpen = Application.FileDialog(msoFileDialogFolderPicker)
        With DlgOpen
            .Filters.Clear
            .AllowMultiSelect = False
            .Title = "Select Destination Folder"
            .Show
        End With
        FilePath = DlgOpen.SelectedItems(1)
        
        'clear Summary sheet
        ShtSummary.ClearSummary
        
        'get module recordset for lookup
        Set RstModule = database.SQLQuery("module")
        
        'get crew no of candidate for supervision File
'        CrewNo = Me.GetCrewNo
        StrCrewNo = "'" & crewno & "'"
        'update candidate details on summary sheet
'        Set Candidate = Candidate.GetCandidateClass(CrewNo)
        ShtSummary.FillOutSummaryCandidate Candidate
        
        'add new folder for candidate
        FilePath = FilePath & "/" & Candidate.crewno & " " & Candidate.Name
        
        If FSO.FolderExists(FilePath) Then
            FSO.DeleteFolder (FilePath)
        End If
        
        FSO.CreateFolder FilePath
         
        'get Daily Log objects
        For i = 1 To 32
            
            'lookup module no for each day
            With RstModule
                .MoveFirst
                .FindFirst "[dayno] = " & i
                ModuleNo = !ModuleNo
            End With
            
'            Set Module = Module.GetModuleClass(ModuleNo)
            Set Course = Course.GetCourseClass(Candidate.CourseNo)
            Set DailyLog = DailyLog.GetDailyLogClass(crewno, ModuleNo)
               
            'fill out front sheet
            ShtCover.PopulateFrontSheet Candidate, Course
            
            'fill out assessment sheet
            ShtAssessment.PopulateAssessmentSummary Candidate

            If Not DailyLog Is Nothing Then
            
                'fill out summarysheet
                ShtSummary.FillOutSummary DailyLog
                   
                'fill out daily log template
                ShtDailyLog.FillOutForm DailyLog, Module, Candidate, Course
                
                'print Daily log
                With ShtDailyLog
                    .Visible = xlSheetVisible
                    If Globals.ENABLE_PRINT = True Then
                        Library.PrintPDF ShtDailyLog, FilePath & "/" & "5 - Daily Log " & Format(i, "00")
                    End If
                   .Visible = xlSheetHidden
                End With
            Else
                Set DailyLog = New ClsDailyLog
            End If
        Next
        
        'process Development Plans
        Set DevelopmentPlans = DevelopmentPlans.AddDevelopmentPlans(Candidate.crewno)
        
        If Not DevelopmentPlans Is Nothing Then
        
            'get number of Development plans
            NoDPs = DevelopmentPlans.Count
    
            For i = 1 To NoDPs
                
                Set WshtDP = ShtDPTemplate.FillOutDP(DevelopmentPlans.GetDP(i), Candidate)
                
                'Print DP Sheet
                With WshtDP
                    If Globals.ENABLE_PRINT = True Then
                        .PageSetup.Orientation = xlLandscape
                        Library.PrintPDF WshtDP, FilePath & "/" & "6 - Development Plan " & i
                    End If
                    Application.DisplayAlerts = False
                    WshtDP.Delete
                    Application.DisplayAlerts = True
                End With
            Next
        End If
        
        'Print Summary Sheet
        With ShtSummary
            .Visible = xlSheetVisible
            If Globals.ENABLE_PRINT = True Then
                If Globals.ENABLE_PRINT = True Then
                    .PageSetup.Orientation = xlLandscape
                    Library.PrintPDF ShtSummary, FilePath & "/" & "2 - Summary"
                End If
            End If
           .Visible = xlSheetHidden
        End With
        
        'Print assessment Sheet
        With ShtAssessment
            .Visible = xlSheetVisible
            If Globals.ENABLE_PRINT = True Then
                If Globals.ENABLE_PRINT = True Then
                    .PageSetup.Orientation = xlLandscape
                    Library.PrintPDF ShtAssessment, FilePath & "/" & "4 - Assessments"
                End If
            End If
           .Visible = xlSheetHidden
        End With
        
        'Print grading Sheet
        With ShtGrading
            .Visible = xlSheetVisible
            If Globals.ENABLE_PRINT = True Then
                Library.PrintPDF ShtGrading, FilePath & "/" & "3 - Grading Guide"
            End If
           .Visible = xlSheetHidden
        End With
        
        'print front sheets
        With ShtCover
            .Visible = xlSheetVisible
            If Globals.ENABLE_PRINT = True Then
                Library.PrintPDF ShtCover, FilePath & "/" & "1 - Front Sheets"
            End If
           .Visible = xlSheetHidden
        End With
        
        'print blank sheet
        With ShtBlank
            .Visible = xlSheetVisible
            If Globals.ENABLE_PRINT = True Then
                Library.PrintPDF ShtBlank, FilePath & "/" & "7 - Blank Sheet"
            End If
           .Visible = xlSheetHidden
        End With
        
        MsgBox "Supervision File is complete"
        'Set database = Nothing
        Set Module = Nothing
        Set Candidate = Nothing
        Set Course = Nothing
        Set DailyLog = Nothing
        Set RstModule = Nothing
        Set DlgOpen = Nothing
        Set FSO = Nothing
        Set DevelopmentPlan = Nothing
    End If

    BtnSupervisionFile = True

Exit Function

ErrorExit:

    Set Module = Nothing
    Set Candidate = Nothing
    Set Course = Nothing
    Set DailyLog = Nothing
    Set RstModule = Nothing
    Set DlgOpen = Nothing
    Set FSO = Nothing
    Set DevelopmentPlan = Nothing
    BtnSupervisionFile = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function


