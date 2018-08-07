Attribute VB_Name = "SupervisionFile"
Private Const StrMODULE As String = "SupervisionFile"

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
    Dim CrewNo As String
    Dim Response As Integer
    Dim i As Integer
    

    Set Module = New ClsModule
    Set Course = New ClsCourse
    Set DailyLog = New ClsDailyLog
    Set FSO = New FileSystemObject
    Set DevelopmentPlans = New ClsDevelopmentPlans
    Set DevelopmentPlan = New ClsDevelopmentPlan
    
    If Not Candidate Is Nothing Then
        FilePath = "T:\Training\COURSES\FF Phase 1\Phase 1 DB\Supervision Files\" & Replace(Candidate.Parent.CourseNo, "/", "-") & "\"
        
        If Not FSO.FolderExists(FilePath) Then
            FSO.CreateFolder FilePath
        End If
        
        'clear Summary sheet
        ShtSummary.ClearSummary
        
        'get module recordset for lookup
        Set RstModule = ModDatabase.SQLQuery("module")
        
        'get crew no of candidate for supervision File
'        CrewNo = Me.GetCrewNo
'        StrCrewNo = "'" & crewno & "'"
        'update candidate details on summary sheet
'        Set Candidate = Candidate.GetCandidateClass(CrewNo)
        ShtSummary.FillOutSummaryCandidate Candidate
        
        'add new folder for candidate
        FilePath = FilePath & Candidate.CrewNo & " " & Candidate.Name & "\"
        
        Debug.Print FilePath
        
        If Not FSO.FolderExists(FilePath) Then
            FSO.CreateFolder FilePath
        End If
          
        'fill out front sheet
        ShtCover.PopulateFrontSheet Candidate
        
        'fill out assessment sheet
        ShtAssessment.PopulateAssessmentSummary Candidate

        'get Daily Log objects
        For i = 1 To 32
            
            Set DailyLog = Candidate.Dailylogs.FindItem(i)

            If Not DailyLog Is Nothing Then
            
                'fill out summarysheet
                ShtSummary.FillOutSummary DailyLog
                   
                'fill out daily log template
                ShtDailyLog.FillOutForm DailyLog
                
                'print Daily log
                With ShtDailyLog
                    .Visible = xlSheetVisible
                    If ModGlobals.ENABLE_PRINT = True Then
                        ModLibrary.PrintPDF ShtDailyLog, FilePath & "/" & "5 - Daily Log " & Format(DailyLog.Module.DayNo, "00")
                    End If
                   .Visible = xlSheetHidden
                End With
            Else
                Set DailyLog = New ClsDailyLog
            End If
        Next
        
        'process Development Plans
'        Set DevelopmentPlans = DevelopmentPlans.AddDevelopmentPlans(Candidate.crewno)
        
        If Not DevelopmentPlans Is Nothing Then
        
            'get number of Development plans
            NoDPs = DevelopmentPlans.Count
    
            For i = 1 To NoDPs
                
'                Set WshtDP = ShtDPTemplate.FillOutDP(DevelopmentPlans.GetDP(i), Candidate)
                
                'Print DP Sheet
                With WshtDP
                    If ModGlobals.ENABLE_PRINT = True Then
                        .PageSetup.Orientation = xlLandscape
                        ModLibrary.PrintPDF WshtDP, FilePath & "/" & "6 - Development Plan " & i
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
            If ModGlobals.ENABLE_PRINT = True Then
                If ModGlobals.ENABLE_PRINT = True Then
                    .PageSetup.Orientation = xlLandscape
                    ModLibrary.PrintPDF ShtSummary, FilePath & "/" & "2 - Summary"
                End If
            End If
           .Visible = xlSheetHidden
        End With
        
        'Print assessment Sheet
        With ShtAssessment
            .Visible = xlSheetVisible
            If ModGlobals.ENABLE_PRINT = True Then
                If ModGlobals.ENABLE_PRINT = True Then
                    .PageSetup.Orientation = xlLandscape
                    ModLibrary.PrintPDF ShtAssessment, FilePath & "/" & "4 - Assessments"
                End If
            End If
           .Visible = xlSheetHidden
        End With
        
        'Print grading Sheet
        With ShtGrading
            .Visible = xlSheetVisible
            If ModGlobals.ENABLE_PRINT = True Then
                ModLibrary.PrintPDF ShtGrading, FilePath & "/" & "3 - Grading Guide"
            End If
           .Visible = xlSheetHidden
        End With
        
        'print front sheets
        With ShtCover
            .Visible = xlSheetVisible
            If ModGlobals.ENABLE_PRINT = True Then
                ModLibrary.PrintPDF ShtCover, FilePath & "/" & "1 - Front Sheets"
            End If
           .Visible = xlSheetHidden
        End With
        
        'print blank sheet
        With ShtBlank
            .Visible = xlSheetVisible
            If ModGlobals.ENABLE_PRINT = True Then
                ModLibrary.PrintPDF ShtBlank, FilePath & "/" & "7 - Blank Sheet"
            End If
           .Visible = xlSheetHidden
        End With
        
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


