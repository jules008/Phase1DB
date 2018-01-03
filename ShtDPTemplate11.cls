VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtDPTemplate11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Const StrMODULE As String = "ShtDPTemplate"
Option Explicit


Public Function CopySheet() As Worksheet
    Const StrPROCEDURE As String = "CopySheet()"

    Dim TestSht As Worksheet
    
    On Error GoTo ErrorHandler
    
    CleanUp
    
    With Me
        .Copy after:=Me
        .Visible = xlSheetHidden
    End With
    
    Set CopySheet = ShtDPTemplate1
    
    CopySheet = True

Exit Function

ErrorExit:
    ShtDPTemplate1.Delete
    CopySheet = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function


'Public Sub FillOutDP(DevelopmentPlan As ClsDevelopmentPlan, Candidate As ClsCandidate)
Public Function FillOutDP(DevelopmentPlan As Object, Candidate As Object) As Worksheet
    Dim DPNo As Integer
    Dim NoRows As Integer
    Dim ShtDP As Worksheet
    Dim DevArea As ClsDevelopmentArea
    Dim i As Integer
    Dim x As Integer
    
    Set DevArea = New ClsDevelopmentArea
    
    'Application.EnableCancelKey = xlDisabled
    
    'copy template to new DP sheet
    With Me
        .Visible = xlSheetVisible
        .Copy after:=ShtDPTemplate
        .Visible = xlSheetHidden
    End With
    
    Set ShtDP = Application.ActiveSheet
    
    With ShtDP
        .Range("A9:F9").ClearContents
        .Range("dpno") = DevelopmentPlan.DPNo
        .Range("DPDate") = DevelopmentPlan.DPDate
        .Range("Issuer") = DevelopmentPlan.Issuer
        .Range("Outcome") = DevelopmentPlan.Outcome
        .Range("followonap") = DevelopmentPlan.FollowOnDP
        .Range("crewno") = Candidate.crewno
        .Range("name") = Candidate.Name
        .Range("crseno") = Candidate.Course.CourseNo
    
        NoRows = DevelopmentPlan.DevelopmentAreas.Count
        
        For i = 1 To NoRows - 1
            'insert row
            .Range("9:9").Rows.Insert
            .Range("9:9").Cells.Interior.ColorIndex = 0
            .Range("9:9").Cells.Font.Bold = False
            .Range("9:9").Cells.Font.Size = 9
            .Range("B9:E9").Cells.HorizontalAlignment = xlLeft
            .Range("F9").NumberFormat = "dd mmmm yy"
        Next
    
        For i = 1 To NoRows
            
            'get copy of development area
            Set DevArea = DevelopmentPlan.DevelopmentAreas.Item(i)
            
            .Range("A8").Offset(i, 0) = DevArea.Reference
            .Range("A8").Offset(i, 1) = DevArea.DevArea
            .Range("A8").Offset(i, 2) = DevArea.CurrPerfLvl
            .Range("A8").Offset(i, 3) = DevArea.ImproveLvl
            .Range("A8").Offset(i, 4) = DevArea.Support
            .Range("A8").Offset(i, 5) = DevArea.RevDate
            .Range("A8").Offset(i, 0).Rows.AutoFit
        Next
    
        'check if the DP has gone onto two pages
        If .HPageBreaks.Count > 0 Then
            .ResetAllPageBreaks
            .Range("PageBreak").PageBreak = xlPageBreakManual
        End If
        Set FillOutDP = ShtDP
    End With
    FrmDevelopmentPlanList.Hide

End Function


Public Sub PrintDP()
    FrmPrintCopies.Show
    
    For x = 1 To FrmPrintCopies.CmoNoCopies
        If Globals.ENABLE_PRINT = True Then
            Worksheets("ShtDP").PrintOut
        End If
    Next

End Sub

Public Sub DeleteDPSheet()
    'delete DP sheet
    Application.DisplayAlerts = False
    Worksheets("ShtDP").Delete
    Application.DisplayAlerts = True

End Sub


Public Sub CleanUp()
    Dim WSheet As Worksheet
    
    For Each WSheet In Application.ActiveWorkbook
        If WSheet.Name = "ShtDpTemplate" Then
            Application.DisplayAlerts = False
            WSheet.Delete
            Application.DisplayAlerts = True
        End If
    Next
End Sub
