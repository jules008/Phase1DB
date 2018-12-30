VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmGrade 
   Caption         =   "Candidate Grade"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6120
   OleObjectBlob   =   "FrmGrade.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmGrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 10 Oct 16
'===============================================================
Option Explicit
Private Const StrMODULE As String = "FrmGrade"
Dim DailyLog As ClsDailyLog

Public Function ShowForm(LocalDailyLog As ClsDailyLog) As Boolean
    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
    
    ResetForm
    If Not LocalDailyLog Is Nothing Then
        Set DailyLog = LocalDailyLog
    End If
    
    If Not PopulateForm Then Err.Raise HANDLED_ERROR
    
    Show
    ShowForm = True

Exit Function

ErrorExit:
    FormTerminate
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
Public Function PopulateForm() As Boolean
    Const StrPROCEDURE As String = "PopulateForm()"

    On Error GoTo ErrorHandler
    
    TxtGrade = DailyLog.OverallGrade
    Select Case TxtGrade
    
        Case Is < 2
            Me.TxtGrade.ForeColor = COLOUR_3
            Me.TxtGradeDesc.Caption = "A very high standard has been achieved or demonstrated"
            Me.TxtGradeDesc.ForeColor = COLOUR_3
            ImgGrade.BackColor = COLOUR_1
            Me.TxtDPDesc.Caption = "A Development Plan is not required"
            Me.BtnNo.Visible = False
            Me.BtnYes.Enabled = True
            Me.BtnYes.Caption = "OK"
            
        Case 2
            Me.TxtGrade.ForeColor = COLOUR_3
            Me.TxtGradeDesc.Caption = "The candidate has achieved the required standard"
            Me.TxtGradeDesc.ForeColor = COLOUR_3
            ImgGrade.BackColor = COLOUR_1
            Me.TxtDPDesc.Caption = "A Development Plan is not required"
            Me.BtnNo.Visible = False
            Me.BtnYes.Enabled = True
            Me.BtnYes.Caption = "OK"
        
        Case 3
            Me.TxtGrade.ForeColor = COLOUR_4
            Me.TxtGradeDesc.Caption = "The candidate has under achieved in one specific area, advice or development required"
            ImgGrade.BackColor = COLOUR_2
            Me.TxtGradeDesc.ForeColor = COLOUR_4
            Me.TxtDPDesc.Caption = "Does the candidate's performance require a Development Plan"
            Me.BtnNo.Visible = True
            Me.BtnNo.Enabled = True
            Me.BtnYes.Enabled = True
            Me.BtnYes.Caption = "Yes"
            
        Case 4
            Me.TxtGrade.ForeColor = COLOUR_6
            Me.TxtGradeDesc.Caption = "The candidate has under achieved in more than one area, further development is required"
            ImgGrade.BackColor = COLOUR_7
            Me.TxtGradeDesc.ForeColor = COLOUR_6
            Me.TxtDPDesc.Caption = "Development Plan(s) are required, do you want to raise one now?"
            Me.BtnNo.Visible = True
            Me.BtnNo.Enabled = True
            Me.BtnYes.Enabled = True
            Me.BtnYes.Caption = "Yes"
            
        Case 5
            Me.TxtGrade.ForeColor = COLOUR_6
            Me.TxtGradeDesc.Caption = "The candidate has under achieved in all areas, further development is required"
            ImgGrade.BackColor = COLOUR_7
            Me.TxtGradeDesc.ForeColor = COLOUR_6
            Me.TxtDPDesc.Caption = "Development Plan(s) are required, do you want to raise one now?"
            Me.BtnNo.Visible = True
            Me.BtnNo.Enabled = True
            Me.BtnYes.Enabled = True
            Me.BtnYes.Caption = "Yes"
            
            
    End Select

    PopulateForm = True

Exit Function

ErrorExit:
    FormTerminate
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
Public Sub ResetForm()
    On Error Resume Next
    
    Me.TxtDPDesc = ""
    Me.TxtGrade = ""
    Me.TxtGradeDesc = ""
    
End Sub

Private Sub BtnNo_Click()
    On Error Resume Next
    
    Me.Hide
End Sub

Private Sub BtnYes_Click()
    Const StrPROCEDURE As String = "BtnYes_Click()"
    
    Dim Candidate As ClsCandidate
    Dim DevelopmentPlan As ClsDevelopmentPlan
    Dim DevArea As ClsDevelopmentArea
    Dim i As Integer
    
    On Error GoTo ErrorHandler

    If DailyLog.OverallGrade > 2 Then
        Set DevelopmentPlan = New ClsDevelopmentPlan
        Set Candidate = DailyLog.Parent
        
        Candidate.DevelopmentPlans.AddItem DevelopmentPlan
        DevelopmentPlan.NewDB
        
        With DevelopmentPlan
            .DPDate = Now
            
            If DailyLog.Score1 > 2 Then
                Set DevArea = New ClsDevelopmentArea
                
                With DevArea
                    .DevArea = "Attitude"
                    .Module = DailyLog.Module
                    .CurrPerfLvl = DailyLog.Comments1
                    .ReviewStatus = "Draft"
                End With
                
                DevelopmentPlan.DevelopmentAreas.AddItem DevArea
                DevArea.NewDB
                DevArea.UpdateDB
                Set DevArea = Nothing
            End If
            
            If DailyLog.Score2 > 2 Then
                Set DevArea = New ClsDevelopmentArea
                
                With DevArea
                    .DevArea = "Practical Ability"
                    .Module = DailyLog.Module
                    .CurrPerfLvl = DailyLog.Comments2
                    .ReviewStatus = "Draft"
                End With
                
                DevelopmentPlan.DevelopmentAreas.AddItem DevArea
                DevArea.NewDB
                DevArea.UpdateDB
                Set DevArea = Nothing
            End If
            
            If DailyLog.Score3 > 2 Then
                Set DevArea = New ClsDevelopmentArea
                
                With DevArea
                    .DevArea = "Knowledge"
                    .Module = DailyLog.Module
                    .CurrPerfLvl = DailyLog.Comments3
                    .ReviewStatus = "Draft"
                End With
                
                DevelopmentPlan.DevelopmentAreas.AddItem DevArea
                DevArea.NewDB
                DevArea.UpdateDB
                Set DevArea = Nothing
            End If
            
            If DailyLog.Score4 > 2 Then
                Set DevArea = New ClsDevelopmentArea
                
                With DevArea
                    .DevArea = "Safety"
                    .Module = DailyLog.Module
                    .CurrPerfLvl = DailyLog.Comments4
                    .ReviewStatus = "Draft"
                End With
                
                DevelopmentPlan.DevelopmentAreas.AddItem DevArea
                DevArea.NewDB
                DevArea.UpdateDB
                Set DevArea = Nothing
            End If
        End With
        DevelopmentPlan.UpdateDB
        
        If Not FrmDevelopmentPlan.ShowForm(DevelopmentPlan) Then Err.Raise HANDLED_ERROR
    
    End If
    Me.Hide
    Set Candidate = Nothing
    Set DevelopmentPlan = Nothing
    Set DevArea = Nothing
Exit Sub

ErrorExit:
    
    Me.Hide
    FormTerminate
    Set Candidate = Nothing
    Set DevelopmentPlan = Nothing
    Set DevArea = Nothing

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

Public Sub FormTerminate()
    On Error Resume Next

    Set DailyLog = Nothing
End Sub
