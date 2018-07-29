VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmAssessment 
   Caption         =   "Add Assessment"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5910
   OleObjectBlob   =   "FrmAssessment.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmAssessment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 04 Nov 16
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmAssessment"

Private Course As ClsCourse

Public Function ShowForm(ExistCourse As ClsCourse) As Boolean
    
    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
    
    Set Course = ExistCourse
    
    ResetForm
        
    If Not FormActivate Then Err.Raise HANDLED_ERROR
    
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

Private Sub ResetForm()

    On Error Resume Next
    
    Me.CmoModule = ""
    Me.CmoName = ""
    Me.TxtScore1 = ""
    Me.TxtScore2 = ""
    Me.TxtScore3 = ""
    Me.TxtScore4 = ""

End Sub

Private Sub BtnAdd_Click()

    Const StrPROCEDURE As String = "BtnAdd_Click()"

    Dim LocAssessment As ClsAssessment
    Dim LocCandidate As ClsCandidate
    Dim UpdateFail As Boolean
    Dim i As Integer
    
    On Error GoTo ErrorHandler

    UpdateFail = False
    
    If ValidateData Then
        
        Set LocCandidate = New ClsCandidate
        
        Set LocCandidate = Course.Candidates.FindItem(Left(CmoName, 4))
        
        For i = 1 To 4
            If Me.Controls("TxtScore" & i).Visible = True Then
            
                Set LocAssessment = New ClsAssessment
                
                With LocAssessment
                
                    .NewDB
                    LocCandidate.Assessments.AddItem LocAssessment
                    .AssessType = Controls("Lbl" & i).Caption
                    .Module.DayNo = Mid(CmoModule.Value, 5, 2)
                    .Module.LoadDB
                    .Attempt = LocCandidate.Assessments.NextAttemptNo(.Module.DayNo, .AssessType)
                    
                    If .Attempt = 99 Then
                    
                        LocCandidate.Assessments.RemoveItem CStr(.AssessmentNo)
                        UpdateFail = True
                        
                    Else
                        If Me.Controls("TxtScore" & i) <> "" Then
                        
                            .Score = Me.Controls("TxtScore" & i).Value
                            
                        End If
                        
                        If Me.Controls("TxtScore" & i) = 0 Then
                        
                            .Score = 1
                            
                        End If
                        
                        .UpdateDB
                    
                    End If
                End With
                
                Set LocAssessment = Nothing
            
            End If
        Next
        If Not ShtAssess.PopulateSheet Then Err.Raise HANDLED_ERROR
        
        If UpdateFail = True Then
            MsgBox "Only 5 attempts can be entered"
        Else
            MsgBox "Assessment added successfully"
        End If
        Set LocCandidate = Nothing
        Set LocAssessment = Nothing
    End If
Exit Sub

ErrorExit:
    Set LocCandidate = Nothing
    Set LocAssessment = Nothing

Exit Sub

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

Private Sub BtnCancel_Click()

    On Error Resume Next
    
    CleanUp
    Me.Hide
End Sub

Private Sub BtnClose_Click()

    On Error Resume Next
    
    CleanUp
    Me.Hide
End Sub

Private Sub CmoModule_Change()
    Const StrPROCEDURE As String = "CmoModule_Change()"

    Dim Assessment As Integer
    
    On Error GoTo ErrorHandler

    Assessment = Me.CmoModule.ListIndex
    
    Select Case Assessment
        Case -1
            Lbl1.Visible = False
            Lbl1.Caption = ""
            Lbl2.Visible = False
            Lbl2.Caption = ""
            Lbl3.Visible = False
            Lbl3.Caption = ""
            Lbl4.Visible = False
            Lbl4.Caption = ""
            TxtScore1.Visible = False
            TxtScore2.Visible = False
            TxtScore3.Visible = False
            TxtScore4.Visible = False
        Case 0
            Lbl1.Visible = True
            Lbl1.Caption = "Written"
            Lbl2.Visible = False
            Lbl2.Caption = ""
            Lbl3.Visible = False
            Lbl3.Caption = ""
            Lbl4.Visible = False
            Lbl4.Caption = ""
            TxtScore1.Visible = True
            TxtScore2.Visible = False
            TxtScore3.Visible = False
            TxtScore4.Visible = False
        Case 1
            Lbl1.Visible = True
            Lbl1.Caption = "Practical"
            Lbl2.Visible = False
            Lbl2.Caption = ""
            Lbl3.Visible = False
            Lbl3.Caption = ""
            Lbl4.Visible = False
            Lbl4.Caption = ""
            TxtScore1.Visible = True
            TxtScore2.Visible = False
            TxtScore3.Visible = False
            TxtScore4.Visible = False
        Case 2
            Lbl1.Visible = True
            Lbl1.Caption = "Written"
            Lbl2.Visible = True
            Lbl2.Caption = "Practical"
            Lbl3.Visible = False
            Lbl3.Caption = ""
            Lbl4.Visible = False
            Lbl4.Caption = ""
            TxtScore1.Visible = True
            TxtScore2.Visible = True
            TxtScore3.Visible = False
            TxtScore4.Visible = False
        Case 3
            Lbl1.Visible = True
            Lbl1.Caption = "Written"
            Lbl2.Visible = True
            Lbl2.Caption = "Practical"
            Lbl3.Visible = False
            Lbl3.Caption = ""
            Lbl4.Visible = False
            Lbl4.Caption = ""
            TxtScore1.Visible = True
            TxtScore2.Visible = True
            TxtScore3.Visible = False
            TxtScore4.Visible = False
        Case 4
            Lbl1.Visible = True
            Lbl1.Caption = "Written"
            Lbl2.Visible = True
            Lbl2.Caption = "Practical"
            Lbl3.Visible = True
            Lbl3.Caption = "BA Board"
            Lbl4.Visible = True
            Lbl4.Caption = "FB Oral"
            TxtScore1.Visible = True
            TxtScore2.Visible = True
            TxtScore3.Visible = True
            TxtScore4.Visible = True
         Case 5
            Lbl1.Visible = True
            Lbl1.Caption = "Written"
            Lbl2.Visible = True
            Lbl2.Caption = "Practical"
            Lbl3.Visible = False
            Lbl3.Caption = ""
            Lbl4.Visible = False
            Lbl4.Caption = ""
            TxtScore1.Visible = True
            TxtScore2.Visible = True
            TxtScore3.Visible = False
            TxtScore4.Visible = False
         Case 6
            Lbl1.Visible = True
            Lbl1.Caption = "Written"
            Lbl2.Visible = True
            Lbl2.Caption = "Oral"
            Lbl3.Visible = True
            Lbl3.Caption = "Knots"
            Lbl4.Visible = True
            Lbl4.Caption = "Practical"
            TxtScore1.Visible = True
            TxtScore2.Visible = True
            TxtScore3.Visible = True
            TxtScore4.Visible = True
        Case 7
            Lbl1.Visible = True
            Lbl1.Caption = "Written"
            Lbl2.Visible = False
            Lbl2.Caption = ""
            Lbl3.Visible = False
            Lbl3.Caption = ""
            Lbl4.Visible = False
            Lbl4.Caption = ""
            TxtScore1.Visible = True
            TxtScore2.Visible = False
            TxtScore3.Visible = False
            TxtScore4.Visible = False
    End Select
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
Private Sub CmoName_Change()
    
    On Error Resume Next
    
    CmoModule.Value = ""
    TxtScore1 = ""
    TxtScore2 = ""
    TxtScore3 = ""
    TxtScore4 = ""
End Sub

Public Function ValidateData() As Boolean

    On Error Resume Next

    If Me.CmoModule = "" Then
        MsgBox "Please enter an assessment module"
        ValidateData = False
        Exit Function
    End If
    
    If Me.CmoName = "" Then
        MsgBox "Please enter a name"
        ValidateData = False
        Exit Function
    End If
    
    If Me.TxtScore1 <> "" Then
    
        If Not IsNumeric(Me.TxtScore1) Then
            MsgBox "Score must be a number between 0 and 100"
            ValidateData = False
            Exit Function
        End If
        
        If Me.TxtScore1 < 0 Or Me.TxtScore1 > 100 Then
            MsgBox "Score must be a number between 0 and 100"
            ValidateData = False
            Exit Function
        End If
    End If
        
    If Me.TxtScore2 <> "" Then
    
        If Not IsNumeric(Me.TxtScore2) Then
            MsgBox "Score must be a number between 0 and 100"
            ValidateData = False
            Exit Function
        End If
        
        If Me.TxtScore2 < 0 Or Me.TxtScore2 > 100 Then
            MsgBox "Score must be a number between 0 and 100"
            ValidateData = False
            Exit Function
        End If
    End If
        
    If Me.TxtScore3 <> "" Then
    
        If Not IsNumeric(Me.TxtScore3) Then
            MsgBox "Score must be a number between 0 and 100"
            ValidateData = False
            Exit Function
        End If
        
        If Me.TxtScore3 < 0 Or Me.TxtScore3 > 100 Then
            MsgBox "Score must be a number between 0 and 100"
            ValidateData = False
            Exit Function
        End If
    End If
        
    If Me.TxtScore4 <> "" Then
    
        If Not IsNumeric(Me.TxtScore4) Then
            MsgBox "Score must be a number between 0 and 100"
            ValidateData = False
            Exit Function
        End If
        
        If Me.TxtScore4 < 0 Or Me.TxtScore4 > 100 Then
            MsgBox "Score must be a number between 0 and 100"
            ValidateData = False
            Exit Function
        End If
    End If
    
    ValidateData = True
    
End Function

Public Sub CleanUp()

    On Error Resume Next
    
    Dim LocCandidate As ClsCandidate
    
    Set LocCandidate = New ClsCandidate
        
    LocCandidate.Assessments.CleanUp
    
    Set LocCandidate = Nothing
End Sub

Public Function FormActivate() As Boolean
    Const StrPROCEDURE As String = "FormActivate()"

    Dim Candidate As ClsCandidate
    Dim Module As ClsModule
    Dim i As Integer
    
    On Error GoTo ErrorHandler

    CmoName.Clear
    CmoModule.Clear
    
    With Course.Candidates
        For i = 1 To .Count
            Set Candidate = .FindItem(i)
            
            Me.CmoName.AddItem Candidate.CrewNo & " " & Candidate.Name
        
        Next
    End With
    
    With ModGlobals.Modules
        For i = 1 To .Count
            Set Module = .FindItem(i)
            
            With Module
                If .Assessment = True Then
                    Me.CmoModule.AddItem "Day " & .DayNo & " - " & .Module
                End If
            End With
        Next
    End With
    Set Candidate = Nothing
    Set Module = Nothing
    

    FormActivate = True

Exit Function

ErrorExit:
    Set Candidate = Nothing
    Set Module = Nothing
    FormActivate = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

Private Sub UserForm_Terminate()
    Set Course = Nothing
End Sub
