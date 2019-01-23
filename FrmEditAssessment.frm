VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmEditAssessment 
   Caption         =   "Add Assessment"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5700
   OleObjectBlob   =   "FrmEditAssessment.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmEditAssessment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===============================================================
' v0,0 - Initial version
' v0,1 - Added Validation to BtnUpdate
' v0,2 - Checked for "" before updating scores
' v0,3 - Updated exam types
'---------------------------------------------------------------
' Date - 15 Jan 17
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmEditAssessment"

Private Candidate As ClsCandidate
Private Course As ClsCourse
Private Module As ClsModule
Private Try1 As ClsAssessment
Private Try2 As ClsAssessment
Private Try3 As ClsAssessment
Private Try4 As ClsAssessment
Private Try5 As ClsAssessment

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
    Me.TxtTry1 = ""
    Me.TxtTry2 = ""
    Me.TxtTry3 = ""
    Me.TxtTry4 = ""
    Me.TxtTry5 = ""

End Sub

Private Sub BtnUpdate_Click()
    
    Const StrPROCEDURE As String = "BtnUpdate_Click()"
    
    On Error GoTo ErrorHandler
    
    ' V0,2 changes************************************************
    If ValidateScores Then
        If Not Try1 Is Nothing Then
        
            If TxtTry1.Value = "" Then
                Try1.Score = 0
            Else
                Try1.Score = TxtTry1.Value
            End If
            
            Try1.UpdateDB
        End If
        
        If Not Try2 Is Nothing Then
            
            If TxtTry2.Value = "" Then
                Try2.Score = 0
            Else
                Try2.Score = TxtTry2.Value
            End If
                
            Try2.UpdateDB
        
        End If
        
        If Not Try3 Is Nothing Then
            
            If TxtTry3.Value = "" Then
                Try3.Score = 0
            Else
                Try3.Score = TxtTry3.Value
            End If
        
            Try3.UpdateDB
        End If
        
        If Not Try4 Is Nothing Then
            
            If TxtTry4.Value = "" Then
                Try4.Score = 0
            Else
                Try4.Score = TxtTry4.Value
            End If
        
            Try4.UpdateDB
        End If
        
        If Not Try5 Is Nothing Then
            
            If TxtTry5.Value = "" Then
                Try5.Score = 0
            Else
                Try5.Score = TxtTry5.Value
            End If
        
            Try5.UpdateDB
        End If
                
        If Not ShtAssess.PopulateSheet Then Err.Raise HANDLED_ERROR
    End If
    '**************************************************************
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

Private Sub BtnClose_Click()
    On Error Resume Next
    Me.Hide
End Sub

Private Sub BtnGetAssessment_Click()
    Const StrPROCEDURE As String = "BtnGetAssessment_Click()"

    On Error GoTo ErrorHandler

    If Not PopulateForm Then Err.Raise HANDLED_ERROR
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
Private Sub CmoModule_Change()
    Const StrPROCEDURE As String = "CmoModule_Change()"

    Dim ModuleName() As String
    
    On Error GoTo ErrorHandler

    If CmoModule.Value <> "" Then
        ModuleName = Split(CmoModule.Value, " - ")
        Set Module = Modules.FindItem(ModuleName(0))
        
        With Me.CmoType
            .Clear
            .Value = ""
            
            Select Case Me.CmoModule.ListIndex
                Case 0
                    .AddItem "Written"
                Case 1
                    .AddItem "BCS Written"
                    .AddItem "BCS Practical"
                    .AddItem "Pump Written"
                    .AddItem "Pump Practical"
                Case 2
                    .AddItem "Written"
                    .AddItem "Practical"
                    .AddItem "BA Board"
                    .AddItem "FB Oral"
                Case 3
                    .AddItem "Written"
                    .AddItem "Practical"
                Case 4
                    .AddItem "Written"
                    .AddItem "Practical"
                Case 5
                    .AddItem "Written"
                    .AddItem "Practical"
                Case 6
                    .AddItem "Written"
                    .AddItem "Practical"
                Case 7
                    .AddItem "Written"
                    .AddItem "Oral"
                    .AddItem "Practical"
            End Select
        End With
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
Private Sub CmoName_Change()
    On Error Resume Next
    
    CmoModule = ""
    CmoType = ""
    
    If CmoName.Value <> "" Then
        Set Candidate = Course.Candidates.FindItem(Left(CmoName, 4))
    End If
End Sub

Public Function ValidateData() As Boolean
    On Error Resume Next
    
    If Me.CmoModule = "" Then
        MsgBox "Please enter an Assessment Module"
        ValidateData = False
        Exit Function
    End If
    
    If Me.CmoName = "" Then
        MsgBox "Please enter a name"
        ValidateData = False
        Exit Function
    End If
    
    If Me.CmoType = "" Then
        MsgBox "Please enter an Assessment Type"
        ValidateData = False
        Exit Function
    End If
    
    ValidateData = True
    
End Function

Private Function ValidateScores() As Boolean
    Dim i As Integer
    Dim Cntrl As Control
    Dim HighTry As Integer
    
    On Error Resume Next
    
    HighTry = 0
    
    For i = 5 To 1 Step -1
        
        Set Cntrl = Me.Controls("TxtTry" & i)
        
        If Cntrl <> "" Then
            
            If Not IsNumeric(Cntrl) Then
                MsgBox "Score must be a number between 0 and 100"
                ValidateScores = False
                Exit Function
            End If
            
            If Cntrl < 0 Or Cntrl > 100 Then
                MsgBox "Score must be a number between 0 and 100"
                ValidateScores = False
                Exit Function
            End If
        End If
            
        If HighTry = 0 Then
            If Cntrl <> "" Then
                HighTry = i
            End If
        Else
            If Cntrl = "" Then
                MsgBox "There is an error with Score " & i
                ValidateScores = False
                Exit Function
            End If
        End If
    Next
    
    ValidateScores = True
    
End Function

Public Function PopulateForm() As Boolean
    Const StrPROCEDURE As String = "PopulateForm()"

    Dim Assessment As ClsAssessment
    Dim i As Integer
    
    On Error GoTo ErrorHandler

    ClearScores
    
    For i = 1 To Candidate.Assessments.Count
        Set Assessment = Candidate.Assessments.FindItem(i)
        
        With Assessment
            If .Module.ModuleNo = Module.ModuleNo And _
                        .AssessType = CmoType.Value Then
                
                Select Case .Attempt
                    Case 1
                        Set Try1 = Assessment
                        TxtTry1 = Try1.Score
                    Case 2
                        Set Try2 = Assessment
                        TxtTry2 = Try2.Score
                    Case 3
                        Set Try3 = Assessment
                        TxtTry3 = Try3.Score
                    Case 4
                        Set Try4 = Assessment
                        TxtTry4 = Try4.Score
                    Case 5
                        Set Try5 = Assessment
                        TxtTry5 = Try5.Score
                
                End Select
        
            End If
        End With
    Next
    If Try1 Is Nothing Then TxtTry1.Enabled = False
    If Try2 Is Nothing Then TxtTry2.Enabled = False
    If Try3 Is Nothing Then TxtTry3.Enabled = False
    If Try4 Is Nothing Then TxtTry4.Enabled = False
    If Try5 Is Nothing Then TxtTry5.Enabled = False
    
    PopulateForm = True
    Set Assessment = Nothing
Exit Function

ErrorExit:
    Set Assessment = Nothing
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
Private Sub UserForm_Terminate()
    
    On Error Resume Next
    
    Set Candidate = Nothing
    Set Course = Nothing
    Set Module = Nothing
    Set Try1 = Nothing
    Set Try2 = Nothing
    Set Try3 = Nothing
    Set Try4 = Nothing
    Set Try5 = Nothing
End Sub

Public Sub ClearScores()
    On Error Resume Next
    
    Set Try1 = Nothing
    Set Try2 = Nothing
    Set Try3 = Nothing
    Set Try4 = Nothing
    Set Try5 = Nothing
    
    TxtTry1.Value = ""
    TxtTry2.Value = ""
    TxtTry3.Value = ""
    TxtTry4.Value = ""
    TxtTry5.Value = ""
    
    TxtTry1.Enabled = True
    TxtTry2.Enabled = True
    TxtTry3.Enabled = True
    TxtTry4.Enabled = True
    TxtTry5.Enabled = True

End Sub

Public Function FormActivate() As Boolean
    Const StrPROCEDURE As String = "FormActivate()"

    Dim LocCandidate As ClsCandidate
    Dim LocModule As ClsModule
    
    Dim i As Integer
        
    On Error GoTo ErrorHandler

    CmoName.Clear
    CmoModule.Clear
    
    With Course.Candidates
        For i = 1 To .Count
            Set LocCandidate = .FindItem(i)
            
            Me.CmoName.AddItem LocCandidate.CrewNo & " " & LocCandidate.Name
        
        Next
    End With
    
    With ModGlobals.Modules
        For i = 1 To .Count
            Set Module = .FindItem(i)
            
            With Module
                If .Assessment = True Then
                    Me.CmoModule.AddItem .ModuleNo & " - " & .Module
                End If
            End With
        Next
    End With
    Set LocCandidate = Nothing
    Set LocModule = Nothing
    

    FormActivate = True

Exit Function

ErrorExit:
    Set LocCandidate = Nothing
    Set LocModule = Nothing
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
