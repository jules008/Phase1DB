VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmDevelopmentPlan 
   Caption         =   "New Action Plan"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9885
   OleObjectBlob   =   "FrmDevelopmentPlan.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmDevelopmentPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===============================================================
' v0,0 - Initial version
' v0,1 - WT2019 Version
'---------------------------------------------------------------
' Date - 30 Dec 18
'===============================================================
Option Explicit
Private Const StrMODULE As String = "FrmDevelopmentPlan"
Private Candidate As ClsCandidate
Private DevelopmentPlan As ClsDevelopmentPlan
Private FormChanged As Boolean

Public Function ShowForm(Optional LocalDevelopmentPlan As ClsDevelopmentPlan, Optional LocalCandidate) As Boolean
    
    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
    
    ResetForm
    
    If LocalDevelopmentPlan Is Nothing Then
    
        Set DevelopmentPlan = New ClsDevelopmentPlan
        If Not LocalCandidate Is Nothing Then
            Set Candidate = LocalCandidate
        End If
        
        Candidate.DevelopmentPlans.AddItem DevelopmentPlan
    Else
        Set DevelopmentPlan = LocalDevelopmentPlan
        Set Candidate = DevelopmentPlan.Parent
    End If
    
    If Not PopulateForm Then Err.Raise HANDLED_ERROR
    FormChanged = False
    ShowForm = True
    Show

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

    Const StrPROCEDURE As String = "BtnClose_Click()"
   
   Dim Response As Integer
    
    On Error GoTo ErrorHandler
    
    If Not Candidate Is Nothing Then Candidate.DevelopmentPlans.CleanUp
    
    FormTerminate
    
    If Me.Visible = True Then Me.Hide
Exit Sub

ErrorExit:
    FormTerminate
    Terminate
    If Me.Visible = True Then Me.Hide
Exit Sub

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

Private Sub BtnEditDev_Click()

    Const StrPROCEDURE As String = "BtnEditDev_Click()"
    
    Dim i As Integer
    Dim Area As String
    Dim LocDevArea As ClsDevelopmentArea
    
    On Error GoTo ErrorHandler
   
   If LstDevAreas.ListIndex = -1 Then
        MsgBox "Please select a development area"
    Else
        i = LstDevAreas.ListIndex
        Area = LstDevAreas.List(i, 1)
        Set LocDevArea = DevelopmentPlan.DevelopmentAreas.FindItem(Area)
        
        If Not FrmDPDevArea.ShowForm(LocalDevArea:=LocDevArea) Then Err.Raise HANDLED_ERROR
        
    End If
    
    If Not PopulateForm Then Err.Raise HANDLED_ERROR
    Set LocDevArea = Nothing
Exit Sub

ErrorExit:
    FormTerminate
    Terminate
    Set LocDevArea = Nothing
Exit Sub

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub
Private Sub BtnNewDevArea_Click()
    Const StrPROCEDURE As String = "BtnNewDevArea_Click()"
    
    On Error GoTo ErrorHandler

    If Not FrmDPDevArea.ShowForm(LocalDevelopmentPlan:=DevelopmentPlan) Then Err.Raise HANDLED_ERROR
    If Not PopulateForm Then Err.Raise HANDLED_ERROR
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

Private Sub BtnPrintOLD()
    Const StrPROCEDURE As String = "BtnPrint_Click()"
    
    Dim PrintPages As Integer
    Dim i As Integer
    Dim x As Integer
    Dim NoRows As Integer
    Dim WshtDP As Worksheet
    Dim LocDevArea As ClsDevelopmentArea
    
    On Error GoTo ErrorHandler
                
        For i = 1 To DevelopmentPlan.DevelopmentAreas.Count
            
            Set LocDevArea = DevelopmentPlan.DevelopmentAreas.FindItem(i)
            Set WshtDP = ShtDPTemplate.FillOutDP(LocDevArea, Candidate)
            
            With WshtDP
                If ModGlobals.ENABLE_PRINT = True Then
                    .PageSetup.Orientation = xlLandscape
'                    ModLibrary.PrintPDF WshtDP, FilePath & "/" & "6 - Development Plan " & i
                End If
                Application.DisplayAlerts = False
                WshtDP.Delete
                Application.DisplayAlerts = True
            End With
        Next
'        FrmPrintCopies.Show
'        .Visible = xlSheetVisible
'
'        For x = 1 To FrmPrintCopies.CmoNoCopies
''            .PrintOut
'        Next
'
'        'delete AP sheet
'        Application.DisplayAlerts = False
'        .Delete
'        Application.DisplayAlerts = True
'    End With
'    Set LocDevArea = Nothing
'
'    Me.Hide
    
Exit Sub

ErrorExit:
    Set LocDevArea = Nothing
    FormTerminate
    Terminate
    WshtDP.Delete
    ShtDPTemplate.Visible = xlSheetHidden
    Application.DisplayAlerts = True

Exit Sub

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

Private Sub BtnPrint_Click()
    
    DevelopmentPlan.PrintForm
    Hide
End Sub

Private Sub BtnUpdate_Click()
    On Error Resume Next
    
    UpdateDevelopmentPlan
End Sub

Private Sub UpdateDevelopmentPlan()
    Dim Response As Integer

    Const StrPROCEDURE As String = "UpdateDevelopmentPlan()"
    
    On Error GoTo ErrorHandler
    
    If ValidateData Then
        With DevelopmentPlan
            If Not UpdateClass Then Err.Raise HANDLED_ERROR
            
            If .Status = "Failed" Then
                Response = MsgBox("The candidate has failed the assessment." _
                                    & " Do you want to raise a new Development Plan?", vbYesNo)
                
                If Response = 6 Then
                    If Not CreateFollowOnDP Then Err.Raise HANDLED_ERROR
                End If
            End If
            .UpdateDB
        End With
        FormTerminate
        Hide
    End If
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
Private Sub BtnSpellChk_Click()
    On Error Resume Next
    
    Dim Cntrls As Collection
    Dim i As Integer
    Dim Cntrl As Control
    
    Set Cntrls = New Collection
    
    For i = 0 To Me.Controls.Count - 1
        Cntrls.Add Controls(i)
    Next
    ModLibrary.SpellCheck Cntrls
    Set Cntrls = Nothing
End Sub

Private Sub ResetForm()
    On Error Resume Next
    
    FormChanged = False
    TxtDPDate = ""
    TxtCourseNo = ""
    TxtCrewNo = ""
    TxtIssuer = ""
    TxtLocalDpNo = ""
    TxtName = ""
    TxtOutcome = ""
    TxtReviewDate = ""
    TxtStatus = ""
    CmoIssuer.Value = ""
    LstDevAreas.Clear
    LstDevAreas.Value = ""
End Sub

Private Function PopulateForm() As Boolean
    
    Const StrPROCEDURE As String = "PopulateForm()"
    
    Dim DevArea As ClsDevelopmentArea
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    With DevelopmentPlan
        TxtLocalDpNo = .LocalDPNo
        If .DPDate = 0 Then TxtDPDate = Format(Now, "DD/MM/YY") Else TxtDPDate = Format(.DPDate, "dd/mm/yy")
        CmoIssuer = .Issuer
        TxtLocalDpNo = .LocalDPNo
        TxtOutcome = .OutcomeIfNotMet
        If .ReviewDate <> 0 Then TxtReviewDate = .ReviewDate
        TxtStatus = .Status
        CmoIssuer.Value = .Issuer
    End With
    
    With Candidate
        TxtCourseNo = .Parent.CourseNo
        TxtCrewNo = .CrewNo
        TxtName = .Name
    End With
    
    With DevelopmentPlan.DevelopmentAreas
        LstDevAreas.Clear
        
        For i = 1 To .Count
            Set DevArea = .FindItem(i)
            
            With LstDevAreas
                .AddItem
                .List(i - 1, 0) = i
                .List(i - 1, 1) = DevArea.DevArea
                .List(i - 1, 2) = DevArea.Module.ModuleNo
                .List(i - 1, 3) = DevArea.ReviewStatus
            End With
        
        Next
    End With
    Set DevArea = Nothing

    PopulateForm = True

Exit Function

ErrorExit:
    FormTerminate
    Terminate
    Set DevArea = Nothing
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
Private Sub CmoIssuer_Change()
    On Error Resume Next
    FormChanged = True
End Sub

Private Sub TxtOutcome_Change()
    On Error Resume Next
    FormChanged = True
End Sub

Private Sub TxtReviewDate_Change()
    On Error Resume Next
    FormChanged = True
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    
    FormActivate

End Sub

Private Function ValidateData() As Boolean
    On Error Resume Next
    
    If Me.TxtOutcome = "" Then
        MsgBox "Please enter an outcome if the candidate fails the assessment"
        ValidateData = False
        Exit Function
    End If
    
    If Me.TxtDPDate = "" Then
        MsgBox "Please enter a review date"
        ValidateData = False
        Exit Function
    End If

    If Not IsDate(Me.TxtReviewDate) Then
        MsgBox "Please enter a valid date"
        ValidateData = False
        Exit Function
    End If
    
    If Me.CmoIssuer = "" Then
        MsgBox "Please select the issuer's name"
        ValidateData = False
        Exit Function
    End If
    
    ValidateData = True
End Function


Private Sub BtnDeleteDev_Click()

    Const StrPROCEDURE As String = "BtnDeleteDev_Click()"
    
    Dim i As Integer
    Dim Area As String
    Dim LocalDevArea As ClsDevelopmentArea
    
    On Error GoTo ErrorHandler
    
    If LstDevAreas.ListIndex = -1 Then
        MsgBox "Please select a development area"

    Else
        i = LstDevAreas.ListIndex
        Area = LstDevAreas.List(i, 1)
        
        Set LocalDevArea = DevelopmentPlan.DevelopmentAreas.FindItem(Area)
        LocalDevArea.DeleteDB
        DevelopmentPlan.DevelopmentAreas.RemoveItem (Area)
        
        If Not PopulateForm Then Err.Raise HANDLED_ERROR
        
    End If
    Set LocalDevArea = Nothing
Exit Sub

ErrorExit:
    FormTerminate
    Terminate
    Set LocalDevArea = Nothing

Exit Sub

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

Public Sub FormActivate()
    
    Const StrPROCEDURE As String = "FormInitialise()"
    
    Dim cell As Range
    Dim RstUsers As Recordset
    
    On Error GoTo ErrorHandler
    
    Set RstUsers = GetAccessList
    
    Me.CmoIssuer.Clear
    
    With RstUsers
        Do
        Me.CmoIssuer.AddItem !UserName
        .MoveNext
        Loop While Not .EOF
    End With
        
    With Me.LstHeadings
        .Clear
        .AddItem
        .List(0, 0) = "No"
        .List(0, 1) = "Area"
        .List(0, 2) = "Module"
        .List(0, 3) = "Status"
    End With
    Set RstUsers = Nothing
Exit Sub

ErrorExit:
    Set RstUsers = Nothing
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

Public Function CreateFollowOnDP() As Boolean
    Const StrPROCEDURE As String = "CreateFollowOnDP()"
    
    Dim FollowOnDP As ClsDevelopmentPlan
    Dim LocalDevArea As ClsDevelopmentArea
    Dim i As Integer
    
    On Error GoTo ErrorHandler

    Set FollowOnDP = New ClsDevelopmentPlan

    'create follow on DP
    Candidate.DevelopmentPlans.AddItem FollowOnDP
    
    With FollowOnDP
        .DPDate = Format(Now, "dd/mm/yy")
        
        'copy failed areas
        For i = 1 To DevelopmentPlan.DevelopmentAreas.Count
            Set LocalDevArea = DevelopmentPlan.DevelopmentAreas.FindItem(i)
            With LocalDevArea
                If .StandardMet = False Then
                    .ImproveLvl = ""
                    .Support = ""
                    FollowOnDP.DevelopmentAreas.AddItem LocalDevArea
                End If
            End With
        Next
        
        'save
        .NewDB
        .UpdateDB
        
        'add DP no to old
        DevelopmentPlan.FollowOnDP = .DPNo
        
        Set DevelopmentPlan = FollowOnDP
        If Not PopulateForm Then Err.Raise HANDLED_ERROR
        
    End With


    Set LocalDevArea = Nothing
    Set FollowOnDP = Nothing
    
    CreateFollowOnDP = True

Exit Function

ErrorExit:
    FormTerminate
    Terminate
    Set LocalDevArea = Nothing
    Set FollowOnDP = Nothing
    CreateFollowOnDP = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function


Public Function UpdateClass() As Boolean
    Const StrPROCEDURE As String = "UpdateClass()"
    
    On Error GoTo ErrorHandler
    
    With DevelopmentPlan
        .DPDate = TxtDPDate
        .Issuer = CmoIssuer
        .LocalDPNo = TxtLocalDpNo
        .OutcomeIfNotMet = TxtOutcome
        
        If TxtReviewDate <> "" Then
            .ReviewDate = TxtReviewDate
        Else
            .ReviewDate = .DPDate + 7
        End If
        
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

Public Sub FormTerminate()
    On Error Resume Next
    
    Candidate.DevelopmentPlans.CleanUp
    Set Candidate = Nothing
    Set DevelopmentPlan = Nothing
End Sub
