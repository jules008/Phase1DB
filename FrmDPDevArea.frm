VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmDPDevArea 
   Caption         =   "Add Development Area"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11340
   OleObjectBlob   =   "FrmDPDevArea.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmDPDevArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'===============================================================
' v0,0 - Initial version
' v0,1 - Date Picker removal
' v0,2 - WT2019 Version
'---------------------------------------------------------------
' Date - 30 Dec 18
'===============================================================
Option Explicit
Private Const StrMODULE As String = "FrmDPDevArea"
Private Candidate As ClsCandidate
Private DevelopmentPlan As ClsDevelopmentPlan
Private DevArea As ClsDevelopmentArea
Private FormChanged As Boolean

Public Function ShowForm(Optional LocalDevArea As ClsDevelopmentArea, Optional LocalDevelopmentPlan As ClsDevelopmentPlan) As Boolean
    
    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
    
    
    If LocalDevArea Is Nothing Then
    
        Set DevArea = New ClsDevelopmentArea
        
        If Not LocalDevelopmentPlan Is Nothing Then
            Set Candidate = LocalDevelopmentPlan.Parent
            Set DevelopmentPlan = LocalDevelopmentPlan
        End If
        
        CmoArea.Enabled = True
    Else
        Set DevArea = LocalDevArea
        Set DevelopmentPlan = DevArea.Parent
        Set Candidate = DevelopmentPlan.Parent
        CmoArea.Enabled = False
    End If
    ResetForm
    If Not PopulateForm Then Err.Raise HANDLED_ERROR
    
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


Private Sub BtnUpdate_Click()
    Const StrPROCEDURE As String = "BtnUpdate_Click()"
    
    Dim AStrModule() As String
    
    On Error GoTo ErrorHandler
    
    AStrModule = Split(CmoModule.Value, " - ")
    
    If ValidateData Then
        With DevArea
            .Assessor = CmoAssesor
            .CurrPerfLvl = TxtCurrLvl
            .DevArea = CmoArea
            .ImproveLvl = TxtImprovReqd
            .Module = ModGlobals.Modules.FindItem(AStrModule(0))
            .Reference = TxtRef
            .RevComments = TxtComments
            If TxtReviewDate <> "" Then .RevDate = TxtReviewDate
            .StandardMet = ChkAchieved
            .Support = TxtSupport
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

Private Sub ChkAchieved_Click()
    On Error Resume Next
    
    FormChanged = True
End Sub

Private Sub CmdClose_Click()
    On Error Resume Next
    
    FormTerminate
    If Me.Visible = True Then Me.Hide
End Sub

Private Sub ResetForm()
    On Error Resume Next
    
    FormChanged = False
    Me.CmoAssesor.Value = ""
    Me.CmoArea.Value = ""
    Me.CmoModule.Value = ""
    Me.TxtRef = ""
    Me.TxtCurrLvl = ""
    Me.TxtImprovReqd = ""
    Me.TxtSupport = ""
    Me.TxtComments = ""
    Me.TxtReviewDate = ""
    Me.ChkAchieved = False
End Sub

Private Function PopulateForm() As Boolean
    
    Const StrPROCEDURE As String = "PopulateForm()"
    
    On Error GoTo ErrorHandler
    
    With DevelopmentPlan
        TxtLocalDPNo = .LocalDPNo
        
    End With
    
    With Candidate
        TxtCourseNo = .Parent.CourseNo
        TxtCrewNo = .CrewNo
        TxtName = .Name
        
    End With
    
    With DevArea
        TxtCurrLvl = .CurrPerfLvl
        TxtImprovReqd = .ImproveLvl
        TxtRef = .Reference
        If .RevDate <> 0 Then TxtReviewDate = .RevDate
        TxtSupport = .Support
        TxtComments = .RevComments
        CmoArea = .DevArea
        CmoAssesor = .Assessor
        CmoModule = .Module.ModuleNo & " - " & .Module.Module
        ChkAchieved = .StandardMet
    End With

    PopulateForm = True

Exit Function

ErrorExit:
    PopulateForm = False
    FormTerminate
    Terminate

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
Private Sub CmoArea_Change()
    Const StrPROCEDURE As String = "CmoArea_Change()"
    
    On Error GoTo ErrorHandler
    
    If DevArea.DevArea = "" Then
        DevArea.DevArea = CmoArea.Value
        DevelopmentPlan.DevelopmentAreas.AddItem DevArea
    End If
    FormChanged = True
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
Private Sub CmoAssesor_Change()
    On Error Resume Next
    
    FormChanged = True
End Sub

Private Sub CmoModule_Change()
    On Error Resume Next
    
    FormChanged = True
End Sub

Private Sub TxtComments_Change()
    On Error Resume Next
    
    FormChanged = True
End Sub

Private Sub TxtCurrLvl_Change()
    On Error Resume Next
    
    FormChanged = True
End Sub

Private Sub TxtImprovReqd_Change()
    On Error Resume Next
    
    FormChanged = True
End Sub

Private Sub TxtRef_Change()
    On Error Resume Next
    
    FormChanged = True
End Sub

Private Sub TxtReviewDate_Change()
    On Error Resume Next
    
    FormChanged = True
End Sub

Private Sub TxtSupport_Change()
    On Error Resume Next
    
    FormChanged = True
End Sub

Private Sub UserForm_Initialize()
    On Error Resume Next
    
    FormInitialise
End Sub
    
Private Function ValidateData() As Boolean
    On Error Resume Next
    
    If Me.CmoArea = "" Then
        MsgBox "Please enter a Development Area"
        ValidateData = False
        Exit Function
    End If
    
    If Me.CmoModule = "" Then
        MsgBox "Please enter a Module"
        ValidateData = False
        Exit Function
    End If

   
    If Me.TxtRef = "" Then
        MsgBox "Please enter a reference.  If this is not applicable, enter 'None'"
        ValidateData = False
        Exit Function
    End If

    ValidateData = True
End Function


Public Sub FormInitialise()
    Const StrPROCEDURE As String = "FormInitialise()"
    
    On Error GoTo ErrorHandler
    
    Dim cell As Range
    Dim screenheight As Integer
    Dim Module As ClsModule
    Dim i As Integer
    Dim RstUsers As Recordset
        
    Set Module = New ClsModule
    Set RstUsers = GetAccessList
    
    With CmoArea
        .AddItem "Attitude"
        .AddItem "Practical Ability"
        .AddItem "Knowledge"
        .AddItem "Safety"
    End With
    
    Me.CmoAssesor.Clear
    
    With RstUsers
        Do
        Me.CmoAssesor.AddItem !UserName
        .MoveNext
        Loop While Not .EOF
    End With
    
    For i = 1 To ModGlobals.Modules.Count
        Set Module = Modules.FindItem(i)
        Me.CmoModule.AddItem Module.ModuleNo & " - " & Module.Module
    Next
   
    'set form height dependant on screen size
    screenheight = GetScreenHeight
      
    If screenheight = 0 Then Err.Raise HANDLED_ERROR
      
    If screenheight < 900 Then
        Me.Height = screenheight - 300
        Me.ScrollHeight = 600
    End If
    Set Module = Nothing
    Set RstUsers = Nothing
Exit Sub

ErrorExit:
    FormTerminate
    Terminate
    Set Module = Nothing
    Set RstUsers = Nothing

Exit Sub

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

Public Sub FormTerminate()
    On Error Resume Next
    
    Set Candidate = Nothing
    Set DevelopmentPlan = Nothing
    Set DevArea = Nothing
End Sub
