VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmDevelopmentPlanList 
   Caption         =   "Action Plan"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10425
   OleObjectBlob   =   "FrmDevelopmentPlanList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmDevelopmentPlanList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
' v0,1 - Added debug object checks
'---------------------------------------------------------------
' Date - 18 Jan 19
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmDevelopmentPlanList"

Private Candidate As ClsCandidate

Public Function ShowForm(LocalCandidate As ClsCandidate) As Boolean
    
    Const StrPROCEDURE As String = "ShowForm()"

    On Error GoTo ErrorHandler

    ResetForm
    
    If LocalCandidate Is Nothing Then
        MsgBox "Please select a candidate"
    Else
        Set Candidate = LocalCandidate
    End If
    
    If Not PopulateForm Then Err.Raise HANDLED_ERROR
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
Private Sub ResetForm()
    On Error Resume Next
    TxtClosedDPs = ""
    TxtCourseNo = ""
    TxtCrewNo = ""
    TxtName = ""
    TxtOpenDPs = ""
    TxtOverdueDPs = ""
End Sub

Private Function PopulateForm() As Boolean
    Const StrPROCEDURE As String = "PopulateForm()"
    
    Dim i As Integer
    Dim DevelopmentPlan As ClsDevelopmentPlan

    On Error GoTo ErrorHandler
    
    If Candidate Is Nothing Then Err.Raise HANDLED_ERROR, , "No Candidate Object"

    With Candidate
        TxtCrewNo = .CrewNo
        TxtCourseNo = .Parent.CourseNo
        TxtName = .Name
    End With
    
    With Candidate.DevelopmentPlans
        LstDPList.Clear
        
        For i = 1 To .Count
            
            Set DevelopmentPlan = .FindItem(i)
            
            
            With LstDPList
                .AddItem
                .List(i - 1, 0) = DevelopmentPlan.DPNo
                .List(i - 1, 1) = DevelopmentPlan.LocalDPNo
                .List(i - 1, 2) = Format(DevelopmentPlan.DPDate, "dd mmm yy")
                
                If developmentmentplan Is Nothing Then Err.Raise HANDLED_ERROR, , "No Development Plan available"
                
                With DevelopmentPlan
                    If .ReviewDate <> 0 Then LstDPList.List(i - 1, 3) = Format(DevelopmentPlan.ReviewDate, "dd mmm yy")
                    If .FollowOnDP <> 0 Then LstDPList.List(i - 1, 5) = .FollowOnDP
                End With
                .List(i - 1, 4) = DevelopmentPlan.Status
       
            End With
        Next
        
        TxtClosedDPs = .NoClosed
        TxtOpenDPs = .NoOpen
        TxtOverdueDPs = .NoOverDue
    End With
    Set DevelopmentPlan = Nothing

    PopulateForm = True

Exit Function

ErrorExit:
    FormTerminate
    Terminate
    Set DevelopmentPlan = Nothing
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

Private Sub BtnClose_Click()
    On Error Resume Next
    FormTerminate
    Me.Hide
End Sub
    
Private Sub BtnDelete_Click()

    Const StrPROCEDURE As String = "BtnDelete_Click()"
    
    Dim DPIndex As Integer
    Dim DPNo As Integer
    Dim DevelopmentPlan As ClsDevelopmentPlan
    
    On Error GoTo ErrorHandler
    
    DPIndex = LstDPList.ListIndex
    
    If DPIndex = -1 Then
        MsgBox "Please select an DP"
    Else
        DPNo = Me.LstDPList.List(DPIndex, 0)
        
        Set DevelopmentPlan = Candidate.DevelopmentPlans.FindItem(CStr(DPNo))
        DevelopmentPlan.DeleteDB
        Candidate.DevelopmentPlans.RemoveItem (CStr(DPNo))
        If Not PopulateForm Then Err.Raise HANDLED_ERROR
    End If
    
    Set DevelopmentPlan = Nothing
Exit Sub

ErrorExit:
    Set DevelopmentPlan = Nothing
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

Private Sub BtnEdit_Click()

    Const StrPROCEDURE As String = "BtnEdit_Click()"
    
    Dim DPIndex As Integer
    Dim DPNo As Integer
    Dim DevelopmentPlan As ClsDevelopmentPlan
    
    On Error GoTo ErrorHandler
    
    DPIndex = LstDPList.ListIndex
    
    If DPIndex = -1 Then
        MsgBox "Please select an DP"
    Else
        DPNo = Me.LstDPList.List(DPIndex, 0)
        
        Set DevelopmentPlan = Candidate.DevelopmentPlans.FindItem(CStr(DPNo))
        
        If Not FrmDevelopmentPlan.ShowForm(LocalDevelopmentPlan:=DevelopmentPlan) Then Err.Raise HANDLED_ERROR
    
        If Not PopulateForm Then Err.Raise HANDLED_ERROR
        Set DevelopmentPlan = Nothing
    End If
Exit Sub

ErrorExit:
    FormTerminate
    Terminate
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

Private Sub BtnNew_Click()
    Const StrPROCEDURE As String = "BtnNew_Click()"

    On Error GoTo ErrorHandler

    FrmDevelopmentPlan.ShowForm LocalCandidate:=Candidate
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
Private Sub UserForm_Activate()
    On Error Resume Next
    'list headings
    With Me.LstHeadings
        .Clear
        .AddItem
        .List(0, 0) = "DP No"
        .List(0, 1) = "Issue Date"
        .List(0, 2) = "Review Date"
        .List(0, 3) = "Status"
        .List(0, 4) = "Follow On DP"
    End With
    
End Sub

Public Sub FormTerminate()
    On Error Resume Next

    Set Candidate = Nothing
End Sub
