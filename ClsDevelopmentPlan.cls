VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsDevelopmentPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
' v0,1 - Bug Fix - Check if DP worksheet is already open
'---------------------------------------------------------------
' Date - 09/01/17
'===============================================================
' Methods
'---------------------------------------------------------------
' LoadDB - Load Class from DB
' NewDB - Creates a new Class in the DB
' UpdateDB - Update Class to DB
' DeleteDB - Delete Class from DB
' SetDPStatus - processes DP Status
' PrintForm - Prints DevelopmentPlan
' GetParentFromPtr - Private routine to copy memory address of parent class
'===============================================================
Private Const StrMODULE As String = "ClsDevelopmentPlan"
Option Explicit
Private pDPNo As Integer
Private pLocalDPNo As String
Private pDPDate As Date
Private pIssuer As String
Private pStatus As String
Private pOutcomeIfNotMet As String
Private pReviewDate As Date
Private pFollowOnDP As Integer
Private pDevelopmentAreas As ClsDevelopmentAreas
Private pDevPlanTemplate As Worksheet
Private pDevPlanSheet As Worksheet
Private pParent As Long

'-------------------------------------------------------------------------
Public Property Get DPNo() As Integer
    DPNo = pDPNo
End Property

Public Property Let DPNo(ByVal vNewValue As Integer)
    pDPNo = vNewValue
End Property

'-------------------------------------------------------------------------
Public Property Get LocalDPNo() As String
    LocalDPNo = pLocalDPNo
End Property

Public Property Let LocalDPNo(ByVal vNewValue As String)
    pLocalDPNo = vNewValue
End Property

'-------------------------------------------------------------------------
Public Property Get DPDate() As Date
    DPDate = pDPDate
End Property

Public Property Let DPDate(ByVal vNewValue As Date)
    pDPDate = vNewValue
End Property

'-------------------------------------------------------------------------
Public Property Get Issuer() As String
    Issuer = pIssuer
End Property

Public Property Let Issuer(ByVal vNewValue As String)
    pIssuer = vNewValue
End Property

'-------------------------------------------------------------------------
Public Property Get Status() As String
    Dim i As Integer
    Dim StatusPassed As Integer
    Dim StatusFailed As Integer
    Dim StatusDraft As Integer
    Dim StatusOpen As Integer
    Dim DevArea As ClsDevelopmentArea
    
    For i = 1 To pDevelopmentAreas.Count
        Set DevArea = pDevelopmentAreas.FindItem(i)
        
        Select Case DevArea.ReviewStatus
            Case "Open"
                StatusOpen = StatusOpen + 1
            Case "Draft"
                StatusDraft = StatusDraft + 1
            Case "Passed"
                StatusPassed = StatusPassed + 1
            Case "Failed"
                StatusFailed = StatusFailed + 1
        End Select
    Next
    
    pStatus = SetDPStatus(StatusDraft, StatusPassed, StatusOpen, StatusFailed)
    
    Status = pStatus
    
    Set DevArea = Nothing
End Property

Public Property Let Status(ByVal vNewValue As String)
    pStatus = vNewValue
End Property

'-------------------------------------------------------------------------
Public Property Get OutcomeIfNotMet() As String
    OutcomeIfNotMet = pOutcomeIfNotMet
End Property

Public Property Let OutcomeIfNotMet(ByVal vNewValue As String)
    pOutcomeIfNotMet = vNewValue
End Property

'-------------------------------------------------------------------------
Public Property Get ReviewDate() As Date
    ReviewDate = pReviewDate
End Property

Public Property Let ReviewDate(ByVal vNewValue As Date)
    pReviewDate = vNewValue
End Property

'-------------------------------------------------------------------------
Public Property Get FollowOnDP() As Integer
    FollowOnDP = pFollowOnDP
End Property

Public Property Let FollowOnDP(ByVal vNewValue As Integer)
    pFollowOnDP = vNewValue
End Property

'-------------------------------------------------------------------------
Public Property Get DevelopmentAreas() As ClsDevelopmentAreas
    Set DevelopmentAreas = pDevelopmentAreas
End Property

Public Property Let DevelopmentAreas(ByVal vNewValue As ClsDevelopmentAreas)
    Set pDevelopmentAreas = vNewValue
End Property

'-------------------------------------------------------------------------
Public Property Get DevPlanTemplate() As Worksheet
    Set DevPlanTemplate = pDevPlanTemplate
End Property

Public Property Let DevPlanTemplate(ByVal vNewValue As Worksheet)
    Set pDevPlanTemplate = vNewValue
End Property

'-------------------------------------------------------------------------
Public Property Get DevPlanSheet() As Worksheet
    Set DevPlanSheet = pDevPlanSheet
End Property

Public Property Let DevPlanSheet(ByVal vNewValue As Worksheet)
    Set pDevPlanSheet = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Parent() As ClsCandidate
    If pParent <> 0 Then
        Set Parent = GetParentFromPtr(pParent)
    End If
End Property

Friend Function SetParent(ByVal Ptr As Long) As Boolean
    pParent = Ptr
End Function

'===============================================================
'Method LoadDB
'load training session from database
'---------------------------------------------------------------
Public Sub LoadDB()
    Dim RstDevelopmentPlan As Recordset
    
    Set RstDevelopmentPlan = ModDatabase.SQLQuery("SELECT * FROM DevelopmentPlan WHERE dpno = " & pDPNo & _
                                        " AND deleted IS NULL")
    
    With RstDevelopmentPlan
        If .RecordCount > 0 Then
            If Not IsNull(!DPNo) Then pDPNo = !DPNo
            If Not IsNull(!DPDate) Then pDPDate = !DPDate
            If Not IsNull(!LocalDPNo) Then pLocalDPNo = !LocalDPNo
            If Not IsNull(!Issuer) Then pIssuer = !Issuer
            If Not IsNull(!Status) Then pStatus = !Status
            If Not IsNull(!OutcomeIfNotMet) Then pOutcomeIfNotMet = !OutcomeIfNotMet
            If Not IsNull(!ReviewDate) Then pReviewDate = !ReviewDate
            If Not IsNull(!FollowOnDP) Then pFollowOnDP = !FollowOnDP
            pDevelopmentAreas.GetCollection
        End If
    End With
    Set RstDevelopmentPlan = Nothing
End Sub

'===============================================================
'Method NewDB
'Creates a new DP in the DB
'---------------------------------------------------------------
Public Function NewDB() As Integer
    Dim RstDevelopmentPlan As Recordset
    
    Set RstDevelopmentPlan = ModDatabase.SQLQuery("DevelopmentPlan")
        
    With RstDevelopmentPlan
        .AddNew
        !CrewNo = Parent.CrewNo
        .Update
        .MoveLast
        pDPNo = !DPNo
        Parent.DevelopmentPlans.SetLocalNos
    End With
    
    With Parent.DevelopmentPlans
        .RemoveItem ("0")
        .AddItem Me
    End With
    
    Set RstDevelopmentPlan = Nothing
End Function

'===============================================================
'Method UpdateDB
'Update DP to database
'---------------------------------------------------------------
Public Function UpdateDB() As Boolean
    Dim RstDevelopmentPlan As Recordset
    
    Set RstDevelopmentPlan = ModDatabase.SQLQuery("SELECT * FROM DevelopmentPlan WHERE dpno = " & pDPNo & _
                                        " AND deleted IS NULL")
    
    With RstDevelopmentPlan
        If .RecordCount = 0 Then
            UpdateDB = False
        Else
            .Edit
            !CrewNo = Parent.CrewNo
            If pLocalDPNo <> "" Then !LocalDPNo = pLocalDPNo
            !DPDate = pDPDate
            !Issuer = pIssuer
            !Status = pStatus
            !OutcomeIfNotMet = pOutcomeIfNotMet
            If pReviewDate <> 0 Then !ReviewDate = pReviewDate
            If pFollowOnDP <> 0 Then !FollowOnDP = pFollowOnDP
            .Update
            UpdateDB = True
        End If
    End With
    Set RstDevelopmentPlan = Nothing
End Function

'===============================================================
'Method DeleteDB
' Delete DP from database
'---------------------------------------------------------------
Public Sub DeleteDB()
    Dim RstDevelopmentPlan As Recordset
    
    Set RstDevelopmentPlan = ModDatabase.SQLQuery("SELECT * FROM DevelopmentPlan WHERE Dpno = " & pDPNo & _
                                        " AND deleted IS NULL")
        
    With RstDevelopmentPlan
        If .RecordCount <> 0 Then
            .Edit
            !Deleted = Now
            .Update
            Parent.DevelopmentPlans.SetLocalNos
            pDevelopmentAreas.DeleteCollection
        End If
    End With
    Set RstDevelopmentPlan = Nothing
End Sub

'===============================================================
' Method SetDPStatus
' processes DP Status
'---------------------------------------------------------------
Private Function SetDPStatus(SDraft As Integer, SPassed As Integer, SOpen As Integer, SFailed As Integer) As String
    Dim STotal As Integer
    
    STotal = SDraft + SPassed + SOpen + SFailed
    
    If STotal = 0 Then
        SetDPStatus = "Draft"
        Exit Function
    End If
        
    If SDraft <> 0 Then
        SetDPStatus = "Draft"
        Exit Function
    End If
      
    If SPassed = STotal Then
        SetDPStatus = "Passed"
        Exit Function
    End If
              
    If SOpen <> 0 Then
        SetDPStatus = "Open"
        Exit Function
    End If
   
    If SPassed + SFailed = STotal Then
        SetDPStatus = "Failed"
        Exit Function
    End If

End Function

'===============================================================
' Method PrintForm
' Prints DevelopmentPlan
'---------------------------------------------------------------
Public Sub PrintForm()
    Const StrPROCEDURE As String = "PrintForm()"

    Dim LocDevArea As ClsDevelopmentArea
    Dim DPWorkbook As Workbook
    Dim i As Integer
    Dim NoCopies As Integer
    
    On Error GoTo ErrorHandler

    
    'V0.1 - changes ***************************************
    
    Application.Workbooks.Add (xlWBATWorksheet)
    Set DPWorkbook = ActiveWorkbook
    '******************************************************
   
    ShtDPTemplate.Visible = xlSheetVisible
    ShtDPTemplate.Copy after:=ActiveWorkbook.Sheets(1)
    Set pDevPlanSheet = ActiveSheet
    ShtDPTemplate.Visible = xlSheetHidden
    
    
    With pDevPlanSheet
        .Visible = xlSheetVisible
        .Activate
        .Range("A9:F9").ClearContents
        .Range("dpno") = pLocalDPNo
        .Range("DPDate") = pDPDate
        .Range("Issuer") = pIssuer
        .Range("Outcome") = pOutcomeIfNotMet
        .Range("followonap") = pFollowOnDP
        .Range("crewno") = Parent.CrewNo
        .Range("name") = Parent.Name
        .Range("crseno") = Parent.Parent.CourseNo
            
        For i = 1 To pDevelopmentAreas.Count
    
            'insert row
            .Range("9:9").Rows.Insert
            .Range("9:9").Cells.Interior.ColorIndex = 0
            .Range("9:9").Cells.Font.Bold = False
            .Range("9:9").Cells.Font.Size = 9
            .Range("B9:E9").Cells.HorizontalAlignment = xlLeft
            .Range("F9").NumberFormat = "dd mmmm yy"
            
            'get copy of development area
            Set LocDevArea = pDevelopmentAreas.FindItem(i)
            
            .Range("A9").Offset(0, 0) = LocDevArea.Reference
            .Range("A9").Offset(0, 1) = LocDevArea.DevArea
            .Range("A9").Offset(0, 2) = LocDevArea.CurrPerfLvl
            .Range("A9").Offset(0, 3) = LocDevArea.ImproveLvl
            .Range("A9").Offset(0, 4) = LocDevArea.Support
            If LocDevArea.RevDate <> 0 Then .Range("A9").Offset(0, 5) = LocDevArea.RevDate
            .Range("A9").Offset(0, 0).Rows.AutoFit
            
        Next
        .Range("A9").Offset(i - 1, 1).EntireRow.Delete
    
        'check if the DP has gone onto two pages
        If .HPageBreaks.Count > 0 Then
            .ResetAllPageBreaks
            .Range("PageBreak").PageBreak = xlPageBreakManual
        End If
    End With
        
    Set LocDevArea = Nothing
    Set pDevPlanSheet = Nothing
Exit Sub

ErrorExit:
    ShtDPTemplate.Visible = xlSheetHidden
    
    Set LocDevArea = Nothing
    Set pDevPlanSheet = Nothing
    
Exit Sub

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub
'===============================================================
'Method GetParentFromPtr
' Private routine to copy memory address of parent class
'---------------------------------------------------------------
Private Function GetParentFromPtr(ByVal Ptr As Long) As ClsDailyLog
    Dim tmp As ClsDailyLog
    
    CopyMemory tmp, Ptr, 4
    Set GetParentFromPtr = tmp
    CopyMemory tmp, 0&, 4
End Function

'=========================================================================
Private Sub Class_Initialize()
    Set pDevelopmentAreas = New ClsDevelopmentAreas
    Set pDevPlanTemplate = ShtDPTemplate
    pDevelopmentAreas.SetParent ObjPtr(Me)
End Sub

Private Sub Class_Terminate()
    pDevelopmentAreas.SetParent 0
    Set pDevPlanTemplate = Nothing
    Set pDevPlanSheet = Nothing
    Set pDevelopmentAreas = Nothing
End Sub
'=========================================================================















