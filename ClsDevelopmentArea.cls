VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsDevelopmentArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 16 Sep 16
'===============================================================
' Methods
'---------------------------------------------------------------
' LoadDB - Load Class from DB
' NewDB - Creates a new Class in the DB
' UpdateDB - Update Class to DB
' DeleteDB - Delete Class from DB
' GetParentFromPtr - Private routine to copy memory address of parent class
'===============================================================
Option Explicit
Private pReference As String
Private pCurrPerfLvl As String
Private pImproveLvl As String
Private pSupport As String
Private pAssessor As String
Private pDevArea As String
Private pReviewStatus As String
Private pModule As ClsModule
Private pStandardMet As Boolean
Private pRevComments As String
Private pRevDate As Date
Private pParent As Long

'-----------------------------------------------------------------------
Public Property Get Reference() As String
    Reference = pReference
End Property

Public Property Let Reference(ByVal vNewValue As String)
    pReference = vNewValue
End Property

'-----------------------------------------------------------------------
Public Property Get CurrPerfLvl() As String
    CurrPerfLvl = pCurrPerfLvl
End Property

Public Property Let CurrPerfLvl(ByVal vNewValue As String)
    pCurrPerfLvl = vNewValue
End Property

'-----------------------------------------------------------------------
Public Property Get ImproveLvl() As String
    ImproveLvl = pImproveLvl
End Property

Public Property Let ImproveLvl(ByVal vNewValue As String)
    pImproveLvl = vNewValue
End Property

'-----------------------------------------------------------------------
Public Property Get Support() As String
    Support = pSupport
End Property

Public Property Let Support(ByVal vNewValue As String)
    pSupport = vNewValue
End Property

'-----------------------------------------------------------------------
Public Property Get Assessor() As String
    Assessor = pAssessor
End Property

Public Property Let Assessor(ByVal vNewValue As String)
    pAssessor = vNewValue
End Property

'-----------------------------------------------------------------------
Public Property Get DevArea() As String
    DevArea = pDevArea
End Property

Public Property Let DevArea(ByVal vNewValue As String)
    pDevArea = vNewValue
End Property

'-----------------------------------------------------------------------
Public Property Get ReviewStatus() As String
    Dim LocalStatus As String
           
    If pCurrPerfLvl = "" Then
        pReviewStatus = "Draft"
        ReviewStatus = pReviewStatus
        Exit Function
    End If
    
    If pImproveLvl = "" Then
        pReviewStatus = "Draft"
        ReviewStatus = pReviewStatus
    Exit Function
    End If
    
    If pSupport = "" Then
        pReviewStatus = "Draft"
        ReviewStatus = pReviewStatus
        Exit Function
    End If
            
    If pAssessor <> "" Then
                
        If pRevDate <> 0 Then
        
            If pStandardMet = True Then
                pReviewStatus = "Passed"
                ReviewStatus = pReviewStatus
                Exit Function
            Else
                pReviewStatus = "Failed"
                ReviewStatus = pReviewStatus
                Exit Function
            End If
        End If
    End If
                     
    pReviewStatus = "Open"
    
    ReviewStatus = pReviewStatus
End Property

Public Property Let ReviewStatus(ByVal vNewValue As String)
    pReviewStatus = vNewValue
End Property

'-----------------------------------------------------------------------
Public Property Get Module() As ClsModule
    Set Module = pModule
End Property

Public Property Let Module(ByVal Module As ClsModule)
    Set pModule = Module
End Property

'-----------------------------------------------------------------------
Public Property Get StandardMet() As Boolean
    StandardMet = pStandardMet
End Property

Public Property Let StandardMet(ByVal vNewValue As Boolean)
    pStandardMet = vNewValue
End Property

'-----------------------------------------------------------------------
Public Property Get RevComments() As String
    RevComments = pRevComments
End Property

Public Property Let RevComments(ByVal vNewValue As String)
    pRevComments = vNewValue
End Property

'-----------------------------------------------------------------------
Public Property Get RevDate() As Date
    RevDate = pRevDate
End Property

Public Property Let RevDate(ByVal vNewValue As Date)
    pRevDate = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Parent() As ClsDevelopmentPlan
    If pParent <> 0 Then
        Set Parent = GetParentFromPtr(pParent)
    End If
End Property

Friend Function SetParent(ByVal Ptr As Long) As Boolean
    pParent = Ptr
End Function

'===============================================================
'Method LoadDB
'Load Class from DB
'---------------------------------------------------------------
Public Sub LoadDB()
    Dim RstDevArea As Recordset
    Dim DPNo As Integer
    Dim StrDevArea As String
    
    DPNo = Parent.DPNo
    StrDevArea = "'" & pDevArea & "'"
    
    Set RstDevArea = database.SQLQuery("SELECT * FROM DPDevItems WHERE  DPNo = " & DPNo & _
                                        " AND DevArea = " & StrDevArea & _
                                        " AND deleted IS NULL")

    With RstDevArea
        If .RecordCount > 0 Then
            If Not IsNull(!DevArea) Then pDevArea = !DevArea
            If Not IsNull(!Reference) Then pReference = !Reference
            If Not IsNull(!CurrPerfLvl) Then pCurrPerfLvl = !CurrPerfLvl
            If Not IsNull(!ImproveLvl) Then pImproveLvl = !ImproveLvl
            If Not IsNull(!Support) Then pSupport = !Support
            If Not IsNull(!Assessor) Then pAssessor = !Assessor
            If Not IsNull(!ReviewStatus) Then pReviewStatus = !ReviewStatus
            If Not IsNull(!StandardMet) Then pStandardMet = !StandardMet
            If Not IsNull(!RevDate) Then pRevDate = !RevDate
            If Not IsNull(!RevComments) Then pRevComments = !RevComments
            
            pModule.DayNo = !DayNo
            pModule.LoadDB
        End If
        Set RstDevArea = Nothing
    End With
End Sub
   
'===============================================================
'Method NewDB
' Creates a new Class in the DB
'---------------------------------------------------------------
Public Sub NewDB()
    Dim RstStrDevArea As Recordset
    Dim DPNo As Integer
    Dim StrDevArea As String
    
    DPNo = Parent.DPNo
    StrDevArea = "'" & pDevArea & "'"
    
    Set RstStrDevArea = database.SQLQuery("SELECT * FROM DPDevItems WHERE  DPNo = " & DPNo & _
                                        " AND DevArea = " & StrDevArea & _
                                        " AND deleted IS NULL")
        
    With RstStrDevArea
        If .RecordCount <> 0 Then
            Debug.Print "Dev Area already in the database"
        Else
            .AddNew
            !DPNo = Parent.DPNo
            !DevArea = pDevArea
            .Update
        End If
    End With
    
    Set RstStrDevArea = Nothing
End Sub

'===============================================================
'Method UpdateDB
'Update Class to DB
'---------------------------------------------------------------
Public Function UpdateDB() As Boolean
    Dim RstDevArea As Recordset
    Dim DPNo As Integer
    Dim StrDevArea As String
    
    DPNo = Parent.DPNo
    StrDevArea = "'" & pDevArea & "'"
    
    Set RstDevArea = database.SQLQuery("SELECT * FROM DPDevItems WHERE  DPNo = " & DPNo & _
                                        " AND DevArea = " & StrDevArea & _
                                        " AND deleted IS NULL")
        
    With RstDevArea
        If .RecordCount = 0 Then
            UpdateDB = False
        Else
            .Edit
            !DevArea = pDevArea
            !Reference = pReference
            !CurrPerfLvl = pCurrPerfLvl
            !ImproveLvl = pImproveLvl
            !Support = pSupport
            !Assessor = pAssessor
            !ReviewStatus = pReviewStatus
            !DayNo = pModule.DayNo
            !StandardMet = pStandardMet
            !RevDate = pRevDate
            !RevComments = pRevComments
            .Update
            UpdateDB = True
            
        End If
    End With
    
    Set RstDevArea = Nothing
            
End Function

'===============================================================
'Method DeleteDB
' Delete Class from DB
'---------------------------------------------------------------
Public Sub DeleteDB()
    Dim RstDevArea As Recordset
    Dim DPNo As Integer
    Dim StrDevArea As String
    
    DPNo = Parent.DPNo
    StrDevArea = "'" & pDevArea & "'"
    
    Set RstDevArea = database.SQLQuery("SELECT * FROM DPDevItems WHERE  DPNo = " & DPNo & _
                                        " AND DevArea = " & StrDevArea & _
                                        " AND deleted IS NULL")

    With RstDevArea
        If .RecordCount <> 0 Then
            .Edit
            !Deleted = Now
            .Update
        End If
    End With
    Set RstDevArea = Nothing
End Sub

'===============================================================
'Method GetParentFromPtr
' Private routine to copy memory address of parent class
'---------------------------------------------------------------
Private Function GetParentFromPtr(ByVal Ptr As Long) As ClsDevelopmentArea
    Dim tmp As ClsDevelopmentArea
    
    CopyMemory tmp, Ptr, 4
    Set GetParentFromPtr = tmp
    CopyMemory tmp, 0&, 4
End Function

'===============================================================
Private Sub Class_Initialize()
    Set pModule = New ClsModule
End Sub

Private Sub Class_Terminate()
    Set pModule = Nothing
End Sub
'===============================================================


