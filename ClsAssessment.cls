VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsAssessment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 03 Nov 16
'===============================================================
' Methods
'---------------------------------------------------------------
' LoadDB - Load Item from DB
' NewDB - creates a new Item in the DB
' UpdateDB - Update Item to database
' DeleteDB - Delete Item from database
' GetParentFromPtr - Private routine to copy memory address of parent class
'===============================================================
Option Explicit
Private pAssessmentNo As Integer
Private pModule As ClsModule
Private pAssessType As String
Private pScore As Integer
Private pAttempt As Integer
Private pParent As Long

'---------------------------------------------------------------
Public Property Get AssessmentNo() As Integer
    AssessmentNo = pAssessmentNo
End Property

Public Property Let AssessmentNo(ByVal vNewValue As Integer)
    pAssessmentNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Module() As ClsModule
    Set Module = pModule
End Property

Public Property Let DModuleayNo(ByVal vNewValue As ClsModule)
    Set pModule = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get AssessType() As String
    AssessType = pAssessType
End Property

Public Property Let AssessType(ByVal vNewValue As String)
    pAssessType = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Score() As Integer
    Score = pScore
End Property

Public Property Let Score(ByVal vNewValue As Integer)
    pScore = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Attempt() As Integer
    Attempt = pAttempt
End Property

Public Property Let Attempt(ByVal vNewValue As Integer)
    pAttempt = vNewValue
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
'load Item from database
'---------------------------------------------------------------
Public Sub LoadDB()
    Dim RstAssessment As Recordset
    
    Set RstAssessment = ModDatabase.SQLQuery("SELECT * FROM Assessment WHERE AssessmentNo = " & pAssessmentNo & _
                                        " AND deleted IS NULL")
                                        
    With RstAssessment
        If .RecordCount > 0 Then
            If Not IsNull(!AssessmentNo) Then pAssessmentNo = !AssessmentNo
            If Not IsNull(!AssessType) Then pAssessType = !AssessType
            If Not IsNull(!Score) Then pScore = !Score
            If Not IsNull(!Attempt) Then pAttempt = !Attempt
                        
            Set pModule = Modules.FindItem(!DayNo)
        End If
    End With
    Set RstAssessment = Nothing
End Sub

'===============================================================
'Method NewDB
'Creates a new item in the DB
'---------------------------------------------------------------
Public Sub NewDB()
    Dim RstAssessment As Recordset
    
    Set RstAssessment = ModDatabase.SQLQuery("Assessment")
        
    With RstAssessment
        .AddNew
        pAssessmentNo = !AssessmentNo
        .Update
    End With
    
    Set RstAssessment = Nothing
End Sub

'===============================================================
'Method UpdateDB
'Update session to database
'---------------------------------------------------------------
Public Function UpdateDB() As Boolean
    Dim RstAssessment As Recordset

    Set RstAssessment = ModDatabase.SQLQuery("SELECT * FROM Assessment WHERE AssessmentNo = " & pAssessmentNo & _
                                        " AND deleted IS NULL")
    
    With RstAssessment
        If .RecordCount = 0 Then
            UpdateDB = False
        Else
            .Edit
            If pAssessType <> "" Then !AssessType = pAssessType
            !Score = pScore
            !CrewNo = Parent.CrewNo
            !DayNo = pModule.DayNo
            !Attempt = pAttempt
            .Update
            UpdateDB = True
        End If
    End With
    
    Set RstAssessment = Nothing

End Function

'===============================================================
'Method DeleteDB
' Delete session from database
'---------------------------------------------------------------
Public Sub DeleteDB()
    Dim RstAssessment As Recordset
    
    Set RstAssessment = ModDatabase.SQLQuery("SELECT * FROM Assessment WHERE AssessmentNo = " & pAssessmentNo & " AND deleted IS NULL")
        
    With RstAssessment
        If .RecordCount <> 0 Then
            .Edit
            !Deleted = Now
            .Update
        End If
    End With
    Set RstAssessment = Nothing
End Sub

'===============================================================
'Method GetParentFromPtr
' Private routine to copy memory address of parent class
'---------------------------------------------------------------
Private Function GetParentFromPtr(ByVal Ptr As Long) As ClsCandidate
    Dim tmp As ClsCandidate
    
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

