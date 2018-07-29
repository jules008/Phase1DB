VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsXTrainingSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 14 Sep 16
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
Private pExtraTrainingNo As Integer
Private pLocalNo As Integer
Private pTrainingDate As Date
Private pTrainingDetails As String
Private pTrainingTaken As Boolean
Private pTrainingResults As String
Private pParent As Long

'---------------------------------------------------------------
Public Property Get ExtraTrainingNo() As Integer
    ExtraTrainingNo = pExtraTrainingNo
End Property

Public Property Let ExtraTrainingNo(ByVal vNewValue As Integer)
    pExtraTrainingNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get LocalNo() As Integer
    LocalNo = pLocalNo
End Property

Public Property Let LocalNo(ByVal vNewValue As Integer)
    pLocalNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get TrainingDate() As Date
    TrainingDate = pTrainingDate
End Property

Public Property Let TrainingDate(ByVal vNewValue As Date)
    pTrainingDate = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get TrainingDetails() As String
    TrainingDetails = pTrainingDetails
End Property

Public Property Let TrainingDetails(ByVal vNewValue As String)
    pTrainingDetails = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get TrainingTaken() As Boolean
    TrainingTaken = pTrainingTaken
End Property

Public Property Let TrainingTaken(ByVal vNewValue As Boolean)
    pTrainingTaken = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get TrainingResults() As String
    TrainingResults = pTrainingResults
End Property

Public Property Let TrainingResults(ByVal vNewValue As String)
    pTrainingResults = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Parent() As ClsDailyLog
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
    Dim RstTrgSession As Recordset
    
    Set RstTrgSession = ModDatabase.SQLQuery("SELECT * FROM ExtraTraining WHERE Extratrainingno = " & pExtraTrainingNo & _
                                        " AND deleted IS NULL")
                                        
    With RstTrgSession
        If .RecordCount > 0 Then
            If Not IsNull(!ExtraTrainingNo) Then pExtraTrainingNo = !ExtraTrainingNo
            If Not IsNull(!TrainingDate) Then pTrainingDate = !TrainingDate
            If Not IsNull(!TrainingDetails) Then pTrainingDetails = !TrainingDetails
            If Not IsNull(!TrainingTaken) Then pTrainingTaken = !TrainingTaken
            If Not IsNull(!TrainingResults) Then pTrainingResults = !TrainingResults
        End If
    End With
    Set RstTrgSession = Nothing
End Sub

'===============================================================
'Method NewDB
'Creates a new item in the DB
'---------------------------------------------------------------
Public Function NewDB() As Integer
    Dim RstTrgSession As Recordset
    
    Set RstTrgSession = ModDatabase.SQLQuery("ExtraTraining")
        
    With RstTrgSession
        .AddNew
        NewDB = !ExtraTrainingNo
        .Update
    End With
    
    Set RstTrgSession = Nothing
End Function

'===============================================================
'Method UpdateDB
'Update session to database
'---------------------------------------------------------------
Public Function UpdateDB() As Boolean
    Dim RstTrgSession As Recordset

    Set RstTrgSession = ModDatabase.SQLQuery("SELECT * FROM ExtraTraining WHERE extratrainingno = " & pExtraTrainingNo & _
                                        " AND deleted IS NULL")
    
    With RstTrgSession
        If .RecordCount = 0 Then
            UpdateDB = False
        Else
            .Edit
            If pTrainingDate <> 0 Then !TrainingDate = pTrainingDate
            If pTrainingDetails <> "" Then !TrainingDetails = pTrainingDetails
            !TrainingTaken = pTrainingTaken
            If pTrainingResults <> "" Then !TrainingResults = pTrainingResults
            !CrewNo = Parent.Parent.CrewNo
            !DayNo = Parent.Module.DayNo
            .Update
            UpdateDB = True
        End If
    End With
    
    Set RstTrgSession = Nothing

End Function

'===============================================================
'Method DeleteDB
' Delete session from database
'---------------------------------------------------------------
Public Sub DeleteDB()
    Dim RstTrgSession As Recordset
    
    Set RstTrgSession = ModDatabase.SQLQuery("SELECT * FROM ExtraTraining WHERE ExtraTrainingNo = " & pExtraTrainingNo & " AND deleted IS NULL")
        
    With RstTrgSession
        If .RecordCount <> 0 Then
            .Edit
            !Deleted = Now
            .Update
        End If
    End With
    Set RstTrgSession = Nothing
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
