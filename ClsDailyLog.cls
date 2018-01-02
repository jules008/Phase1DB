VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsDailyLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 24 Aug 16
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
Private pLogID As Integer
Private pDLDate As Date
Private pAssessor As String
Private pScore1 As Single
Private pScore2 As Single
Private pScore3 As Single
Private pScore4 As Single
Private pComments1 As String
Private pComments2 As String
Private pComments3 As String
Private pComments4 As String
Private pCommentsMisc As String
Private pOverallGrade As Single
Private pDPNo As Integer
Private pModule As ClsModule
Private pXtrainingSessions As ClsXTrainingSessions
Private pParent As Long

'-----------------------------------------------------------------
Public Property Get LogID() As Integer
    LogID = pLogID
End Property

Public Property Let LogID(ByVal vNewValue As Integer)
    pLogID = vNewValue
End Property

'-----------------------------------------------------------------
Public Property Get DLDate() As Date
    DLDate = pDLDate
End Property

Public Property Let DLDate(ByVal vNewValue As Date)
    pDLDate = vNewValue
End Property

'------------------------------------------------------------------------------
Public Property Get Module() As ClsModule
    Set Module = pModule
End Property

Public Property Let Module(ByVal vNewValue As ClsModule)
    Set pModule = vNewValue
End Property

'-----------------------------------------------------------------
Public Property Get Assessor() As String
    Assessor = pAssessor
End Property

Public Property Let Assessor(ByVal vNewValue As String)
    pAssessor = vNewValue
End Property

'-----------------------------------------------------------------
Public Property Get Score1() As Single
    Score1 = pScore1
End Property

Public Property Let Score1(ByVal vNewValue As Single)
    pScore1 = vNewValue
End Property

'-----------------------------------------------------------------
Public Property Get Score2() As Single
    Score2 = pScore2
End Property

Public Property Let Score2(ByVal vNewValue As Single)
    pScore2 = vNewValue
End Property

'-----------------------------------------------------------------
Public Property Get Score3() As Single
    Score3 = pScore3
End Property

Public Property Let Score3(ByVal vNewValue As Single)
    pScore3 = vNewValue
End Property

'-----------------------------------------------------------------
Public Property Get Score4() As Single
    Score4 = pScore4
End Property

Public Property Let Score4(ByVal vNewValue As Single)
    pScore4 = vNewValue
End Property

'-----------------------------------------------------------------
Public Property Get Comments1() As String
    Comments1 = pComments1
End Property

Public Property Let Comments1(ByVal vNewValue As String)
    pComments1 = vNewValue
End Property

'-----------------------------------------------------------------
Public Property Get Comments2() As String
    Comments2 = pComments2
End Property

Public Property Let Comments2(ByVal vNewValue As String)
    pComments2 = vNewValue
End Property

'-----------------------------------------------------------------
Public Property Get Comments3() As String
    Comments3 = pComments3
End Property

Public Property Let Comments3(ByVal vNewValue As String)
    pComments3 = vNewValue
End Property

'-----------------------------------------------------------------
Public Property Get Comments4() As String
    Comments4 = pComments4
End Property

Public Property Let Comments4(ByVal vNewValue As String)
    pComments4 = vNewValue
End Property

'-----------------------------------------------------------------
Public Property Get CommentsMisc() As String
    CommentsMisc = pCommentsMisc
End Property

Public Property Let CommentsMisc(ByVal vNewValue As String)
    pCommentsMisc = vNewValue
End Property

'-----------------------------------------------------------------
Public Property Get OverallGrade() As Single
    OverallGrade = pOverallGrade
End Property

Public Property Let OverallGrade(ByVal vNewValue As Single)
    pOverallGrade = vNewValue
End Property

'-----------------------------------------------------------------
Public Property Get DPNo() As Integer
    DPNo = pDPNo
End Property

Public Property Let DPNo(ByVal vNewValue As Integer)
    pDPNo = vNewValue
End Property

'------------------------------------------------------------------------------
Public Property Get XtrainingSessions() As ClsXTrainingSessions
    Set XtrainingSessions = pXtrainingSessions
End Property

Public Property Let XtrainingSessions(ByVal vNewValue As ClsXTrainingSessions)
    Set pXtrainingSessions = vNewValue
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
'Load Class from DB
'---------------------------------------------------------------
Public Sub LoadDB()
    Dim RstDailyLog As Recordset
    Dim StrCrewNo As String
    Dim DayNo As Integer
    
    StrCrewNo = "'" & Parent.CrewNo & "'"
    DayNo = pModule.DayNo
    
    Set RstDailyLog = database.SQLQuery("SELECT * FROM dailylog WHERE  CrewNo = " & StrCrewNo & " AND DayNo = " & DayNo & " AND deleted IS NULL")

    With RstDailyLog
        If .RecordCount > 0 Then
            If Not IsNull(!LogID) Then pLogID = !LogID
            If Not IsNull(!assessdate) Then pDLDate = !assessdate
            If Not IsNull(!Assessor) Then pAssessor = !Assessor
            If Not IsNull(!Score_1) Then pScore1 = !Score_1
            If Not IsNull(!Score_2) Then pScore2 = !Score_2
            If Not IsNull(!Score_3) Then pScore3 = !Score_3
            If Not IsNull(!Score_4) Then pScore4 = !Score_4
            If Not IsNull(!OverallGrade) Then pOverallGrade = !OverallGrade
            If Not IsNull(!comments_1) Then pComments1 = !comments_1
            If Not IsNull(!comments_2) Then pComments2 = !comments_2
            If Not IsNull(!comments_3) Then pComments3 = !comments_3
            If Not IsNull(!comments_4) Then pComments4 = !comments_4
            If Not IsNull(!comments_misc) Then pCommentsMisc = !comments_misc
            If Not IsNull(!DevPlanNo) Then pDPNo = !DevPlanNo
    
            Module.LoadDB
            pXtrainingSessions.GetCollection
            
        End If
        Set RstDailyLog = Nothing
    End With
End Sub

'===============================================================
'Method NewDB
' Creates a new Class in the DB
'---------------------------------------------------------------
Public Sub NewDB()
    Dim RstDailyLog As Recordset
    Dim StrCrewNo As String
    Dim DayNo As Integer
        
    StrCrewNo = "'" & Parent.CrewNo & "'"
    DayNo = pModule.DayNo
    
    Set RstDailyLog = database.SQLQuery("SELECT * FROM dailylog WHERE  CrewNo = " & StrCrewNo & " AND DayNo = " & DayNo & " AND deleted IS NULL")
        
    With RstDailyLog
        If .RecordCount <> 0 Then
            Debug.Print "Daily log already in the database"
        Else
            .AddNew
            !CrewNo = Parent.CrewNo
            !DayNo = pModule.DayNo
            .Update
        End If
    End With
    
    Set RstDailyLog = Nothing
End Sub

'===============================================================
'Method UpdateDB
'Update Class to DB
'---------------------------------------------------------------
Public Function UpdateDB() As Boolean
    Dim RstDailyLog As Recordset
    Dim StrCrewNo As String
    Dim DayNo As Integer
        
    StrCrewNo = "'" & Parent.CrewNo & "'"
    DayNo = pModule.DayNo
    
    Set RstDailyLog = database.SQLQuery("SELECT * FROM dailylog WHERE  CrewNo = " & StrCrewNo & " AND DayNo = " & DayNo & " AND deleted IS NULL")

    With RstDailyLog
        If .RecordCount = 0 Then
            UpdateDB = False
        Else
            .Edit
            !assessdate = pDLDate
            !Assessor = pAssessor
            !Score_1 = pScore1
            !Score_2 = pScore2
            !Score_3 = pScore3
            !Score_4 = pScore4
            !comments_1 = pComments1
            !comments_2 = pComments2
            !comments_3 = pComments3
            !comments_4 = pComments4
            !comments_misc = pCommentsMisc
            !OverallGrade = pOverallGrade
            !DevPlanNo = pDPNo
            .Update
            UpdateDB = True
        End If
    End With
    
    Set RstDailyLog = Nothing
            
End Function

'===============================================================
'Method DeleteDB
' Delete Class from DB
'---------------------------------------------------------------
Public Sub DeleteDB()
    Dim RstDailyLog As Recordset
    
    Set RstDailyLog = database.SQLQuery("SELECT * FROM dailylog WHERE LogID = " & pLogID & " AND deleted IS NULL")

    With RstDailyLog
        If .RecordCount <> 0 Then
            .Edit
            !Deleted = Now
            .Update
            pXtrainingSessions.DeleteCollection
        End If
    End With
    Set RstDailyLog = Nothing
End Sub

'===============================================================
'Method GetParentFromPtr
' Private routine to copy memory address of parent class
'---------------------------------------------------------------
Private Function GetParentFromPtr(ByVal Ptr As Long) As ClsCandidate
    Dim tmp As ClsDailyLog
    
    CopyMemory tmp, Ptr, 4
    Set GetParentFromPtr = tmp
    CopyMemory tmp, 0&, 4
End Function

'===============================================================
Private Sub Class_Initialize()
    Set pXtrainingSessions = New ClsXTrainingSessions
    pXtrainingSessions.SetParent ObjPtr(Me)
    Set pModule = New ClsModule
End Sub

Private Sub Class_Terminate()
    pXtrainingSessions.SetParent 0
    Set pXtrainingSessions = Nothing
    Set pModule = Nothing
End Sub
'===============================================================

