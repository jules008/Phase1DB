VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsDailyLogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 11 Oct 16
'===============================================================
' Methods
'---------------------------------------------------------------
' FindItem - finds Item from index
' AddItem - Add Item to collection
' RemoveItem - Delete Item from collection
' Count - returns number of Items in collection
' GetCollection - Find all Items from Crew no and module no
' CleanUp - deletes unfinished items
' DeleteCollection - deletes all training session from Crew No and Module No
' GetParentFromPtr - Private routine to copy memory address of parent class
'===============================================================
Option Explicit
Private pDailyLogs As Collection
Private pETOffered As Integer
Private pETTaken As Integer
Private pETRefused As Integer
Private pParent As Long

'---------------------------------------------------------------
Public Property Get ETOffered() As Integer
    Dim DailyLog As ClsDailyLog
    
    For Each DailyLog In pDailyLogs
        pETOffered = pETOffered + DailyLog.XtrainingSessions.Offered
    Next
    
    ETOffered = pETOffered
    
    Set DailyLog = Nothing
End Property

'---------------------------------------------------------------
Public Property Get ETTaken() As Integer
    Dim DailyLog As ClsDailyLog
    
    For Each DailyLog In pDailyLogs
        pETTaken = pETTaken + DailyLog.XtrainingSessions.Taken
    Next
    
    ETTaken = pETTaken
    
    Set DailyLog = Nothing
End Property

'---------------------------------------------------------------
Public Property Get ETRefused() As Integer
    Dim DailyLog As ClsDailyLog
    
    For Each DailyLog In pDailyLogs
        pETRefused = pETRefused + DailyLog.XtrainingSessions.Refused
    Next
    
    ETRefused = pETRefused
    
    Set DailyLog = Nothing
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
'Method FindItem
'finds Item from index
'---------------------------------------------------------------
Public Function FindItem(DayNo As Variant) As ClsDailyLog
    On Error Resume Next
    Set FindItem = pDailyLogs.Item(DayNo)
End Function

'===============================================================
'Method AddItem
'Add Item to collection
'---------------------------------------------------------------
Public Sub AddItem(DailyLog As ClsDailyLog)
    DailyLog.SetParent ObjPtr(Me.Parent)
    pDailyLogs.Add DailyLog, Key:=CStr(DailyLog.Module.DayNo)
End Sub

'===============================================================
'Method RemoveItem
'deletes Item from index
'---------------------------------------------------------------
Public Sub RemoveItem(DayNo As Variant)
    pDailyLogs.Remove DayNo
End Sub

'===============================================================
'Method count
'returns number of Items in collection
'---------------------------------------------------------------
Public Function Count() As Integer
    Count = pDailyLogs.Count
End Function

'===============================================================
'Method GetCollection
'Find all Items for parent candidate
'---------------------------------------------------------------
Public Sub GetCollection()
    Dim DailyLog As ClsDailyLog
    Dim RstDailyLog As Recordset
    Dim StrCrewNo As String
    Dim i As Integer

    StrCrewNo = "'" & Parent.CrewNo & "'"
    
    Set RstDailyLog = ModDatabase.SQLQuery("SELECT * FROM DailyLog WHERE CrewNo = " & StrCrewNo & _
                                        " AND deleted IS NULL")
    
    With RstDailyLog
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            For i = 1 To .RecordCount
                Set DailyLog = New ClsDailyLog
                DailyLog.Module.DayNo = !DayNo
                DailyLog.Module.LoadDB
                DailyLog.LogID = !LogID
                AddItem DailyLog
                DailyLog.LoadDB
                .MoveNext
            Next
       End If
       Set DailyLog = Nothing
       Set RstDailyLog = Nothing
    End With
End Sub

'===============================================================
'Method DeleteCollection
'deletes all Items matching Crew No
'---------------------------------------------------------------
Public Sub DeleteCollection()
    Dim DailyLog As ClsDailyLog
    Dim i As Integer
    
    For i = Me.Count To 1 Step -1
        Set DailyLog = pDailyLogs.Item(i)
        DailyLog.DeleteDB
        pDailyLogs.Remove (i)
    Next
    Set DailyLog = Nothing
End Sub

'===============================================================
'Method CleanUp
'deletes unfinished items
'---------------------------------------------------------------
Public Sub CleanUp()
    Dim DailyLog As ClsDailyLog
    
    For Each DailyLog In pDailyLogs
        If DailyLog.OverallGrade = 0 Then
            pDailyLogs.Remove CStr(DailyLog.Module.DayNo)
            DailyLog.DeleteDB
        End If
    Next
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

'===============================================================
Private Sub Class_Initialize()
    Set pDailyLogs = New Collection
End Sub

Private Sub Class_Terminate()
    Set pDailyLogs = Nothing
End Sub


