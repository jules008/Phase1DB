VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsXTrainingSessions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 13 Oct 16
'===============================================================
' Methods
'---------------------------------------------------------------
' FindItem - finds  item from index
' AddItem - Add item to collection
' removeItem - Delete item from collection
' Count - returns number of items in collection
' GetCollection - Find all items sessions
' DeleteCollection - deletes all items
' GetParentFromPtr - Private routine to copy memory address of parent class
'===============================================================
Option Explicit

Private pTrainingSessions As Collection
Private pOffered As Integer
Private pTaken As Integer
Private pRefused As Integer
Private pParent As Long

'---------------------------------------------------------------
Public Property Get Offered() As Variant
    pOffered = pTrainingSessions.Count
    Offered = pOffered
End Property

'---------------------------------------------------------------
Public Property Get Taken() As Variant
    Dim Session As ClsXTrainingSession
    
    For Each Session In pTrainingSessions
        If Session.TrainingTaken = True Then pTaken = pTaken + 1
    Next
    Taken = pTaken
    Set Session = Nothing
End Property

'---------------------------------------------------------------
Public Property Get Refused() As Variant
    Dim Session As ClsXTrainingSession
    
    For Each Session In pTrainingSessions
        If Session.TrainingTaken = False Then pRefused = pRefused + 1
    Next
    Refused = pRefused
    Set Session = Nothing
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
'Method FindItem
'finds training item in collection
'---------------------------------------------------------------
Public Function FindItem(Index As Variant) As ClsXTrainingSession
    On Error Resume Next
    Set FindItem = pTrainingSessions.Item(Index)
End Function

'===============================================================
'Method AddItem
'Add training sesssion to collection
'---------------------------------------------------------------
Public Sub AddItem(TrainingSession As ClsXTrainingSession)
    TrainingSession.SetParent ObjPtr(Me.Parent)
    pTrainingSessions.Add TrainingSession, Key:=CStr(TrainingSession.ExtraTrainingNo)
End Sub

'===============================================================
'Method removeItem
'deletes training item from index
'---------------------------------------------------------------
Public Sub RemoveItem(Index As Variant)
    pTrainingSessions.Remove Index
End Sub

'===============================================================
'Method count
'returns number of training sessions in collection
'---------------------------------------------------------------
Public Function Count() As Integer
    Count = pTrainingSessions.Count
End Function

'===============================================================
'Method GetCollection
'Find all extra training sessions from Crew no and module no
'---------------------------------------------------------------
Public Sub GetCollection()
    Dim TrainingSession As ClsXTrainingSession
    Dim RstTrgSession As Recordset
    Dim StrCrewNo As String
    Dim DayNo As Integer
    Dim i As Integer
    
    StrCrewNo = "'" & Parent.Parent.CrewNo & "'"
    DayNo = Parent.Module.DayNo
    
    Set RstTrgSession = database.SQLQuery("SELECT ExtraTrainingNo FROM ExtraTraining WHERE CrewNo = " & StrCrewNo & " AND DayNo = " & DayNo & " AND deleted IS NULL")
    
    With RstTrgSession
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            For i = 1 To .RecordCount
                Set TrainingSession = New ClsXTrainingSession
                TrainingSession.ExtraTrainingNo = !ExtraTrainingNo
                AddItem TrainingSession
                TrainingSession.LoadDB
                .MoveNext
            Next
        End If
    End With
    Set TrainingSession = Nothing
    Set RstTrgSession = Nothing
End Sub

'===============================================================
'Method DeleteCollection
'deletes all Training sessions in collection
'---------------------------------------------------------------
Public Sub DeleteCollection()
    Dim TrainingSession As ClsXTrainingSession
    Dim i As Integer
    
    For i = pTrainingSessions.Count To 1 Step -1
        Set TrainingSession = pTrainingSessions.Item(i)
        TrainingSession.DeleteDB
        pTrainingSessions.Remove (i)
    Next
    
    Set TrainingSession = Nothing
End Sub

'===============================================================
'Method CleanUp
'deletes unfinished items
'---------------------------------------------------------------
Public Sub CleanUp()
    Dim TrainingSession As ClsXTrainingSession
    
    On Error Resume Next
    
    For Each TrainingSession In pTrainingSessions
        If TrainingSession.TrainingDetails = "" Then
            TrainingSession.DeleteDB
            pTrainingSessions.Remove (TrainingSession.ExtraTrainingNo)
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
    Set pTrainingSessions = New Collection
End Sub

Private Sub Class_Terminate()
    Set pTrainingSessions = Nothing
End Sub
'===============================================================





