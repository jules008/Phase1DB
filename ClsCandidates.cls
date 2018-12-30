VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsCandidates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
' v0,1 - WT2019 Version
'---------------------------------------------------------------
' Date - 30 Dec 18
'===============================================================
' Methods
'---------------------------------------------------------------
' FindItem - finds item from index
' AddItem - Add item to collection
' removeItem - Delete item from collection
' Count - returns number of items in collection
' GetCollection - Find all items
' DeleteCollection - deletes all items
' CleanUp - deletes unfinished items
' GetParentFromPtr - Private routine to copy memory address of parent class
'===============================================================
Option Explicit

Private pCandidates As Collection
Private pParent As Long

'---------------------------------------------------------------
Public Property Get Parent() As ClsCourse
    If pParent <> 0 Then
        Set Parent = GetParentFromPtr(pParent)
    End If
End Property

Friend Function SetParent(ByVal Ptr As Long) As Boolean
    pParent = Ptr
End Function


'===============================================================
'Method FindItem
'finds item in collection
'---------------------------------------------------------------
Public Function FindItem(CrewNo As Variant) As ClsCandidate
    On Error Resume Next
    Set FindItem = pCandidates.Item(CrewNo)
End Function

'===============================================================
'Method AddItem
'Add item to collection
'---------------------------------------------------------------
Public Sub AddItem(Candidate As ClsCandidate)
    On Error Resume Next

    Candidate.SetParent ObjPtr(Me.Parent)
    pCandidates.Add Candidate, Key:=CStr(Candidate.CrewNo)
End Sub

'===============================================================
'Method removeItem
'deletes item from index
'---------------------------------------------------------------
Public Sub RemoveItem(CrewNo As Variant)
    On Error Resume Next
    pCandidates.Remove CrewNo
End Sub

'===============================================================
'Method count
'returns number items in collection
'---------------------------------------------------------------
Public Function Count() As Integer
    Count = pCandidates.Count
End Function

'===============================================================
'Method GetCollection
'Find all items from courseno
'---------------------------------------------------------------
Public Sub GetCollection()
    Dim Candidate As ClsCandidate
    Dim RstCandidates As Recordset
    Dim i As Integer
    
    Set RstCandidates = ModDatabase.SQLQuery("SELECT CrewNo FROM candidate WHERE courseno = 'WT2018'")
    
    With RstCandidates
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            For i = 1 To .RecordCount
                Set Candidate = New ClsCandidate
                Candidate.CrewNo = !CrewNo
                AddItem Candidate
                Candidate.LoadDB
                .MoveNext
            Next
        End If
    End With
    Set Candidate = Nothing
    Set RstCandidates = Nothing
End Sub

'===============================================================
'Method DeleteCollection
'deletes all Training sessions in collection
'---------------------------------------------------------------
Public Sub DeleteCollection()
    Dim Candidate As ClsCandidate
    Dim i As Integer
    
    For i = pCandidates.Count To 1 Step -1
        Set Candidate = pCandidates.Item(i)
        Candidate.DeleteDB
        pCandidates.Remove (i)
    Next
    
    Set Candidate = Nothing
End Sub


'===============================================================
'Method CleanUp
'deletes unfinished items
'---------------------------------------------------------------
Public Sub CleanUp()
    Dim Candidate As ClsCandidate
    Dim RstCandidate As Recordset
    Dim i As Integer
    
    Set RstCandidate = ModDatabase.SQLQuery("Candidate")
    
    With RstCandidate
        If .RecordCount > 0 Then
            Do While Not .EOF
                If IsNull(!Name) Then .Delete
                .MoveNext
            Loop
        End If
    End With
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
    Set pCandidates = New Collection
End Sub

Private Sub Class_Terminate()
    Set pCandidates = Nothing
End Sub
'===============================================================





