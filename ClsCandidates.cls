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
' Class ClsCandidates
' v0,0 - Initial version
' v0,1 - WT2018 Version
'---------------------------------------------------------------
' Date - 24 Dec 18
'===============================================================
' Methods
'---------------------------------------------------------------
' FindItem - finds item from index
' AddItem - Add item to collection
' RemoveItem - Deletes item from collection
' ItemCount - Returns number of items in collection
' GetCollection - Gets all items in collection
' DeleteCollection - Deletes all items in collection
' CleanUp - deletes unfinished items
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

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pCandidates.[_newenum]
End Function

' ===============================================================
' Method FindItem
' Finds Item from index
'---------------------------------------------------------------
Public Function FindItem(CrewNo As Variant) As ClsCandidate
Attribute FindItem.VB_UserMemId = 0
    On Error Resume Next
    Set FindItem = pCandidates.Item(CrewNo)
End Function

' ===============================================================
' Method AddItem
' Add item to Collection
'---------------------------------------------------------------

Public Sub AddItem(Candidate As ClsCandidate)
    Candidate.SetParent ObjPtr(Me)
    pCandidates.Add Candidate, Key:=CStr(Candidate.CrewNo)
End Sub

' ===============================================================
' Method RemoveItem
' Deletes item from collection
'---------------------------------------------------------------
Public Sub RemoveItem(CrewNo As Variant)
    pCandidates.Remove CrewNo
End Sub

' ===============================================================
' Method ItemCount
' Returns number of items in collection
'---------------------------------------------------------------
Public Function Count() As Integer
    Count = pCandidates.Count
End Function

' ===============================================================
' Method GetCollection
' Gets all items in collection
'---------------------------------------------------------------
Public Sub GetCollection()
    Dim RstCandidate As Recordset
    Dim Candidate As ClsCandidate
    Dim i As Integer

    Set RstCandidate = ModDatabase.SQLQuery("SELECT * FROM Candidate WHERE Deleted IS NULL")
    With RstCandidate
        .MoveLast
        .MoveFirst
        For i = 1 To .recordcount
            Set Candidate = New ClsCandidate
            Candidate.CrewNo = !CrewNo
            Candidate.LoadDB
            Me.AddItem Candidate
            .Movenext
        Next
    End With
End Sub

' ===============================================================
' Method DeleteCollection
' Deletes all items in collection
'---------------------------------------------------------------
Public Sub DeleteCollection()
    Dim Candidate As ClsCandidate
    For Each Candidate In pCandidates
        pCandidates.Remove CStr(Candidate.CrewNo)
        Candidate.DeleteDB
    Next
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

' ===============================================================
' Method GetParentFromPtr
' Private routine to copy memory address of parent class
' ---------------------------------------------------------------
Private Function GetParentFromPtr(ByVal Ptr As Long) As ClsCourse
    Dim tmp As ClsCourse

    CopyMemory tmp, Ptr, 4
    Set GetParentFromPtr = tmp
    CopyMemory tmp, 0&, 4
End Function

' ===============================================================
Private Sub Class_Initialize()
    Set pCandidates = New Collection
End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pCandidates = Nothing

End Sub

'---------------------------------------------------------------
