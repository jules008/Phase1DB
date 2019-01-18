VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsAssessments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
' v0,1 - Check for zero before dividing average
' v0,2 - Improved clean up routine
'---------------------------------------------------------------
' Date - 18 Jan 19
'===============================================================
' Methods
'---------------------------------------------------------------
' FindItem - finds  item from index
' AddItem - Add item to collection
' removeItem - Delete item from collection
' Count - returns number of items in collection
' GetCollection - Find all items sessions
' DeleteCollection - deletes all items
' NextAttemptNo - Returns next attempt number for given assessment
' GetParentFromPtr - Private routine to copy memory address of parent class
'===============================================================
Option Explicit

Private pAssessments As Collection
Private pAverageScore As Single
Private pParent As Long

'---------------------------------------------------------------
Public Property Get AverageScore() As Single
    Dim Assessment As ClsAssessment
    Dim TotalScore As Integer
    Dim AssCount As Integer
    
    If pAssessments.Count > 0 Then
        For Each Assessment In pAssessments
            With Assessment
                If .Attempt = 1 Then
                    TotalScore = TotalScore + .Score
                    AssCount = AssCount + 1
                End If
            End With
        Next
        If AssCount > 0 Then pAverageScore = TotalScore / AssCount Else pAverageScore = 0
    Else
        pAverageScore = 0
    End If
    AverageScore = pAverageScore
    
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
'finds training item in collection
'---------------------------------------------------------------
Public Function FindItem(Index As Variant) As ClsAssessment
    On Error Resume Next
    Set FindItem = pAssessments.Item(Index)
End Function

'===============================================================
'Method AddItem
'Add training sesssion to collection
'---------------------------------------------------------------
Public Sub AddItem(Assessment As ClsAssessment)
    Assessment.SetParent ObjPtr(Me.Parent)
    pAssessments.Add Assessment, Key:=CStr(Assessment.AssessmentNo)
End Sub

'===============================================================
'Method removeItem
'deletes training item from index
'---------------------------------------------------------------
Public Sub RemoveItem(Index As Variant)
    pAssessments.Remove Index
End Sub

'===============================================================
'Method count
'returns number of training sessions in collection
'---------------------------------------------------------------
Public Function Count() As Integer
    Count = pAssessments.Count
End Function

'===============================================================
'Method GetCollection
'Find all extra training sessions from Crew no
'---------------------------------------------------------------
Public Sub GetCollection()
    Dim Assessment As ClsAssessment
    Dim RstAssessment As Recordset
    Dim StrCrewNo As String
    Dim i As Integer
    
    StrCrewNo = "'" & Parent.CrewNo & "'"
    
    Set RstAssessment = ModDatabase.SQLQuery("SELECT AssessmentNo FROM assessment WHERE " & _
                                            " CrewNo = " & StrCrewNo & _
                                            " AND deleted IS NULL")
    
    With RstAssessment
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            For i = 1 To .RecordCount
                Set Assessment = New ClsAssessment
                Assessment.AssessmentNo = !AssessmentNo
                AddItem Assessment
                Assessment.LoadDB
                .MoveNext
            Next
        End If
    End With
    Set Assessment = Nothing
    Set RstAssessment = Nothing
End Sub

'===============================================================
'Method DeleteCollection
'deletes all Training sessions in collection
'---------------------------------------------------------------
Public Sub DeleteCollection()
    Dim Assessment As ClsAssessment
    Dim i As Integer
    
    For i = pAssessments.Count To 1 Step -1
        Set Assessment = pAssessments.Item(i)
        Assessment.DeleteDB
        pAssessments.Remove (i)
    Next
    
    Set Assessment = Nothing
End Sub

'===============================================================
'Method CleanUp
'deletes unfinished items
'---------------------------------------------------------------
Public Sub CleanUp()
    Dim Assessment As ClsAssessment
    Dim RstAssessment As Recordset
    Dim StrCrewNo As String
    Dim i As Integer
    
    StrCrewNo = "'" & Parent.CrewNo & "'"
    
    Set RstAssessment = ModDatabase.SQLQuery("Assessment")
    
    With RstAssessment
        If .RecordCount > 0 Then
            Do While Not .EOF
                If IsNull(!CrewNo) Then .Delete
                If !Score = 0 Then .Delete
                .MoveNext
            Loop
        End If
    End With
End Sub

'===============================================================
'Method NextAttemptNo
'Returns next attempt number for given assessment
'---------------------------------------------------------------
Public Function NextAttemptNo(DayNo As Integer, AssessType As String) As Integer
    Dim Assessment As ClsAssessment
    Dim TryCount As Integer
    
    On Error Resume Next
    
    TryCount = 0
    
    For Each Assessment In pAssessments
        If Assessment.Module.DayNo = DayNo And _
                Assessment.AssessType = AssessType Then
            TryCount = TryCount + 1
        End If
    Next
    If TryCount = 6 Then
        NextAttemptNo = 99
    Else
        NextAttemptNo = TryCount
    End If
End Function
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
    Set pAssessments = New Collection
End Sub

Private Sub Class_Terminate()
    Set pAssessments = Nothing
End Sub
'===============================================================








