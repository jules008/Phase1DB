VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtAssessment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


'Public Sub PopulateAssessmentSummary(Candidate As ClsCandidate)
Public Sub PopulateAssessmentSummary(Candidate As Object)
    Dim RstAssessment As Recordset
    Dim StrCrewNo As String
    Dim RngAssessIndex As Range
    
    StrCrewNo = "'" & Candidate.CrewNo & "'"
    
    Set RngAssessIndex = Me.Range("A16")
    Set RstAssessment = database.SQLQuery("SELECT * FROM Assessment WHERE crewno = " & StrCrewNo & " AND deleted IS NULL")
    
    With RstAssessment
        If .RecordCount > 0 Then
            RngAssessIndex.CopyFromRecordset RstAssessment
        
        
        End If
    End With
    
End Sub
