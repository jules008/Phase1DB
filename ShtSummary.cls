VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


'Public Sub FillOutSummary(DailyLog As ClsDailyLog)
Public Sub FillOutSummary(DailyLog As Object)
    Dim RngArea As Range
    Dim RngModules As Range
    Dim ColNo As Integer
    Dim RowNo As Integer
    Dim RngIndex As Range
    
    Set RngModules = Me.Range("B6:AG6")
    Set RngIndex = RngModules.Find(DailyLog.ModuleNo)
    Set RngArea = Range("A8")
    
    ColNo = RngIndex.Column
    
    'Scores
    RngArea.Offset(0, ColNo - 1) = DailyLog.Score1
    RngArea.Offset(1, ColNo - 1) = DailyLog.Score2
    RngArea.Offset(2, ColNo - 1) = DailyLog.Score3
    RngArea.Offset(3, ColNo - 1) = DailyLog.Score4
    RngArea.Offset(5, ColNo - 1) = DailyLog.OverallGrade
    
End Sub

Public Sub ClearSummary()
    Range("B8:AG13").ClearContents
End Sub

Public Sub FillOutSummaryCandidate(Candidate As Object)
'Public Sub FillOutSummaryCandidate(Candidate As ClsCandidate)
    With Candidate
        Me.Range("Sup_CourseNo") = .CourseNo
        Me.Range("sup_Crewno") = .CrewNo
        Me.Range("sup_division") = .Division
        Me.Range("Sup_name") = .Name
        Me.Range("sup_station") = .StationNo
    End With
End Sub
