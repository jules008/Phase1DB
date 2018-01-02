VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtDailyLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Public Sub FillOutForm(DailyLog As Object, Module As Object, Candidate As Object, Course As Object)
'Public Sub FillOutForm(DailyLog As ClsDailyLog, Module As ClsModule, Candidate As ClsCandidate, Course As ClsCourse)
    Dim RngComments As Range
    Dim i As Integer
    
    With DailyLog
        Range("DL_CrewNo") = .CrewNo
        Range("DL_Module") = .ModuleNo
        Range("DL_Assessor") = .Assessor
        Range("DL_OverallScore") = .OverallGrade
        Range("DL_Comments1") = .Comments1
        Range("DL_Size1") = .Comments1
        Range("DL_Score1") = .Score1
        Range("DL_Comments2") = .Comments2
        Range("DL_Size2") = .Comments2
        Range("DL_Score2") = .Score2
        Range("DL_Comments3") = .Comments3
        Range("DL_Size3") = .Comments3
        Range("DL_Score3") = .Score3
        Range("DL_Comments4") = .Comments4
        Range("DL_Size4") = .Comments4
        Range("DL_Comments5") = .CommentsMisc
        Range("DL_Size5") = .CommentsMisc
        Range("DL_Date") = .DLDate
    End With
    
    With Course
        Range("DL_CourseNo") = .CourseNo
        Range("DL_StartDate") = .StartDate
    End With
        
    With Candidate
        Range("DL_Name") = .Name
        Range("DL_Station") = .StationNo
        Range("DL_Division") = .Division
    End With
    
    With Module
        Range("DL_Module") = .ModuleNo & " - " & .ModuleDesc
        Range("DL_DayNo") = .DayNo
    End With
    
    For i = 1 To 5
        ShtDailyLog.Activate
        Set RngComments = Range("DL_size" & i)
        With RngComments
            .Rows.AutoFit
            If .RowHeight < 40 Then .RowHeight = 40
        End With
    Next
End Sub
