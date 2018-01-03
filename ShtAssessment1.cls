VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtAssessment1"
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
    Dim OffsetX As Integer
    Dim OffsetY As Integer
    
    StrCrewNo = "'" & Candidate.crewno & "'"
    
    Set RngAssessIndex = Me.Range("A7")
    Set RstAssessment = Database.SQLQuery("SELECT * FROM Assessment WHERE crewno = " & StrCrewNo & " AND deleted IS NULL")
    
    With RstAssessment
        Do While Not .EOF
            Select Case !DayNo
                
                Case Is = 3
                    OffsetY = 1
                Case Is = 9
                    OffsetY = 2
                Case Is = 11
                    
                    Select Case !AssessType
                        Case Is = "Written"
                            OffsetY = 3
                        Case Is = "Practical"
                            OffsetY = 4
                    End Select
                    
                Case Is = 17
                
                    Select Case !AssessType
                        Case Is = "Written"
                            OffsetY = 5
                        Case Is = "Practical"
                            OffsetY = 6
                    End Select
                    
                Case Is = 20
                
                    Select Case !AssessType
                        Case Is = "Written"
                            OffsetY = 7
                        Case Is = "Practical"
                            OffsetY = 8
                        Case Is = "BA Board"
                            OffsetY = 9
                        Case Is = "FB Oral"
                            OffsetY = 10
                    End Select
                    
                Case Is = 27
                    
                    Select Case !AssessType
                        Case Is = "Written"
                            OffsetY = 11
                        Case Is = "Practical"
                            OffsetY = 12
                    End Select
                
                Case Is = 28
                    
                    Select Case !AssessType
                        Case Is = "Written"
                            OffsetY = 13
                        Case Is = "Oral"
                            OffsetY = 14
                        Case Is = "Knots"
                            OffsetY = 15
                        Case Is = "Practical"
                            OffsetY = 16
                    End Select
                    
                Case Is = 29
                    OffsetY = 17
            End Select
            
            OffsetX = !Attempt
            
            RngAssessIndex.Offset(OffsetX, OffsetY) = !Score
            .MoveNext
        Loop
    End With
    
End Sub
