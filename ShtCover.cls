VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShtCover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


'Public Sub PopulateFrontSheet(Candidate As ClsCandidate, Course As ClsCourse)
Public Sub PopulateFrontSheet(Candidate As Object)
    Me.Shapes("name").TextFrame.Characters.Text = "Name: " & Candidate.Name
    Me.Shapes("Number").TextFrame.Characters.Text = "Number: " & Candidate.crewno
    Me.Shapes("Course").TextFrame.Characters.Text = "Course: " & Candidate.Parent.CourseNo
    Me.Shapes("StartDate").TextFrame.Characters.Text = "Start Date: " & Candidate.Parent.StartDate


End Sub
