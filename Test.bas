Attribute VB_Name = "Test"
Option Explicit



Public Sub test()
    Dim Target As Range
    Set Target = ShtAssess.EnterAssessment("4323", 3, "Written %", 1)
    Target.Select
End Sub
