Attribute VB_Name = "Library"
Option Explicit

Public Function ConvertHoursIntoDecimal(TimeIn As Date)
    On Error Resume Next
    
    Dim TB, Result As Single
    
    TB = Split(TimeIn, ":")
    ConvertHoursIntoDecimal = TB(0) + ((TB(1) * 100) / 60) / 100
    
End Function
Function EndOfMonth(InputDate As Date) As Variant
    On Error Resume Next
    
    EndOfMonth = Day(DateSerial(Year(InputDate), Month(InputDate) + 1, 0))
End Function
Public Sub PerfSettingsOn()
    On Error Resume Next
    
    'turn off some Excel functionality so your code runs faster
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

End Sub

Public Sub PerfSettingsOff()
    On Error Resume Next
        
    'turn off some Excel functionality so your code runs faster
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
'====================Spell check=====================================
Public Sub SpellCheck(ByRef Cntrls As Collection)
    On Error Resume Next
    
    Dim RngSpell As Range
    Dim Cntrl As Control
    
    Set RngSpell = ShtWorking.Range("A1")
    
    For Each Cntrl In Cntrls
        
        If Left(Cntrl.Name, 3) = "Txt" Then
            Debug.Print Cntrl.Name
            RngSpell = Cntrl
            RngSpell.CheckSpelling
            Cntrl = RngSpell
        End If
    Next
    
End Sub

'=========debug print contents of recordset

Public Sub RecordsetPrint(RST As Recordset)
    On Error Resume Next
    
    Dim DBString As String
    Dim RSTField As Field
    Dim i As Integer

    ReDim AyFields(RST.Fields.Count)
    
    Do Until RST.EOF
        For i = 0 To RST.Fields.Count - 1
             DBString = DBString & RST.Fields(i).Value & ", "
        Next
        RST.MoveNext
        Debug.Print DBString
        DBString = ""
    Loop

End Sub

Public Sub PrintPDF(WSheet As Worksheet, PathAndFileName As String)
    On Error Resume Next
    
    Dim strPath As String
    Dim myFile As Variant
    Dim strFile As String
    On Error GoTo errHandler
    
    strFile = PathAndFileName & ".pdf"
    
    WSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=strFile, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    
exitHandler:
        Exit Sub
errHandler:
        MsgBox "Could not create PDF file"
        Resume exitHandler

End Sub
