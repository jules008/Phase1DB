VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Const StrMODULE As String = "WorkBook"

Private Sub Workbook_Deactivate()
    On Error Resume Next
    Globals.Terminate
End Sub

Private Sub Workbook_Open()
    Const StrPROCEDURE As String = "Workbook_Open()"
    
    On Error GoTo ErrorHandler
    
    If Not Initialise Then Err.Raise HANDLED_ERROR
    If Not ShtCourse.WorksheetActivate Then Err.Raise HANDLED_ERROR
Exit Sub

ErrorExit:
    Globals.Terminate
    
Exit Sub
    
ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub
