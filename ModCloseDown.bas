Attribute VB_Name = "ModCloseDown"
'===============================================================
' Module ModCloseDown
'===============================================================
' v1.0.0 - Initial Version
' v1.1.0 - WT2019 Version
'---------------------------------------------------------------
' Date - 30 Dec 18
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModCloseDown"

' ===============================================================
' Terminate
' Closedown processing
' ---------------------------------------------------------------
Public Function Terminate() As Boolean
    Const StrPROCEDURE As String = "Terminate()"

    On Error GoTo ErrorHandler

    ModDatabase.DBTerminate
    
    Application.DisplayFullScreen = False
    
    Terminate = True

Exit Function

ErrorExit:

    ModDatabase.DBTerminate

    
    Terminate = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

