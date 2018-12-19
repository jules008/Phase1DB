Attribute VB_Name = "ModCloseDown"
'===============================================================
' Module ModCloseDown
'===============================================================
' v1.0.0 - Initial Version
' v0,1 - WT2018 Version
'---------------------------------------------------------------
' Date - 19 Dec 18
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

