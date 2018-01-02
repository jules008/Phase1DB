Attribute VB_Name = "APICalls"
Option Explicit

Private Const StrMODULE As String = "ModuleName"
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Const SM_CXSCREEN = 0
Const SM_CYSCREEN = 1

Public Function GetScreenHeight() As Integer
    Const StrPROCEDURE As String = "GetScreenHeight()"

    On Error GoTo ErrorHandler
    Dim x  As Long
    Dim y  As Long
   
    x = GetSystemMetrics(SM_CXSCREEN)
    y = GetSystemMetrics(SM_CYSCREEN)

    GetScreenHeight = y
    
Exit Function

ErrorExit:
    GetScreenHeight = 0

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

