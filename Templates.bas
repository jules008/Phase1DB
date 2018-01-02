Attribute VB_Name = "Templates"
'Option Explicit
'
'module label constant----------------------------------
'Private Const StrMODULE As String = "ModuleName"
'-------------------------------------------------------
'
'Procedure label constant----------------------------------
'    Const StrPROCEDURE As String = "ProcName()"
'
'-------------------------------------------------------
'
'On Error Goto ----------------------------------
'    On Error GoTo ErrorHandler
'
'-------------------------------------------------------
'
'Function Error Handling----------------------------------
'
'    FunctionName = True
'
'Exit Function
'
'ErrorExit:
'    CleanUpCode
'    FunctionName = False
'
'Exit Function
'
'ErrorHandler:
'    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
'        Stop
'        Resume
'    Else
'        Resume ErrorExit
'    End If
'End Function
'-------------------------------------------------------
'
'Entry point Error Handling----------------------------------
'Exit Sub
'
'ErrorExit:
'    CleanUpCode
'
'Exit Sub
'
'ErrorHandler:
'    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
'        Stop
'        Resume
'    Else
'        Resume ErrorExit
'    End If
'End Sub
'
'-------------------------------------------------------
'
'************************************************************************
'        Err.Raise 12
'************************************************************************
'
