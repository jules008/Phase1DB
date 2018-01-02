Attribute VB_Name = "Globals"
Option Explicit
Private Const StrMODULE As String = "Globals"

Public Const DEBUG_MODE As Boolean = False   ' TRUE / FALSE
Public Const OUTPUT_MODE As String = "Debug"  ' "Log" / "Debug"
Public Const ENABLE_PRINT = True           ' TRUE / FALSE
Public Const APP_NAME As String = "Phase 1 Database"
Public Const HANDLED_ERROR As Long = 9999
Public Const USER_CANCEL As Long = 18
Public Const FILE_ERROR_LOG As String = "Error.log"
Public Const VERSION = "2.1"
Public Const VER_DATE = "10/11/16"
Public Const MedGreen As Long = 12379352
Public Const LightGreen As Long = 14610923
Public Const DarkGreen As Long = 2646607
Public Const DarkAmber As Long = 26012
Public Const LightAmber As Long = 10284031
Public Const DarkRed As Long = 393372
Public Const LightRed As Long = 13551615
Public Modules As ClsModules
Public Courses As ClsCourses
Public MailSystem As ClsMailSystem

Type Supervisor
    Username As String
    Forename As String
    Surname As String
    Admin As Boolean
    AccessLvl As Integer
    CrewNo As String
    Role As String
    Rank As String
    email As String
End Type

Enum Role
    Admin = 1
    Trainer = 2
    WCS
    

End Enum
'API Calls
'------------------------------------------------------------------------
Public Declare Sub CopyMemory _
Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)


Public Function Initialise() As Boolean
    Const StrPROCEDURE As String = "Initialise()"
    
    On Error GoTo ErrorHandler
    
    Terminate
    If Not database.DBConnect Then Err.Raise HANDLED_ERROR
    
    Set Modules = New ClsModules
    Set Courses = New ClsCourses
    Set MailSystem = New ClsMailSystem
    Initialise = True

Exit Function

ErrorExit:
    Initialise = False
    
Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

Public Sub Terminate()
    
    On Error Resume Next
    
    database.DBTerminate
    Set Modules = Nothing
    Set Courses = Nothing
    Set MailSystem = Nothing
    
End Sub
