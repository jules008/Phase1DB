Attribute VB_Name = "ModGlobals"
'===============================================================
' Module ModGlobals
'===============================================================
' v1.0.0 - Initial Version
' v1,0 - WT2019 Version
' v1,1 - Updated colours
'---------------------------------------------------------------
' Date - 18 Jan 19
'===============================================================
Private Const StrMODULE As String = "ModGlobals"

Option Explicit

' ===============================================================
' Global Constants
' ---------------------------------------------------------------
Public Const PROJECT_FILE_NAME As String = "Phase 1 Candidate Database v2"
Public Const APP_NAME As String = "Phase 1 Database"
Public Const EXPORT_FILE_PATH As String = "E:\Development Areas\Phase 1DB\Library\"
Public Const LIBRARY_FILE_PATH As String = "E:\Development Areas\Phase 1DB\Library\"
Public Const DB_FILE_NAME As String = "Phase 1 Live DB v1,34"
Public Const INI_FILE_PATH As String = "\System Files\"
Public Const INI_FILE_NAME As String = "System.ini"
Public Const STOP_FLAG As Boolean = False
Public Const MAINT_MSG As String = ""
Public Const SEND_ERR_MSG As Boolean = False
Public Const TEST_PREFIX As String = "TEST - "
Public Const FILE_ERROR_LOG As String = "Error.log"
Public Const VERSION = "2.0.0"
Public Const DB_VER = "V1.1.1"
Public Const VER_DATE = "15 Feb 19"

' ===============================================================
' Error Constants
' ---------------------------------------------------------------
Public Const HANDLED_ERROR As Long = 9999
Public Const UNKNOWN_USER As Long = 1000
Public Const SYSTEM_RESTART As Long = 1001
Public Const NO_DATABASE_FOUND As Long = 1002
Public Const ACCESS_DENIED As Long = 1003
Public Const NO_INI_FILE As Long = 1004
Public Const DB_WRONG_VER As Long = 1005
Public Const GENERIC_ERROR As Long = 1006
Public Const USER_CANCEL As Long = 18

' ===============================================================
' Error Variables
' ---------------------------------------------------------------
Public FaultCount1002 As Integer
Public FaultCount1008 As Integer

' ===============================================================
' Global Variables
' ---------------------------------------------------------------
Public DEBUG_MODE As Boolean
Public SEND_EMAILS As Boolean
Public ENABLE_PRINT As Boolean
Public DB_PATH As String
Public DEV_MODE As Boolean
Public SYS_PATH As String

' ===============================================================
' Global Class Declarations
' ---------------------------------------------------------------
Public Modules As ClsModules
Public Courses As ClsCourses
Public MailSystem As ClsMailSystem

' ---------------------------------------------------------------
' Others
' ---------------------------------------------------------------

' ===============================================================
' Colours
' ---------------------------------------------------------------
Public Const COLOUR_1 = 12379352
Public Const COLOUR_2 = 6737151
Public Const COLOUR_3 = 2646607
Public Const COLOUR_4 = 4626167
Public Const Colour_5 = 9305182
Public Const COLOUR_6 = 393372
Public Const COLOUR_7 = 13551615
Public Const Colour_8 = 9617978
Public Const COLOUR_9 = 4626167

' ===============================================================
' Enum Declarations
' ---------------------------------------------------------------
Enum Role
    Admin = 1
    Trainer = 2
    WCS
End Enum

Enum EnumFormValidation
    FormOK = 2
    ValidationError = 1
    FunctionalError = 0
End Enum

' ===============================================================
' Type Declarations
' ---------------------------------------------------------------
Type Supervisor
    UserName As String
    Forename As String
    Surname As String
    Admin As Boolean
    AccessLvl As Integer
    CrewNo As String
    Role As String
    Rank As String
    email As String
End Type


