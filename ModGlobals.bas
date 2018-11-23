Attribute VB_Name = "ModGlobals"
'===============================================================
' Module ModGlobals
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 23 Apr 18
'===============================================================
Private Const StrMODULE As String = "ModGlobals"

Option Explicit

' ===============================================================
' Global Constants
' ---------------------------------------------------------------
Public Const PROJECT_FILE_NAME As String = "Phase 1 Candidate Database v2"
Public Const APP_NAME As String = "Phase 1 Database"
Public Const EXPORT_FILE_PATH As String = "\\lincsfire.lincolnshire.gov.uk\folderredir$\Documents\julian.turner\Documents\RDS Project\Phase 1 DB\Library\"
Public Const LIBRARY_FILE_PATH As String = "\\lincsfire.lincolnshire.gov.uk\folderredir$\Documents\julian.turner\Documents\RDS Project\Phase 1 DB\Library\"
Public Const DB_FILE_NAME As String = "Phase 1 Live DB v1,34"
Public Const INI_FILE_PATH As String = "\System Files\"
Public Const INI_FILE_NAME As String = "System.ini"
Public Const STOP_FLAG As Boolean = False
Public Const MAINT_MSG As String = ""
Public Const SEND_ERR_MSG As Boolean = False
Public Const TEST_PREFIX As String = "TEST - "
Public Const FILE_ERROR_LOG As String = "Error.log"
Public Const VERSION = "2.0.0"
Public Const DB_VER = "V1.0.0"
Public Const VER_DATE = "29 Jul 18"

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
Public Const MedGreen As Long = 12379352
Public Const LightGreen As Long = 14610923
Public Const DarkGreen As Long = 2646607
Public Const DarkAmber As Long = 26012
Public Const LightAmber As Long = 10284031
Public Const DarkRed As Long = 393372
Public Const LightRed As Long = 13551615

' ===============================================================
' Enum Declarations
' ---------------------------------------------------------------
Enum Role
    Admin = 1
    Trainer = 2
    WCS
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


