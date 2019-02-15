Attribute VB_Name = "ModDatabase"
'===============================================================
' Module ModDatabase
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 19 Apr 18
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModDatabase"

Public DB As DAO.Database
Public MyQueryDef As DAO.QueryDef

' ===============================================================
' SQLQuery
' Queries database with given SQL script
' ---------------------------------------------------------------
Public Function SQLQuery(SQL As String) As Recordset
    Dim RstResults As Recordset
    
    Const StrPROCEDURE As String = "SQLQuery()"

    On Error GoTo ErrorHandler
      
Restart:
    Application.StatusBar = ""

    If DB Is Nothing Then
        Err.Raise NO_DATABASE_FOUND, Description:="Unable to connect to database"
    Else
        If FaultCount1008 > 0 Then FaultCount1008 = 0
    
        Set RstResults = DB.OpenRecordset(SQL, dbOpenDynaset)
        Set SQLQuery = RstResults
    End If
    
    Set RstResults = Nothing
    
Exit Function

ErrorExit:

    Set RstResults = Nothing

    Set SQLQuery = Nothing
    Terminate

Exit Function

ErrorHandler:
    
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        If CustomErrorHandler(Err.Number) Then
            If Not Initialise Then Err.Raise HANDLED_ERROR
            Resume Restart
        Else
            Err.Raise HANDLED_ERROR
        End If
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' DBConnect
' Provides path to database
' ---------------------------------------------------------------
Public Function DBConnect() As Boolean
    Const StrPROCEDURE As String = "DBConnect()"

    On Error GoTo ErrorHandler

    Set DB = OpenDatabase(DB_PATH, False, False, "MS Access;pwd=W£8df34JC")
  
    DBConnect = True

Exit Function

ErrorExit:

    DBConnect = False

Exit Function

ErrorHandler:

If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function


' ===============================================================
' DBTerminate
' Disconnects and closes down DB connection
' ---------------------------------------------------------------
Public Function DBTerminate() As Boolean
    Const StrPROCEDURE As String = "DBTerminate()"

    On Error GoTo ErrorHandler

    If Not DB Is Nothing Then DB.Close
    Set DB = Nothing

    DBTerminate = True

Exit Function

ErrorExit:

    DBTerminate = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' SelectDB
' Selects DB to connect to
' ---------------------------------------------------------------
Public Function SelectDB() As Boolean
    Const StrPROCEDURE As String = "SelectDB()"

    On Error GoTo ErrorHandler
    Dim DlgOpen As FileDialog
    Dim FileLoc As String
    Dim NoFiles As Integer
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    'open files
    Set DlgOpen = Application.FileDialog(msoFileDialogOpen)
    
     With DlgOpen
        .Filters.Clear
        .Filters.Add "Access Files (*.accdb)", "*.accdb"
        .AllowMultiSelect = False
        .Title = "Connect to Database"
        .Show
    End With
    
    'get no files selected
    NoFiles = DlgOpen.SelectedItems.Count
    
    'exit if no files selected
    If NoFiles = 0 Then
        MsgBox "There was no database selected", vbOKOnly + vbExclamation, "No Files"
        SelectDB = True
        Exit Function
    End If
  
    'add files to array
    For i = 1 To NoFiles
        FileLoc = DlgOpen.SelectedItems(i)
    Next
    
    DB_PATH = FileLoc
    
    Set DlgOpen = Nothing

    SelectDB = True

Exit Function

ErrorExit:

    Set DlgOpen = Nothing
    SelectDB = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' UpdateDBScript
' Script to update DB
' ---------------------------------------------------------------
Public Sub UpdateDBScript()
    Dim TableDef As DAO.TableDef
    Dim Ind As DAO.Index
    Dim RstTable As Recordset
    Dim i As Integer
    Dim Binary As String
    
    Dim Fld As DAO.Field
    
    SYS_PATH = ThisWorkbook.Path & INI_FILE_PATH
    ModStartUp.ReadINIFile
    DBConnect
    
    Set RstTable = SQLQuery("TblDBVersion")
    
    'check preceding DB Version
    If RstTable.Fields(0) <> "V1.1.0" Then
        MsgBox "Database needs to be upgraded to V1.1.0 to continue", vbOKOnly + vbCritical
        Exit Sub
    End If
    
    DB.Execute " SELECT * INTO ModuleBAK FROM module"
    
    DB.Execute "UPDATE Module SET Module = 'EDBA' WHERE DayNo = 22"
    DB.Execute "UPDATE Module SET Module = 'BA consolidation' WHERE DayNo = 23"
    DB.Execute "UPDATE Module SET Module = 'PPV' WHERE DayNo = 24"
    DB.Execute "UPDATE Module SET Module = 'BA Assessments' WHERE DayNo = 25"
    
    DB.Execute "UPDATE Module SET Module = 'Consolidation BCS' WHERE DayNo = 31"
    DB.Execute "UPDATE Module SET Module = 'Consolidation BCS' WHERE DayNo = 32"
    DB.Execute "UPDATE Module SET Module = 'Tech Rope Rescue' WHERE DayNo = 33"
    
    DB.Execute "UPDATE Module SET Module = 'CFS/CFP/Cobra' WHERE DayNo = 43"
    DB.Execute "UPDATE Module SET Module = 'CFS/CFP/Cobra' WHERE DayNo = 44"
    DB.Execute "UPDATE Module SET Module = 'Phase 2/ Safeguarding/ MDT' WHERE DayNo = 45"
    
    DB.Execute "UPDATE Module SET Module = 'Final Assessments' WHERE DayNo = 47"
    DB.Execute "UPDATE Module SET Module = 'Final Assessments and Debriefs' WHERE DayNo = 48"
    DB.Execute "INSERT INTO Module VALUES (49, 49, 49,'Course Review', 0, NULL)"
    DB.Execute "INSERT INTO Module VALUES (50, 50, 50,'WT Pass out', 0, NULL)"
        
    'update DB Version
    Set RstTable = SQLQuery("TblDBVersion")
    
    With RstTable
        .Edit
        .Fields(0) = "V1.1.1"
        .Update
    End With
        
    MsgBox "Database successfully updated", vbOKOnly + vbInformation
    
    Set RstTable = Nothing
    Set TableDef = Nothing
    Set Fld = Nothing
    
End Sub
              
' ===============================================================
' UpdateDBScriptUndo
' Script to update DB
' ---------------------------------------------------------------
Public Sub UpdateDBScriptUndo()
    Dim TableDef As DAO.TableDef
    Dim Ind As DAO.Index
    Dim RstTable As Recordset
    Dim i As Integer
        
    Dim Fld As DAO.Field
        
    SYS_PATH = ThisWorkbook.Path & INI_FILE_PATH
    ModStartUp.ReadINIFile
    DBConnect
    
    Set RstTable = SQLQuery("TblDBVersion")

    If RstTable.Fields(0) <> "V1.1.1" Then
        MsgBox "Database needs to be upgraded to V1.1.1 to continue", vbOKOnly + vbCritical
        Exit Sub
    End If

    DB.Execute "DROP TABLE Module"
    DB.Execute "SELECT * INTO Module FROM moduleBAK"
    DB.Execute "DROP TABLE ModuleBAK"
    
    Set RstTable = SQLQuery("TblDBVersion")
    
    With RstTable
        .Edit
        .Fields(0) = "V1.1.0"
        .Update
    End With
    
    MsgBox "Database reset successfully", vbOKOnly + vbInformation
    
    Set RstTable = Nothing
    Set TableDef = Nothing
    Set Fld = Nothing

End Sub

' ===============================================================
' GetDBVer
' Returns the version of the DB
' ---------------------------------------------------------------
Public Function GetDBVer() As String
    Dim DBVer As Recordset
    
    Const StrPROCEDURE As String = "GetDBVer()"

    On Error GoTo ErrorHandler

    Set DBVer = SQLQuery("TblDBVersion")

    GetDBVer = DBVer.Fields(0)

    Debug.Print DBVer.Fields(0)
    Set DBVer = Nothing
Exit Function

ErrorExit:

    GetDBVer = ""
    
    Set DBVer = Nothing

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' UpdateSysMsg
' Updates the system message and resets read flags
' ---------------------------------------------------------------
Public Sub UpdateSysMsg()
    Dim RstMessage As Recordset
    
    Set RstMessage = SQLQuery("TblMessage")
    
    With RstMessage
        If .RecordCount = 0 Then
            .AddNew
        Else
            .Edit
        End If
        
        .Fields("SystemMessage") = "Version " & VERSION & " - What's New" _
                    & Chr(13) & "(See Release Notes on Support tab for further information)" _
                    & Chr(13) & "" _
                    & Chr(13) & " - Bug Fix - Hidden Assets" _
                    & Chr(13) & ""
        
        .Fields("ReleaseNotes") = "Software Version: " & VERSION _
                    & Chr(13) & "Database Version: " & DB_VER _
                    & Chr(13) & "Date: " & VER_DATE _
                    & Chr(13) & "" _
                    & Chr(13) & "- Bug Fix - Hidden Assets - Had ANOTHER go at fixing the hidden assets bug.  Hopefully fixed now" _
                    & Chr(13) & ""
        .Update
    End With
    
    'reset read flags
    DB.Execute "UPDATE TblPerson SET MessageRead = False WHERE MessageRead = True"
    
    Set RstMessage = Nothing

End Sub

' ===============================================================
' ShowUsers
' Show users logged onto system
' ---------------------------------------------------------------
Public Sub ShowUsers()
    Dim RstUsers As Recordset
    
    Set RstUsers = SQLQuery("TblUsers")
    
    With RstUsers
        Debug.Print
        Do While Not .EOF
            Debug.Print "User: " & .Fields(0) & " - Logged on: " & .Fields(1)
            .MoveNext
        Loop
    End With
    
    Set RstUsers = Nothing
End Sub
