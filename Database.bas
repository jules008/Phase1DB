Attribute VB_Name = "Database"
'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 24 Oct 16
'===============================================================
' Methods
'---------------------------------------------------------------
' SQLQuery - Query database
' DBConnect - Connects to Database
' DBTerminate - Disconnects Database
' SelectDB - Selects Database for use
'===============================================================

Option Explicit
Private Const StrMODULE As String = "Database"

Private DBPath As String
Public DB As DAO.database
Public MyQueryDef As DAO.QueryDef

'===============================================================
' Method SQLQuery
' Query database
'---------------------------------------------------------------
Public Function SQLQuery(SQL As String) As Recordset
    On Error Resume Next
    
    Dim Results As Recordset
    
    If DB Is Nothing Then Initialise
    
    Set Results = DB.OpenRecordset(SQL, dbOpenDynaset)
    Set SQLQuery = Results
    
End Function

'===============================================================
' Method DBConnect
' Connects to Database
'---------------------------------------------------------------
Public Function DBConnect()
    Const StrPROCEDURE As String = "DBConnect()"
    
    On Error GoTo ErrorHandler
    
    DBPath = ShtCourse.Range("DBpath")
    
    If DBPath = "" Then
        MsgBox ("No database has been selected")
        database.SelectDB
    Else
        Set DB = OpenDatabase(DBPath, False, False, "MS Access;pwd=W£8df34JC")
    End If
    
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

'===============================================================
' Method DBTerminate
' Disconnects Database
'---------------------------------------------------------------
Public Function DBTerminate()
    On Error Resume Next
    
    If Not DB Is Nothing Then DB.Close
    Set DB = Nothing
End Function

'===============================================================
' Method SelectDB
' Selects Database for use
'---------------------------------------------------------------
Public Function SelectDB()
    
    Const StrPROCEDURE As String = "SelectDB()"
    
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
        MsgBox "There was no database selected", vbOKOnly, "No Files"
        SelectDB = True
        Exit Function
    End If
  
    'add files to array
    For i = 1 To NoFiles
        FileLoc = DlgOpen.SelectedItems(i)
    Next
    
    'save database location
    ShtCourse.Unprotect
    Range("dbpath") = FileLoc
    Set DlgOpen = Nothing
    ShtCourse.Protect
    SelectDB = True
Exit Function

ErrorExit:
    Set DlgOpen = Nothing
    ShtCourse.Protect
    SelectDB = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

