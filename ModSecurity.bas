Attribute VB_Name = "ModSecurity"
'===============================================================
' Module ModSecurity
'===============================================================
' v1.0.0 - Initial Version
'---------------------------------------------------------------
' Date - 17 Jan 17
'===============================================================

Option Explicit

Private Const StrMODULE As String = "ModSecurity"

' ===============================================================
' CourseAccessCheck
' Returns whether person is on access list
' ---------------------------------------------------------------
Public Function CourseAccessCheck(CourseNo As String) As Boolean
    Dim StrUserName As String
    Dim StrCourseNo As String
    Dim RstUserList As Recordset
    
    Const StrPROCEDURE As String = "CourseAccessCheck()"

    On Error GoTo ErrorHandler

    StrUserName = "'" & Application.UserName & "'"
    StrCourseNo = "'" & CourseNo & "'"
    
    Set RstUserList = ModDatabase.SQLQuery("SELECT * FROM useraccess WHERE " & _
                            " CourseNo = " & StrCourseNo & _
                            " AND UserName = " & StrUserName)
    
    If RstUserList.RecordCount = 0 Then
        CourseAccessCheck = False
    Else
        CourseAccessCheck = True
    End If
    
    Set RstUserList = Nothing

    CourseAccessCheck = True

Exit Function

ErrorExit:

    Set RstUserList = Nothing
    CourseAccessCheck = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' RemoveUser
' Removes user from access list for course
' ---------------------------------------------------------------
Public Function RemoveUser(UserName As String, Optional CourseNo As String) As Boolean
    Dim StrUserName As String
    Dim StrCourseNo As String
    Dim RstUserList As Recordset
    Dim RstCourseUserLst As Recordset
    
    Const StrPROCEDURE As String = "RemoveUser()"

    On Error GoTo ErrorHandler

    StrUserName = "'" & UserName & "'"
    
    If ModDatabase.DB Is Nothing Then
        If Not Initialise Then Err.Raise HANDLED_ERROR
    End If
    
    'if courseno is not included, then delete the user from both the user list tables
    'and the course access table
    If CourseNo = "" Then
        Set RstUserList = ModDatabase.SQLQuery("SELECT * FROM UserList WHERE " & _
                                                "UserName = " & StrUserName)
        
        Set RstCourseUserLst = ModDatabase.SQLQuery("SELECT * FROM useraccess WHERE " & _
                                                "UserName = " & StrUserName)
    Else
    
        'if course no is included, then only delete the user from the course access table
        StrCourseNo = "'" & CourseNo & "'"
        
        Set RstUserList = ModDatabase.SQLQuery("SELECT * FROM useraccess WHERE " & _
                                " CourseNo = " & StrCourseNo & _
                                " AND UserName = " & StrUserName)
        
    End If
    
    With RstCourseUserLst
        If Not RstCourseUserLst Is Nothing Then
            If .RecordCount > 0 Then
                Do While Not .EOF
                    .Delete
                    .MoveNext
                Loop
            End If
        End If
    End With
        
    With RstUserList
        If .RecordCount > 0 Then
            Do While Not .EOF
                .Delete
                .MoveNext
            Loop
        End If
    End With
    
    
    Set RstUserList = Nothing
    Set RstCourseUserLst = Nothing

    RemoveUser = True
    
Exit Function

ErrorExit:

    Set RstUserList = Nothing
    Set RstCourseUserLst = Nothing
    RemoveUser = False

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
' IsAdmin
' Checks whether person is an admin
' ---------------------------------------------------------------
Public Function IsAdmin() As Boolean
    Const StrPROCEDURE As String = "IsAdmin()"

    Dim RstUserList As Recordset
    Dim StrUserName As String
    
    On Error GoTo ErrorHandler

    
    If ModDatabase.DB Is Nothing Then
        If Not Initialise Then Err.Raise HANDLED_ERROR
    End If
    
    StrUserName = "'" & Application.UserName & "'"
    
    Set RstUserList = ModDatabase.SQLQuery("SELECT * FROM userlist WHERE " & _
                            " UserName = " & StrUserName _
                            & "AND admin = TRUE")
    
    With RstUserList
        If .RecordCount > 0 Then
            IsAdmin = True
        Else
            IsAdmin = False
        End If
    End With
    
    Set RstUserList = Nothing

    IsAdmin = True

Exit Function

ErrorExit:

    Set RstUserList = Nothing
    IsAdmin = False

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function


' ===============================================================
' GetAccessList
' Returns access list for course
' ---------------------------------------------------------------
Public Function GetAccessList() As Recordset
    Const StrPROCEDURE As String = "GetAccessList()"
    
    Dim StrUserName As String
    Dim StrCourseNo As String
    Dim CourseNo As String
    Dim RstUserList As Recordset
    Dim RstCourseUserLst As Recordset

    On Error GoTo ErrorHandler
    
    If CourseNo = "" Then
        Set RstUserList = ModDatabase.SQLQuery("userlist")
    Else
        StrCourseNo = "'" & CourseNo & "'"
        
        Set RstUserList = ModDatabase.SQLQuery("SELECT * FROM useraccess WHERE " & _
                                " CourseNo = " & StrCourseNo)
    End If
    
    If RstUserList.RecordCount <> 0 Then
        Set GetAccessList = RstUserList
    End If
    
    Set RstUserList = Nothing

Exit Function

ErrorExit:

    Set RstUserList = Nothing
    Set GetAccessList = Nothing

Exit Function

ErrorHandler:   If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

' ===============================================================
' AddUpdateUser
' Adds or updates user
' ---------------------------------------------------------------
Public Function AddUpdateUser(User As Supervisor, Optional CourseNo As String) As Boolean
    Const StrPROCEDURE As String = "AddUpdateUser()"

    Dim RstUserList As Recordset
    Dim StrUserName As String
    Dim StrCourseNo As String
    
    On Error GoTo ErrorHandler

    If User.UserName = "" Then
        User.UserName = User.Forename & " " & User.Surname
    End If
    
    StrUserName = "'" & User.UserName & "'"
    
    If CourseNo <> "" Then
        
        StrCourseNo = "'" & CourseNo & "'"
        Set RstUserList = ModDatabase.SQLQuery("SELECT * FROM useraccess WHERE " & _
                                            "UserName = " & StrUserName & _
                                            " AND courseno = " & StrCourseNo)
        With RstUserList
            If .RecordCount = 0 Then
                .AddNew
                !UserName = User.UserName
                !CourseNo = CourseNo
                .Update
            End If
        End With
    Else
    
        Set RstUserList = ModDatabase.SQLQuery("SELECT * FROM userlist WHERE " & _
                                            "UserName = " & StrUserName)
        With RstUserList
            If .RecordCount = 0 Then
                .AddNew
            Else
                .Edit
            End If
            
            !CrewNo = User.CrewNo
            !Rank = User.Rank
            !Admin = User.Admin
            !Forename = User.Forename
            !Surname = User.Surname
            !AccessLvl = User.AccessLvl
            !Role = User.Role
            !email = User.email
            
            .Update
        
        End With
    End If
    Set RstUserList = Nothing
    AddUpdateUser = True
Exit Function

ErrorExit:
    Set RstUserList = Nothing
    AddUpdateUser = False

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
' GetUserDetails
' Returns user details in Recordset
' ---------------------------------------------------------------
Public Function GetUserDetails(UserName As String) As Recordset
    Dim StrUserName As String
    Dim RstUserList As Recordset
    
    On Error Resume Next
    
    StrUserName = "'" & UserName & "'"
    
    Set RstUserList = ModDatabase.SQLQuery("SELECT * FROM userlist WHERE " & _
                            " UserName = " & StrUserName)
                            
    If RstUserList.RecordCount <> 0 Then
        Set GetUserDetails = RstUserList
    End If
    
    Set RstUserList = Nothing
    
End Function
