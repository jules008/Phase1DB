Attribute VB_Name = "Security"
Option Explicit
Private Const StrMODULE As String = "Security"

Public Function CourseAccessCheck(CourseNo As String) As Boolean
    On Error Resume Next
    
    Dim StrUsername As String
    Dim StrCourseNo As String
    Dim RstUserList As Recordset
    
    StrUsername = Application.Username
    StrCourseNo = "'" & CourseNo & "'"
    
    Set RstUserList = database.SQLQuery("SELECT * FROM useraccess WHERE " & _
                            " CourseNo = " & StrCourseNo & _
                            " AND username = " & StrUsername)
    
    If RstUserList.RecordCount = 0 Then
        CourseAccessCheck = False
    Else
        CourseAccessCheck = True
    End If
    
    Set RstUserList = Nothing
    
End Function


Public Function RemoveUser(Username As String, Optional CourseNo As String) As Boolean
    Const StrPROCEDURE As String = "RemoveUser()"

    Dim StrUsername As String
    Dim StrCourseNo As String
    Dim RstUserList As Recordset
    Dim RstCourseUserLst As Recordset
    
    On Error GoTo ErrorHandler

    StrUsername = "'" & Username & "'"
    
    If database.DB Is Nothing Then
        If Not Initialise Then Err.Raise HANDLED_ERROR
    End If
    
    'if courseno is not included, then delete the user from both the user list tables
    'and the course access table
    If CourseNo = "" Then
        Set RstUserList = database.SQLQuery("SELECT * FROM UserList WHERE " & _
                                                "Username = " & StrUsername)
        
        Set RstCourseUserLst = database.SQLQuery("SELECT * FROM useraccess WHERE " & _
                                                "Username = " & StrUsername)
    Else
    
        'if course no is included, then only delete the user from the course access table
        StrCourseNo = "'" & CourseNo & "'"
        
        Set RstUserList = database.SQLQuery("SELECT * FROM useraccess WHERE " & _
                                " CourseNo = " & StrCourseNo & _
                                " AND username = " & StrUsername)
        
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
Public Function IsAdmin() As Boolean
    On Error Resume Next
    
    Dim RstUserList As Recordset
    Dim StrUsername As String
    
    If database.DB Is Nothing Then
        If Not Initialise Then Err.Raise HANDLED_ERROR
    End If
    
    StrUsername = Application.Username
    
    Set RstUserList = database.SQLQuery("SELECT * FROM userlist WHERE " & _
                            " username = " & StrUsername _
                            & " AND admin = TRUE")
    
    With RstUserList
        If .RecordCount > 0 Then
            IsAdmin = True
        Else
            IsAdmin = False
        End If
    End With
    
    Set RstUserList = Nothing

End Function

Public Function GetAccessList(Optional CourseNo As String) As Recordset
    On Error Resume Next
    
    Dim StrCourseNo As String
    Dim RstUserList As Recordset
    
    If CourseNo = "" Then
        Set RstUserList = database.SQLQuery("userlist")
    Else
        StrCourseNo = "'" & CourseNo & "'"
        
        Set RstUserList = database.SQLQuery("SELECT * FROM useraccess WHERE " & _
                                " CourseNo = " & StrCourseNo)
    End If
    
    If RstUserList.RecordCount <> 0 Then
        Set GetAccessList = RstUserList
    End If
    
    Set RstUserList = Nothing
    
End Function

Public Function GetUserDetails(Username As String) As Recordset
    Dim StrUsername As String
    Dim RstUserList As Recordset
    
    On Error Resume Next
    
    StrUsername = "'" & Username & "'"
    
    Set RstUserList = database.SQLQuery("SELECT * FROM userlist WHERE " & _
                            " UserName = " & StrUsername)
                            
    If RstUserList.RecordCount <> 0 Then
        Set GetUserDetails = RstUserList
    End If
    
    Set RstUserList = Nothing
    
End Function

Public Function AddUpdateUser(User As Supervisor, Optional CourseNo As String) As Boolean
    Const StrPROCEDURE As String = "AddUpdateUser()"

    Dim RstUserList As Recordset
    Dim StrUsername As String
    Dim StrCourseNo As String
    
    On Error GoTo ErrorHandler

    If User.Username = "" Then
        User.Username = User.Forename & " " & User.Surname
    End If
    
    StrUsername = "'" & User.Username & "'"
    
    If CourseNo <> "" Then
        
        StrCourseNo = "'" & CourseNo & "'"
        Set RstUserList = database.SQLQuery("SELECT * FROM useraccess WHERE " & _
                                            "username = " & StrUsername & _
                                            " AND courseno = " & StrCourseNo)
        With RstUserList
            If .RecordCount = 0 Then
                .AddNew
                !Username = User.Username
                !CourseNo = CourseNo
                .Update
            End If
        End With
    Else
    
        Set RstUserList = database.SQLQuery("SELECT * FROM userlist WHERE " & _
                                            "username = " & StrUsername)
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

