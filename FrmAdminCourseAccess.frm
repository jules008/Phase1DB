VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmAdminCourseAccess 
   Caption         =   "Action Plan"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6105
   OleObjectBlob   =   "FrmAdminCourseAccess.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmAdminCourseAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 09 Nov 16
'===============================================================
Option Explicit

Private Const StrMODULE As String = "FrmAdminCourseAccess"

Private Course As ClsCourse
Private ActiveUserName As String

Public Function ShowForm() As Boolean
    
   Const StrPROCEDURE As String = "ShowForm()"
   
   On Error GoTo ErrorHandler
   
    ResetForm
    If Not UserformActivate Then Err.Raise HANDLED_ERROR
    Show

    ShowForm = True
Exit Function

ErrorExit:
    
    ShowForm = False
    Terminate
Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

Private Sub BtnAdd_Click()
    Const StrPROCEDURE As String = "BtnAdd_Click()"

    Dim UserName As String
    Dim User As Supervisor
    
    On Error GoTo ErrorHandler

    User.UserName = CmoUsers
    
    If ValidateData Then
        If Not Security.AddUpdateUser(User, CmoCourseNo) Then Err.Raise HANDLED_ERROR
        If Not RefreshUserList Then Err.Raise HANDLED_ERROR
    End If
Exit Sub

ErrorExit:
    Terminate

Exit Sub

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub
Private Sub BtnClose_Click()
    On Error Resume Next
    
    Me.Hide
End Sub


Private Sub BtnNew_Click()
    On Error Resume Next
    
    ResetForm
End Sub


Private Sub BtnRemove_Click()
    Const StrPROCEDURE As String = "BtnRemove_Click()"
    
    Dim Response As Integer
    Dim UserName As String
    Dim SelUser As Integer
    
    On Error GoTo ErrorHandler
        
    If CmoCourseNo <> "" Then
                
        SelUser = LstAccessList.ListIndex
        If SelUser <> -1 Then
            
            UserName = LstAccessList.List(SelUser, 0)
        
            Response = MsgBox("Are you sure you want to remove access for " _
                        & UserName & " from Course " & CmoCourseNo & "?", 36)
        
            If Response = 6 Then
            
                If Not Security.RemoveUser(UserName, CmoCourseNo) Then Err.Raise HANDLED_ERROR
            
            End If
            
            RefreshUserList
        Else
            MsgBox "Please select a user"
        End If
    End If
Exit Sub

ErrorExit:
    Terminate

Exit Sub

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If

End Sub

Private Sub CmoCourseNo_Change()
    On Error Resume Next
    
    ResetForm
    If Not RefreshUserList Then Err.Raise HANDLED_ERROR
End Sub

Private Sub ResetForm()
    On Error Resume Next
    LstAccessList.Clear
    
End Sub

Public Function RefreshCourses() As Boolean
    Const StrPROCEDURE As String = "RefreshCourses()"
    
    Dim i As Integer
    Dim LocCourse As ClsCourse
    
    On Error GoTo ErrorHandler
    
    With CmoCourseNo
        
        .Clear
        
        For i = 1 To Courses.Count
        
            Set LocCourse = Courses.FindItem(i)
            .AddItem LocCourse.CourseNo
            
            Debug.Print LocCourse.CourseNo & " - " & [CourseNo]
            
        Next
    End With

    RefreshCourses = True

Exit Function

ErrorExit:
    Set LocCourse = Nothing
    RefreshCourses = False
    Terminate
Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function


Public Function RefreshUserList() As Boolean
    Const StrPROCEDURE As String = "RefreshUserList()"

    Dim RstUserList As Recordset
    Dim i As Integer
    
    On Error GoTo ErrorHandler

    Set RstUserList = GetAccessList(CmoCourseNo)
    
    LstAccessList.Clear
    
    If Not RstUserList Is Nothing Then
        With RstUserList
            Do
                LstAccessList.AddItem
                LstAccessList.List(i, 0) = RstUserList!UserName
                .MoveNext
                i = i + 1
             Loop While Not .EOF
        End With
    End If
    Set RstUserList = Nothing
    RefreshUserList = True

Exit Function

ErrorExit:
    Terminate
    Set RstUserList = Nothing
    RefreshUserList = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
Public Function RefreshUserDropDown() As Boolean
    Const StrPROCEDURE As String = "RefreshUserDropDown()"

    Dim RstUserList As Recordset
    Dim i As Integer
    
    On Error GoTo ErrorHandler

    Set RstUserList = GetAccessList
    
    CmoUsers.Clear
    
    If Not RstUserList Is Nothing Then
        With RstUserList
            Do
                CmoUsers.AddItem RstUserList!UserName
                .MoveNext
             Loop While Not .EOF
        End With
    End If
    
    Set RstUserList = Nothing

    RefreshUserDropDown = True
Exit Function

ErrorExit:
    Terminate
    Set RstUserList = Nothing
    RefreshUserDropDown = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
Public Function ValidateData() As Boolean

    On Error Resume Next
    
    If CmoCourseNo = "" Then
        MsgBox "Please select a Course"
        ValidateData = False
        Exit Function
    End If
    
    If CmoUsers = "" Then
        MsgBox "Please select a user"
        ValidateData = False
        Exit Function
    End If
    
    ValidateData = True
        

End Function

Public Function UserformActivate() As Boolean
    Const StrPROCEDURE As String = "UserformActivate()"

    On Error GoTo ErrorHandler
    
    With LstHeadings
        .Clear
        .AddItem
        .List(0, 0) = "Users"
    End With
    If Not RefreshCourses Then Err.Raise HANDLED_ERROR
    If Not RefreshUserDropDown Then Err.Raise HANDLED_ERROR

    UserformActivate = True

Exit Function

ErrorExit:
    
    Terminate
    UserformActivate = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

