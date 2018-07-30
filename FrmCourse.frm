VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmCourse 
   Caption         =   "Add Course"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5055
   OleObjectBlob   =   "FrmCourse.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmCourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 26 Sep 16
'===============================================================
Option Explicit
Private Const StrMODULE As String = "FrmCourse"
Private Course As ClsCourse
Private FormChanged As Boolean

Public Function ShowForm(Optional LocalCourse As ClsCourse) As Boolean
    Const StrPROCEDURE As String = "ShowForm()"
    
    On Error GoTo ErrorHandler
    
    ResetForm
    
    If LocalCourse Is Nothing Then
        Set Course = New ClsCourse
        TxtCourseNo.Enabled = True
    Else
        Set Course = LocalCourse
        TxtCourseNo.Enabled = False
    End If
    
    If Not PopulateForm Then Err.Raise HANDLED_ERROR
    FormChanged = False
    Show
    
    ShowForm = True

Exit Function

ErrorExit:
    FormTerminate
    Terminate
    ShowForm = False

Exit Function

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

Private Sub ResetForm()

    On Error Resume Next
    
    FormChanged = False
    Me.CmoCourseDirector.Value = ""
    Me.CmoStatus.Value = ""
    Me.TxtCourseNo = ""
    Me.TxtPassOutDate = ""
    Me.TxtStrtDate = ""
End Sub

Private Function PopulateForm() As Boolean

    Const StrPROCEDURE As String = "PopulateForm()"

    On Error GoTo ErrorHandler

    With Course
        CmoCourseDirector = .CourseDirector
        CmoStatus = .Status
        TxtCourseNo = .CourseNo
        If .PassOutDate <> 0 Then TxtPassOutDate = .PassOutDate
        If .StartDate <> 0 Then TxtStrtDate = .StartDate
    End With
    PopulateForm = True
Exit Function

ErrorExit:
    PopulateForm = False
    FormTerminate
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

Private Sub BtnClose_Click()
    Dim Response As Integer
    
    On Error Resume Next
    
    If FormChanged = True Then
        Response = MsgBox("The form has been changed, would you like to save these changes?", vbYesNo)
        
        If Response = 6 Then BtnUpdate_Click
        FormChanged = False
    End If
    FormTerminate
    Me.Hide
End Sub

Private Sub BtnDelete_Click()
    
    Const StrPROCEDURE As String = "BtnDelete_Click()"
    
    Dim Response As Integer
    
    On Error GoTo ErrorHandler
    
    Response = MsgBox("Are you sure you want to delete the course?", vbYesNo)
    
    If Response = 6 Then
        Courses.RemoveItem Course.CourseNo
        Course.DeleteDB
        ResetForm
        FormChanged = False
    End If

Exit Sub

ErrorExit:
    FormTerminate
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

Private Sub BtnNew_Click()
    
    On Error Resume Next
    
    ResetForm
    Set Course = New ClsCourse
    TxtCourseNo.Enabled = True
    
End Sub

Private Sub BtnUpdate_Click()

    Const StrPROCEDURE As String = "BtnUpdate_Click()"
    
    On Error GoTo ErrorHandler
    
    If ValidateData Then
        With Course
            .CourseDirector = CmoCourseDirector
            .CourseNo = TxtCourseNo
            .PassOutDate = TxtPassOutDate
            .StartDate = TxtStrtDate
            .Status = CmoStatus
            
            If Not .UpdateDB Then
                .NewDB
                .UpdateDB
                ModGlobals.Courses.AddItem Course
                ShtCourse.CmoCourseNo = TxtCourseNo
            End If
            FormTerminate
            Hide
        End With
    End If
Exit Sub
    
ErrorExit:
    FormTerminate
    Terminate
    Hide

Exit Sub

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

Private Sub CmoCourseDirector_Change()
    FormChanged = True
End Sub

Private Sub CmoStatus_Change()
    FormChanged = True
End Sub

Private Sub TxtCourseNo_Change()
    FormChanged = True
End Sub

Private Sub TxtPassOutDate_Change()
    FormChanged = True
End Sub

Private Sub TxtStrtDate_Change()
    FormChanged = True
End Sub

Private Sub UserForm_Initialize()
    On Error Resume Next
    FormInitialise
End Sub

Private Function ValidateData() As Boolean
    
    On Error Resume Next
    
    If Me.TxtCourseNo = "" Then
        MsgBox "Please enter a course no"
        ValidateData = False
        Exit Function
    End If
        
    If Me.TxtPassOutDate = "" Then
        MsgBox "Please enter a pass out date"
        ValidateData = False
        Exit Function
    End If
        
    If Not IsDate(Me.TxtStrtDate) Then
        MsgBox "Please enter a valid Start Date"
        ValidateData = False
        Exit Function
    End If
    
    If Not IsDate(Me.TxtPassOutDate) Then
        MsgBox "Please enter a valid Pass Out Date"
        ValidateData = False
        Exit Function
    End If
    
    If Me.CmoCourseDirector = "" Then
        MsgBox "Please enter a course director"
        ValidateData = False
        Exit Function
    End If
        
    If Me.CmoStatus = "" Then
        MsgBox "Please enter a status"
        ValidateData = False
        Exit Function
    End If
        
    ValidateData = True
End Function

Public Sub FormInitialise()
    
    Const StrPROCEDURE As String = "FormInitialise()"
    
    Dim cell As Range
    Dim RstUsers As Recordset
    
    On Error GoTo ErrorHandler
    
    Set RstUsers = GetAccessList

    'get Course director list
    CmoCourseDirector.Clear
    
    With RstUsers
        Do
        Me.CmoCourseDirector.AddItem !UserName
        .MoveNext
        Loop While Not .EOF
    End With
    
    'get Status list
    CmoStatus.Clear
    
    For Each cell In ShtLists.Range("CourseStatus")
        Me.CmoStatus.AddItem cell
    Next
    Set RstUsers = Nothing
Exit Sub

ErrorExit:
    FormTerminate
    Terminate
    Set RstUsers = Nothing

Exit Sub

ErrorHandler:
    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

Public Sub FormTerminate()
    On Error Resume Next
    Set Course = Nothing
End Sub

Private Sub UserForm_Terminate()
    On Error Resume Next
    FormTerminate
End Sub
