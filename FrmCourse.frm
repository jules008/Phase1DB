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
' v1,0 - WT2019 Version
'---------------------------------------------------------------
' Date - 29 Dec 18
'===============================================================
Option Explicit
Private Const StrMODULE As String = "FrmCourse"
Private Course As ClsCourse
Private FormChanged As Boolean

' ===============================================================
' ShowForm
' Shows Course form
' ---------------------------------------------------------------
Public Function ShowForm(Optional LocalCourse As ClsCourse) As Boolean
    Const StrPROCEDURE As String = "ShowForm()"

    On Error GoTo ErrorHandler

    If Not ResetForm Then Err.Raise HANDLED_ERROR
           
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

' ===============================================================
' ResetForm
' Resets form and clears fields
' ---------------------------------------------------------------
Private Function ResetForm() As Boolean
    Const StrPROCEDURE As String = "ResetForm()"

    On Error GoTo ErrorHandler

    FormChanged = False
    Me.CmoCourseDirector.Value = ""
    Me.CmoStatus.Value = ""
    Me.TxtCourseNo = ""
    Me.TxtPassOutDate = ""
    Me.TxtStrtDate = ""

    ResetForm = True

Exit Function

ErrorExit:

    ResetForm = False

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
' PopulateForm
' Polulates course form with values
' ---------------------------------------------------------------
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
' BtnClose_Click
' Close event of form
' ---------------------------------------------------------------
Private Sub BtnClose_Click()
    Dim ErrNo As Integer
    Dim Response As Integer
    
    Const StrPROCEDURE As String = "BtnClose_Click()"

    On Error GoTo ErrorHandler

Restart:
    
    If Course Is Nothing Then Err.Raise SYSTEM_RESTART
      
    unload Me
    
GracefulExit:

Exit Sub

ErrorExit:

Exit Sub

ErrorHandler:
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        ErrNo = Err.Number
        CustomErrorHandler (Err.Number)
        If ErrNo = SYSTEM_RESTART Then Resume Restart Else Resume GracefulExit
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If

End Sub

' ===============================================================
' BtnDelete_Click
' Deletes Course
' ---------------------------------------------------------------
Private Sub BtnDelete_Click()
    Dim ErrNo As Integer
    
    Const StrPROCEDURE As String = "BtnDelete_Click()"
    
    Dim Response As Integer
    
    On Error GoTo ErrorHandler
    
Restart:
    
    If Course Is Nothing Then Err.Raise SYSTEM_RESTART

    Response = MsgBox("Are you sure you want to delete the course?", vbYesNo)
    
    If Response = 6 Then
        Courses.RemoveItem Course.CourseNo
        Course.DeleteDB
        ResetForm
        FormChanged = False
    End If

GracefulExit:

Exit Sub

ErrorExit:

Exit Sub
ErrorHandler:
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        ErrNo = Err.Number
        CustomErrorHandler (Err.Number)
        If ErrNo = SYSTEM_RESTART Then Resume Restart Else Resume GracefulExit
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' BtnNew_Click
' Creates a new course
' ---------------------------------------------------------------
Private Sub BtnNew_Click()
    
    On Error Resume Next
    
    ResetForm
    Set Course = New ClsCourse
    TxtCourseNo.Enabled = True
    
End Sub

' ===============================================================
' BtnUpdate_Click
' Updates any changes to the course
' ---------------------------------------------------------------
Private Sub BtnUpdate_Click()
    Dim ErrNo As Integer

    Const StrPROCEDURE As String = "BtnUpdate_Click()"

    On Error GoTo ErrorHandler

Restart:

    If Course Is Nothing Then Err.Raise SYSTEM_RESTART

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
            
            unload me
        End With
    End If
    
GracefulExit:

Exit Sub
    
ErrorExit:

Exit Sub

ErrorHandler:
    If Err.Number >= 1000 And Err.Number <= 1500 Then
        ErrNo = Err.Number
        CustomErrorHandler (Err.Number)
        If ErrNo = SYSTEM_RESTART Then Resume Restart Else Resume GracefulExit
    End If

    If CentralErrorHandler(StrMODULE, StrPROCEDURE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

' ===============================================================
' CmoCourseDirector_Change
' Detects changes to the form
' ---------------------------------------------------------------
Private Sub CmoCourseDirector_Change()
    FormChanged = True
End Sub

' ===============================================================
' CmoStatus_Change
' Detects changes to the form
' ---------------------------------------------------------------
Private Sub CmoStatus_Change()
    FormChanged = True
End Sub

' ===============================================================
' TxtCourseNo_Change
' Detects changes to the form
' ---------------------------------------------------------------
Private Sub TxtCourseNo_Change()
    FormChanged = True
End Sub

' ===============================================================
' TxtPassOutDate_Change
' Detects changes to the form
' ---------------------------------------------------------------
Private Sub TxtPassOutDate_Change()
    FormChanged = True
End Sub

' ===============================================================
' TxtStrtDate_Change
' Detects changes to the form
' ---------------------------------------------------------------
Private Sub TxtStrtDate_Change()
    FormChanged = True
End Sub

' ===============================================================
' UserForm_Initialize
' Trigger Initialise form function
' ---------------------------------------------------------------
Private Sub UserForm_Initialize()
    On Error Resume Next
    
    FormInitialise
    
End Sub

' ===============================================================
' CmoCourseDirector_Change
' Detects changes to the form
' ---------------------------------------------------------------
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

' ===============================================================
' FormInitialise
' Initialise form
' ---------------------------------------------------------------
Public Sub FormInitialise()
    
    Const StrPROCEDURE As String = "FormInitialise()"
    
    Dim cell As Range
    Dim RstUsers As Recordset
    
    On Error GoTo ErrorHandler
    
    Set RstUsers = GetAccessList

    CmoCourseDirector.Clear
    
    If Not RstUsers Is Nothing Then
        With RstUsers
            Do While Not .EOF
                Me.CmoCourseDirector.AddItem !UserName
                .MoveNext
            Loop
        End With
    End If
    
    'get Status list
    CmoStatus.Clear
    
    For Each cell In ShtLists.Range("CourseStatus")
        Me.CmoStatus.AddItem cell
    Next
        
    Set RstUsers = Nothing
Exit Sub

ErrorExit:
    
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

' ===============================================================
' FormTerminate
' Terminates form
' ---------------------------------------------------------------
Public Sub FormTerminate()
    On Error Resume Next
    Set Course = Nothing
End Sub

' ===============================================================
' UserForm_Terminate
' Terminates form
' ---------------------------------------------------------------
Private Sub UserForm_Terminate()
    On Error Resume Next
    FormTerminate
End Sub
