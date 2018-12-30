VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsCourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
' v0,1 - WT2019 Version
'---------------------------------------------------------------
' Date - 30 Dec 18
'===============================================================
' Methods
'---------------------------------------------------------------
' LoadDB - Find course from Extra Training No
' NewDB - creates a new course in the DB
' UpdateDB - Update course to database
' DeleteDB - Delete course from database
' GetParentFromPtr - Private routine to copy memory address of parent class
'===============================================================
Option Explicit
Private pCourseNo As String
Private pStartDate As Date
Private pPassOutDate As Date
Private pCourseDirector As String
Private pCandidates As ClsCandidates
Private pStatus As String

'---------------------------------------------------------------
Public Property Get CourseNo() As String
    CourseNo = pCourseNo
End Property

Public Property Let CourseNo(ByVal vNewValue As String)
    pCourseNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get StartDate() As Date
    StartDate = pStartDate
End Property

Public Property Let StartDate(ByVal vNewValue As Date)
    pStartDate = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get PassOutDate() As Date
    PassOutDate = pPassOutDate
End Property

Public Property Let PassOutDate(ByVal vNewValue As Date)
    pPassOutDate = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get CourseDirector() As String
    CourseDirector = pCourseDirector
End Property

Public Property Let CourseDirector(ByVal vNewValue As String)
    pCourseDirector = vNewValue
End Property

'-----------------------------------------------------------------
Public Property Get Candidates() As ClsCandidates
    Set Candidates = pCandidates
End Property

Public Property Let Candidates(ByVal vNewValue As ClsCandidates)
    Set pCandidates = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Status() As String
    Status = pStatus
End Property

Public Property Let Status(ByVal vNewValue As String)
    pStatus = vNewValue
End Property

'===============================================================
'Method LoadDB
'load course from database
'---------------------------------------------------------------
Public Sub LoadDB()
    Dim RstCourse As Recordset
    
    Set RstCourse = ModDatabase.SQLQuery("SELECT * FROM course WHERE courseno = '" & pCourseNo & "'")
    
    With RstCourse
        If .RecordCount > 0 Then
            If Not IsNull(!CourseNo) Then pCourseNo = !CourseNo
            If Not IsNull(!StartDate) Then pStartDate = !StartDate
            If Not IsNull(!PassOutDate) Then pPassOutDate = !PassOutDate
            If Not IsNull(!CourseDirector) Then pCourseDirector = !CourseDirector
            If Not IsNull(!Status) Then pStatus = !Status
            
            pCandidates.GetCollection
        End If
    End With
    Set RstCourse = Nothing
End Sub

'===============================================================
'Method NewDB
'Creates a new course in the DB
'---------------------------------------------------------------
Public Sub NewDB()
    Dim RstCourse As Recordset
    Dim StrCourseNo As String
    
    StrCourseNo = "'" & pCourseNo & "'"
    
    Set RstCourse = ModDatabase.SQLQuery("SELECT * FROM course WHERE courseno = " & StrCourseNo & " AND deleted IS NULL")
        
    With RstCourse
        If .RecordCount <> 0 Then
            Debug.Print "Course already in the database"
        Else
            .AddNew
            !CourseNo = pCourseNo
            .Update
        End If
    End With
    
    Set RstCourse = Nothing
End Sub
'===============================================================
'Method UpdateDB
'Update course to database
'---------------------------------------------------------------
Public Function UpdateDB() As Boolean
    Dim RstCourse As Recordset
    Dim StrCourseNo As String
    
    StrCourseNo = "'" & pCourseNo & "'"
    
    Set RstCourse = ModDatabase.SQLQuery("SELECT * FROM course WHERE courseno = " & StrCourseNo & _
                                    " AND deleted IS NULL")
    
    With RstCourse
        If .RecordCount = 0 Then
            UpdateDB = False
        Else
            .Edit
            !CourseNo = pCourseNo
            !StartDate = pStartDate
            !PassOutDate = pPassOutDate
            !CourseDirector = pCourseDirector
            !Status = pStatus
            .Update
            UpdateDB = True
        End If
    End With
    Set RstCourse = Nothing
End Function

'===============================================================
'Method DeleteDB
' Delete course from database
'---------------------------------------------------------------
Public Sub DeleteDB()
    Dim RstCourse As Recordset
    Dim StrCourseNo As String
    
    StrCourseNo = "'" & pCourseNo & "'"
    
    Set RstCourse = ModDatabase.SQLQuery("SELECT * FROM course WHERE courseno = " & StrCourseNo & " AND deleted IS NULL")
    
    With RstCourse
        If .RecordCount <> 0 Then
            .Delete
            
            pCandidates.DeleteCollection
        End If
    End With
    Set RstCourse = Nothing
End Sub

'=========================================================================
Private Sub Class_Initialize()
    Set pCandidates = New ClsCandidates
    
    pCandidates.SetParent ObjPtr(Me)
End Sub

Private Sub Class_Terminate()
    pCandidates.SetParent 0
    
    Set pCandidates = Nothing
End Sub
'=========================================================================




