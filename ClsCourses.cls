VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsCourses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 03 Oct 16
'===============================================================
' Methods
'---------------------------------------------------------------
' FindItem - finds Course from index
' AddItem - Add Course to collection
' RemoveItem - remove Course from collection
' Count - returns number of Courses in collection
' GetCollection - Find all Courses
' DeleteCollection - deletes all Courses
'===============================================================
Option Explicit

Private pCourses As Collection

'===============================================================
'Method FindItem
'finds training item in collection
'---------------------------------------------------------------
Public Function FindItem(CourseNo As Variant) As ClsCourse
    On Error Resume Next
    If CourseNo = "" Then CourseNo = 1
    Set FindItem = pCourses.Item(CourseNo)
End Function

'===============================================================
'Method AddItem
'Add training sesssion to collection
'---------------------------------------------------------------
Public Sub AddItem(Course As ClsCourse)
    pCourses.Add Course, Key:=CStr(Course.CourseNo)
End Sub

'===============================================================
'Method RemoveItem
'deletes training item from index
'---------------------------------------------------------------
Public Sub RemoveItem(CourseNo As Variant)
    pCourses.Remove CourseNo
End Sub

'===============================================================
'Method count
'returns number of training sessions in collection
'---------------------------------------------------------------
Public Function Count() As Integer
    Count = pCourses.Count
End Function

'===============================================================
'Method GetCollection
'Find all Courses
'---------------------------------------------------------------
Public Sub GetCollection()
    Dim Course As ClsCourse
    Dim RstCourse As Recordset
    Dim i As Integer
    
    Set RstCourse = ModDatabase.SQLQuery("SELECT * FROM Course WHERE deleted IS NULL")
    
    With RstCourse
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            For i = 1 To .RecordCount
                Set Course = New ClsCourse
                Course.CourseNo = !CourseNo
                AddItem Course
                Course.LoadDB
                .MoveNext
            Next
        End If
    End With
    Set Course = Nothing
    Set RstCourse = Nothing
End Sub

'===============================================================
'Method DeleteCollection
'deletes all Courses
'---------------------------------------------------------------
Public Sub DeleteCollection()
    Dim Course As ClsCourse
    Dim i As Integer
    
    For i = Count To 1 Step -1
        Set Course = pCourses.Item(i)
        Course.DeleteDB
        pCourses.Remove (i)
    Next
    Set Course = Nothing
End Sub

'===============================================================
'Method GetParentFromPtr
' Private routine to copy memory address of parent class
'---------------------------------------------------------------
Private Function GetParentFromPtr(ByVal Ptr As Long) As ClsDailyLog
    Dim tmp As ClsDailyLog
    
    CopyMemory tmp, Ptr, 4
    Set GetParentFromPtr = tmp
    CopyMemory tmp, 0&, 4
End Function

'===============================================================
Private Sub Class_Initialize()
    Set pCourses = New Collection
    GetCollection
End Sub

Private Sub Class_Terminate()
    Set pCourses = Nothing
End Sub
'===============================================================




