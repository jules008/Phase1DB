VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsMailItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 12 Oct 16
'===============================================================
' Methods
'---------------------------------------------------------------
' FindItem - finds  item from index
' AddItem - Add item to collection
' removeItem - Delete item from collection
' Count - returns number of items in collection
' GetCollection - Find all items sessions
' DeleteCollection - deletes all items
' GetParentFromPtr - Private routine to copy memory address of parent class
'===============================================================
Option Explicit

Private pMailItems As Collection
Private pParent As Long

'---------------------------------------------------------------
Public Property Get Parent() As ClsDailyLog
    If pParent <> 0 Then
        Set Parent = GetParentFromPtr(pParent)
    End If
End Property

Friend Function SetParent(ByVal Ptr As Long) As Boolean
    pParent = Ptr
End Function

'===============================================================
'Method FindItem
'finds item in collection
'---------------------------------------------------------------
Public Function FindItem(EmailNo As Variant) As ClsMailItem
    On Error Resume Next
    Set FindItem = pMailItems.Item(EmailNo)
End Function

'===============================================================
'Method AddItem
'Add training sesssion to collection
'---------------------------------------------------------------
Public Sub AddItem(MailItem As ClsMailItem)
    MailItem.SetParent ObjPtr(Me.Parent)
    pMailItems.Add MailItem, Key:=CStr(MailItem.EmailNo)
End Sub

'===============================================================
'Method removeItem
'deletes training item from index
'---------------------------------------------------------------
Public Sub RemoveItem(EmailNo As Variant)
    pMailItems.RemoveItem EmailNo
End Sub

'===============================================================
'Method count
'returns number of training sessions in collection
'---------------------------------------------------------------
Public Function Count() As Integer
    Count = pMailItems.Count
End Function
'
'===============================================================
'Method GetCollection
'Find all extra training sessions from Crew no and module no
'---------------------------------------------------------------
Public Sub GetCollection()
    Dim MailItem As ClsMailItem
    Dim RstMailItems As Recordset
    Dim StrCrewNo As String
    Dim DayNo As Integer
    Dim i As Integer

    StrCrewNo = "'" & Parent.Parent.CrewNo & "'"

    Set RstMailItems = database.SQLQuery("SELECT MailItemNo FROM MailItem WHERE CrewNo = " & StrCrewNo & _
                                        " AND deleted IS NULL")

    With RstMailItems
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            For i = 1 To .RecordCount
                Set MailItem = New ClsMailItem
'                MailItem.MailItemNo = !MailItemNo
                AddItem MailItem
                MailItem.LoadDB
                .MoveNext
            Next
        End If
    End With
    Set MailItem = Nothing
    Set RstMailItems = Nothing
End Sub

'===============================================================
'Method DeleteCollection
'deletes all Training sessions in collection
'---------------------------------------------------------------
Public Sub DeleteCollection()
    Dim MailItem As ClsMailItem
    Dim i As Integer

    For i = pMailItems.Count To 1 Step -1
        Set MailItem = pMailItems.Item(i)
        MailItem.DeleteDB
        pMailItems.Remove (i)
    Next

    Set MailItem = Nothing
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
    Set pMailItems = New Collection
End Sub

Private Sub Class_Terminate()
    Set pMailItems = Nothing
End Sub
'===============================================================





