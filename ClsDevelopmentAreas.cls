VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsDevelopmentAreas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 16 Sep 16
'===============================================================
' Methods
'---------------------------------------------------------------
' FindItem - finds item from index
' AddItem - Add item to collection
' removeItem - Delete item from collection
' Count - returns number of items in collection
' GetCollection - Find all items sessions
' DeleteCollection - deletes all items session
' GetParentFromPtr - Private routine to copy memory address of parent class
'===============================================================
Option Explicit
Private pDevAreas As Collection
Private pParent As Long

'---------------------------------------------------------------
Public Property Get Parent() As ClsDevelopmentPlan
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
Public Function FindItem(DevArea As Variant) As ClsDevelopmentArea
    On Error Resume Next
    Set FindItem = pDevAreas.Item(DevArea)
End Function

'===============================================================
'Method AddItem
'Add item to collection
'---------------------------------------------------------------
Public Sub AddItem(DevArea As ClsDevelopmentArea)
    On Error GoTo ErrorHandler
    
    DevArea.SetParent ObjPtr(Me.Parent)
    pDevAreas.Add DevArea, Key:=CStr(DevArea.DevArea)

ErrorHandler:
    If Err.Number = 457 Then
        MsgBox "There is already an entry for the selected Development Area"
        FrmDPDevArea.CmoArea.Value = ""
    End If
    
End Sub

'===============================================================
'Method removeItem
'deletes item from collection
'---------------------------------------------------------------
Public Sub RemoveItem(Index As Variant)
    pDevAreas.Remove Index
End Sub

'===============================================================
'Method count
'returns number of itms in collection
'---------------------------------------------------------------
Public Function Count() As Integer
    Count = pDevAreas.Count
End Function

'===============================================================
'Method GetCollection
'Find all items
'---------------------------------------------------------------
Public Sub GetCollection()
    Dim DevArea As ClsDevelopmentArea
    Dim RstDevArea As Recordset
    Dim DPNo As Integer
    Dim StrDevArea As String
    Dim i As Integer
    
    DPNo = Parent.DPNo
    
    Set RstDevArea = database.SQLQuery("SELECT * FROM DPDevItems WHERE dpno = " & DPNo & _
                                        " AND deleted IS NULL")
    
    With RstDevArea
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            For i = 1 To .RecordCount
                Set DevArea = New ClsDevelopmentArea
                DevArea.DevArea = !DevArea
                AddItem DevArea
                DevArea.LoadDB
                .MoveNext
            Next
        End If
    End With
    Set DevArea = Nothing
    Set RstDevArea = Nothing
End Sub

'===============================================================
'Method DeleteCollection
'deletes all Items
'---------------------------------------------------------------
Public Sub DeleteCollection()
    Dim DevArea As ClsDevelopmentArea
    Dim i As Integer
    
    For i = Me.Count To 1 Step -1
        Set DevArea = pDevAreas.Item(i)
        DevArea.DeleteDB
        pDevAreas.Remove (i)
    Next
    
    Set DevArea = Nothing
End Sub

'===============================================================
'Method GetParentFromPtr
' Private routine to copy memory address of parent class
'---------------------------------------------------------------
Private Function GetParentFromPtr(ByVal Ptr As Long) As ClsDevelopmentArea
    Dim tmp As ClsDevelopmentArea
    
    CopyMemory tmp, Ptr, 4
    Set GetParentFromPtr = tmp
    CopyMemory tmp, 0&, 4
End Function

'===============================================================
Private Sub Class_Initialize()
    Set pDevAreas = New Collection
End Sub

Private Sub Class_Terminate()
    Set pDevAreas = Nothing
End Sub
'===============================================================




