VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsModules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 20 Sep 16
'===============================================================
' Methods
'---------------------------------------------------------------
' FindItem - finds Item from index
' AddItem - Add Item to collection
' RemoveItem - Delete Item from collection
' Count - returns number of Items in collection
' GetCollection - Find all Items from Crew no and module no
' GetParentFromPtr - Private routine to copy memory address of parent class
'===============================================================
Option Explicit

Private pModules As Collection

'---------------------------------------------------------------
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pModules.[_newenum]
End Function

' ===============================================================
' Method FindItem
' Finds Item from index
'---------------------------------------------------------------

Public Function FindItem(ModuleNo As Variant) As ClsModule
Attribute FindItem.VB_UserMemId = 0
    On Error Resume Next
    Set FindItem = pModules.Item(ModuleNo)
End Function

'===============================================================
'Method AddItem
' Add item to Collection
'---------------------------------------------------------------
Public Sub AddItem(Module As ClsModule)
    Module.SetParent ObjPtr(Me)
    pModules.Add Module, Key:=CStr(Module.DayNo)
End Sub

'===============================================================
'Method RemoveItem
'deletes Item from index
'---------------------------------------------------------------
Public Sub RemoveItem(ModuleNo As Variant)
    pModules.Remove ModuleNo
End Sub

'===============================================================
'Method count
'returns number of Items in collection
'---------------------------------------------------------------
Public Function Count() As Integer
    Count = pModules.Count
End Function

'===============================================================
'Method GetCollection
'Find all Items for parent candidate
'---------------------------------------------------------------
Public Sub GetCollection()
    Dim Module As ClsModule
    Dim RstModule As Recordset
    Dim i As Integer

    Set RstModule = ModDatabase.SQLQuery("Module")
    
    With RstModule
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            For i = 1 To .RecordCount
                Set Module = New ClsModule
                Module.DayNo = !DayNo
                Module.LoadDB
                AddItem Module
                .MoveNext
            Next
       End If
       Set Module = Nothing
       Set RstModule = Nothing
    End With
End Sub

'===============================================================
Private Sub Class_Initialize()
    Set pModules = New Collection
    GetCollection
End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pModules = Nothing
End Sub



