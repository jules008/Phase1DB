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
' Class ClsModules
' v0,0 - Initial version
' v0,1 - WT2018 Version
'---------------------------------------------------------------
' Date - 19 Dec 18
'===============================================================
' Methods
'---------------------------------------------------------------
' FindItem - finds Item from index
' AddItem - Add Item to collection
' RemoveItem - Deletes item from collection
' ItemCount - Returns number of items in collection
' GetCollection - Gets all items in collection
' DeleteCollection - Deletes all items in collection
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

' ===============================================================
' Method AddItem
' Add item to Collection
'---------------------------------------------------------------
Public Sub AddItem(Module As ClsModule)
    Module.SetParent ObjPtr(Me)
    pModules.Add Module, Key:=CStr(Module.ModuleNo)
End Sub

' ===============================================================
' Method RemoveItem
' Deletes item from collection
'---------------------------------------------------------------
Public Sub RemoveItem(ModuleNo As Variant)
    pModules.Remove ModuleNo
End Sub

' ===============================================================
' Method ItemCount
' Returns number of items in collection
'---------------------------------------------------------------
Public Function Count() As Integer
    Count = pModules.Count
End Function

' ===============================================================
' Method GetCollection
' Gets all items in collection
'---------------------------------------------------------------
Public Sub GetCollection()
    Dim RstModule As Recordset
    Dim Module As ClsModule
    Dim i As Integer

    Set RstModule = ModDatabase.SQLQuery("Module")
    With RstModule
        .MoveLast
        .MoveFirst
        For i = 1 To .RecordCount
            Set Module = New ClsModule
            Module.ModuleNo = !ModuleNo
            Module.DBGet
            Me.AddItem Module
            .Movenext
        Next
    End With
End Sub

' ===============================================================
Private Sub Class_Initialize()
    Set pModules = New Collection
    GetCollection
End Sub

'---------------------------------------------------------------
Private Sub Class_Terminate()
    Set pModules = Nothing

End Sub

'---------------------------------------------------------------
