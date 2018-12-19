VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' Class ClsModule
' v0,0 - Initial version
' v0,1 - WT2018 Version
'---------------------------------------------------------------
' Date - 19 Dec 18
'===============================================================
' Methods
'---------------------------------------------------------------
' LoadDB - Get module from database
'===============================================================
Option Explicit
Private pModuleID As String
Private pDayNo As Integer
Private pModuleNo As Integer
Private pModule As String
Private pAssessment As Boolean
Private pParent As Long

'---------------------------------------------------------------
Public Property Get ModuleID() As Integer
    ModuleID = pModuleID
End Property

Public Property Let ModuleID(ByVal vNewValue As Integer)
    pModuleID = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get DayNo() As Integer
    DayNo = pDayNo
End Property

Public Property Let DayNo(ByVal vNewValue As Integer)
    pDayNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get ModuleNo() As String
    ModuleNo = pModuleNo
End Property

Public Property Let ModuleNo(ByVal vNewValue As String)
    pModuleNo = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Module() As String
    Module = pModule
End Property

Public Property Let Module(ByVal vNewValue As String)
    pModule = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Assessment() As Boolean
    Assessment = pAssessment
End Property

Public Property Let Assessment(ByVal vNewValue As Boolean)
    pAssessment = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Parent() As ClsModules
    If pParent <> 0 Then
        Set Parent = GetParentFromPtr(pParent)
    End If
End Property

Friend Function SetParent(ByVal Ptr As Long) As Boolean
    pParent = Ptr
End Function

' ===============================================================
' Method GetParentFromPtr
' Private routine to copy memory address of parent class
' ---------------------------------------------------------------
Private Function GetParentFromPtr(ByVal Ptr As Long) As ClsModules
    Dim tmp As ClsModules

    CopyMemory tmp, Ptr, 4
    Set GetParentFromPtr = tmp
    CopyMemory tmp, 0&, 4
End Function

'===============================================================
'Method DBGet
'Find Module Index no
'---------------------------------------------------------------
Public Sub DBGet()
    Dim RstModule As Recordset
    
    Set RstModule = ModDatabase.SQLQuery("SELECT * FROM module WHERE dayno = " & pModuleNo)
    
    With RstModule
        If .RecordCount > 0 Then
            If Not IsNull(!ModuleID) Then pModuleID = !ModuleID
            If Not IsNull(!DayNo) Then pDayNo = !DayNo
            If Not IsNull(!Module) Then pModule = !Module
            If Not IsNull(!ModuleNo) Then pModuleNo = !ModuleNo
            If !Assessment = True Then pAssessment = True Else pAssessment = False
        End If
    End With
    Set RstModule = Nothing
End Sub


