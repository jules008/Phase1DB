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
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 04 Nov 16
'===============================================================
' Methods
'---------------------------------------------------------------
' LoadDB - Get module from database
'===============================================================
Option Explicit
Private pModuleID As Integer
Private pDayNo As Integer
Private pModuleNo As String
Private pModule As String
Private pAssessment As Boolean

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

'===============================================================
'Method LoadDB
'Find Module Index no
'---------------------------------------------------------------
Public Sub LoadDB()
    Dim RstModule As Recordset
    
    Set RstModule = database.SQLQuery("SELECT * FROM module WHERE dayno = " & pDayNo & " AND deleted IS NULL")
    
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

