VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsDevelopmentPlans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================
' v0,0 - Initial version
'---------------------------------------------------------------
' Date - 27 Sep 16
'===============================================================
' Methods
'---------------------------------------------------------------
' FindItem - finds Item from index
' AddItem - Add Item to collection
' RemoveItem - Delete Item from collection
' Count - returns number of Items in collection
' GetCollection - Find all Items from Crew no and module no
' DeleteCollection - deletes all training session from Crew No and Module No
' SetLocalNos - Resets all local DP Nos
' GetParentFromPtr - Private routine to copy memory address of parent class
'===============================================================
Option Explicit
Private pDevelopmentPlans As Collection
Private pNoOpen As Integer
Private pNoClosed As Integer
Private pNoOverDue As Integer
Private pParent As Long

'---------------------------------------------------------------
Public Property Get NoOpen() As Variant
    Dim DevelopmentPlan As ClsDevelopmentPlan
    
    pNoOpen = 0
    For Each DevelopmentPlan In pDevelopmentPlans
        If DevelopmentPlan.Status <> "Failed" And DevelopmentPlan.Status <> "Passed" Then
            pNoOpen = pNoOpen + 1
        End If
    Next
    
    NoOpen = pNoOpen
    
    Set DevelopmentPlan = Nothing
End Property

Public Property Let NoOpen(ByVal vNewValue As Variant)
    pNoOpen = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get NoClosed() As Variant
    Dim DevelopmentPlan As ClsDevelopmentPlan
    
    pNoClosed = 0
    For Each DevelopmentPlan In pDevelopmentPlans
        If DevelopmentPlan.Status = "Failed" Or DevelopmentPlan.Status = "Passed" Then
            pNoClosed = pNoClosed + 1
        End If
    Next
    
    NoClosed = pNoClosed
    
    Set DevelopmentPlan = Nothing
End Property

Public Property Let NoClosed(ByVal vNewValue As Variant)
    pNoClosed = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get NoOverDue() As Variant
    Dim DevelopmentPlan As ClsDevelopmentPlan
    
    pNoOverDue = 0
    For Each DevelopmentPlan In pDevelopmentPlans
        If DevelopmentPlan.ReviewDate < Now Then
            If DevelopmentPlan.Status <> "Failed" Or DevelopmentPlan.Status <> "Passed" Then
                pNoOverDue = pNoOverDue + 1
            End If
        End If
    Next
    
    NoOverDue = pNoOverDue
    
    Set DevelopmentPlan = Nothing
End Property

Public Property Let NoOverDue(ByVal vNewValue As Variant)
    pNoOverDue = vNewValue
End Property

'---------------------------------------------------------------
Public Property Get Parent() As ClsCandidate
    If pParent <> 0 Then
        Set Parent = GetParentFromPtr(pParent)
    End If
End Property

Friend Function SetParent(ByVal Ptr As Long) As Boolean
    pParent = Ptr
End Function

'===============================================================
'Method FindItem
'finds Item from index
'---------------------------------------------------------------
Public Function FindItem(DPNo As Variant) As ClsDevelopmentPlan
    On Error Resume Next
    Set FindItem = pDevelopmentPlans.Item(DPNo)
End Function

'===============================================================
'Method AddItem
'Add Item to collection
'---------------------------------------------------------------
Public Sub AddItem(DevelopmentPlan As ClsDevelopmentPlan)
    DevelopmentPlan.SetParent ObjPtr(Me.Parent)
    pDevelopmentPlans.Add DevelopmentPlan, Key:=CStr(DevelopmentPlan.DPNo)
End Sub

'===============================================================
'Method RemoveItem
'deletes Item from index
'---------------------------------------------------------------
Public Sub RemoveItem(DPNo As Variant)
    pDevelopmentPlans.Remove DPNo
    SetLocalNos
End Sub

'===============================================================
'Method count
'returns number of Items in collection
'---------------------------------------------------------------
Public Function Count() As Integer
    Count = pDevelopmentPlans.Count
End Function

'===============================================================
'Method GetCollection
'Find all Items for parent candidate
'---------------------------------------------------------------
Public Sub GetCollection()
    Dim DevelopmentPlan As ClsDevelopmentPlan
    Dim RstDevelopmentPlans As Recordset
    Dim StrCrewNo As String
    Dim i As Integer
    
    StrCrewNo = "'" & Parent.CrewNo & "'"
    
    Set RstDevelopmentPlans = ModDatabase.SQLQuery("SELECT * FROM Developmentplan WHERE crewno = " & StrCrewNo & _
                                            " AND deleted IS NULL")
    
    With RstDevelopmentPlans
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            For i = 1 To .RecordCount
                Set DevelopmentPlan = New ClsDevelopmentPlan
                DevelopmentPlan.DPNo = !DPNo
                AddItem DevelopmentPlan
                DevelopmentPlan.LoadDB
                .MoveNext
            Next
        End If
        SetLocalNos
        Set RstDevelopmentPlans = Nothing
        Set DevelopmentPlan = Nothing
    End With
End Sub

'===============================================================
'Method DeleteCollection
'deletes all Items
'---------------------------------------------------------------
Public Sub DeleteCollection()
    Dim DevelopmentPlan As ClsDevelopmentPlan
    Dim i As Integer
    
    For i = Me.Count To 1 Step -1
        Set DevelopmentPlan = pDevelopmentPlans.Item(i)
        DevelopmentPlan.DeleteDB
        pDevelopmentPlans.Remove (i)
    Next
    Set DevelopmentPlan = Nothing
End Sub

'===============================================================
'Method SetLocalNos
'Resets all local DP Nos
'---------------------------------------------------------------
Public Sub SetLocalNos()
    Dim DevelopmentPlan As ClsDevelopmentPlan
    Dim RstDevelopmentPlans As Recordset
    Dim StrCrewNo As String
    Dim i As Integer
    
    i = 1
    For Each DevelopmentPlan In pDevelopmentPlans
        DevelopmentPlan.LocalDPNo = i
        DevelopmentPlan.UpdateDB
        i = i + 1
    Next
    Set DevelopmentPlan = Nothing
End Sub

'===============================================================
'Method CleanUp
'deletes unfinished items
'---------------------------------------------------------------
Public Sub CleanUp()
    Dim DevelopmentPlan As ClsDevelopmentPlan
    
    For Each DevelopmentPlan In pDevelopmentPlans
        If DevelopmentPlan.DPNo = 0 Then
            pDevelopmentPlans.Remove ("0")
        End If
    Next
    SetLocalNos
End Sub
'===============================================================
'Method GetParentFromPtr
' Private routine to copy memory address of parent class
'---------------------------------------------------------------
Private Function GetParentFromPtr(ByVal Ptr As Long) As ClsDevelopmentPlan
    Dim tmp As ClsDevelopmentPlan
    
    CopyMemory tmp, Ptr, 4
    Set GetParentFromPtr = tmp
    CopyMemory tmp, 0&, 4
End Function

'===============================================================
Private Sub Class_Initialize()
    Set pDevelopmentPlans = New Collection
End Sub

Private Sub Class_Terminate()
    Set pDevelopmentPlans = Nothing
End Sub






