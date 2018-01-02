VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmPrintCopies 
   Caption         =   "No of Copies"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2970
   OleObjectBlob   =   "FrmPrintCopies.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmPrintCopies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BtnOk_Click()
    Me.Hide
    
End Sub

Private Sub UserForm_Initialize()
    Dim i As Integer
    
    For i = 1 To 5
    
        Me.CmoNoCopies.AddItem i
    Next
End Sub

