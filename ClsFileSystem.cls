VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsFileSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public FSO As New FileSystemObject

'GetFolders Method
'-----------------
Public Function GetFolders(FileLoc As String) As Folders

    
    Set GetFolders = FSO.GetFolder(FileLoc).SubFolders

    
End Function

'GetFiles Method
'---------------
Public Function GetFiles(FileLoc As String) As Files
    
    
    Set GetFiles = FSO.GetFolder(FileLoc).Files
   
    
End Function








