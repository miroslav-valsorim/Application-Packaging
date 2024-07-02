Option Explicit

Const DeleteReadOnly = True 
Dim oFSO, oDrive, sFileName

Set oFSO   = CreateObject("Scripting.FileSystemObject") 
sFileName  = "Launcher.config"

For Each oDrive In oFSO.Drives 
  If oDrive.DriveType = 2 Then Recurse oDrive.RootFolder
Next 

Sub Recurse(oFolder)
  Dim oSubFolder, oFile

  If IsAccessible(oFolder) Then
    For Each oSubFolder In oFolder.SubFolders
     Recurse oSubFolder
    Next 

    For Each oFile In oFolder.Files
      If oFile.Name = sFileName Then
        'oFile.Delete ' or whatever
      End If
    Next 
  End If
End Sub

Function IsAccessible(oFolder)
  On Error Resume Next
  IsAccessible = oFolder.SubFolders.Count >= 0
End Function