Option Explicit

Dim folderPath, marker
folderPath = "C:\Path\To\Your\Moveit\Backups"
marker = "P.00992001720."

Dim fso, folder, file
Dim startPos, endPos, newName, fullNewPath

Set fso = CreateObject("Scripting.FileSystemObject")

' Silent exit if folder missing
If Not fso.FolderExists(folderPath) Then WScript.Quit

Set folder = fso.GetFolder(folderPath)

For Each file In folder.Files

  ' Only process .dat-related files (including .dat.xxx cases)
  If InStr(1, LCase(file.Name), ".dat") > 0 Then

    ' Find marker
    startPos = InStr(1, file.Name, marker, vbTextCompare)

    If startPos > 0 Then
      
      ' Find FIRST ".dat" after marker
      endPos = InStr(startPos, file.Name, ".dat", vbTextCompare)

      If endPos > 0 Then
        
        ' Extract clean filename: marker → end of ".dat"
        newName = Mid(file.Name, startPos, (endPos - startPos) + 4)

        ' Rename only if needed
        If file.Name <> newName Then
          
          fullNewPath = folder.Path & "\" & newName

          ' Prevent overwrite
          If Not fso.FileExists(fullNewPath) Then
            On Error Resume Next
            file.Name = newName
            On Error GoTo 0
          End If

        End If

      End If

    End If

  End If

Next
