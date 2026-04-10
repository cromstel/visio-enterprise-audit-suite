Option Explicit

Dim folderPath, marker
folderPath = "C:\Path\To\Your\Moveit\Backups"
marker = "P.00992001720."

Dim fso, folder, file
Dim startPos, endPos, newName, fullNewPath

Set fso = CreateObject("Scripting.FileSystemObject")

WScript.Echo "=== SCRIPT START ==="
WScript.Echo "Folder: " & folderPath
WScript.Echo "Marker: " & marker

' Check folder
If Not fso.FolderExists(folderPath) Then
  WScript.Echo "ERROR: Folder not found!"
  WScript.Quit
End If

Set folder = fso.GetFolder(folderPath)

For Each file In folder.Files

  WScript.Echo "----------------------------------"
  WScript.Echo "Processing: " & file.Name

  ' Only .dat files
  If LCase(fso.GetExtensionName(file.Name)) = "dat" Then
    WScript.Echo "✔ .dat file detected"

    ' Find marker
    startPos = InStr(1, file.Name, marker, vbTextCompare)

    If startPos > 0 Then
      WScript.Echo "✔ Marker found at position: " & startPos

      ' Find ".dat" after marker
      endPos = InStr(startPos, file.Name, ".dat", vbTextCompare)

      If endPos > 0 Then
        WScript.Echo "✔ .dat found at position: " & endPos

        newName = Mid(file.Name, startPos, (endPos - startPos) + 4)
        WScript.Echo "New name will be: " & newName

        If file.Name <> newName Then

          fullNewPath = folder.Path & "\" & newName

          If Not fso.FileExists(fullNewPath) Then
            WScript.Echo "Renaming..."

            On Error Resume Next
            file.Name = newName

            If Err.Number <> 0 Then
              WScript.Echo "❌ ERROR renaming: " & Err.Description
              Err.Clear
            Else
              WScript.Echo "✅ Renamed successfully"
            End If

            On Error GoTo 0
          Else
            WScript.Echo "⚠ Skipped: target file already exists"
          End If

        Else
          WScript.Echo "ℹ Already clean, no rename needed"
        End If

      Else
        WScript.Echo "❌ .dat not found after marker"
      End If

    Else
      WScript.Echo "❌ Marker not found"
    End If

  Else
    WScript.Echo "⏭ Skipped (not .dat)"
  End If

Next

WScript.Echo "=== SCRIPT END ==="
