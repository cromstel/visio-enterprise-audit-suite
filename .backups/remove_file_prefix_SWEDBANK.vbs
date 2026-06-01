'==============================================================================
' MOVEit Backup Filename Prefix Removal Script (Silent)
'==============================================================================
' Purpose:      Removes the fixed prefix "G086241FI0910175S04919W0-" from filenames
'               in a specified folder
' Language:     VBScript (Windows Script Host)
' Execution:    Silent - no console output, no user interaction required
' Author:       Samuel Lamptey (cromstel@gmail.com)
' Created:      2026-05-30
' Last Updated:  2026-05-30
' Notes:        - Ensure the folderPath variable is set to the correct directory
'               - The script will exit silently if the folder does not exist
'==============================================================================

Option Explicit

'------------------------------------------------------------------------------
' CONFIGURATION
'------------------------------------------------------------------------------
' Description: Configure the target folder path. The prefix to remove is hardcoded
'              and always remains the same: G086241FI0910175S04919W0-

Dim folderPath
folderPath = "C:\Users\Cardif Nordic\Desktop\demo files"

' Fixed prefix that will be removed from all matching filenames
Const PREFIX_TO_REMOVE = "G086241FI0910175S04919W0-"

'------------------------------------------------------------------------------
' MAIN EXECUTION
'------------------------------------------------------------------------------

' Create FileSystemObject for file operations
Dim fso, folder

Set fso = CreateObject("Scripting.FileSystemObject")

' Validate target folder exists - exit silently if not found
If Not fso.FolderExists(folderPath) Then
    WScript.Quit
End If

Set folder = fso.GetFolder(folderPath)

'--------------------------------------------------------------------------
' PROCESS EACH FILE IN THE FOLDER
'--------------------------------------------------------------------------

Dim files, file, fileName, newName, fileNames, i

Set files = folder.Files

' Collect all filenames first to avoid modification during iteration issues
ReDim fileNames(files.Count)
i = 0
For Each file In files
    fileNames(i) = file.Name
    i = i + 1
Next

' Process collected filenames
For i = 0 To UBound(fileNames)
    fileName = fileNames(i)
    
    ' Check if filename starts with the prefix
    If Left(fileName, Len(PREFIX_TO_REMOVE)) = PREFIX_TO_REMOVE Then
        ' Calculate new filename by removing the prefix
        newName = Mid(fileName, Len(PREFIX_TO_REMOVE) + 1)
        
        ' Only rename if the name would actually change
        If newName <> fileName Then
            On Error Resume Next
            fso.GetFile(fso.BuildPath(folderPath, fileName)).Name = newName
            On Error GoTo 0
        End If
    End If
Next

Set files = Nothing

' Clean up objects
Set folder = Nothing
Set fso = Nothing

' Script completes - exits silently
WScript.Quit
