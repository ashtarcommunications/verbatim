Attribute VB_Name = "Filesystem"
Option Explicit

Public Function FileExists(ByVal FilePath As String) As Boolean
    Dim FSO As Object

    On Error GoTo Handler
    
    FileExists = False

    #If Mac Then
        If AppleScriptTask("Verbatim.scpt", "FileExists", Replace(FilePath, ".localized", vbNullString)) = "true" Then
            FileExists = True
        End If
        Set FSO = Nothing
    #Else
        'Use late binding to avoid needing an FSO reference
        Set FSO = CreateObject("Scripting.FileSystemObject")
        If FSO.FileExists(FileExists) = True Then
            FileExists = True
        End If
        Set FSO = Nothing
    #End If

    Exit Function

Handler:
    Set FSO = Nothing
    Application.StatusBar = "Error checking for file " & FilePath & " - Error " & Err.Number & ": " & Err.Description

End Function

Public Function FolderExists(ByVal FolderPath As String) As Boolean
    Dim FSO As Object

    On Error GoTo Handler
    
    FolderExists = False

    #If Mac Then
        If AppleScriptTask("Verbatim.scpt", "FolderExists", Replace(FolderExists, ".localized", vbNullString)) = "true" Then
            FolderExists = True
        End If
        Set FSO = Nothing
    #Else
        Set FSO = CreateObject("Scripting.FileSystemObject")
        If FSO.FolderExists(FolderPath) = True Then
            FolderExists = True
        End If
        Set FSO = Nothing
    #End If

    Exit Function

Handler:
    Set FSO = Nothing
    Application.StatusBar = "Error checking for folder " & FolderPath & " - Error " & Err.Number & ": " & Err.Description

End Function
