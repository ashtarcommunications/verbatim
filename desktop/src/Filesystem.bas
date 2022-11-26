Attribute VB_Name = "Filesystem"
Option Explicit

Public Function FileExists(ByVal FilePath As String) As Boolean
    On Error GoTo Handler
    
    FileExists = False

    #If Mac Then
        If AppleScriptTask("Verbatim.scpt", "FileExists", Replace(FilePath, ".localized", vbNullString)) = "true" Then
            FileExists = True
        End If
        Set FSO = Nothing
    #Else
        Dim FSO As Object
        ' Use late binding to avoid needing an FSO reference
        Set FSO = CreateObject("Scripting.FileSystemObject")
        If FSO.FileExists(FilePath) = True Then
            FileExists = True
        End If
        Set FSO = Nothing
    #End If

    Exit Function

Handler:
    #If Mac Then
        ' Do Nothing
    #Else
        Set FSO = Nothing
    #End If
    Application.StatusBar = "Error checking for file " & FilePath & " - Error " & Err.Number & ": " & Err.Description
End Function

Public Function FolderExists(ByVal FolderPath As String) As Boolean
    On Error GoTo Handler
    
    FolderExists = False

    #If Mac Then
        If AppleScriptTask("Verbatim.scpt", "FolderExists", Replace(FolderExists, ".localized", vbNullString)) = "true" Then
            FolderExists = True
        End If
        Set FSO = Nothing
    #Else
        Dim FSO As Object
        Set FSO = CreateObject("Scripting.FileSystemObject")
        If FSO.FolderExists(FolderPath) = True Then
            FolderExists = True
        End If
        Set FSO = Nothing
    #End If

    Exit Function

Handler:
    #If Mac Then
        ' Do nothing
    #Else
        Set FSO = Nothing
    #End If
    Application.StatusBar = "Error checking for folder " & FolderPath & " - Error " & Err.Number & ": " & Err.Description
End Function

Function GetSubfoldersInFolder(FolderPath) As String
    #If Mac Then
        Dim Script As String
        
        Script = "tell application ""Finder""" & Chr(13)
        Script = Script & "set r to """"" & Chr(13)
        Script = Script & "set myFolders to folders of folder""" & FolderPath & """" & Chr(13)
        Script = Script & "repeat with f in myFolders" & Chr(13)
        Script = Script & "set r to (r & f as string) & ""\n""" & Chr(13)
        Script = Script & "end repeat" & Chr(13)
        Script = Script & "return r" & Chr(13)
        Script = Script & "end tell"
        
        GetSubfoldersInFolder = AppleScriptTask("Verbatim.scpt", "GetSubfoldersInFolder", FolderPath)
        
        ' Trim trailing newline
        If Right(GetSubfoldersInFolder, 1) = Chr(10) Or Right(GetSubfoldersInFolder, 1) = Chr(13) Then GetSubfoldersInFolder = Left(GetSubfoldersInFolder, Len(GetSubfoldersInFolder) - 1)
    #Else
        Exit Sub
    #End If
End Function
Function GetFilesInFolder(FolderPath) As String
    #If Mac Then
        Dim POSIXPath As String
        Dim Script As String
        
        POSIXPath = MacScript("tell text 1 thru -2 of " & Chr(34) & FolderPath & Chr(34) & " to return quoted form of it's POSIX Path")
        
        Script = "set streamEditorCommand to " & Chr(34) & " |  tr  [/:] [:/] " & Chr(34) & Chr(13)
        Script = Script & "set streamEditorCommand to streamEditorCommand & " & Chr(34)
        Script = Script & " | sed -e " & Chr(34) & "  & quoted form of (" & Chr(34) & " s.:." & Chr(34)
        Script = Script & "  & (POSIX file " & Chr(34) & "/" & Chr(34) & "  as string) & " & Chr(34) & "." & Chr(34) & " )" & Chr(13)
        Script = Script & "do shell script """ & "find -E " & POSIXPath
        Script = Script & " -iregex " & "'.*/[^~][^/]*\\." & "(docx|doc|docm|dot|dotm)" & "$' " & "-maxdepth 1"
        Script = Script & """ & streamEditorCommand without altering line endings"
    
        GetFilesInFolder = AppleScriptTask("Verbatim.scpt", "GetFilesInFolder", POSIXPath)
        
        ' Trim trailing newline
        If Right(GetFilesInFolder, 1) = Chr(10) Or Right(GetFilesInFolder, 1) = Chr(13) Then GetFilesInFolder = Left(GetFilesInFolder, Len(GetFilesInFolder) - 1)
    #Else
        Exit Sub
    #End If
End Function

Sub DeleteFile(Path As String)
    On Error Resume Next
    
    #If Mac Then
        ' Built-in Mac Kill doesn't work with filenames over 28 characters
        AppleScriptTask "Verbatim.scpt", "KillFileOnMac", Path
    #Else
        Kill Path
    #End If
End Sub

Sub DeleteFolder(Path As String)
    On Error Resume Next
    
    #If Mac Then
        AppleScriptTask "Verbatim.scpt", "KillFolderOnMac", Path
    #Else
        Dim FSO As FileSystemObject
        Set FSO = New FileSystemObject
        FSO.DeleteFolder Path
        Set FSO = Nothing
    #End If
End Sub

Public Function GetFile(Path As String) As clsFile
    Set GetFile = New clsFile
    GetFile.Init Path
End Function

Public Function GetFolder(Path As String) As clsFolder
    Set GetFolder = New clsFolder
    GetFolder.Init Path
End Function

Public Function ReadFile(Path As String) As String
    Dim i As Integer
    i = FreeFile
    Open Path For Input As FreeFile
    ReadFile = Input(LOF(i), i)
    Close i
End Function

Public Sub CopyFile(Path As String, NewPath As String)
    On Error GoTo Handler

    #If Mac Then
        Dim Script
        Script = "cp " & Path & " " & NewPath
        AppleScriptTask "Verbatim.scpt", "DoShellScript", Script
    #Else
        Dim FSO As FileSystemObject
        Set FSO = New FileSystemObject
        FSO.CopyFile Path, NewPath
    #End If

    Exit Sub

Handler:
    #If Mac Then
        ' Do Nothing
    #Else
        Set FSO = Nothing
    #End If
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Public Function GetFileAsBase64(Path As String) As String
    On Error GoTo Handler

    #If Mac Then
        Dim Script
        Script = "base64 " & Path
        GetFileAsBase64 = AppleScriptTask("Verbatim.scpt", "DoShellScript", Script)
    #Else
        Dim FileStream As ADODB.Stream
        Dim TempFileName As String
        
        Dim xmlDoc
        Set xmlDoc = CreateObject("Microsoft.XMLDOM")
        
        Dim xmlElem
        Set xmlElem = xmlDoc.createElement("tmp")
                
        ' Create a temporary copy of the current file to upload
        Filesystem.CopyFile Path, Path & ".base64"
        
        ' Create FileStream
        Set FileStream = New ADODB.Stream
        FileStream.Open
        FileStream.Type = adTypeBinary
        FileStream.LoadFromFile FileName:=Path & ".base64"
           
        ' Convert to Base64
        xmlElem.dataType = "bin.base64"
        xmlElem.nodeTypedValue = FileStream.Read
        FileStream.Close
        
        GetFileAsBase64 = Replace(xmlElem.Text, vbLf, "")
            
        Set FileStream = Nothing
        Set xmlDoc = Nothing
        Set xmlElem = Nothing
        
        Filesystem.DeleteFile Path & ".base64"
    #End If
    
    Exit Function
    
Handler:
    #If Mac Then
        ' Do Nothing
    #Else
        FileStream.Close
        Set FileStream = Nothing
        Set xmlDoc = Nothing
        Set xmlElem = Nothing
    #End If
    
    Filesystem.DeleteFile Path & ".base64"
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Function
