Attribute VB_Name = "Filesystem"
Option Explicit

Public Function FileExists(ByVal FilePath As String) As Boolean
    On Error GoTo Handler
    
    FileExists = False

    #If Mac Then
        Dim Script
        Script = "if [ -f '" & FilePath & "' ]; then echo 1; else echo 0; fi;"
        If AppleScriptTask("Verbatim.scpt", "RunShellScript", Script) = "1" Then
            FileExists = True
        Else
            Script = "if [ -f '" & Replace(FilePath, ".localized", "") & "' ]; then echo 1; else echo 0; fi;"
            If AppleScriptTask("Verbatim.scpt", "RunShellScript", Script) = "1" Then
                FileExists = True
            End If
        End If
    #Else
        Dim FSO As Object
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
        Dim Script
        Script = "if [ -d '" & Replace(FolderPath, ".localized", "") & "' ]; then echo 1; else echo 0; fi;"
        If AppleScriptTask("Verbatim.scpt", "RunShellScript", Script) = "1" Then
            FolderExists = True
        End If
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

'@Ignore ProcedureNotUsed
'@Ignore ParameterNotUsed
Public Function GetSubfoldersInFolder(ByVal FolderPath As String) As String
    #If Mac Then
        GetSubfoldersInFolder = AppleScriptTask("Verbatim.scpt", "GetSubfoldersInFolder", FolderPath)
        
        ' Trim trailing newline
        If Right(GetSubfoldersInFolder, 1) = Chr(10) Or Right(GetSubfoldersInFolder, 1) = Chr(13) Then GetSubfoldersInFolder = Left(GetSubfoldersInFolder, Len(GetSubfoldersInFolder) - 1)
    #Else
        GetSubfoldersInFolder = ""
        Exit Function
    #End If
End Function

'@Ignore ProcedureNotUsed
'@Ignore ParameterNotUsed
Public Function GetFilesInFolder(ByVal FolderPath As String) As String
    #If Mac Then
        GetFilesInFolder = AppleScriptTask("Verbatim.scpt", "GetFilesInFolder", FolderPath)
        
        ' Trim trailing newline
        If Right(GetFilesInFolder, 1) = Chr(10) Or Right(GetFilesInFolder, 1) = Chr(13) Then GetFilesInFolder = Left(GetFilesInFolder, Len(GetFilesInFolder) - 1)
    #Else
        GetFilesInFolder = ""
        Exit Function
    #End If
End Function

Public Sub DeleteFile(ByRef Path As String)
    On Error Resume Next
    
    #If Mac Then
        ' Built-in Mac Kill doesn't work with filenames over 28 characters
        AppleScriptTask "Verbatim.scpt", "KillFileOnMac", Path
    #Else
        Kill Path
    #End If

    On Error GoTo 0
End Sub

'@Ignore ProcedureNotUsed
Public Sub DeleteFolder(ByRef Path As String)
    On Error Resume Next
    
    #If Mac Then
        AppleScriptTask "Verbatim.scpt", "KillFolderOnMac", Path
    #Else
        Dim FSO As Object
        Set FSO = CreateObject("Scripting.FileSystemObject")
        FSO.DeleteFolder Path
        Set FSO = Nothing
    #End If
    
    On Error GoTo 0
End Sub

Public Function GetFile(ByRef Path As String) As clsFile
    Set GetFile = New clsFile
    '@Ignore FunctionReturnValueDiscarded
    GetFile.Init Path
End Function

Public Function GetFolder(ByRef Path As String) As clsFolder
    Set GetFolder = New clsFolder
    '@Ignore FunctionReturnValueDiscarded
    GetFolder.Init Path
End Function

Public Function ReadFile(ByVal Path As String) As String
    Dim i As Long
    i = FreeFile
    Open Path For Input As FreeFile
    ReadFile = Input$(LOF(i), i)
    Close i
End Function

Public Sub CopyFile(ByRef Path As String, ByRef NewPath As String)
    On Error GoTo Handler

    #If Mac Then
        Dim Script
        Script = "cp " & Path & " " & NewPath
        AppleScriptTask "Verbatim.scpt", "RunShellScript", Script
    #Else
        Dim FSO As Object
        Set FSO = CreateObject("Scripting.FileSystemObject")
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

Public Function GetFileAsBase64(ByRef Path As String) As String
    On Error GoTo Handler

    #If Mac Then
        Dim Script
        Script = "base64 " & Path
        GetFileAsBase64 = AppleScriptTask("Verbatim.scpt", "RunShellScript", Script)
    #Else
        Dim FileStream As Object
        
        Dim xmlDoc As Object
        Set xmlDoc = CreateObject("Microsoft.XMLDOM")
        
        Dim xmlElem As Object
        Set xmlElem = xmlDoc.createElement("tmp")
                
        ' Create a temporary copy of the current file to upload
        Filesystem.CopyFile Path, Path & ".base64"
        
        ' Create FileStream
        Set FileStream = CreateObject("ADODB.Stream")
        FileStream.Open
        FileStream.Type = 1 'adTypeBinary
        FileStream.LoadFromFile Filename:=Path & ".base64"
           
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

#If Mac Then
Function RequestFolderAccess(RootPath) As Boolean
    Dim Script As String
    Dim Files
    Dim i As Long
    
    On Error Resume Next
    
    ' Get an array of all files in the root path or subfolders to request permission
    Script = "find '" & RootPath & "' -type f"
    Files = Split(AppleScriptTask("Verbatim.scpt", "RunShellScript", Script), Chr(13))
    For i = 0 To UBound(Files)
        Files(i) = Replace(Files(i), "//", "/")
    Next i
    
    RequestFolderAccess = GrantAccessToMultipleFiles(Files)
End Function
#End If
