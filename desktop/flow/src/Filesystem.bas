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
