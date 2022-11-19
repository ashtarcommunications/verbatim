Attribute VB_Name = "Plugins"
Option Explicit

Public Sub InstallPlugin(Plugin As String)
    ' Turn on error checking
    On Error GoTo Handler

    Application.StatusBar = "Installing " & Plugin & " plugin..."
    
    Dim TempFile As String
    TempFile = Environ("TEMP") & Application.PathSeparator & "verbatim-plugin.zip"
    
    Dim DefaultPluginsFolder As String
    DefaultPluginsFolder = Environ("AppData") & Application.PathSeparator & "Verbatim" & Application.PathSeparator & "Plugins"
    
    Dim DestinationFolder As String
    DestinationFolder = GetSetting("Verbatim", "Main", "VerbatimPluginsPath", DefaultPluginsFolder & Application.PathSeparator & Plugin)

    'HTTP.DownloadFile "https://paperlessdebate.com/plugins/" & Plugin & ".zip", TempFile
    HTTP.DownloadFile "https://www.learningcontainer.com/wp-content/uploads/2020/05/sample-zip-file.zip", TempFile

    #If Mac Then
        ' TODO - figure out how to unzip on mac
    #Else
        ' Expand-Archive requires Powershell 5+ on Windows 10+
        CreateObject("WScript.Shell").Run "powershell -command Expand-Archive " & TempFile & " -DestinationPath " & DestinationFolder, 0, True
    #End If
    
    Filesystem.DeleteFile TempFile
    
    Exit Sub

Handler:
    MsgBox "Plugin Installation Failed. Error " & Err.Number & ": " & Err.Description
End Sub

Public Sub UninstallPlugin(Plugin As String)
    Dim DefaultPluginsFolder As String
    DefaultPluginsFolder = Environ("AppData") & Application.PathSeparator & "Verbatim" & Application.PathSeparator & "Plugins"
    Filesystem.DeleteFolder DefaultPluginsFolder & Application.PathSeparator & Plugin
End Sub
