Attribute VB_Name = "Plugins"
Option Explicit

Public Sub StartTimer()
' Starts a user supplied timer.
    On Error GoTo Handler
    Dim TimerPath As String
    
    #If Mac Then
        Dim TimerPOSIX As String
        
        On Error GoTo Handler
        
        ' Get path to timer app
        TimerPath = GetSetting("Verbatim", "Plugins", "TimerPath", "?")
        
        ' If not set, try default
        If TimerPath = "?" Then TimerPath = "/Applications/VerbatimTimer.app"
    
        ' Make sure timer app exists (check both folder and file since .app is actually a folder)
        If Filesystem.FolderExists(TimerPath) = False And Filesystem.FileExists(TimerPath) = False Then
            MsgBox "Timer application not found. Ensure you have installed the Verbatim Timer or entered a custom path to another application in the Verbatim Settings."
            Exit Sub
        Else
            ' If java app selected, run it from the shell
            If Right(TimerPath, 5) = ".jar:" Or Right(TimerPath, 4) = ".jar" Then
                AppleScriptTask "Verbatim.scpt", "RunShellScript", "open '" & TimerPath & "'"
            Else
                AppleScriptTask "Verbatim.scpt", "ActivateTimer", TimerPath
            End If
        End If
    
        Exit Sub
    #Else
        TimerPath = GetSetting("Verbatim", "Plugins", "TimerPath", "")
        If TimerPath = "" Then
            TimerPath = Environ$("ProgramW6432") & Application.PathSeparator & "Verbatim" & Application.PathSeparator & "Plugins" & Application.PathSeparator & "VerbatimTimer.exe"
        End If
        
        ' Check timer exists
        If Filesystem.FileExists(TimerPath) = False Then
            MsgBox "Timer application not found. Ensure you have installed the Verbatim Timer or entered a custom path to another application in the Verbatim Settings."
            Exit Sub
        End If
        
        ' Run Timer
        Shell TimerPath, vbNormalFocus
       
        Exit Sub
    #End If
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Public Sub GetFromCiteCreator()
    On Error GoTo Handler
    
    #If Mac Then
        AppleScriptTask "Verbatim.scpt", "GetFromCiteCreator", ""
        Formatting.PasteText
        Exit Sub
    #Else
        Dim CiteCreatorPath As String
        
        On Error GoTo Handler
        
        ' Check GetFromCiteCreator script exists
        CiteCreatorPath = Environ$("ProgramW6432") & Application.PathSeparator & "Verbatim" & Application.PathSeparator & "Plugins" & Application.PathSeparator & "GetFromCiteCreator.exe"
        If Filesystem.FileExists(CiteCreatorPath) = False Then
            MsgBox "The GetFromCiteCreator plugin does not appear to be installed. Check https://paperlessdebate.com for more information on how to install."
            Exit Sub
        End If
        
        ' Run the script
        Shell CiteCreatorPath, vbMinimizedNoFocus
                
        Exit Sub
    #End If
    
Handler:
    MsgBox "Getting from Cite Creator failed - ensure Google Chrome and the Cite Creator extension are installed and open." & vbCrLf & vbCrLf & "Error " & Err.Number & ": " & Err.Description
End Sub

Public Sub NavPaneCycle()
' Runs the NavPaneCycle program via Shell
    #If Mac Then
        MsgBox "NavPaneCycle is only supported on Windows"
        Exit Sub
    #Else
        On Error Resume Next
    
        Dim NPCPath As String
        NPCPath = Environ$("ProgramW6432") & Application.PathSeparator & "Verbatim" & Application.PathSeparator & "Plugins" & Application.PathSeparator & "NavPaneCycle.exe"
    
        ' Check NPC exists
        If Filesystem.FileExists(NPCPath) = False Then
            MsgBox "The NavPaneCycle plugin does not appear to be installed. Check https://paperlessdebate.com for more information on how to install."
            Exit Sub
        End If

        ' Make sure NavPane is showing
        If ActiveWindow.DocumentMap = False Then Exit Sub
    
        ' Don't run if window is invisible
        If ActiveWindow.Visible = False Then Exit Sub
    
        ' Run NPC
        Shell NPCPath, vbMinimizedNoFocus
    
        On Error GoTo 0
    
        Exit Sub
    #End If
End Sub
