Attribute VB_Name = "Startup"
Option Explicit

Public Sub AutoOpen()
    Startup.Start
End Sub

Public Sub AutoNew()
    On Error Resume Next
    
    ' Add doc variables with name and version number
    ThisDocument.Variables.Add Name:="Creator", Value:=GetSetting("Verbatim", "Profile", "Name", "")
    ThisDocument.Variables.Add Name:="Team", Value:=GetSetting("Verbatim", "Profile", "SchoolName", "")
    ThisDocument.Variables.Add Name:="VerbatimVersion", Value:=Settings.GetVersion
    ThisDocument.Variables.Add Name:="OS", Value:=Application.System.OperatingSystem
    ThisDocument.Variables.Add Name:="OSVersion", Value:=Application.System.Version
    ThisDocument.Variables.Add Name:="WordVersion", Value:=Application.Version
    ThisDocument.Saved = True
    
    Startup.Start
       
    On Error GoTo 0
End Sub

Public Sub AutoClose()
    On Error Resume Next
    
    If ActiveWindow.Visible = True Then
        ' If current doc was active speech doc, clear it
        If Globals.ActiveSpeechDoc = ActiveDocument.Name Then Globals.ActiveSpeechDoc = ""
        
        #If Mac Then
            ' Do nothing
        #Else
            ' Don't run if in Protected View in Windows
            If Not Application.ActiveProtectedViewWindow Is Nothing Then Exit Sub
        #End If
        
        ' Check if current file is a .doc file instead of a .docx and default save settings
        If GetSetting("Verbatim", "Admin", "SuppressDocCheck", False) = False Then
            '@Ignore FunctionReturnValueDiscarded
            Troubleshooting.CheckDocx Notify:=True
            '@Ignore FunctionReturnValueDiscarded
            Troubleshooting.CheckSaveFormat Notify:=True
        End If
    
        ' If last doc, check if audio recording is still on
        If Application.Documents.Count = 1 And Globals.RecordAudioToggle = True Then
            If MsgBox("Audio recording appears to be active. Stop and save recording now? If you answer ""No"", recording will be lost.", vbYesNo) = vbYes Then Audio.SaveRecord
        End If
    End If
    
    On Error GoTo 0
End Sub

Public Sub Start()
    On Error Resume Next
    
    Globals.InitializeGlobals
    
    ' Don't run if in Protected View or file isn't visible on Windows
    #If Mac Then
        ' Do Nothing
    #Else
        If Not Application.ActiveProtectedViewWindow Is Nothing Or ActiveWindow.Visible = False Then Exit Sub
    #End If
    
    ' Set default view and navigation pane
    View.DefaultView
    ActiveWindow.DocumentMap = True
    
    ' Refresh document styles from template if setting checked and not editing template itself
    If GetSetting("Verbatim", "Admin", "AutoUpdateStyles", True) = True And ActiveDocument.FullName <> ActiveDocument.AttachedTemplate.FullName Then ActiveDocument.UpdateStyles
    ActiveDocument.Saved = True
       
    ' Prevent Word making new styles
    If GetSetting("Verbatim", "Admin", "SuppressStyleChecks", False) = False Then
        Application.RestrictLinkedStyles = True
        Options.AutoFormatAsYouTypeDefineStyles = False
    End If
       
    ' Check for NPCStartup setting and call NavPaneCycle if True
    If GetSetting("Verbatim", "View", "NPCStartup", False) = True Then Plugins.NavPaneCycle
    
    ' Refresh screen to solve blank screen bug
    Application.ScreenRefresh

    ' Check if it's the first run
    If GetSetting("Verbatim", "Admin", "FirstRun", True) = True Then
        Startup.FirstRun
    Else
        ' If first document opened and warnings not suppressed, check if template is incorrectly installed
        ' Can't check for script file installation here because running AppleScriptTask during startup prevents the file opening
        If GetSetting("Verbatim", "Admin", "SuppressInstallChecks", False) = False And Application.Documents.Count = 1 Then
            If Troubleshooting.InstallCheckTemplateName = False Or Troubleshooting.InstallCheckTemplateLocation = False Then
                If MsgBox("Verbatim appears to be installed incorrectly. Would you like to open the Troubleshooter? This message can be suppressed in the Verbatim settings.", vbYesNo) = vbYes Then
                    UI.ShowForm "Troubleshooter"
                    Exit Sub
                End If
            End If
        End If
        
        ' Check for updates weekly on Wednesdays
        If GetSetting("Verbatim", "Profile", "AutomaticUpdates", True) = True Then
            If DateDiff("d", GetSetting("Verbatim", "Profile", "LastUpdateCheck"), Now) > 6 Then
                If DatePart("w", Now) = 4 Then
                    Settings.UpdateCheck
                    Exit Sub
                End If
            End If
        End If
    End If

    ' Check for custom code to import
    If GetSetting("Verbatim", "Admin", "ImportCustomCode", False) = True Then
        Settings.ImportCustomCode Notify:=True
    End If
    
    ' Reset keybindings on Mac if PC shortcuts are set
    #If Mac Then
        Dim DebateTemplate As Document
        Set DebateTemplate = ActiveDocument.AttachedTemplate.OpenAsDocument
        If Application.CustomizationContext <> "Debate.dotm" Then Application.CustomizationContext = ActiveDocument.AttachedTemplate
        Dim k As KeyBinding
        For Each k In KeyBindings
            If k.Command = "Verbatim.Formatting.PasteText" Then
                If k.KeyString = "Shift+2" Then
                    Settings.ResetKeyboardShortcuts
                    Exit For
                End If
            End If
        Next k

        '@Ignore MemberNotOnInterface
        DebateTemplate.Close SaveChanges:=True
        Set DebateTemplate = Nothing
        
        '@Ignore ValueRequired
        Application.CustomizationContext = ThisDocument
    #End If

    On Error GoTo 0
End Sub

Public Sub FirstRun()
    ' Set FirstRun to False for future
    SaveSetting "Verbatim", "Admin", "FirstRun", False
    
    ' Unverbatimize Normal to clear out old installs
    Settings.UnverbatimizeNormal
    
    ' Remove old registry keys
    SaveSetting "Verbatim", "Main", "TabroomUsername", ""
    SaveSetting "Verbatim", "Main", "TabroomPassword", ""
    SaveSetting "Verbatim", "Main", "GmailUsername", ""
    SaveSetting "Verbatim", "Main", "GmailPassword", ""
    
    ' Setup keyboard shortcuts (includes tilde fix)
    Settings.ResetKeyboardShortcuts

    ' Run setup wizard
    UI.ShowForm "Setup"
End Sub
