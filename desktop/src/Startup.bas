Attribute VB_Name = "Startup"
'@Folder("Startup")

Option Explicit

Public Sub AutoOpen()
    Startup.Start
End Sub

Public Sub AutoNew()
    On Error Resume Next
    
    ' Add doc variables with name and version number
    ThisDocument.Variables.Add Name:="Creator", Value:=GetSetting("Verbatim", "Main", "Name", vbNullString)
    ThisDocument.Variables.Add Name:="Team", Value:=GetSetting("Verbatim", "Main", "TeamName", vbNullString)
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
        If Globals.ActiveSpeechDoc = ActiveDocument.Name Then Globals.ActiveSpeechDoc = vbNullString
        
        ' Don't run if in Protected View in Windows
        #If Not Mac Then
        If Application.ActiveProtectedViewWindow Is Nothing Then
        #End If
            ' Check if current file is a .doc file instead of a .docx and default save settings
            If GetSetting("Verbatim", "Admin", "SuppressDocCheck", False) = False Then
                Troubleshooting.CheckDocx Notify:=True
                Troubleshooting.CheckSaveFormat Notify:=True
            End If
        
            ' If last doc, check if audio recording is still on
            If Application.Documents.Count = 1 And Globals.RecordAudioToggle = True Then
                If MsgBox("Audio recording appears to be active. Stop and save recording now? If you answer ""No"", recording will be lost.", vbYesNo) = vbYes Then Audio.SaveRecord
            End If
        #If Not Mac Then
        End If
        #End If
    End If
    
    On Error GoTo 0
End Sub

Public Sub Start()

    ' Set Mac global for easier conditionals
    ' TODO - is this necessary?
    #If Mac Then
        IsMac = True
    #Else
        IsMac = False
    #End If
    
    Globals.InitializeGlobals
    
    On Error Resume Next
    
    ' Don't run if in Protected View or file isn't visible on Windows
    #If Not Mac Then
    If Application.ActiveProtectedViewWindow Is Nothing And ActiveWindow.Visible = True Then
    #End If
        ' Set default view and navigation pane
        View.DefaultView
        ActiveWindow.DocumentMap = True
        
        ' Refresh document styles from template if setting checked and not editing template itself
        If GetSetting("Verbatim", "Format", "AutoUpdateStyles", True) = True And ActiveDocument.FullName <> ActiveDocument.AttachedTemplate.FullName Then ActiveDocument.UpdateStyles
        ActiveDocument.Saved = True
           
        ' Check for NPCStartup setting and call NavPaneCycle if True
        If GetSetting("Verbatim", "Admin", "NPCStartup", False) = True Then View.NavPaneCycle
        
        ' Refresh screen to solve blank screen bug
        Application.ScreenRefresh
    
        ' Check if it's the first run
        If GetSetting("Verbatim", "Admin", "FirstRun", True) = True Then
            Startup.FirstRun
        Else
            ' If first document opened and warnings not suppressed, check if template is incorrectly installed.
            If GetSetting("Verbatim", "Admin", "SuppressInstallChecks", False) = False And Application.Documents.Count = 1 Then
                If Troubleshooting.InstallCheckTemplateName = True Or Troubleshooting.InstallCheckTemplateLocation = True Then
                    If MsgBox("Verbatim appears to be installed incorrectly. Would you like to open the Troubleshooter? This message can be suppressed in the Verbatim settings.", vbYesNo) = vbYes Then
                        UI.ShowForm "Settings"
                        Exit Sub
                    End If
                End If
            End If
            
            ' Check for updates weekly on Wednesdays
            If GetSetting("Verbatim", "Admin", "AutoUpdateCheck", True) = True Then
                If DateDiff("d", GetSetting("Verbatim", "Main", "LastUpdateCheck"), Now) > 6 Then
                    If DatePart("w", Now) = 4 Then
                        Settings.UpdateCheck
                        Exit Sub
                    End If
                End If
            End If
            
        End If

        ' Check for custom code to import
        If GetSetting("Verbatim", "Main", "ImportCustomCode", False) = True Then
            Settings.ImportCustomCode Notify:=True
        End If
    #If Not Mac Then
    End If
    #End If

    On Error GoTo 0
End Sub

Public Sub FirstRun()
    ' Set FirstRun to False for future
    SaveSetting "Verbatim", "Admin", "FirstRun", False
    
    ' Unverbatimize Normal to clear out old installs
    ' TODO - this won't work if no VBOM access
    Settings.UnverbatimizeNormal
    
    ' Remove old registry keys
    DeleteSetting "Verbatim", "Main", "TabroomUsername"
    DeleteSetting "Verbatim", "Main", "TabroomPassword"
    DeleteSetting "Verbatim", "Main", "GmailUsername"
    DeleteSetting "Verbatim", "Main", "GmailPassword"
    
    ' Setup keyboard shortcuts (includes tilde fix)
    Settings.ResetKeyboardShortcuts

    ' Run setup wizard
    Settings.ShowSetupWizard
End Sub

