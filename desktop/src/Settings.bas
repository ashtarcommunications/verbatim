Attribute VB_Name = "Settings"
Option Explicit

Public Sub UnverbatimizeNormal(Optional ByVal Notify As Boolean)
' Deprecated except to uninstall old versions
       
    ' Bail without VBOM access
    #If Mac Then
    #Else
        If Registry.RegKeyRead("HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version & "\Word\Security\AccessVBOM") <> 1 Then
            If Notify = True Then
                MsgBox "You must enable VBOM access in your macro security settings to unverbatimize your Normal template.", vbOKOnly
            End If
            Exit Sub
        End If
    #End If
       
    ' Delete module from normal template - turn off error checking in case it doesn't exist
    On Error Resume Next
    Application.OrganizerDelete source:=Application.NormalTemplate.FullName, Name:="AttachVerbatim", Object:=wdOrganizerObjectProjectItems

    ' Delete CustomUI if it exists
    #If Mac Then
        ' Do Nothing
    #Else
        On Error GoTo Handler
        If Filesystem.FileExists(CStr(Environ$("USERPROFILE")) & "\AppData\Local\Microsoft\Office\Word.officeUI") = True Then
            Kill CStr(Environ$("USERPROFILE")) & "\AppData\Local\Microsoft\Office\Word.officeUI"
        End If
    #End If
    
    If Notify = True Then MsgBox "Normal template successfully un-verbatimized!"
    
    Exit Sub
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

'*************************************************************************************
'* IMPORT/EXPORT FUNCTIONS                                                           *
'*************************************************************************************

Public Sub ImportCustomCode(Optional ByVal Notify As Boolean)
    Dim p As Object

    ' Turn on Error Handling
    On Error GoTo Handler

    ' Set registry setting to avoid repeatedly trying to import code
    SaveSetting "Verbatim", "Admin", "ImportCustomCode", False

    ' Check if Access to VBOM allowed
    #If Mac Then
    #Else
        If Registry.RegKeyRead("HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version & "\Word\Security\AccessVBOM") <> 1 Then
            If Notify = True Then MsgBox "Importing custom code requires you to enable ""Trust Access to the VBA project object model"" in your Macro security settings. You can do this manually, or run the Verbatim troubleshooter."
            Exit Sub
        End If
    #End If

    ' Make sure custom code file exists
    #If Mac Then
        If Filesystem.FileExists(ActiveDocument.AttachedTemplate.Path & Application.PathSeparator & "VerbatimCustomCode.bas") = False Then
    #Else
        If Filesystem.FileExists(Application.NormalTemplate.Path & Application.PathSeparator & "VerbatimCustomCode.bas") = False Then
    #End If
            If Notify = True Then MsgBox "No custom code module found in your Templates folder. It must be named ""VerbatimCustomCode.bas"" to import."
            Exit Sub
        End If
    
    ' Warn user
    If MsgBox("Attemping to import custom code - this will overwrite your current custom code module. Proceed?", vbOKCancel) = vbCancel Then Exit Sub
    
    ' Delete current Custom code module - turn off error checking temporarily in case it doesn't exist
    On Error Resume Next
    Application.OrganizerDelete source:=ActiveDocument.AttachedTemplate.FullName, Name:="Custom", Object:=wdOrganizerObjectProjectItems
    On Error GoTo Handler
    
    ' Import the module and delete the file
    #If Mac Then
        Set p = FindVBProject(ActiveDocument.AttachedTemplate.Path & Application.PathSeparator & ActiveDocument.AttachedTemplate)
    #Else
        Set p = ActiveDocument.AttachedTemplate.VBProject
    #End If
    If p Is Nothing Then
        MsgBox "Failed to import custom code."
        Exit Sub
    End If
    p.VBComponents.Import (Application.NormalTemplate.Path & Application.PathSeparator & "VerbatimCustomCode.bas")
    Filesystem.DeleteFile Application.NormalTemplate.Path & Application.PathSeparator & "VerbatimCustomCode.bas"
    
    If Notify = True Then MsgBox "Custom code successfully imported!"

    Set p = Nothing

    Exit Sub

Handler:
    Set p = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Public Sub ExportCustomCode(Optional ByVal Notify As Boolean)
    Dim Module As Object
    
    'Turn on Error Handling
    On Error GoTo Handler
  
    #If Mac Then
        Dim p As Object
        Set p = FindVBProject(ActiveDocument.AttachedTemplate.Path & Application.PathSeparator & ActiveDocument.AttachedTemplate)
        If p Is Nothing Then
            MsgBox "Failed to find the Custom code module."
            Exit Sub
        End If
    
        Set Module = p.VBComponents("Custom")
        If Module.CodeModule.CountOfLines <= 1 Then
            If Notify = True Then MsgBox "No custom code found."
            Exit Sub
        End If
    #Else
        If Registry.RegKeyRead("HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version & "\Word\Security\AccessVBOM") <> 1 Then
            If Notify = True Then MsgBox "Exporting custom code requires you to enable ""Trust Access to the VBA project object model"" in your Macro security settings. You can do this manually, or run the Verbatim troubleshooter."
            Exit Sub
        End If
    #End If

    'Export the Custom module
    Set Module = ActiveDocument.AttachedTemplate.VBProject.VBComponents("Custom")
    If Module.CodeModule.CountOfLines <= 1 Then
        If Notify = True Then MsgBox "No custom code found."
        Exit Sub
    End If
    
    Module.Export Application.NormalTemplate.Path & Application.PathSeparator & "VerbatimCustomCode.bas"
     
    'Set registry for automatic import on startup
    SaveSetting "Verbatim", "Admin", "ImportCustomCode", True
    
    If Notify = True Then MsgBox "Custom code exported as VerbatimCustomCode.bas to your Templates folder."
   
    #If Mac Then
        Set p = Nothing
    #End If
    Set Module = Nothing
    
    Exit Sub

Handler:
    #If Mac Then
        Set p = Nothing
    #End If
    Set Module = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

#If Mac Then
Private Function FindVBProject(d As String) As Object
    
    Dim p As Object
    
    On Error Resume Next
    
    For Each p In Application.VBE.VBProjects
        If (p.Filename = d) Then
            Set FindVBProject = p
            Exit Function
        End If
    Next
    
End Function
#End If

' *************************************************************************************
' * UPDATE FUNCTIONS                                                                  *
' *************************************************************************************

Public Sub UpdateCheck(Optional ByVal Notify As Boolean)
    On Error GoTo Handler

    Application.StatusBar = "Checking for Verbatim updates..."

    ' Create and send HttpReq
    Dim Response As Dictionary
    Set Response = HTTP.GetReq(Globals.UPDATES_URL)
    
    ' Exit if the request fails
    If Response.Item("status") <> 200 Then
        Application.StatusBar = "Update Check Failed"
        SaveSetting "Verbatim", "Profile", "LastUpdateCheck", Now
        If Notify = True Then MsgBox "Update Check Failed."
        Exit Sub
    End If
    
    ' Set LastUpdateCheck
    SaveSetting "Verbatim", "Profile", "LastUpdateCheck", Now
    
    ' If newer version is found
    Dim UpdatedVersion As String
    UpdatedVersion = Response.Item("body")("verbatim")("latest")("desktop")
    
    If Settings.NewerVersion(UpdatedVersion, Settings.GetVersion) Then
    
        ' Prompt to launch website - no longer download automatically because it tends to trip antivirus heuristics
        If MsgBox("There is a newer version of Verbatim available. Download now?", vbYesNo) = vbNo Then Exit Sub
            
        Settings.LaunchWebsite Globals.PAPERLESSDEBATE_URL
    Else
        Application.StatusBar = "No Verbatim updates found."
        If Notify = True Then MsgBox "No Verbatim updates found."
    End If
         
    Set Response = Nothing
    Exit Sub

Handler:
    Set Response = Nothing
    Application.StatusBar = "Update Check Failed. Error " & Err.Number & ": " & Err.Description
    If Notify = True Then MsgBox "Update Check Failed. Error " & Err.Number & ": " & Err.Description

End Sub

Public Function NewerVersion(ByVal Version1 As String, ByVal Version2 As String) As Boolean
' Adapted from https://forum.ozgrid.com/forum/index.php?thread%2F52830-compare-version-number-strings%2F=
' Returns true if Version1 is newer
    Dim i As Long
    Dim Version1Array() As String
    Dim Version2Array() As String
    Version1Array = Split(Version1, ".")
    Version2Array = Split(Version2, ".")
    Dim k As Long
    
    k = UBound(Version1Array)
    If UBound(Version2Array) < k Then k = UBound(Version2Array)
    
    For i = 0 To k
        If Version1Array(i) > Version2Array(i) Then
            NewerVersion = True
            Exit For
        ElseIf Version1Array(i) < Version2Array(i) Then
            NewerVersion = False
            Exit For
        Else
            If UBound(Version1Array) = UBound(Version2Array) Then
                NewerVersion = False
            ElseIf UBound(Version1Array) > UBound(Version2Array) Then
                NewerVersion = True
            Else
                NewerVersion = False
            End If
        End If
    Next i
End Function

' *************************************************************************************
' * KEYBOARD FUNCTIONS                                                                  *
' *************************************************************************************

Public Sub ChangeKeyboardShortcut(ByVal KeyName As WdKey, ByVal MacroName As String)
    ' Change keyboard shortcuts in template
    '@Ignore ImplicitUnboundDefaultMemberAccess
    If Application.CustomizationContext <> "Debate.dotm" Then Application.CustomizationContext = ActiveDocument.AttachedTemplate
    
    Select Case MacroName
        Case Is = "Paste"
            KeyBindings.Add wdKeyCategoryMacro, "Formatting.PasteText", BuildKeyCode(KeyName)
        Case Is = "Condense"
            KeyBindings.Add wdKeyCategoryMacro, "Condense.CondenseAllOrCard", BuildKeyCode(KeyName)
        Case Is = "Pocket"
            KeyBindings.Add wdKeyCategoryStyle, "Pocket", BuildKeyCode(KeyName)
        Case Is = "Hat"
            KeyBindings.Add wdKeyCategoryStyle, "Hat", BuildKeyCode(KeyName)
        Case Is = "Block"
            KeyBindings.Add wdKeyCategoryStyle, "Block", BuildKeyCode(KeyName)
        Case Is = "Tag"
            KeyBindings.Add wdKeyCategoryStyle, "Tag", BuildKeyCode(KeyName)
        Case Is = "Cite"
            KeyBindings.Add wdKeyCategoryStyle, "Cite", BuildKeyCode(KeyName)
        Case Is = "Underline"
            KeyBindings.Add wdKeyCategoryMacro, "Formatting.ToggleUnderline", BuildKeyCode(KeyName)
        Case Is = "Emphasis"
            KeyBindings.Add wdKeyCategoryStyle, "Emphasis", BuildKeyCode(KeyName)
        Case Is = "Highlight"
            KeyBindings.Add wdKeyCategoryMacro, "Formatting.Highlight", BuildKeyCode(KeyName)
        Case Is = "Clear"
            KeyBindings.Add wdKeyCategoryMacro, "Formatting.ClearToNormal", BuildKeyCode(KeyName)
        Case Is = "Shrink Text"
            KeyBindings.Add wdKeyCategoryMacro, "Shrink.ShrinkAllOrCard", BuildKeyCode(KeyName)
        Case Is = "Select Similar"
            KeyBindings.Add wdKeyCategoryMacro, "Formatting.SelectSimilar", BuildKeyCode(KeyName)
        Case Else
            ' Nothing
        
    End Select
    
    '@Ignore ValueRequired
    Application.CustomizationContext = ThisDocument
End Sub

Public Sub ResetKeyboardShortcuts()
    On Error Resume Next
    
    Dim ModifierKey As Long
    #If Mac Then
        ModifierKey = wdKeyCommand
    #Else
        ModifierKey = wdKeyControl
    #End If
    
    ' Clear old keybindings
    Settings.RemoveKeyBindings
    
    ' Save defaults to registry
    SaveSetting "Verbatim", "Keyboard", "F2Shortcut", "Paste"
    SaveSetting "Verbatim", "Keyboard", "F3Shortcut", "Condense"
    SaveSetting "Verbatim", "Keyboard", "F4Shortcut", "Pocket"
    SaveSetting "Verbatim", "Keyboard", "F5Shortcut", "Hat"
    SaveSetting "Verbatim", "Keyboard", "F6Shortcut", "Block"
    SaveSetting "Verbatim", "Keyboard", "F7Shortcut", "Tag"
    SaveSetting "Verbatim", "Keyboard", "F8Shortcut", "Cite"
    SaveSetting "Verbatim", "Keyboard", "F9Shortcut", "Underline"
    SaveSetting "Verbatim", "Keyboard", "F10Shortcut", "Emphasis"
    SaveSetting "Verbatim", "Keyboard", "F11Shortcut", "Highlight"
    SaveSetting "Verbatim", "Keyboard", "F12Shortcut", "Clear"

    ' Save shortcuts in the template
    '@Ignore ImplicitUnboundDefaultMemberAccess
    If Application.CustomizationContext <> "Debate.dotm" Then Application.CustomizationContext = ActiveDocument.AttachedTemplate

    ' Speech shortcuts (tilde key shortcuts are set by FixTilde)
    KeyBindings.Add wdKeyCategoryMacro, "Paperless.SendToSpeechCursor", BuildKeyCode(ModifierKey, wdKeyAlt, vbKeyRight)
    KeyBindings.Add wdKeyCategoryMacro, "Paperless.SendToSpeechEnd", BuildKeyCode(ModifierKey, wdKeyAlt, wdKeyShift, vbKeyRight)
    KeyBindings.Add wdKeyCategoryMacro, "Flow.SendToFlowCell", BuildKeyCode(ModifierKey, wdKeyG)
    KeyBindings.Add wdKeyCategoryMacro, "Flow.SendToFlowColumn", BuildKeyCode(ModifierKey, wdKeyAlt, wdKeyG)
    KeyBindings.Add wdKeyCategoryMacro, "Flow.SendHeadingsToFlowCell", BuildKeyCode(ModifierKey, wdKeyShift, wdKeyG)
    KeyBindings.Add wdKeyCategoryMacro, "Flow.SendHeadingsToFlowColumn", BuildKeyCode(ModifierKey, wdKeyAlt, wdKeyShift, wdKeyG)
    
    KeyBindings.Add wdKeyCategoryMacro, "QuickCards.InsertCurrentQuickCard", BuildKeyCode(ModifierKey, wdKeyAlt, wdKeyShift, wdKeyV)
    
    ' Organize shortcuts
    KeyBindings.Add wdKeyCategoryMacro, "UI.ShowFormHelp", BuildKeyCode(wdKeyF1)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.PasteText", BuildKeyCode(wdKeyF2)
    KeyBindings.Add wdKeyCategoryMacro, "Condense.CondenseAllOrCard", BuildKeyCode(wdKeyF3)
    KeyBindings.Add wdKeyCategoryStyle, "Pocket", BuildKeyCode(wdKeyF4)
    KeyBindings.Add wdKeyCategoryStyle, "Hat", BuildKeyCode(wdKeyF5)
    KeyBindings.Add wdKeyCategoryStyle, "Block", BuildKeyCode(wdKeyF6)
    KeyBindings.Add wdKeyCategoryStyle, "Tag", BuildKeyCode(wdKeyF7)
    KeyBindings.Add wdKeyCategoryStyle, "Cite", BuildKeyCode(wdKeyF8)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.ToggleUnderline", BuildKeyCode(wdKeyF9)
    KeyBindings.Add wdKeyCategoryStyle, "Emphasis", BuildKeyCode(wdKeyF10)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.Highlight", BuildKeyCode(wdKeyF11)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.ClearToNormal", BuildKeyCode(wdKeyF12)
    
    ' Alternate Organize shortcuts for systems with F-key problems, e.g. Mac Word hijacks F6
    KeyBindings.Add wdKeyCategoryMacro, "UI.ShowFormHelp", BuildKeyCode(ModifierKey, wdKeyAlt, wdKey1)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.PasteText", BuildKeyCode(ModifierKey, wdKeyAlt, wdKey2)
    KeyBindings.Add wdKeyCategoryMacro, "Condense.CondenseAllOrCard", BuildKeyCode(ModifierKey, wdKeyAlt, wdKey3)
    KeyBindings.Add wdKeyCategoryStyle, "Pocket", BuildKeyCode(ModifierKey, wdKeyAlt, wdKey4)
    KeyBindings.Add wdKeyCategoryStyle, "Hat", BuildKeyCode(ModifierKey, wdKeyAlt, wdKey5)
    KeyBindings.Add wdKeyCategoryStyle, "Block", BuildKeyCode(ModifierKey, wdKeyAlt, wdKey6)
    KeyBindings.Add wdKeyCategoryStyle, "Tag", BuildKeyCode(ModifierKey, wdKeyAlt, wdKey7)
    KeyBindings.Add wdKeyCategoryStyle, "Cite", BuildKeyCode(ModifierKey, wdKeyAlt, wdKey8)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.ToggleUnderline", BuildKeyCode(ModifierKey, wdKeyAlt, wdKey9)
    KeyBindings.Add wdKeyCategoryStyle, "Emphasis", BuildKeyCode(ModifierKey, wdKeyAlt, wdKey0)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.Highlight", BuildKeyCode(ModifierKey, wdKeyAlt, wdKeyHyphen)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.ClearToNormal", BuildKeyCode(ModifierKey, wdKeyAlt, wdKeyEquals)
       
    ' Format shortcuts
    KeyBindings.Add wdKeyCategoryMacro, "Shrink.ShrinkAllOrCard", BuildKeyCode(wdKeyAlt, wdKeyF3)
    KeyBindings.Add wdKeyCategoryMacro, "Shrink.ShrinkAllOrCard", BuildKeyCode(ModifierKey, wdKey8)
    KeyBindings.Add wdKeyCategoryMacro, "Condense.CondenseWithPilcrows", BuildKeyCode(ModifierKey, wdKeyAlt, wdKeyF3)
    KeyBindings.Add wdKeyCategoryMacro, "Condense.CondenseNoPilcrows", BuildKeyCode(ModifierKey, wdKeyF3)
    KeyBindings.Add wdKeyCategoryMacro, "Condense.Uncondense", BuildKeyCode(ModifierKey, wdKeyAlt, wdKeyShift, wdKeyF3)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.AutoFormatCite", BuildKeyCode(ModifierKey, wdKeyF8)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.CopyPreviousCite", BuildKeyCode(wdKeyAlt, wdKeyF8)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.AutoUnderline", BuildKeyCode(wdKeyAlt, wdKeyF9)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.AutoEmphasizeFirst", BuildKeyCode(ModifierKey, wdKeyAlt, wdKeyF10)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.UpdateStyles", BuildKeyCode(ModifierKey, wdKeyF12)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.SelectSimilar", BuildKeyCode(ModifierKey, wdKeyF2)
    KeyBindings.Add wdKeyCategoryMacro, "Plugins.GetFromCiteCreator", BuildKeyCode(wdKeyAlt, wdKeyF2)
    KeyBindings.Add wdKeyCategoryMacro, "Formatting.AutoNumberTags", BuildKeyCode(ModifierKey, wdKeyShift, wdKey3)
    
    ' Paperless shortcuts
    KeyBindings.Add wdKeyCategoryMacro, "Paperless.SelectHeadingAndContent", BuildKeyCode(ModifierKey, wdKeyAlt, wdKeyA)
    KeyBindings.Add wdKeyCategoryMacro, "Paperless.MoveUp", BuildKeyCode(ModifierKey, wdKeyAlt, vbKeyUp)
    KeyBindings.Add wdKeyCategoryMacro, "Paperless.MoveDown", BuildKeyCode(ModifierKey, wdKeyAlt, vbKeyDown)
    KeyBindings.Add wdKeyCategoryMacro, "Paperless.MoveToBottom", BuildKeyCode(ModifierKey, wdKeyAlt, wdKeyShift, vbKeyDown)
    KeyBindings.Add wdKeyCategoryMacro, "Paperless.DeleteHeading", BuildKeyCode(ModifierKey, wdKeyAlt, vbKeyLeft)
    KeyBindings.Add wdKeyCategoryMacro, "Paperless.NewSpeech", BuildKeyCode(ModifierKey, wdKeyShift, wdKeyN)
    KeyBindings.Add wdKeyCategoryMacro, "Paperless.CopyToUSB", BuildKeyCode(ModifierKey, wdKeyShift, wdKeyU)
    KeyBindings.Add wdKeyCategoryMacro, "UI.ShowFormShare", BuildKeyCode(ModifierKey, wdKeyShift, wdKeyS)

    ' Tools shortcuts
    KeyBindings.Add wdKeyCategoryMacro, "Plugins.StartTimer", BuildKeyCode(ModifierKey, wdKeyShift, wdKeyT)
    KeyBindings.Add wdKeyCategoryMacro, "UI.ShowFormStats", BuildKeyCode(ModifierKey, wdKeyShift, wdKeyI)
    KeyBindings.Add wdKeyCategoryMacro, "Plugins.NavPaneCycle", BuildKeyCode(ModifierKey, wdKeyShift, wdKeyW)
    
    ' View shortcuts
    KeyBindings.Add wdKeyCategoryMacro, "View.ArrangeWindows", BuildKeyCode(ModifierKey, wdKeyShift, wdKeyTab)
    KeyBindings.Add wdKeyCategoryMacro, "View.SwitchWindows", BuildKeyCode(ModifierKey, wdKeyTab)
    KeyBindings.Add wdKeyCategoryMacro, "View.InvisibilityOff", BuildKeyCode(ModifierKey, wdKeyShift, wdKeyV)
    KeyBindings.Add wdKeyCategoryMacro, "View.ToggleReadingView", BuildKeyCode(ModifierKey, wdKeyAlt, wdKeyR)
    
    ' Caselist shortcuts
    KeyBindings.Add wdKeyCategoryMacro, "Caselist.CiteRequestCard", BuildKeyCode(ModifierKey, wdKeyShift, wdKeyQ)
    
    ' Settings shortcuts
    KeyBindings.Add wdKeyCategoryMacro, "UI.ShowFormSettings", BuildKeyCode(wdKeyAlt, wdKeyF1)

    ' Also sets shortcuts that use the tilde key
    Troubleshooting.FixTilde
    
    ' Save template
    ActiveDocument.AttachedTemplate.Save

    ' Reset customization context
    '@Ignore ValueRequired
    Application.CustomizationContext = ThisDocument
    
    On Error GoTo 0
End Sub

Public Sub RemoveKeyBindings()
    Dim k As KeyBinding
    
    For Each k In Application.KeyBindings
        k.Clear
    Next k
End Sub

' *************************************************************************************
' * MISC FUNCTIONS                                                                    *
' *************************************************************************************

Public Sub LaunchWebsite(ByVal URL As String)
    On Error GoTo Handler
    
    #If Mac Then
        AppleScriptTask "Verbatim.scpt", "RunShellScript", "open " & URL
    #Else
        ActiveDocument.FollowHyperlink (URL)
    #End If
    
    Exit Sub
   
Handler:
    If Err.Number = 4198 Then
        MsgBox "Opening website failed. Check your internet connection."
    Else
        MsgBox "Error " & Err.Number & ": " & Err.Description
    End If

End Sub

Public Sub OpenWordHelp()
    #If Mac Then
        Help wdHelp
    #Else
        CommandBars.FindControl(ID:=984).Execute
    #End If
End Sub

Public Sub OpenTemplatesFolder()
    #If Mac Then
    
        AppleScriptTask "Verbatim.scpt", "OpenFolder", Application.NormalTemplate.Path
    #Else
        Shell "explorer.exe " & CStr(Environ$("USERPROFILE")) & "\AppData\Roaming\Microsoft\Templates", vbNormalFocus
    #End If
End Sub

Public Function GetVersion() As String
    ' On mac, this can fail with a single document open, so ensure we always return something
    GetVersion = ""
    On Error Resume Next
    GetVersion = ActiveDocument.AttachedTemplate.BuiltInDocumentProperties(wdPropertyKeywords)
    On Error GoTo 0
End Function
