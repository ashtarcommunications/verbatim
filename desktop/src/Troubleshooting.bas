Attribute VB_Name = "Troubleshooting"
Option Explicit

Public Function InstallCheckTemplateName(Optional ByVal Notify As Boolean) As Boolean
    Dim msg As String
    
    On Error Resume Next

    ' Checks if Verbatim is installed as the wrong filename and optionally notifies the user
    If ActiveDocument.AttachedTemplate.Name <> "Debate.dotm" Then
        InstallCheckTemplateName = False
        
        If Notify = True Then
            msg = "WARNING - Verbatim appears to be installed incorrectly as " _
            & ActiveDocument.AttachedTemplate.Name & ". " _
            & "Verbatim must be named ""Debate.dotm"" or many features will not work correctly " _
            & "and it will break compatibility with others. " _
            & "It is strongly recommended you change the file name back to Debate.dotm. " _
            & "This warning can be suppressed in the Verbatim settings."
            MsgBox (msg)
        End If

    Else
        InstallCheckTemplateName = True
    End If
    
    On Error GoTo 0
End Function

Public Function InstallCheckTemplateLocation(Optional ByVal Notify As Boolean) As Boolean
    Dim msg As String
    Dim NormalPath As String
    Dim MsgPath As String
    
    On Error Resume Next

    #If Mac Then
        NormalPath = LCase(Application.NormalTemplate.Path)
        MsgPath = "~/Library/Group Containers/UBF8T34G9.Office/User Content/Templates"
    #Else
        ' Use LCase because Windows 8 Environ returns uppercase drive letters
        NormalPath = LCase$(CStr(Environ$("USERPROFILE")) & "\AppData\Roaming\Microsoft\Templates")
        MsgPath = "c:\Users\<yourname>\AppData\Roaming\Microsoft\Templates"
    #End If

    ' Checks if Verbatim is installed in the wrong location and optionally notifes the user
    If LCase$(ActiveDocument.AttachedTemplate.Path) <> NormalPath Then
        InstallCheckTemplateLocation = False
        
        If Notify = True Then
            msg = "WARNING - Verbatim appears to be installed in the wrong location. " _
            & "The Verbatim template file (Debate.dotm) should be located in your Word Templates folder, usually located at: " _
            & vbCrLf & MsgPath & vbCrLf _
            & "Using it from a different location will break many features. " _
            & "You can open your templates folder or suppress this warning in the Verbatim settings."
            MsgBox (msg)
        End If

    Else
        InstallCheckTemplateLocation = True
    End If
    
    On Error GoTo 0
End Function

#If Mac Then
Function InstallCheckScptFileExists(Optional Notify As Boolean) As Boolean
    Dim msg As String
    
    On Error Resume Next

    ' Checks if Verbatim.scpt is installed at the correct location
    If Filesystem.FileExists("/Users/" & Environ("USER") & "/Library/Application Scripts/com.Microsoft.Word/Verbatim.scpt") = False Then
        InstallCheckScptFileExists = False
        
        If Notify = True Then
            msg = "WARNING - You do not appear to have Verbatim.scpt installed at " _
            & "/Users/<yourusername>/Library/Application Scripts/com.Microsoft.Word/Verbatim.scpt - " _
            & "Verbatim.scpt must be installed or many features will not work correctly. " _
            & "It is strongly recommended you run the Verbatim installer again, or manually install the file. " _
            & "This warning can be suppressed in the Verbatim settings."
            MsgBox (msg)
        End If

    Else
        InstallCheckScptFileExists = True
    End If
    
    On Error GoTo 0
End Function
#End If

Public Function CheckSaveFormat(Optional ByVal Notify As Boolean) As Boolean
    Dim msg As String
    
    On Error Resume Next

    ' Check if default save format is .docx and optionally notifies the user
    If Application.DefaultSaveFormat <> "Docx" Then
        CheckSaveFormat = False
        
        If Notify = True Then
            msg = "Your default save format is not set to .docx"
            msg = msg & " - It is highly recommended that you use the .docx format instead. "
            msg = msg & "Change automatically?" & vbCrLf & "(This warning can be supressed in the Verbatim options)"
            If MsgBox(msg, vbYesNo) = vbYes Then Application.DefaultSaveFormat = "Docx"
        End If
    Else
        CheckSaveFormat = True
    End If
    
    On Error GoTo 0
End Function

'@Ignore FunctionReturnValueAlwaysDiscarded
Public Function CheckDocx(Optional ByVal Notify As Boolean) As Boolean
    Dim msg As String
    
    On Error Resume Next
    
    ' Check if current document is a .doc instead of a docx
    If Right$(ActiveDocument.Name, 3) = "doc" Then
        CheckDocx = False
        
        If Notify = True Then
            msg = "This file is saved as .doc instead of .docx"
            msg = msg & " - It is highly recommended that you use the .docx format instead. "
            msg = msg & "Save as .docx automatically? This will overwrite any current file in the same directory with the same name." & vbCrLf & "(This warning can be supressed in the Verbatim options)"
            If MsgBox(msg, vbYesNo) = vbYes Then
                ActiveDocument.SaveAs Filename:=Left$(ActiveDocument.FullName, InStrRev(ActiveDocument.FullName, ".") - 1), FileFormat:=wdFormatXMLDocument
            End If
        End If
    Else
        CheckDocx = True
    End If
    
    On Error GoTo 0
End Function

Public Function CheckDuplicateTemplates() As Boolean
    Dim DesktopPath As String
    Dim DownloadsPath As String
    
    On Error Resume Next
    
    #If Mac Then
        DesktopPath = "/Users/" & Environ("USER") & "/Desktop/Debate.dotm"
        DownloadsPath = "/Users/" & Environ("USER") & "/Downloads/Debate.dotm"
    #Else
        DesktopPath = Environ$("USERPROFILE") & "\Desktop\Debate.dotm"
        DownloadsPath = Environ$("USERPROFILE") & "\Downloads\Debate.dotm"
    #End If
    
    CheckDuplicateTemplates = Filesystem.FileExists(DesktopPath) = True Or Filesystem.FileExists(DownloadsPath) = True
    
    On Error GoTo 0
End Function

Public Function CheckAddins() As Boolean
    On Error Resume Next

    #If Mac Then
        CheckAddins = True
        Exit Function
    #Else
        Dim Addin As COMAddIn
        
        CheckAddins = True
        
        For Each Addin In Application.COMAddIns
            If Addin.Description = "Send to Bluetooth" And Addin.Connect = True Then
                CheckAddins = False
            End If
        Next Addin
    #End If
    
    On Error GoTo 0
End Function

' *************************************************************************************
' * FIX FUNCTIONS                                                                     *
' *************************************************************************************

Public Sub DeleteDuplicateTemplates()
    Dim DesktopPath As String
    Dim DownloadsPath As String
    
    On Error Resume Next
    
    ' Check for "Debate.dotm" in the Desktop and Downloads folders, prompt to delete if found
    If Troubleshooting.CheckDuplicateTemplates = False Then
        If MsgBox("A duplicate copy of Debate.dotm was found on your Desktop or in your Downloads folder - this can cause interoperability issues. Attempt to delete automatically?", vbYesNo) = vbYes Then
            #If Mac Then
                DesktopPath = "/Users/" & Environ("USER") & "/Desktop/Debate.dotm"
                DownloadsPath = "/Users/" & Environ("USER") & "/Downloads/Debate.dotm"
            #Else
                DesktopPath = Environ$("USERPROFILE") & "\Desktop\Debate.dotm"
                DownloadsPath = Environ$("USERPROFILE") & "\Downloads\Debate.dotm"
            #End If
            
            If Filesystem.FileExists(DesktopPath) Then Filesystem.DeleteFile DesktopPath
            If Filesystem.FileExists(DownloadsPath) Then Filesystem.DeleteFile DownloadsPath
        End If
    End If
    
    On Error GoTo 0
End Sub

Public Sub SetDefaultSave()
    Application.DefaultSaveFormat = "Docx"
End Sub

Public Sub DisableAddins()
    #If Mac Then
        MsgBox "This function doesn't work on Mac"
        Exit Sub
    #Else
        Dim Addin As COMAddIn
        For Each Addin In Application.COMAddIns
            ' Disable problematic bluetooth addin
            If Addin.Description = "Send to Bluetooth" Then Addin.Connect = False
        Next Addin
    #End If
End Sub

Public Sub FixTilde()
    '@Ignore VariableNotUsed
    Dim ModifierKey As Long
    
    #If Mac Then
        ModifierKey = wdKeyCommand
    #Else
        '@Ignore AssignmentNotUsed
        ModifierKey = wdKeyControl
    #End If
    
    ' VkKeyScan should usually return 192 - on models where it incorrectly returns 223, use 96 instead
    ' Keycodes: 96 = `, 192 = A`, 223 = Beta
    #If Mac Then
        KeyBindings.Add wdKeyCategoryMacro, "Paperless.SendToSpeechCursor", 192
        KeyBindings.Add wdKeyCategoryMacro, "Paperless.SendToSpeechEnd", BuildKeyCode(wdKeyAlt, 192)
    #Else
        If VkKeyScan(Asc("`")) = 192 Then
            KeyBindings.Add wdKeyCategoryMacro, "Paperless.SendToSpeechCursor", VkKeyScan(Asc("`"))
            KeyBindings.Add wdKeyCategoryMacro, "Paperless.SendToSpeechEnd", BuildKeyCode(wdKeyAlt, VkKeyScan(Asc("`")))
            
            KeyBindings.Add wdKeyCategoryMacro, "Flow.SendToFlowCell", BuildKeyCode(ModifierKey, VkKeyScan(Asc("`")))
            KeyBindings.Add wdKeyCategoryMacro, "Flow.SendToFlowColumn", BuildKeyCode(ModifierKey, wdKeyAlt, VkKeyScan(Asc("`")))
            KeyBindings.Add wdKeyCategoryMacro, "Flow.SendHeadingsToFlowCell", BuildKeyCode(ModifierKey, wdKeyShift, VkKeyScan(Asc("`")))
            KeyBindings.Add wdKeyCategoryMacro, "Flow.SendHeadingsToFlowColumn", BuildKeyCode(ModifierKey, wdKeyAlt, wdKeyShift, VkKeyScan(Asc("`")))
        Else
            KeyBindings.Add wdKeyCategoryMacro, "Paperless.SendToSpeechCursor", VkKeyScan(96)
            KeyBindings.Add wdKeyCategoryMacro, "Paperless.SendToSpeechEnd", BuildKeyCode(wdKeyAlt, VkKeyScan(96))
            
            KeyBindings.Add wdKeyCategoryMacro, "Flow.SendToFlowCell", BuildKeyCode(ModifierKey, VkKeyScan(96))
            KeyBindings.Add wdKeyCategoryMacro, "Flow.SendToFlowColumn", BuildKeyCode(ModifierKey, wdKeyAlt, VkKeyScan(96))
            KeyBindings.Add wdKeyCategoryMacro, "Flow.SendHeadingsToFlowCell", BuildKeyCode(ModifierKey, wdKeyShift, VkKeyScan(96))
            KeyBindings.Add wdKeyCategoryMacro, "Flow.SendHeadingsToFlowColumn", BuildKeyCode(ModifierKey, wdKeyAlt, wdKeyShift, VkKeyScan(96))
        End If
    #End If
End Sub
