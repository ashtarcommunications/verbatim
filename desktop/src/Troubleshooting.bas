Attribute VB_Name = "Troubleshooting"
Option Explicit

' *************************************************************************************
' * REGISTRY FUNCTIONS                                                                *
' *************************************************************************************

#If Mac Then
    ' Do Nothing
#Else
Function RegKeyRead(RegKey As String) As String
    Dim WS As Object
    On Error Resume Next
    Set WS = CreateObject("WScript.Shell")
    RegKeyRead = WS.RegRead(RegKey)
    Set WS = Nothing
End Function

Function RegKeyExists(RegKey As String) As Boolean
    Dim WS As Object
    On Error GoTo Handler
    Set WS = CreateObject("WScript.Shell")
    WS.RegRead RegKey
    RegKeyExists = True
    Set WS = Nothing
    Exit Function
  
Handler:
    ' If key isn't found, it will throw an error
    RegKeyExists = False
    Set WS = Nothing
End Function

Sub RegKeySave(RegKey As String, Value As String, Optional ValueType As String = "REG_SZ")
    Dim WS As Object
    On Error Resume Next
    Set WS = CreateObject("WScript.Shell")
    WS.RegWrite RegKey, Value, ValueType
    Set WS = Nothing
End Sub
#End If

' *************************************************************************************
' * INSTALL CHECK FUNCTIONS                                                           *
' *************************************************************************************
Function InstallCheckTemplateName(Optional Notify As Boolean) As Boolean
    Dim msg As String
    
    On Error Resume Next

    ' Checks if Verbatim is installed as the wrong filename and optionally notifies the user
    If ActiveDocument.AttachedTemplate.Name <> "Debate.dotm" Then
        InstallCheckTemplateName = True
        
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
        InstallCheckTemplateName = False
    End If
End Function

Function InstallCheckTemplateLocation(Optional Notify As Boolean) As Boolean
    Dim msg As String
    Dim NormalPath
    Dim MsgPath
    
    On Error Resume Next

    #If Mac Then
        NormalPath = LCase(Application.NormalTemplate.Path)
        MsgPath = "~/Library/Group Containers/UBF8T34G9.Office/User Content/Templates"
    #Else
        ' Use LCase because Windows 8 Environ returns uppercase drive letters
        NormalPath = LCase(CStr(Environ("USERPROFILE")) & "\AppData\Roaming\Microsoft\Templates")
        MsgPath = "c:\Users\<yourname>\AppData\Roaming\Microsoft\Templates"
    #End If

    ' Checks if Verbatim is installed in the wrong location and optionally notifes the user
    If LCase(ActiveDocument.AttachedTemplate.Path) <> NormalPath Then
        InstallCheckTemplateLocation = True
        
        If Notify = True Then
            msg = "WARNING - Verbatim appears to be installed in the wrong location. " _
            & "The Verbatim template file (Debate.dotm) should be located in your Word Templates folder, usually located at: " _
            & vbCrLf & MsgPath & vbCrLf _
            & "Using it from a different location will break many features. " _
            & "You can open your templates folder or suppress this warning in the Verbatim settings."
            MsgBox (msg)
        End If

    Else
        InstallCheckTemplateLocation = False
    End If
End Function

#If Mac Then
Function InstallCheckScptFileExists(Optional Notify As Boolean) As Boolean
    Dim msg As String
    
    On Error Resume Next

    ' Checks if Verbatim.scpt is installed at the correct location
    If Filesystem.FileExists("~/Library/Application Scripts/com.Microsoft.Word/Verbatim.scpt") = False Then
        InstallCheckScptFileExists = True
        
        If Notify = True Then
            msg = "WARNING - You do not appear to have Verbatim.scpt installed at " _
            & "~/Library/Application Scripts/com.Microsoft.Word/Verbatim.scpt - " _
            & "Verbatim.scpt must be installed or many features will not work correctly. " _
            & "It is strongly recommended you run the Verbatim installer again, or manually install the file. " _
            & "This warning can be suppressed in the Verbatim settings."
            MsgBox (msg)
        End If

    Else
        InstallCheckScptFileExists = False
    End If
End Function
#End If
Function CheckSaveFormat(Optional Notify As Boolean) As Boolean
    Dim msg As String
    
    On Error Resume Next

    ' Check if default save format is .docx and optionally notifies the user
    If Application.DefaultSaveFormat = "Doc" Or Application.DefaultSaveFormat = "Doc97" Then
        CheckSaveFormat = True
        
        If Notify = True Then
            msg = "Your default save format appears to be set to .doc instead of .docx"
            msg = msg & " - It is highly recommended that you use the .docx format instead. "
            msg = msg & "Change automatically?" & vbCrLf & "(This warning can be supressed in the Verbatim options)"
            If MsgBox(msg, vbYesNo) = vbYes Then Application.DefaultSaveFormat = "Docx"
        End If
    Else
        CheckSaveFormat = False
    End If
End Function

Function CheckDocx(Optional Notify As Boolean) As Boolean
    Dim msg As String
    
    On Error Resume Next
    
    ' Check if current document is a .doc instead of a docx
    If Right(ActiveDocument.Name, 3) = "doc" Then
        CheckDocx = True
        
        If Notify = True Then
            msg = "This file is saved as .doc instead of .docx"
            msg = msg & " - It is highly recommended that you use the .docx format instead. "
            msg = msg & "Save as .docx automatically? This will overwrite any current file in the same directory with the same name." & vbCrLf & "(This warning can be supressed in the Verbatim options)"
            If MsgBox(msg, vbYesNo) = vbYes Then
                ActiveDocument.SaveAs FileName:=Left(ActiveDocument.FullName, InStrRev(ActiveDocument.FullName, ".") - 1), FileFormat:=wdFormatXMLDocument
            End If
        End If
    Else
        CheckDocx = False
    End If
End Function

Public Function CheckDefaultSave() As Boolean
    If Application.DefaultSaveFormat = "Docx" Then
        CheckDefaultSave = True
    Else
        CheckDefaultSave = False
    End If
End Function

Public Function CheckDuplicateTemplates() As Boolean
    On Error Resume Next
    
    If Filesystem.FileExists(Environ("USERPROFILE") & "\Desktop\Debate.dotm") = False And _
    Filesystem.FileExists(Environ("USERPROFILE") & "\Downloads\Debate.dotm") = False Then
        CheckDuplicateTemplates = True
    Else
        CheckDuplicateTemplates = False
    End If
End Function

Public Function CheckAddins() As Boolean
    On Error Resume Next

    #If Mac Then
        CheckAddins = True
        Exit Sub
    #Else
        Dim Addin As COMAddIn
               
        For Each Addin In Application.COMAddIns
            If Addin.Description = "Send to Bluetooth" And Addin.Connect = True Then
                CheckAddins = False
            End If
        Next Addin
    #End If
End Function

' *************************************************************************************
' * FIX FUNCTIONS                                                                     *
' *************************************************************************************
Sub DeleteDuplicateTemplates()
    Dim FilePath As String
    
    On Error Resume Next
    
    ' Check for "Debate.dotm" in the Desktop and Downloads folders, prompt to delete if found
    #If Mac Then
        FilePath = "/Users/" & Environ("USER") & "/Desktop/Debate.dotm"
        If Filesystem.FileExists(FilePath) = True Then
            If MsgBox("A duplicate copy of Debate.dotm was found on your Desktop - this can cause interoperability issues. Attempt to delete automatically?", vbYesNo) = vbYes Then
                Filesystem.DeleteFile FilePath
            End If
        End If
        
        FilePath = "/Users/" & Environ("USER") & "/Downloads/Debate.dotm"
      
        If Filesystem.FileExists(FilePath) = True Then
            If MsgBox("A duplicate copy of Debate.dotm was found in your Downloads folder - this can cause interoperability issues. Attempt to delete automatically?", vbYesNo) = vbYes Then
                Filesystem.DeleteFile FilePath
            End If
        End If
    #Else
        FilePath = Environ("USERPROFILE") & "\Desktop\Debate.dotm"
        If Filesystem.FileExists(FilePath) = True Then
            If MsgBox("A duplicate copy of Debate.dotm was found on your Desktop - this can cause interoperability issues. Attempt to delete automatically?", vbYesNo) = vbYes Then
                Filesystem.DeleteFile FilePath
            End If
        End If

        FilePath = Environ("USERPROFILE") & "\Downloads\Debate.dotm"
        If Filesystem.FileExists(FilePath) = True Then
            If MsgBox("A duplicate copy of Debate.dotm was found in your Downloads folder - this can cause interoperability issues. Attempt to delete automatically?", vbYesNo) = vbYes Then
                Filesystem.DeleteFile FilePath
            End If
        End If
    #End If
End Sub

Sub SetDefaultSave()
    Application.DefaultSaveFormat = "Docx"
End Sub

Sub DisableAddins()
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

Sub FixTilde()

    Dim ModifierKey
    
    #If Mac Then
        ModifierKey = wdKeyCommand
    #Else
        ModifierKey = wdKeyControl
    #End If
    
    ' VkKeyScan should usually return 192 - on models where it incorrectly returns 223, use 96 instead
    ' Keycodes: 96 = `, 192 = A`, 223 = Beta
    If VkKeyScan(Asc("`")) = 192 Then
        KeyBindings.Add wdKeyCategoryMacro, "Paperless.SendToSpeechCursor", VkKeyScan(Asc("`"))
        KeyBindings.Add wdKeyCategoryMacro, "Paperless.SendToSpeechEnd", BuildKeyCode(wdKeyAlt, VkKeyScan(Asc("`")))
        KeyBindings.Add wdKeyCategoryMacro, "Flow.SendToFlow", BuildKeyCode(ModifierKey, VkKeyScan(Asc("`")))
        KeyBindings.Add wdKeyCategoryMacro, "Paperless.SelectHeadingAndContent", BuildKeyCode(ModifierKey, wdKeyShift, VkKeyScan(Asc("`")))
    Else
        KeyBindings.Add wdKeyCategoryMacro, "Paperless.SendToSpeechCursor", VkKeyScan(96)
        KeyBindings.Add wdKeyCategoryMacro, "Paperless.SendToSpeechEnd", BuildKeyCode(wdKeyAlt, VkKeyScan(96))
        KeyBindings.Add wdKeyCategoryMacro, "Flow.SendToFlow", BuildKeyCode(ModifierKey, VkKeyScan(96))
        KeyBindings.Add wdKeyCategoryMacro, "Paperless.SelectHeadingAndContent", BuildKeyCode(ModifierKey, wdKeyShift, VkKeyScan(("`")))
    End If
End Sub
