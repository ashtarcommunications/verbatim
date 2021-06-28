Attribute VB_Name = "AttachVerbatim"
'@Folder("AttachVerbatim")
' Verbatim Startup
' Should be installed as a global template in the Word Startup folder to enable "Always On" mode. Verbatim must also be installed to work.
' Copyright © 2021 Aaron Hardy
' https://paperlessdebate.com
' support@paperlessdebate.com
'
' Verbatim is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License 3.0 as published by
' the Free Software Foundation.
'
' Verbatim is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License 3 for more details.
'
' For a copy of the GNU General Public License 3 see:
' http://www.gnu.org/licenses/gpl-3.0.txt

Option Explicit

'Windows API declarations to get command line
#If Win64 Then
    Public Declare PtrSafe Function GetCommandLine Lib "kernel32" Alias "GetCommandLineW" () As Long
    Public Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (MyDest As Any, MySource As Any, ByVal MySize As Long)
#ElseIf Not Mac Then
    Public Declare Function GetCommandLine Lib "kernel32" Alias "GetCommandLineW" () As Long
    Public Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (MyDest As Any, MySource As Any, ByVal MySize As Long)
#End If

#If Not Mac Then
    Public Function CmdToStr(ByRef CommandLine As Long) As String
        Dim Buffer() As Byte
        Dim StrLen As Long
        
        'Converts pointer to command line to a string
        If CommandLine Then
           StrLen = lstrlenW(CommandLine) * 2
           If StrLen Then
              ReDim Buffer(0 To (StrLen - 1)) As Byte
              CopyMemory Buffer(0), ByVal CommandLine, StrLen
              CmdToStr = Buffer
           End If
        End If
    End Function
#End If

Public Sub AutoExec()
    If GetSetting("Verbatim", "Admin", "AlwaysOn", False) = True Then
        #If Mac Then
            AttachVerbatim
        #Else
            'If Word is opening a document, command line will contain "/n", so don't add a new doc
            If InStr(CmdToStr(GetCommandLine), "/n") = 0 Then
                AttachVerbatim
            End If
        #End If
    End If
End Sub

Sub RibbonStartup(ByVal control As IRibbonControl)
    Select Case control.ID
        Case Is = "Verbatimize"
            AttachVerbatim
        Case Else
            'Do Nothing
    End Select
End Sub

Public Function CheckVerbatimExists() As Boolean
    Dim FSO As Object

    On Error GoTo Handler
    
    CheckVerbatimExists = False

    #If Mac Then
        If AppleScriptTask("Verbatim.scpt", "FileExists", "Macintosh HD" & Replace(Replace(Application.NormalTemplate.Path & Application.PathSeparator & "Debate.dotm", ".localized", ""), "/", ":")) = "true" Then
            CheckVerbatimExists = True
        End If
        Set FSO = Nothing
    #Else
        'Use late binding to avoid needing a reference in Normal.dotm
        Set FSO = CreateObject("Scripting.FileSystemObject")
        If FSO.FileExists(Application.NormalTemplate.Path & Application.PathSeparator & "Debate.dotm") = True Then
            CheckVerbatimExists = True
        End If
        Set FSO = Nothing
    #End If

    Exit Function

Handler:
    Set FSO = Nothing
    Application.StatusBar = "Error checking Verbatim template. Error " & Err.Number & ": " & Err.Description

End Function

Public Sub AttachVerbatim()
    On Error GoTo Handler
    
    If CheckVerbatimExists = False Then
        Application.StatusBar = "Debate.dotm not found in your Templates folder - it must be installed correctly to attach it."
        Exit Sub
    End If
    
    'If starting Word from scratch, add a new doc based on the template - will suppress Word's built-in doc
    If Application.Documents.Count = 0 Then
        Application.Documents.Add Template:=Application.NormalTemplate.Path & Application.PathSeparator & "Debate.dotm"
    Else
        'Attach Verbatim to the current doc
        ActiveDocument.AttachedTemplate = Application.NormalTemplate.Path & Application.PathSeparator & "Debate.dotm"
        ActiveDocument.UpdateStyles
        #If Mac Then
            Application.AddIns(Application.NormalTemplate.Path & Application.PathSeparator & "Debate.dotm").Installed = True
        #End If
    End If
    
    'TODO - is this necessary anymore, and is it necessary on Mac?
    #If Not Mac Then
        'Add and close another document to fake Word into refreshing ribbon
        Application.Run "'ActiveDocument.AttachedTemplate'!Paperless.NewDocument"
        ActiveDocument.Close wdDoNotSaveChanges
    #End If
    
    Exit Sub
    
Handler:
    Application.StatusBar = "Error Attaching Verbatim. Error " & Err.Number & ": " & Err.Description

End Sub

