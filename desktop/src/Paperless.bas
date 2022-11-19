Attribute VB_Name = "Paperless"
Option Explicit

'*************************************************************************************
'* RIBBON FUNCTIONS                                                                  *
'*************************************************************************************

Sub AutoOpenFolder(control As IRibbonControl, pressed As Boolean)
' Runs in the background to automatically open all documents in the speech folder.

    Dim AutoOpenDir As String
    
    #If Mac Then
        Dim Files
        Dim f
    #Else
        Dim Files As Scripting.Files
        Dim FSO As Scripting.FileSystemObject
        Dim f As Scripting.File
    #End If
    Dim d As Document
    Dim IsOpen As Boolean
    
    On Error GoTo Handler
    
    ' If pressed, turn on the listener
    If pressed Then
    
        ' Check for auto open folder
        AutoOpenDir = GetSetting("Verbatim", "Paperless", "AutoOpenDir", "")
        If AutoOpenDir = "" Or AutoOpenDir = "?" Then
            If MsgBox("You have not set an Auto Open folder. Open settings now?", vbYesNo) = vbYes Then
                UI.ShowForm "Settings"
                Globals.AutoOpenFolderToggle = False
                Ribbon.RefreshRibbon
                Exit Sub
            Else
                Globals.AutoOpenFolderToggle = False
                Ribbon.RefreshRibbon
                Exit Sub
            End If
        End If
        
        #If Mac Then
            ' Ensure a trailing :
            If Right(AutoOpenDir, 1) <> ":" Then AutoOpenDir = AutoOpenDir & ":"
        #End If
        
        ' Prompt that listener is on
        If MsgBox("This will start a listener that automatically opens all documents in the root of your Auto Open folder:" & vbCrLf & AutoOpenDir, vbOKCancel) = vbCancel Then
            Globals.AutoOpenFolderToggle = False
            Ribbon.RefreshRibbon
            Exit Sub
        End If
        
        Globals.AutoOpenFolderToggle = True
        
        #If Not Mac Then
            Set FSO = New Scripting.FileSystemObject
        #End If
        
        ' Loop until unpressed
        Do
            DoEvents
            #If Mac Then
                Files = Split(Filesystem.GetFilesInFolder(AutoOpenDir), Chr(10))
            #Else
                Set Files = FSO.GetFolder(AutoOpenDir).Files
            #End If
            
            ' Loop all files - if not open, open it
            For Each f In Files
                IsOpen = False
                For Each d In Application.Documents
                    If d.FullName = f.Path Then IsOpen = True
                Next d
                
                If IsOpen = False _
                    And Left(f.Name, 1) <> "~" _
                    And (Right(f.Path, 3) = "doc" _
                        Or Right(f.Path, 4) = "docx" _
                        Or Right(f.Path, 3) = "rtf" _
                    ) Then Documents.Open f.Path
            Next f
        Loop Until Globals.AutoOpenFolderToggle = False
    
    Else
        Globals.AutoOpenFolderToggle = False
        MsgBox "Stopped listening to the Auto Open folder.", vbInformation
    End If
    
    Ribbon.RefreshRibbon
    #If Not Mac Then
        Set FSO = Nothing
        Set Files = Nothing
    #End If
    
    Exit Sub
    
Handler:
    #If Not Mac Then
        Set FSO = Nothing
        Set Files = Nothing
    #End If
    Ribbon.RefreshRibbon
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Sub GetSpeeches(control As IRibbonControl, ByRef returnedVal)

    Dim xml As String
    
    On Error Resume Next
        
    ' Set Mouse Pointer
    System.Cursor = wdCursorWait

    ' Initialize XML
    xml = "<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">"

    Dim Response As Dictionary
    'Set Response = HTTP.GetReq(Globals.CASELIST_URL & "/tabroom/rounds?current=true")
    Set Response = HTTP.GetReq(Globals.MOCK_ROUNDS & "?current=true")
    
    If Response("status") = 401 Then
        UI.ShowForm "Login"
        Exit Sub
    End If

    Dim Round
    Dim Tournament As String
    Dim RoundName As String
    Dim Side As String
    Dim SideName As String
    Dim Opponent As String
    
    Dim i As Long
    i = 0

    For Each Round In Response("body")
        i = i + 1
        Tournament = Round("tournament")
        Side = Round("side")
        RoundName = Round("round")
        Opponent = Round("opponent")
        Tournament = Trim(ScrubString(Tournament))
        Side = Trim(ScrubString(Side))
        SideName = Strings.DisplaySide(Side)
        RoundName = Strings.RoundName(Trim(ScrubString(RoundName)))
        Opponent = Trim(ScrubString(Opponent))
            
        If Side = "A" Then
            xml = xml & "<button id=""Speech2AC" & i & """ label=""2AC" & " " & Tournament & " " & RoundName & " vs " & Opponent & """ tag=""2AC" & " " & Tournament & " " & RoundName & " vs " & Opponent & """ onAction=""Paperless.NewSpeechFromMenu"" />"
            xml = xml & "<button id=""Speech1AR" & i & """ label=""1AR" & " " & Tournament & " " & RoundName & " vs " & Opponent & """ tag=""1AR" & " " & Tournament & " " & RoundName & " vs " & Opponent & """ onAction=""Paperless.NewSpeechFromMenu"" />"
            xml = xml & "<button id=""Speech2AR" & i & """ label=""2AR" & " " & Tournament & " " & RoundName & " vs " & Opponent & """ tag=""2AR" & " " & Tournament & " " & RoundName & " vs " & Opponent & """ onAction=""Paperless.NewSpeechFromMenu"" />"
            xml = xml & "<button id=""Speech1AC" & i & """ label=""1AC" & " " & Tournament & " " & RoundName & " vs " & Opponent & """ tag=""1AC" & " " & Tournament & " " & RoundName & " vs " & Opponent & """ onAction=""Paperless.NewSpeechFromMenu"" />"
            xml = xml & "<menuSeparator id=""separator" & i & """ />"
        Else
            xml = xml & "<button id=""Speech1NC" & i & """ label=""1NC" & " " & Tournament & " " & RoundName & " vs " & Opponent & """ tag=""1NC" & " " & Tournament & " " & RoundName & " vs " & Opponent & """ onAction=""Paperless.NewSpeechFromMenu"" />"
            xml = xml & "<button id=""Speech2NC" & i & """ label=""2NC" & " " & Tournament & " " & RoundName & " vs " & Opponent & """ tag=""2NC" & " " & Tournament & " " & RoundName & " vs " & Opponent & """ onAction=""Paperless.NewSpeechFromMenu"" />"
            xml = xml & "<button id=""Speech1NR" & i & """ label=""1NR" & " " & Tournament & " " & RoundName & " vs " & Opponent & """ tag=""1NR" & " " & Tournament & " " & RoundName & " vs " & Opponent & """ onAction=""Paperless.NewSpeechFromMenu"" />"
            xml = xml & "<button id=""Speech2NR" & i & """ label=""2NR" & " " & Tournament & " " & RoundName & " vs " & Opponent & """ tag=""2NR" & " " & Tournament & " " & RoundName & " vs " & Opponent & """ onAction=""Paperless.NewSpeechFromMenu"" />"
            xml = xml & "<menuSeparator id=""separator" & i & """ />"
        End If
    Next Round
        
    ' Add default speech options
    xml = xml & "<button id=""Speech2AC"" label=""2AC"" tag=""2AC"" onAction=""Paperless.NewSpeechFromMenu"" />"
    xml = xml & "<button id=""Speech1AR"" label=""1AR"" tag=""1AR"" onAction=""Paperless.NewSpeechFromMenu"" />"
    xml = xml & "<button id=""Speech2AR"" label=""2AR"" tag=""2AR"" onAction=""Paperless.NewSpeechFromMenu"" />"
    xml = xml & "<button id=""Speech1AC"" label=""1AC"" tag=""1AC"" onAction=""Paperless.NewSpeechFromMenu"" />"
    xml = xml & "<menuSeparator id=""separator"" />"
    xml = xml & "<button id=""Speech1NC"" label=""1NC"" tag=""1NC"" onAction=""Paperless.NewSpeechFromMenu"" />"
    xml = xml & "<button id=""Speech2NC"" label=""2NC"" tag=""2NC"" onAction=""Paperless.NewSpeechFromMenu"" />"
    xml = xml & "<button id=""Speech1NR"" label=""1NR"" tag=""1NR"" onAction=""Paperless.NewSpeechFromMenu"" />"
    xml = xml & "<button id=""Speech2NR"" label=""2NR"" tag=""2NR"" onAction=""Paperless.NewSpeechFromMenu"" />"
    
    ' Close XML
    xml = xml & "</menu>"
                 
    returnedVal = xml
    
    Set Response = Nothing

    System.Cursor = wdCursorNormal
    Exit Sub

Handler:
    Set Response = Nothing
    System.Cursor = wdCursorNormal
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Sub NewSpeechFromMenu(control As IRibbonControl)
    Dim AutoSaveDirectory As String
    Dim FileName As String
    Dim h

    ' Add a new document based on the template
    Paperless.NewDocument

    ' Get filename from control tag
    FileName = control.Tag
    
    ' If Tag is just the speech name, add a date
    If Len(FileName) = 3 Then
        If Hour(Now) > 12 Then h = Hour(Now) - 12 & "PM"
        If Hour(Now) <= 12 Then h = Hour(Now) & "AM"
        FileName = FileName & " " & Month(Now) & "-" & Day(Now) & " " & h
    End If
    
    ' Add speech to the name
    FileName = "Speech " & FileName
    
    ' Autosave or open save dialog
    If GetSetting("Verbatim", "Paperless", "AutoSaveSpeech", False) = True Then
        AutoSaveDirectory = GetSetting("Verbatim", "Paperless", "AutoSaveDir", CurDir())
        If Right(AutoSaveDirectory, 1) <> "\" Then AutoSaveDirectory = AutoSaveDirectory & "\"
        FileName = AutoSaveDirectory & "\" & FileName
        ActiveDocument.SaveAs FileName:=FileName, FileFormat:=wdFormatXMLDocument
    Else
        With Application.Dialogs(wdDialogFileSaveAs)
            .Name = FileName
            If .Show = 0 Then Exit Sub
        End With
    End If

    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

'*************************************************************************************
'* MOVE AND SELECT FUNCTIONS                                                         *
'*************************************************************************************

Public Sub SelectHeadingAndContent()
    Dim OLevel As Integer
    
    ' Move to start of current paragraph and collapse the selection
    Selection.StartOf Unit:=wdParagraph
    Selection.Collapse
        
    ' Move backwards through each paragraph to find the first tag, block title, hat, pocket or the top of the document
    Do While True
        If Selection.Paragraphs.outlineLevel < wdOutlineLevel5 Then Exit Do ' Headings 1-4
        If Selection.Start <= ActiveDocument.Range.Start Then ' Top of document
            Application.StatusBar = "Nothing found to select"
            Exit Sub
        End If
        Selection.Move Unit:=wdParagraph, Count:=-1
    Loop
        
    ' Get current outline level
    OLevel = Selection.Paragraphs.outlineLevel
    
    ' Extend selection until hitting the bottom or a bigger outline level
    Selection.MoveEnd Unit:=wdParagraph, Count:=1
    Do While True And Selection.End <> ActiveDocument.Range.End
        Selection.MoveEnd Unit:=wdParagraph, Count:=1
        If Selection.Paragraphs.Last.outlineLevel <= OLevel Then
            Selection.MoveEnd Unit:=wdParagraph, Count:=-1
            Exit Do ' Bigger Outline Level
        End If
    Loop

End Sub

Public Function SelectHeadingAndContentRange(p As Paragraph) As Range
    Dim r As Range
    Set r = p.Range
    
    ' Extend selection until hitting the bottom or a bigger outline level
    r.MoveEnd Unit:=wdParagraph, Count:=1
    Do While True And r.End <> ActiveDocument.Range.End
        r.MoveEnd Unit:=wdParagraph, Count:=1
        If r.Paragraphs.Last.outlineLevel <= p.outlineLevel Then
            r.MoveEnd Unit:=wdParagraph, Count:=-1
            Exit Do ' Bigger Outline Level
        End If
    Loop
    
    Set SelectHeadingAndContentRange = r
End Function


Sub MoveUp()
' Moves the current pocket, hat, block, or tag, up one level in the document outline
    Dim OLevel As Long
    Dim CurrentView As Long
    Dim StartLocation As Long
    
    On Error GoTo Handler
    
    Application.ScreenUpdating = False
    
    ' Save current view
    CurrentView = ActiveWindow.ActivePane.View.Type
    
    ' Move to start of current paragraph and collapse the selection
    Selection.StartOf Unit:=wdParagraph
    Selection.Collapse
    
    ' Move backwards through each paragraph to find the first tag, block title, hat, pocket, or the top of the document
    Do While True
        If Selection.Start <= ActiveDocument.Range.Start Then Exit Sub ' Top of doc
        If Selection.Paragraphs.outlineLevel < wdOutlineLevel5 Then Exit Do ' Headings 1-4
        Selection.Move Unit:=wdParagraph, Count:=-1
    Loop
        
    ' Get current outline level
    OLevel = Selection.Paragraphs.outlineLevel
    
    ' Check to make sure you're not moving a card above a block
    If OLevel = 4 Then
        StartLocation = Selection.Start ' Save current location
        Do While True
            Selection.Move Unit:=wdParagraph, Count:=-1
            If Selection.Start <= ActiveDocument.Range.Start Then
                Selection.Start = StartLocation
                Exit Sub
            End If
            If Selection.Paragraphs.outlineLevel = wdOutlineLevel4 Then
                Selection.Start = StartLocation
                Exit Do
            End If
            If Selection.Paragraphs.outlineLevel < wdOutlineLevel4 Then
                Application.StatusBar = "Already the first card on this block"
                Selection.Start = StartLocation
                Exit Sub
            End If
        Loop
    End If
    
    ' Switch to outline view and collapse to current level
    ActiveWindow.ActivePane.View.Type = wdOutlineView
    ActiveWindow.View.ShowHeading OLevel
    
    ' Move up
    ' Selection.Range.Relocate wdRelocateUp - CRASHES WORD 2013
    Application.Run "OutlineMoveUp"
    Selection.Collapse

    ' Switch back to previous view
    ActiveWindow.ActivePane.View.Type = CurrentView
    
    Application.ScreenUpdating = True

    Exit Sub
    
Handler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Sub MoveDown()
' Moves the current pocket, hat, block, or tag down one level in the document outline

    Dim OLevel As Long
    Dim CurrentView As Long
    Dim StartLocation As Long
    
    On Error GoTo Handler
    Application.ScreenUpdating = False
    
    ' Save current view
    CurrentView = ActiveWindow.ActivePane.View.Type
    
    ' Move to start of current paragraph and collapse the selection
    Selection.StartOf Unit:=wdParagraph
    Selection.Collapse
    
    ' Move backwards through each paragraph to find the first tag, block title, hat, pocket, or the top of the document
    Do While True
        If Selection.Paragraphs.outlineLevel < wdOutlineLevel5 Then
            Exit Do ' Headings 1-4
        Else
            Application.StatusBar = "Nothing found to move"
            Exit Sub
        End If
        Selection.Move Unit:=wdParagraph, Count:=-1
    Loop
        
    ' Get current outline level
    OLevel = Selection.Paragraphs.outlineLevel
    
    ' Check to make sure you're not already at the bottom
    StartLocation = Selection.Start ' Save current location
    Do While True
        Selection.Move Unit:=wdParagraph, Count:=1
        If Selection.End + 1 >= ActiveDocument.Range.End Then
                Selection.Start = StartLocation
                Selection.Collapse
                Application.StatusBar = "Already at the bottom"
                Exit Sub
        End If
        If Selection.Paragraphs.outlineLevel <= OLevel Then
            Selection.Start = StartLocation
            Selection.Collapse
            Exit Do
        End If
    Loop
    
    ' Check to make sure you're not moving a card off a block or the bottom
    If OLevel = 4 Then
        StartLocation = Selection.Start ' Save current location
        Do While True
            Selection.Move Unit:=wdParagraph, Count:=1
            If Selection.End + 1 >= ActiveDocument.Range.End Then
                Selection.Start = StartLocation
                Selection.Collapse
                Exit Sub
            End If
            If Selection.Paragraphs.outlineLevel = wdOutlineLevel4 Then
                Selection.Start = StartLocation
                Selection.Collapse
                Exit Do
            End If
            If Selection.Paragraphs.outlineLevel < wdOutlineLevel4 Then
                Application.StatusBar = "Already the last card on this block"
                Selection.Start = StartLocation
                Selection.Collapse
                Exit Sub
            End If
        Loop
    End If
    
    ' Switch to outline view and collapse to current level
    ActiveWindow.ActivePane.View.Type = wdOutlineView
    ActiveWindow.View.ShowHeading OLevel

    ' Move down
    ' Selection.Range.Relocate wdRelocateDown - CRASHES WORD 2013
    Application.Run "OutlineMoveDown"
    Selection.Collapse

    ' Switch back to previous view
    ActiveWindow.ActivePane.View.Type = CurrentView
    
    Application.ScreenUpdating = True
    
    Exit Sub
    
Handler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Sub DeleteHeading()
' Deletes the current card, block, hat, or pocket
    Paperless.SelectHeadingAndContent
    Selection.Delete

End Sub

'*************************************************************************************
'* SEND FUNCTIONS                                                                    *
'*************************************************************************************

Sub SendToSpeech()
' Sends content to the Speech doc.  Sends currently selected text,
' or if nothing is selected, the current tag, block, hat, or pocket
' If in reading view, enters a stopped reading marker at the current location

    Dim CurrentDoc As String
    Dim SpeechDoc As Document
    Dim d As Document
    Dim FoundDoc As Long

    On Error GoTo Handler

    ' If in reading mode, enter a stopped reading marker
    If ActiveWindow.View.ReadingLayout Then
        
        ' For Word 2013, use a comment
        If Application.Version >= "15.0" Then
            
            Selection.Collapse
            
            ' Insert a comment and close reviewing pane
            Selection.Comments.Add Range:=Selection.Range
            ActiveWindow.ActivePane.Close
            Exit Sub
        
        ' Previous versions, use a marker
        Else
            ActiveWindow.View.ReadingLayoutAllowEditing = True
            Selection.Collapse
            If Selection.Words(1).End <> Selection.End Then Selection.MoveRight wdWord
            Selection.Font.Color = wdColorRed
            Selection.Font.Size = 18
            Selection.TypeText Chr(167) & " Marked " & FormatDateTime(Time, 4) & " " & Chr(167) & " "
            Exit Sub
        End If
    End If

    ' Save active document name
    CurrentDoc = ActiveDocument.Name

SpeechDocCheck:

    ' If there's an active speech doc, use it
    If Globals.ActiveSpeechDoc <> "" Then
        For Each d In Application.Documents
            If d.Name = Globals.ActiveSpeechDoc Then
                Set SpeechDoc = Application.Documents(Globals.ActiveSpeechDoc)
            End If
        Next d
    Else
        ' Look for a document with "speech" in the title
        For Each d In Application.Documents
            If InStr(LCase(d.Name), "speech") Then
                FoundDoc = FoundDoc + 1
                If FoundDoc = 1 Then Set SpeechDoc = d
            End If
        Next d
        
        ' If no Speech doc is found, prompt whether to create one.
        ' If yes, create a new document based on the current template to save, then retry
        If FoundDoc = 0 Then
            If MsgBox("Speech document is not open - create one?", vbYesNo, "Create Speech?") = vbNo Then
                Exit Sub
            Else
                ' Create New Speech Doc
                Paperless.NewSpeech
            
                ' Switch focus back after save
                Documents(CurrentDoc).Activate
                GoTo SpeechDocCheck:
            End If
        End If
    
        ' If multiple Speech docs are open, warn the user.
        If FoundDoc > 1 Then
            UI.ShowForm "ChooseSpeechDoc"
            Exit Sub
        End If
    End If
    
    ' Turn off screen updating for the heavy-lifting
    Application.ScreenUpdating = False
    
    ' If text is selected, copy and send it.  Add a return if not in the selection.
    If Selection.End > Selection.Start Then
        Selection.Copy
        
        ' Trap for sending to middle of text
        If SpeechDoc.ActiveWindow.Selection.Start <> SpeechDoc.ActiveWindow.Selection.Paragraphs(1).Range.Start Then
            If MsgBox("Sending to the middle of text. Are you sure?", vbOKCancel) = vbCancel Then Exit Sub
        End If
        
        SpeechDoc.ActiveWindow.Selection.Paste
        If Selection.Characters.Last.Text <> Chr(13) Then
            SpeechDoc.ActiveWindow.Selection.TypeParagraph
        End If
        Exit Sub
    End If
    
    ' If nothing is selected, select the current card, block, hat or pocket
    Paperless.SelectHeadingAndContent
        
    ' If still nothing selected, exit
    If Selection.Start = Selection.End Then Exit Sub
        
    ' Copy the unit
    Selection.Copy
    
    ' Trap for sending to middle of text or sending a card into a block/hat
    If SpeechDoc.ActiveWindow.Selection.Start <> SpeechDoc.ActiveWindow.Selection.Paragraphs(1).Range.Start Then
       If MsgBox("Sending to the middle of text. Are you sure?", vbOKCancel) = vbCancel Then Exit Sub
    End If
    If Selection.Paragraphs(1).outlineLevel = 4 Then
        If SpeechDoc.ActiveWindow.Selection.Paragraphs.outlineLevel < wdOutlineLevel4 Then
            If MsgBox("Sending a card into a block, hat, or pocket.  Are you sure?", vbOKCancel) = vbCancel Then Exit Sub
        End If
    End If
    
    ' Paste it
    SpeechDoc.ActiveWindow.Selection.Paste
        
    ' Reset Selection
    Selection.Collapse
    
    Set SpeechDoc = Nothing
    
    Application.ScreenUpdating = True
    
    Exit Sub

Handler:
    Application.ScreenUpdating = True
    Set SpeechDoc = Nothing
    
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

'*************************************************************************************
'* DOCUMENT FUNCTIONS                                                                *
'*************************************************************************************

Sub NewDocument()
' Adds a new document based on the debate template
    Application.Documents.Add Template:=Application.NormalTemplate.Path & Application.PathSeparator & "Debate.dotm"
End Sub

Sub NewSpeech()
' Creates a new Speech document
    Dim SpeechName As String
    Dim FileName As String
    Dim h
    Dim AutoSaveDirectory As String
 
    On Error GoTo Handler
    
SpeechName:
    ' Get input for which Speech to name it
    SpeechName = InputBox("Which Speech (1NC, 2AC, etc...)? You can also add extra info about the round.", "New Speech", "e.g. 2AC Round 3 vs Hogwarts")
    If SpeechName = "" Then Exit Sub
    If SpeechName = "e.g. 2AC Round 3 vs Hogwarts" Then GoTo SpeechName
    SpeechName = Trim(ScrubString(SpeechName))
    SpeechName = Replace(SpeechName, "/", "")
    SpeechName = Replace(SpeechName, "\", "")
    
    ' Create filename
    If Hour(Now) > 12 Then h = Hour(Now) - 12 & "PM"
    If Hour(Now) <= 12 Then h = Hour(Now) & "AM"
    FileName = "Speech " & SpeechName & " " & Month(Now) & "-" & Day(Now) & " " & h

    ' Add new document based on template
    Paperless.NewDocument
 
    ' If AutoSave is set, save the doc - otherwise bring up Save As dialogue with default name set
    If GetSetting("Verbatim", "Paperless", "AutoSaveSpeech", False) = True Then
        AutoSaveDirectory = GetSetting("Verbatim", "Paperless", "AutoSaveDir", CurDir())
        If Right(AutoSaveDirectory, 1) <> "\" Then AutoSaveDirectory = AutoSaveDirectory & "\"
        FileName = AutoSaveDirectory & "\" & FileName
        ActiveDocument.SaveAs FileName:=FileName, FileFormat:=wdFormatXMLDocument
    Else
        With Application.Dialogs(wdDialogFileSaveAs)
            .Name = FileName
            If .Show = 0 Then Exit Sub
        End With
    End If
    
    Exit Sub
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

'*************************************************************************************
'* TOOL FUNCTIONS                                                                    *
'*************************************************************************************

Sub CopyToUSB()
' Copies the current file to the root folder of the first found USB drive

    #If Mac Then
        Dim POSIXActive
        Dim FileName As String
        
        Dim MountPoints As String
        Dim MountPointArray
        Dim m
    
        On Error GoTo Handler
        
        ' Get full POSIX path of current file
        POSIXActive = MacScript("return POSIX path of """ & ActiveDocument.FullName & """")
        
        ' Get list of mounted USB drives - throws an error if none plugged in, so turn off error checking temporarily
        On Error Resume Next
        MountPoints = AppleScriptTask("Verbatim.scpt", "RunShellScript", "system_profiler SPUSBDataType | grep 'Mount Point'")
        On Error GoTo Handler
        
        ' Exit if no USB drives found
        If MountPoints = "" Then
            MsgBox "No USB drives found!"
            Exit Sub
        End If
        
        ' Split into array and loop each drive
        MountPointArray = Split(MountPoints, Chr(13))
        For Each m In MountPointArray
            m = Trim(Replace(m, "Mount Point: ", "")) & "/" ' Get just the mount path and add a trailing /
            
            ' Strip "Speech" if option set
            If GetSetting("Verbatim", "Paperless", "StripSpeech", True) = True And Len(ActiveDocument.Name) > 11 Then
                FileName = Trim(Replace(ActiveDocument.Name, "speech", "", 1, -1, vbTextCompare))
            Else
                FileName = ActiveDocument.Name
            End If
            
            ' Check if file already exists on USB
            If AppleScriptTask("Verbatim.scpt", "RunShellScript", "test -e '" & m & FileName & "'; echo $?") = "0" Then
                If MsgBox("File Exists.  Overwrite?", vbOKCancel) = vbCancel Then Exit Sub
            End If
            
            ' Save File locally
            ActiveDocument.Save
            
            ' Copy To USB
            AppleScriptTask "Verbatim.scpt", "RunShellScript", "cp '" & POSIXActive & "' '" & m & FileName & "'"
            MsgBox "Sucessfully copied to USB!"
            
        Next m
        
        Exit Sub
    
    #Else
        Dim FSO As Scripting.FileSystemObject
        Set FSO = New Scripting.FileSystemObject
        Dim Drv As Drives
        Set Drv = FSO.Drives
        Dim d
        Dim USB
        Dim FileName As String
        
        On Error GoTo Handler
        
        ' Find USB Drive
        For Each d In Drv
            If d.DriveType = 1 Then
                USB = d
                Exit For
            End If
        Next
    
        ' If no drive found, exit
        If USB = 0 Then
            MsgBox "No USB Drive Found."
            Exit Sub
        End If
        
        ' Strip "Speech" if option set
        If GetSetting("Verbatim", "Paperless", "StripSpeech", True) = True Then
            FileName = Replace(ActiveDocument.Name, "speech", "", 1, -1, vbTextCompare)
        Else
            FileName = ActiveDocument.Name
        End If
        
        ' Check if file already exists on USB
        If FSO.FileExists(USB & "\" & FileName) = True Then
            If MsgBox("File Exists.  Overwrite?", vbOKCancel) = vbCancel Then Exit Sub
        End If
        
        ' Save File locally
        ActiveDocument.Save
        
        ' Copy To USB
        FSO.CopyFile ActiveDocument.FullName, USB & "\" & FileName
    
        MsgBox "Sucessfully copied to USB!"
    
        Set FSO = Nothing
        Set Drv = Nothing
    
        Exit Sub
    #End If
    
Handler:
    #If Not Mac Then
        Set FSO = Nothing
        Set Drv = Nothing
    #End If
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Sub StartTimer()
' Starts a user supplied timer.
' To install, place any timer in the same folder as Debate.dotm, and name the executable "Timer.exe"

    #If Mac Then
        Dim TimerApp As String
        Dim TimerAppPOSIX As String
        
        On Error GoTo Handler
        
        ' Get path to timer app
        TimerApp = GetSetting("Verbatim", "Paperless", "TimerApp", "?")
        
        ' If not set, try default
        If TimerApp = "?" Then TimerApp = MacScript("return path to applications folder as string") & "Debate Timer for Mac.app"
    
        ' Make sure timer app exists
        #If MAC_OFFICE_VERSION >= 15 Then
        If AppleScriptTask("Verbatim.scpt", "FileExists", TimerApp) = "false" Then
        #Else
        If MacScript("tell application ""Finder""" & Chr(13) & "exists file """ & TimerApp & """" & Chr(13) & "end tell") = "false" Then
        #End If
            MsgBox "Timer application not found. Ensure you have one installed and enter the correct path to the application in the Verbatim Settings." & vbCrLf & vbCrLf & "See the Verbatim manual on paperlessdebate.com for suggestions of Mac timer programs."
            Exit Sub
        Else
            ' If java app selected, run it from the shell
            If Right(TimerApp, 5) = ".jar:" Or Right(TimerApp, 4) = ".jar" Then
                TimerAppPOSIX = MacScript("return POSIX path of """ & TimerApp & """")
                #If MAC_OFFICE_VERSION >= 15 Then
                    AppleScriptTask "Verbatim.scpt", "RunShellScript", "open '" & TimerAppPOSIX & "'"
                #Else
                    MacScript ("do shell script ""open '" & TimerAppPOSIX & "'""")
                #End If
            Else
                #If MAC_OFFICE_VERSION >= 15 Then
                    AppleScriptTask "Verbatim.scpt", "ActivateTimer", TimerApp
                #Else
                    MacScript ("tell application """ & TimerApp & """ to activate")
                #End If
            End If
        End If
    
        Exit Sub
    #Else
        Dim FSO As Scripting.FileSystemObject
        Set FSO = New Scripting.FileSystemObject
        
        On Error GoTo Handler
        
        ' Check timer exists
        If FSO.FileExists(ActiveDocument.AttachedTemplate.Path & "\Timer.exe") = False Then
            MsgBox "Timer not found.  Make sure Timer.exe is in the same folder as your template."
            Exit Sub
        End If
        
        ' Run Timer
        Shell ActiveDocument.AttachedTemplate.Path & "\Timer.exe", vbNormalFocus
    
        Set FSO = Nothing
    
        Exit Sub
    #End If
    
Handler:
    #If Not Mac Then
        Set FSO = Nothing
    #End If
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

'*************************************************************************************
'* WARRANT FUNCTIONS                                                                 *
'*************************************************************************************

Sub NewWarrant()
    Selection.Comments.Add Range:=Selection.Range
End Sub

Sub DeleteAllWarrants()
    Dim c As Comment
    For Each c In ActiveDocument.Comments
        c.Delete
    Next c
End Sub

