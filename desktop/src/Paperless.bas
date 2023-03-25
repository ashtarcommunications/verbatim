Attribute VB_Name = "Paperless"
Option Explicit

'*************************************************************************************
'* RIBBON FUNCTIONS                                                                  *
'*************************************************************************************

Sub AutoOpenFolder(c As IRibbonControl, pressed As Boolean)
' Runs in the background to automatically open all documents in the speech folder.

    Dim AutoOpenDir As String
    
    #If Mac Then
        Dim Files
        Dim f
    #Else
        Dim Files
        Dim FSO As Object
        Dim f
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
            If Right(AutoOpenDir, 1) <> Application.PathSeparator Then AutoOpenDir = AutoOpenDir & Application.PathSeparator
        #End If
        
        ' Prompt that listener is on
        If MsgBox("This will start a listener that automatically opens all documents in the root of your Auto Open folder:" & vbCrLf & AutoOpenDir, vbOKCancel) = vbCancel Then
            Globals.AutoOpenFolderToggle = False
            Ribbon.RefreshRibbon
            Exit Sub
        End If
        
        Globals.AutoOpenFolderToggle = True
        
        #If Mac Then
            ' Do Nothing
        #Else
            Set FSO = CreateObject("Scripting.FileSystemObject")
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
    #If Mac Then
        ' Do Nothing
    #Else
        Set FSO = Nothing
        Set Files = Nothing
    #End If
    
    Exit Sub
    
Handler:
    #If Mac Then
        ' Do Nothing
    #Else
        Set FSO = Nothing
        Set Files = Nothing
    #End If
    Ribbon.RefreshRibbon
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Sub GetSpeeches(c As IRibbonControl, ByRef returnedVal)

    Dim xml As String
    
    On Error Resume Next
        
    ' Set Mouse Pointer
    System.Cursor = wdCursorWait

    ' Initialize XML
    xml = "<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">"

    If GetSetting("Verbatim", "Profile", "DisableTabroom", False) = False Then
    
        Application.StatusBar = "Retrieving rounds from openCaselist..."
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
    
        Application.StatusBar = "Retrieved rounds from openCaselist"
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
    End If
        
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

Sub NewSpeechFromMenu(c As IRibbonControl)
    Dim AutoSaveDirectory As String
    Dim FileName As String
    Dim h

    ' Add a new document based on the template
    Paperless.NewDocument

    ' Get filename from control tag
    FileName = c.Tag
    
    ' If Tag is just the speech name, add a date
    If Len(FileName) = 3 Then
        If Hour(Now) = 12 Then h = "12PM"
        If Hour(Now) > 12 Then h = Hour(Now) - 12 & "PM"
        If Hour(Now) < 12 Then h = Hour(Now) & "AM"
        If Hour(Now) = 0 Then h = "12AM"
        FileName = FileName & " " & Month(Now) & "-" & Day(Now) & " " & h
    End If
    
    ' Add speech to the name
    FileName = "Speech " & FileName
    
    ' Autosave or open save dialog
    If GetSetting("Verbatim", "Paperless", "AutoSaveSpeech", False) = True Then
        AutoSaveDirectory = GetSetting("Verbatim", "Paperless", "AutoSaveDir", CurDir())
        If Right(AutoSaveDirectory, 1) <> Application.PathSeparator Then AutoSaveDirectory = AutoSaveDirectory & Application.PathSeparator
        FileName = AutoSaveDirectory & Application.PathSeparator & FileName
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

Public Sub SelectCardText()

    Paperless.SelectHeadingAndContent
    
    Do While True
        If Selection.Paragraphs.Count < 2 Then Exit Do

        If Selection.Paragraphs(1).outlineLevel <> wdOutlineLevelBodyText Then
            Selection.MoveStart Unit:=wdParagraph, Count:=1
        ' Ignore paragraphs starting with [, (, or < as they're likely 2-line cites
        ElseIf Left(Selection.Paragraphs(1).Range.Text, 1) Like "[\[(<]" Then
            Selection.MoveStart Unit:=wdParagraph, Count:=1
        Else
            With Selection.Paragraphs(1).Range.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = ""
                .Forward = True
                .Wrap = wdFindStop
                .Format = True
                .Style = "Cite"
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
                
                .Execute
                
                If .Found Then
                    Selection.MoveStart Unit:=wdParagraph, Count:=1
                Else
                    Exit Do
                End If
                
                .ClearFormatting
                .Replacement.ClearFormatting
            End With
        End If
    Loop
End Sub

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
Sub SendToSpeechCursor()
    Paperless.SendToSpeech PasteAtEnd:=False
End Sub
Sub SendToSpeechEnd()
    Paperless.SendToSpeech PasteAtEnd:=True
End Sub

Sub SendToSpeech(Optional PasteAtEnd As Boolean)
' Sends content to the Speech doc.  Sends currently selected text,
' or if nothing is selected, the current tag, block, hat, or pocket
' If in reading view, enters a stopped reading marker at the current location

    Dim CurrentPage As Long
    Dim CurrentDoc As String
    Dim SpeechDoc As Document
    Dim d As Document
    Dim FoundDoc As Long

    On Error GoTo Handler

    ' If in reading mode, enter a stopped reading marker
    If ActiveWindow.View.ReadingLayout Then
        Application.ScreenUpdating = False
        CurrentPage = ActiveWindow.ActivePane.Selection.Information(wdActiveEndPageNumber)
        
        ActiveWindow.View = wdWebView
        Selection.Collapse Direction:=wdCollapseEnd
        If Selection.Words(1).End <> Selection.End Then Selection.MoveRight wdWord
        Selection.Font.Color = wdColorRed
        Selection.Font.Size = 18
        Selection.TypeText Chr(167) & " Marked " & FormatDateTime(Time, 4) & " " & Chr(167) & " "
        ActiveWindow.View = wdReadingView
        
        With ActiveWindow.ActivePane
            If CurrentPage > .Selection.Information(wdActiveEndPageNumber) Then
                .PageScroll Down:=CurrentPage - .Selection.Information(wdActiveEndPageNumber)
            ElseIf CurrentPage < .Selection.Information(wdActiveEndPageNumber) Then
                .PageScroll Up:=CurrentPage - .Selection.Information(wdActiveEndPageNumber)
            End If
        End With
        Application.ScreenUpdating = True
        
        Exit Sub
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
       
    ' Use selection or the current heading
    If Selection.End > Selection.Start Then
        Selection.Copy
    Else
        Paperless.SelectHeadingAndContent
        Selection.Copy
    End If
                
    ' If still nothing selected, exit
    If Selection.Start = Selection.End Then Exit Sub
          
    Dim SpeechDocStart
    Dim SpeechDocEnd
          
    If PasteAtEnd = True Then
        SpeechDocStart = SpeechDoc.ActiveWindow.Selection.Start
        SpeechDocEnd = SpeechDoc.ActiveWindow.Selection.End
        
        SpeechDoc.ActiveWindow.Selection.EndKey Unit:=wdStory
        SpeechDoc.ActiveWindow.Selection.InsertParagraph
    Else
        ' Trap for sending to middle of text or sending a card into a block/hat
        If SpeechDoc.ActiveWindow.Selection.Start <> SpeechDoc.ActiveWindow.Selection.Paragraphs(1).Range.Start Then
           If MsgBox("Sending to the middle of text. Are you sure?", vbOKCancel) = vbCancel Then Exit Sub
        End If
        If Selection.Paragraphs(1).outlineLevel = 4 Then
            If SpeechDoc.ActiveWindow.Selection.Paragraphs.outlineLevel < wdOutlineLevel4 Then
                If MsgBox("Sending a card into a block, hat, or pocket.  Are you sure?", vbOKCancel) = vbCancel Then Exit Sub
            End If
        End If
    End If
   
    ' Paste it and add a return if necessary
    SpeechDoc.ActiveWindow.Selection.Paste
    If Selection.Characters.Last.Text <> Chr(13) Then
        SpeechDoc.ActiveWindow.Selection.TypeParagraph
    End If
        
    ' Reset Selection
    Selection.Collapse
    If PasteAtEnd = True Then
        SpeechDoc.ActiveWindow.Selection.Start = SpeechDocStart
        SpeechDoc.ActiveWindow.Selection.End = SpeechDocEnd
    End If
    
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
    SpeechName = Replace(SpeechName, Application.PathSeparator, "")
    
    ' Create filename
    If Hour(Now) = 12 Then h = "12PM"
    If Hour(Now) > 12 Then h = Hour(Now) - 12 & "PM"
    If Hour(Now) < 12 Then h = Hour(Now) & "AM"
    If Hour(Now) = 0 Then h = "12AM"
    FileName = "Speech " & SpeechName & " " & Month(Now) & "-" & Day(Now) & " " & h

    ' Add new document based on template
    Paperless.NewDocument
 
    ' If AutoSave is set, save the doc - otherwise bring up Save As dialogue with default name set
    If GetSetting("Verbatim", "Paperless", "AutoSaveSpeech", False) = True Then
        AutoSaveDirectory = GetSetting("Verbatim", "Paperless", "AutoSaveDir", CurDir())
        If Right(AutoSaveDirectory, 1) <> Application.PathSeparator Then AutoSaveDirectory = AutoSaveDirectory & Application.PathSeparator
        FileName = AutoSaveDirectory & FileName
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

    Dim FileName As String
    
    ' Strip "Speech" if option set
    If GetSetting("Verbatim", "Paperless", "StripSpeech", True) = True And Len(ActiveDocument.Name) > 11 Then
        FileName = Trim(Replace(ActiveDocument.Name, "speech", "", 1, -1, vbTextCompare))
    Else
        FileName = Trim(ActiveDocument.Name)
    End If
    
    ' Save File locally
    ActiveDocument.Save
            
    #If Mac Then
        Dim MountPoints As String
        Dim MountPointArray
        Dim m
    
        On Error GoTo Handler
        
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
            
            ' Check if file already exists on USB
            If AppleScriptTask("Verbatim.scpt", "RunShellScript", "test -e '" & m & FileName & "'; echo $?") = "0" Then
                If MsgBox("File Exists.  Overwrite?", vbOKCancel) = vbCancel Then Exit Sub
            End If
            
            ' Copy To USB
            AppleScriptTask "Verbatim.scpt", "RunShellScript", "cp '" & ActiveDocument.FullName & "' '" & m & FileName & "'"
            MsgBox "Sucessfully copied to USB!"
        Next m
        
        Exit Sub
    
    #Else
        Dim FSO As Object
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Dim Drv
        Set Drv = FSO.Drives
        Dim d
        Dim USB
        
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
               
        ' Check if file already exists on USB
        If Filesystem.FileExists(USB & Application.PathSeparator & FileName) = True Then
            If MsgBox("File Exists.  Overwrite?", vbOKCancel) = vbCancel Then Exit Sub
        End If
        
        
        ' Copy To USB
        FSO.CopyFile ActiveDocument.FullName, USB & Application.PathSeparator & FileName
    
        MsgBox "Sucessfully copied to USB!"
    
        Set FSO = Nothing
        Set Drv = Nothing
    
        Exit Sub
    #End If
    
Handler:
    #If Mac Then
        ' Do Nothing
    #Else
        Set FSO = Nothing
        Set Drv = Nothing
    #End If
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Sub StartTimer()
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
    
        ' Make sure timer app exists
        If AppleScriptTask("Verbatim.scpt", "FileExists", TimerPath) = "false" Then
            MsgBox "Timer application not found. Ensure you have one installed and enter the correct path to the application in the Verbatim Settings." & vbCrLf & vbCrLf & "See the Verbatim manual on paperlessdebate.com for suggestions of Mac timer programs."
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
            TimerPath = Environ("ProgramW6432" & Application.PathSeparator & "Verbatim\Plugins\Timer.exe")
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

'*************************************************************************************
'* QUICK CARD FUNCTIONS                                                              *
'*************************************************************************************
Public Sub AddQuickCard()
    Dim t As Template
    Dim Name As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    If Selection.Start = Selection.End Then
        MsgBox "You must select some text to save a Quick Card", vbOKOnly
        Exit Sub
    End If
    
    Name = InputBox("What shortcut word/phrase do you want to use for your Quick Card? Usually this is the author's last name.", "Add Quick Card", "")
    If Name = "" Then Exit Sub

    Set t = ActiveDocument.AttachedTemplate
          
    For i = 1 To t.BuildingBlockTypes.Count
        If t.BuildingBlockTypes(i).Name = "Custom 1" Then
            For j = 1 To t.BuildingBlockTypes(i).Categories.Count
                If t.BuildingBlockTypes(i).Categories(j).Name = "Verbatim" Then
                    For k = 1 To t.BuildingBlockTypes(i).Categories(j).BuildingBlocks.Count
                        If LCase(t.BuildingBlockTypes(i).Categories(j).BuildingBlocks(k).Name) = LCase(Name) Then
                            MsgBox "There's already a Quick Card with that name, try again with a different name!", vbOKOnly, "Failed To Add Quick Card"
                            Exit Sub
                        End If
                    Next k
                End If
            Next j
        End If
    Next i
    
    t.BuildingBlockEntries.Add Name, wdTypeCustom1, "Verbatim", Selection.Range
    
    't.Save
    Ribbon.RefreshRibbon
    
    MsgBox "Successfully created Quick Card with the shortcut """ & Name & """"

    Set t = Nothing
    Exit Sub

Handler:
    Set t = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Public Sub InsertCurrentQuickCard()
    Selection.Range.InsertAutoText
End Sub

Public Sub InsertQuickCard(QuickCardName As String)
    Dim t As Template
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Set t = ActiveDocument.AttachedTemplate
    
    For i = 1 To t.BuildingBlockTypes.Count
        If t.BuildingBlockTypes(i).Name = "Custom 1" Then
            For j = 1 To t.BuildingBlockTypes(i).Categories.Count
                If t.BuildingBlockTypes(i).Categories(j).Name = "Verbatim" Then
                    For k = 1 To t.BuildingBlockTypes(i).Categories(j).BuildingBlocks.Count
                        If LCase(t.BuildingBlockTypes(i).Categories(j).BuildingBlocks(k).Name) = LCase(QuickCardName) Then
                            t.BuildingBlockTypes(i).Categories(j).BuildingBlocks(k).Insert Selection.Range, True
                        End If
                    Next k
                End If
            Next j
        End If
    Next i
    
    Set t = Nothing
    Exit Sub

Handler:
    Set t = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Public Sub DeleteQuickCard(Optional QuickCardName As String)
    Dim t As Template
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    If QuickCardName <> "" Or IsNull(QuickCardName) Then
        If MsgBox("Are you sure you want to delete the Quick Card """ & QuickCardName & """? This cannot be reversed.", vbYesNo, "Are you sure?") = vbNo Then Exit Sub
    Else
        If MsgBox("Are you sure you want to delete all saved Quick Cards? This cannot be reversed.", vbYesNo, "Are you sure?") = vbNo Then Exit Sub
    End If
    
    
    Set t = ActiveDocument.AttachedTemplate

    ' Delete all building blocks in the Custom 1/Verbatim category
    For i = 1 To t.BuildingBlockTypes.Count
        If t.BuildingBlockTypes(i).Name = "Custom 1" Then
            For j = 1 To t.BuildingBlockTypes(i).Categories.Count
                If t.BuildingBlockTypes(i).Categories(j).Name = "Verbatim" Then
                    For k = t.BuildingBlockTypes(i).Categories(j).BuildingBlocks.Count To 1 Step -1
                        ' If name provided, delete just that building block, otherwise delete everything in the category
                        If QuickCardName <> "" Or IsNull(QuickCardName) Then
                            If t.BuildingBlockTypes(i).Categories(j).BuildingBlocks(k).Name = QuickCardName Then
                                t.BuildingBlockTypes(i).Categories(j).BuildingBlocks(k).Delete
                            End If
                        Else
                            t.BuildingBlockTypes(i).Categories(j).BuildingBlocks(k).Delete
                        End If
                    Next k
                End If
            Next j
        End If
    Next i

    ' t.Save
    Set t = Nothing
        
    Exit Sub
Handler:
    Set t = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Public Sub GetQuickCardsContent(c As IRibbonControl, ByRef returnedVal)
' Get content for dynamic menu for Quick Cards
    Dim t As Template
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim xml As String
    Dim QuickCardName As String
    Dim DisplayName As String
    Dim Description As String
       
    On Error Resume Next
        
    Set t = ActiveDocument.AttachedTemplate

    ' Start the menu
    xml = "<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">"
    
    ' Populate the list of current Quick Cards in the Custom 1 / Verbatim gallery
    For i = 1 To t.BuildingBlockTypes.Count
        If t.BuildingBlockTypes(i).Name = "Custom 1" Then
            For j = 1 To t.BuildingBlockTypes(i).Categories.Count
                If t.BuildingBlockTypes(i).Categories(j).Name = "Verbatim" Then
                    For k = 1 To t.BuildingBlockTypes(i).Categories(j).BuildingBlocks.Count
                         QuickCardName = t.BuildingBlockTypes(i).Categories(j).BuildingBlocks(k).Name
                         DisplayName = Strings.OnlySafeChars(QuickCardName)
                        xml = xml & "<button id=""QuickCard" & Replace(DisplayName, " ", "") & """ label=""" & DisplayName & """ tag=""" & QuickCardName & """ onAction=""Paperless.InsertQuickCardFromRibbon"" imageMso=""AutoSummaryResummarize"" />"
                    Next k
                End If
            Next j
        End If
    Next i
    
    ' Close the menu
    xml = xml & "<button id=""QuickCardSettings"" label=""Quick Card Settings"" onAction=""Ribbon.RibbonMain"" imageMso=""AddInManager""" & " />"
    xml = xml & "</menu>"
    
    Set t = Nothing
    
    returnedVal = xml
        
    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub
Sub InsertQuickCardFromRibbon(c As IRibbonControl)
    Paperless.InsertQuickCard c.Tag
End Sub
