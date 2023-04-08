Attribute VB_Name = "Formatting"
'@IgnoreModule ProcedureNotUsed
Option Explicit

'@Ignore ParameterNotUsed
Public Sub UnderlineMode(ByVal c As IRibbonControl, ByVal pressed As Boolean)
' Ribbon callback for onAction of UnderlineMode togglebutton
    ' If ToggleButton is turned on
    If pressed Then
        Globals.UnderlineModeToggle = True
        MsgBox "Underline Mode is turned ON. Click the button again to turn off."
        Application.StatusBar = "Underline Mode on - press button on ribbon to cancel."
        
        ' Turns on a listener to automatically underline text until button pressed again
        Do
          DoEvents ' Give control back to application
          If Selection.Type > 1 Then
              If Selection.Paragraphs.OutlineLevel = wdOutlineLevelBodyText Then ' Only affect cards
                  If Selection.Font.Underline = wdUnderlineNone Then ' Testing for style here instead doesn't work
                      Selection.Style = "Underline"
                  Else
                      Selection.ClearFormatting
                  End If
                  Selection.Collapse 0 ' 0 Direction allows keyboard to underline to the right
              End If
          End If
        Loop Until Globals.UnderlineModeToggle = False ' Loop until button is pressed again
    Else
        Globals.UnderlineModeToggle = False
        MsgBox "Underline Mode is turned OFF."
    End If
End Sub

Public Sub ToggleUnderline()
' Toggles any style underlined text to Normal and back to Underline style
    
    ' Check for all underlinining, not a specific style, to be more universal
    If Selection.Font.Underline = 1 Then
        Selection.ClearFormatting
    Else
        Selection.Style = "Underline"
    End If
End Sub

Public Sub PasteText()
' Pastes unformatted text
    #If Mac Then
        ' Normal Clipboard DataObject is unreliable in Mac VBA and pastes extra characters,
        ' so use the built-in method instead
        Selection.PasteSpecial dataType:=wdPasteText
    #Else
        Dim Clipboard As DataObject
        Set Clipboard = New DataObject
        Dim Text As String
        
        ' Assign clipboard contents to string if text to stop screen moving while pasting
        Clipboard.GetFromClipboard
        If Clipboard.GetFormat(1) = False Then Exit Sub
        Text = Clipboard.GetText(1)
            
        Selection.Text = Text
        
        If GetSetting("Verbatim", "Format", "CondenseOnPaste", False) = True Then
            Condense.CondenseCard
        End If
    
        Selection.Collapse 0
    #End If
End Sub

Public Sub Highlight()
    WordBasic.Highlight
End Sub

Public Sub ClearToNormal()
' Clears all formatting if text is selected, otherwise sets the paragraph style to Normal
    If Selection.End > Selection.Start Then
        Selection.ClearFormatting
    Else
        Selection.Paragraphs.Style = ActiveDocument.Styles.Item("Normal").NameLocal
    End If
End Sub

Public Sub CopyPreviousCite()
' Duplicates previous cite - only works with one-line cites
    Dim StartLocation As Long
    
    ' Save Current Location
    StartLocation = Selection.Start
    
    ' Find previous cite
    With Selection.Find
      .ClearFormatting
      .Text = ""
      .Wrap = wdFindStop
      .Format = True
      .Style = ActiveDocument.Styles.Item("Cite").NameLocal
      .Forward = False
      .Execute
    End With
    
    ' If nothing found, exit
    If Selection.Find.Found = False Then
        Application.StatusBar = "No Cite Found"
        Exit Sub
    End If
    
    ' If found, select the whole paragraph
    Selection.Collapse
    Selection.StartOf Unit:=wdParagraph
    Selection.MoveEnd Unit:=wdParagraph, Count:=1
    Selection.Copy
    
    ' Return to original location and paste
    Selection.Start = StartLocation
    Selection.Collapse
    Selection.Paste
End Sub

Public Sub UniHighlight()
' Replaces all highlighting in the document with the selected color
    Dim r As Range
    Set r = ActiveDocument.Range
    
    With r.Find
        .ClearFormatting
        .Highlight = True
        .Replacement.ClearFormatting
        .Replacement.Highlight = True
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .Replacement.ClearFormatting
    End With
    
    Set r = Nothing
End Sub

Public Sub UniHighlightWithException()
    Dim r As Range
    Set r = ActiveDocument.Range
    
    Dim ExceptionColor As String
    ExceptionColor = GetSetting("Verbatim", "Format", "HighlightingException", "None")
    
    If ExceptionColor = "" Or ExceptionColor = "None" Then
        MsgBox "You don't have a highlighter exception color configured in the settings. Please set one and try again.", vbOKOnly
        Exit Sub
    End If
    
    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Highlight = True
        .Replacement.Highlight = True
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
       
        Do While .Execute(Forward:=True) = True
            If r.HighlightColorIndex = Formatting.HighlightColorToEnum(ExceptionColor) Then
                r.Collapse Direction:=wdCollapseEnd
            Else
                r.HighlightColorIndex = Options.DefaultHighlightColorIndex
            End If
        Loop
    
        .ClearFormatting
        .Replacement.ClearFormatting
    End With
    
    Set r = Nothing
End Sub

Public Function HighlightColorToEnum(ByVal Color As String) As Long
    Select Case Color
        Case Is = "None"
            HighlightColorToEnum = wdNoHighlight
        Case Is = "Black"
            HighlightColorToEnum = wdBlack
        Case Is = "Blue"
            HighlightColorToEnum = wdBlue
        Case Is = "Bright Green"
            HighlightColorToEnum = wdBrightGreen
        Case Is = "Dark Blue"
            HighlightColorToEnum = wdDarkBlue
        Case Is = "Dark Red"
            HighlightColorToEnum = wdDarkRed
        Case Is = "Dark Yellow"
            HighlightColorToEnum = wdDarkYellow
        Case Is = "Light Gray"
            HighlightColorToEnum = wdGray25
        Case Is = "Dark Gray"
            HighlightColorToEnum = wdGray50
        Case Is = "Green"
            HighlightColorToEnum = wdGreen
        Case Is = "Pink"
            HighlightColorToEnum = wdPink
        Case Is = "Red"
            HighlightColorToEnum = wdRed
        Case Is = "Teal"
            HighlightColorToEnum = wdTeal
        Case Is = "Turquoise"
            HighlightColorToEnum = wdTurquoise
        Case Is = "Violet"
            HighlightColorToEnum = wdViolet
        Case Is = "White"
            HighlightColorToEnum = wdWhite
        Case Is = "Yellow"
            HighlightColorToEnum = wdYellow
        Case Else
            HighlightColorToEnum = wdNoHighlight
    End Select
End Function

Public Sub RemoveBlanks()
' Removes blank lines from appearing in the Navigation Pane by setting them to Normal text
    Dim p As Paragraph

    ' Prompt user to confirm
    If MsgBox("Removing blank lines from the Nav Pane is irreversible. Are you sure?", vbOKCancel) = vbCancel Then Exit Sub

    For Each p In ActiveDocument.Paragraphs
        If p.OutlineLevel < wdOutlineLevel5 And p.Range.End - p.Range.Start <= 1 Then
            p.Style = "Normal"
        End If
    Next p
End Sub

Public Sub ShowComments()
' Toggles showing comments
    With ActiveWindow.View
        If .ShowRevisionsAndComments Then
            .ShowRevisionsAndComments = False
        Else
            .ShowRevisionsAndComments = True
            .MarkupMode = wdBalloonRevisions
        End If
    End With
End Sub

Public Sub InsertHeader()
' Inserts a custom header based on team/user information in Verbatim settings
    ActiveDocument.Sections.Item(1).Headers.Item(wdHeaderFooterPrimary).Range.Text = GetSetting("Verbatim", "Profile", "SchoolName", "") _
        & vbCrLf & "File Title" & vbTab & vbTab & GetSetting("Verbatim", "Profile", "Name", "")
    ActiveDocument.Sections.Item(1).Headers.Item(wdHeaderFooterPrimary).PageNumbers.Add (wdAlignPageNumberRight)
End Sub

Public Sub UpdateStyles()
' Updates document styles from template
    ActiveDocument.UpdateStyles
End Sub

Public Sub SelectSimilar()
    On Error Resume Next
    
    Application.ScreenUpdating = False
    
    If Selection.Font.Underline = wdUnderlineNone And Selection.Font.size = ActiveDocument.Styles.Item("Normal").Font.size Then
        ActiveDocument.Content.Font.Shrink
        WordBasic.SelectSimilarFormatting
        ActiveDocument.Content.Font.Grow
    Else
        WordBasic.SelectSimilarFormatting
    End If
    
    Application.ScreenUpdating = True
    
    On Error GoTo 0
End Sub

Public Sub RemoveHyperlinks()
' Remove all hyperlinks from document
    Dim i As Long
    Dim Count As Long
    
    For i = ActiveDocument.Hyperlinks.Count To 1 Step -1
        ActiveDocument.Hyperlinks.Item(i).Delete
        Count = Count + 1
    Next i

    Application.StatusBar = Count & " hyperlinks removed."
End Sub

Public Sub AutoFormatCite()
    Dim r As Range
       
    ' Set range to current paragraph
    Set r = Selection.Paragraphs.Item(1).Range
    
    ' Find first comma
    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Text = ","
        .Wrap = wdFindStop
        .Execute
    End With
    
    ' Select word before comma
    r.MoveStart Unit:=wdWord, Count:=-1
    r.MoveEnd Unit:=wdCharacter, Count:=-1
    
    ' If it's numeric, it's the year, so expand backwards one word to catch last name, make a cite, and exit
    If IsNumeric(r.Text) = True Then
        r.MoveStart Unit:=wdWord, Count:=-1
        r.Style = "Cite"
        Exit Sub
    Else ' If non-numeric, it's likely the last name, so make a cite
        r.Style = "Cite"
        r.Collapse 0
    End If
    
    ' Move right in paragraph until finding a digit - should be the start of the date, then extend to get whole date
    r.MoveStartUntil Cset:="0123456789", Count:=Len(Selection.Paragraphs.Item(1).Range.Text)
    r.MoveEndWhile Cset:="-/0123456789", Count:=Len(Selection.Paragraphs.Item(1).Range.Text)

    ' If end of date doesn't match current year, make year portion a cite
    If Right$(r.Text, 2) <> Right$(Year(Date), 2) Then
        r.Collapse 0
        r.MoveStartWhile Cset:="0123456789", Count:=-4
        r.Style = "Cite"
    
    ' Otherwise, skip year potion and extend backwards to make rest of date a cite
    Else
        r.Collapse 0
        r.MoveStartWhile Cset:="0123456789", Count:=-4
        r.Collapse
        r.MoveStartWhile Cset:="-/", Count:=-1
        r.Collapse
        r.MoveStartWhile Cset:="-/0123456789", Count:=-5
        r.Style = "Cite"
    End If

    Set r = Nothing
End Sub

Public Sub ReformatAllCites()
    Dim SelectionStart As Long
    Dim SelectionEnd As Long
    SelectionStart = Selection.Start
    SelectionEnd = Selection.End
    
    ' Go to top of document
    Selection.HomeKey Unit:=wdStory
    Selection.Collapse
    
    ' Find each occurrence of the Cite style
    ' Have to use selection instead of a range to have access to ClearFormatting
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .Style = "Cite"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute
        
        ' Find all matches
        Do Until .Found = False
            ' Select existing cite and clear formatting, then re-format
            Selection.Collapse
            Selection.StartOf Unit:=wdParagraph
            Selection.MoveEnd Unit:=wdParagraph, Count:=1
            Selection.ClearFormatting
            Formatting.AutoFormatCite
            
            ' Move down to avoid getting stuck
            Selection.MoveDown Unit:=wdParagraph, Count:=1
            .Execute
        Loop
        
        .ClearFormatting
        .Replacement.ClearFormatting
    End With
    
    Selection.Start = SelectionStart
    Selection.End = SelectionEnd
End Sub

Public Sub AutoUnderline()
    Dim TagWord As Range
    
    Dim SI As SynonymInfo
    Dim Synonyms() As String
    Dim TagSynonyms As Collection
    Set TagSynonyms = New Collection
    
    Dim i As Long
    Dim j As Long
    Dim k As Variant
    
    Dim w As Range
    Dim CardEnd As Long
    
    Dim IntersectionCount As Long
    
    Dim dict As Dictionary
    Set dict = New Dictionary
    
    On Error GoTo Handler
    
    ' If cursor isn't on a tag, exit
    Selection.Collapse
    If Selection.Paragraphs.OutlineLevel <> wdOutlineLevel4 Or Len(Selection.Paragraphs.Item(1).Range.Text) < 2 Then
        MsgBox "Cursor must be in a tag to automatically underline a card."
        Exit Sub
    End If
    
    ' Select tag and loop words, add each word and all synonyms if adjective, noun, adverb, or verb
    Selection.Paragraphs.Item(1).Range.Select
    For Each TagWord In Selection.Words
        TagSynonyms.Add TagWord.Text
        Set SI = SynonymInfo(TagWord.Text)
        If SI.MeaningCount > 0 Then
            For i = 1 To SI.MeaningCount
                If SI.PartOfSpeechList(i) < 4 Then ' 0=Adjective, 1=Noun, 2=Adverb, 3=Verb
                    Synonyms = SI.SynonymList(i)
                    For j = 1 To UBound(Synonyms)
                        TagSynonyms.Add Synonyms(j)
                    Next j
                End If
            Next i
        End If
    Next TagWord
        
    ' Select card, then deselect tag - exit if no more content
    Paperless.SelectCardText
    
    ' Set end point, then collapse and start processing card
    CardEnd = Selection.End
    Selection.Collapse
    Do While Selection.End <> CardEnd
        Selection.MoveEnd wdWord, 1
        Select Case Trim$(Selection.Words.Item(Selection.Words.Count).Text)
            ' Extend the chunk until we find punctuation, then roll back 1 word
            Case Is = ".", ",", """", "(", ")", "),", ":", ";", " - ", Chr$(151) ' 151 = em dash
                Selection.MoveEnd wdWord, -1
                
                ' Loop all words in chunk and compare to all words in TagSynonyms - count it if it matches
                For Each w In Selection.Words
                    For i = 1 To TagSynonyms.Count
                        If LCase$(Trim$(w.Text)) = LCase$(Trim$(TagSynonyms.Item(i))) Then IntersectionCount = IntersectionCount + 1
                    Next i
                Next w
                
                ' Add the range for the chunk and the normalized chunk score to the dictionary, then reset counter
                dict.Add Selection.Range, IntersectionCount / Selection.Words.Count
                IntersectionCount = 0
                
                ' Move one word right to skip punctuation and start new chunk
                Selection.MoveEnd wdWord, 1
                Selection.Collapse Direction:=0
        End Select
    Loop
    Selection.Collapse
    
    ' Loop through dictionary - key will be the range, so set style if chunk score is high enough
    ' 0.1 threshold is pretty good for underline relevance, 0.25 is pretty good for emphasis
    For Each k In dict.Keys
        If dict.Item(k) >= 0.1 Then k.Style = "Underline"
        If GetSetting("Verbatim", "Format", "AutoUnderlineEmphasis", False) = True Then
            If dict.Item(k) >= 0.25 Then k.Style = "Emphasis"
        End If
        ' Debug.Print "Range: " & k & vbTab & "Score: " & dict.Item(k)
    Next k
    
    ' Clean up
    Set TagSynonyms = Nothing
    Set dict = Nothing
    Set SI = Nothing

    Exit Sub
    
Handler:
    Set TagSynonyms = Nothing
    Set dict = Nothing
    Set SI = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Public Sub AutoEmphasizeFirst()
    Dim w As Range
    For Each w In Selection.Words
        w.Characters.Item(1).Style = "Emphasis"
    Next w
End Sub

Public Sub RemoveEmphasis()
    Dim r As Range
    Set r = ActiveDocument.Range
    
    If MsgBox("Are you sure you want to convert all emphasized text to underlined?", vbYesNo) = vbNo Then Exit Sub
        
    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .Style = "Emphasis"
        .Replacement.Style = "Underline"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .Replacement.ClearFormatting
    End With
    
    Set r = Nothing
End Sub

Public Sub AutoNumberTags()
    Dim p As Paragraph
    Dim i As Long
    
    ' Remove pre-existing numbers
    Formatting.DeNumberTags
    
    ' Loop through each tag and insert a number - restart numbering on any larger heading
    For Each p In ActiveDocument.Paragraphs
        Select Case p.OutlineLevel
            Case Is = 1, 2, 3
                i = 0
            Case Is = 4
                If Len(p.Range.Text) > 1 Then
                    i = i + 1
                    p.Range.InsertBefore i & ". "
                End If
            Case Else
                ' Do Nothing
        End Select
    Next p
End Sub

Public Sub DeNumberTags()
    Dim p As Paragraph
    Dim r As Range
    
    ' Loop through each tag
    For Each p In ActiveDocument.Paragraphs
        If p.OutlineLevel = 4 Then
            
            ' Delete numbers from beginning of line if there's a delimiter, then trim
            Set r = p.Range
            r.Collapse
            r.MoveEndWhile Cset:="0123456789.():-", Count:=5
            If Left$(r.Text, 5) Like "*[.():-]*" And r.Text <> "" Then
                r.Delete
                r.Collapse
                r.MoveEndWhile Cset:=" "
                If r.Text <> "" Then r.Delete
            End If
        End If
    Next p
    
    Set r = Nothing
End Sub

Public Sub FixFakeTags()
    Dim p As Paragraph
    
    For Each p In ActiveDocument.Paragraphs
        If p.OutlineLevel = wdOutlineLevelBodyText And p.Range.Bold = True And p.Range.Font.size > ActiveDocument.Styles.Item("Underline").Font.size Then
            p.Range.Select
            Selection.ClearFormatting
            p.Style = "Tag"
        End If
    Next p
End Sub

Public Sub ConvertAnalyticsToTags()
    Dim p As Paragraph
    
    For Each p In ActiveDocument.Paragraphs
        If LCase$(Left$(p.Style, 8)) = "analytic" Then p.Style = "Tag"
    Next p
End Sub

Public Sub RemoveExtraStyles()
    Dim s As Style
       
    Dim i As Long
    i = 0
    Dim StyleCount As Long
    StyleCount = ActiveDocument.Styles.Count
    Dim ProgressForm As frmProgress
    
    On Error GoTo Handler
    
    If MsgBox("WARNING: Removing extra styles can result in lost formatting, especially without running 'Convert To Default Styles' first, and can take a long time. Proceed?", vbYesNo, "Remove extra styles?") = vbNo Then Exit Sub
    
    Application.ScreenUpdating = False
    
    ' Show progress bar
    Set ProgressForm = New frmProgress
    ProgressForm.Caption = "Fixing style names..."
    ProgressForm.lblCaption.Caption = ActiveDocument.Styles.Count & " Remaining Styles..."
    ProgressForm.lblProgress.Width = 0
    ProgressForm.Show
    
    ' Visibility = True actually hides the style in the style pane
    ' Have to change the name before deleting to avoid Word crashing on long style names
    For Each s In ActiveDocument.Styles
        If ProgressForm.Visible = False Then Exit Sub
        
        i = i + 1
        ProgressForm.lblCaption.Caption = ActiveDocument.Styles.Count & " Remaining Styles..."
        ProgressForm.lblFile.Caption = "Processing " & s.NameLocal
        ProgressForm.lblProgress.Width = (i / StyleCount) * ProgressForm.fProgress.Width
        If ProgressForm.lblProgress.Width > 15 Then ProgressForm.lblProgress.Width = ProgressForm.lblProgress.Width - 15
    
        DoEvents ' Necessary for Progress form to update
    
        If s.Type = wdStyleTypeParagraph Or s.Type = wdStyleTypeLinked Then
            If Left$(s.NameLocal, 16) = "Heading 1,Pocket" And s.ParagraphFormat.OutlineLevel = wdOutlineLevel1 Then
                s.NameLocal = "Heading 1,Pocket"
                s.Visibility = False
            ElseIf Left$(s.NameLocal, 13) = "Heading 2,Hat" And s.ParagraphFormat.OutlineLevel = wdOutlineLevel2 Then
                s.NameLocal = "Heading 2,Hat"
                s.Visibility = False
            ElseIf Left$(s.NameLocal, 15) = "Heading 3,Block" And s.ParagraphFormat.OutlineLevel = wdOutlineLevel3 Then
                s.NameLocal = "Heading 3,Block"
                s.Visibility = False
            ElseIf Left$(s.NameLocal, 13) = "Heading 4,Tag" And s.ParagraphFormat.OutlineLevel = wdOutlineLevel4 Then
                s.NameLocal = "Heading 4,Tag"
                s.Visibility = False
            ElseIf Left$(s.NameLocal, 18) = "Normal,Normal/Card" And s.ParagraphFormat.OutlineLevel = wdOutlineLevelBodyText Then
                s.NameLocal = "Normal,Normal/Card"
                s.Visibility = False
            Else
                s.Visibility = True
                s.UnhideWhenUsed = False
                If s.BuiltIn = False Then
                    s.NameLocal = "DeleteMe"
                    s.Delete
                    ActiveDocument.UndoClear
                End If
            End If
        ElseIf s.Type = wdStyleTypeCharacter Then
            If Left$(s.NameLocal, 21) = "Style 13 pt Bold,Cite" Then
                s.NameLocal = "Style 13 pt Bold,Cite"
                s.Visibility = False
            ElseIf Left$(s.NameLocal, 25) = "Style Underline,Underline" Then
                s.NameLocal = "Style Underline,Underline"
                s.Visibility = False
            ElseIf s.NameLocal = "Emphasis" Or Left$(s.NameLocal, 9) = "Emphasis," Then
                s.NameLocal = "Emphasis"
                s.Visibility = False
            Else
                s.Visibility = True
                s.UnhideWhenUsed = False
                If s.BuiltIn = False Then
                    s.NameLocal = "DeleteMe"
                    s.Delete
                    ' The undo stack will crash Word when there's too many styles
                    ActiveDocument.UndoClear
                End If
            End If
        Else
            s.Visibility = True
            s.UnhideWhenUsed = False
            
            If s.BuiltIn = False Then
                s.NameLocal = "DeleteMe"
                s.Delete
                ActiveDocument.UndoClear
            End If
        End If
    Next s

    Application.ScreenUpdating = True
    
    ' Update progress form as complete
    ProgressForm.lblCaption.Caption = "Processing complete."
    ProgressForm.lblFile.Caption = ""
    ProgressForm.lblProgress.Width = ProgressForm.fProgress.Width - 6
    Unload ProgressForm
    Set ProgressForm = Nothing
    Exit Sub
    
Handler:
    If Not ProgressForm Is Nothing Then Unload ProgressForm
    Set ProgressForm = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Public Sub ConvertToDefaultStyles()
    Dim p As Paragraph
    Dim s As Style
    Dim r As Range

    Dim i As Long
    i = 0
    Dim StyleCount As Long
    StyleCount = ActiveDocument.Styles.Count * 2
    Dim ProgressForm As frmProgress
    
    On Error Resume Next
    
    If MsgBox("WARNING: Converting to default styles can take a long time, and sometimes loses formatting. Proceed?", vbYesNo, "Convert to Default Styles?") = vbNo Then Exit Sub
    
    Application.ScreenUpdating = False
    
    ' Show progress bar
    Set ProgressForm = New frmProgress
    ProgressForm.Caption = "Converting to default styles..."
    ProgressForm.lblProgress.Width = 0
    ProgressForm.lblCaption.Caption = "Converting headings..."
    ProgressForm.Show

    ' Convert all headings to built-in styles
    For Each p In ActiveDocument.Paragraphs
        ' Trap for cancel button on Progress Form
        If ProgressForm.Visible = False Then Exit Sub
        If p.OutlineLevel = wdOutlineLevel1 Then
            p.Style = wdStyleHeading1
        ElseIf p.OutlineLevel = wdOutlineLevel2 Then
            p.Style = wdStyleHeading2
        ElseIf p.OutlineLevel = wdOutlineLevel3 Then
            p.Style = wdStyleHeading3
        ElseIf p.OutlineLevel = wdOutlineLevel4 Then
            p.Style = wdStyleHeading4
        End If
    Next p
    
    ProgressForm.lblCaption.Caption = "Fixing style names..."
    
    ' Fix style names for built-in styles to prevent other styles overwriting the name
    For Each s In ActiveDocument.Styles
        If ProgressForm.Visible = False Then Exit Sub
        
        i = i + 1
        ProgressForm.lblFile.Caption = "Processing " & s.NameLocal & " (" & i & " of " & StyleCount & ")"
        ProgressForm.lblProgress.Width = (i / StyleCount) * ProgressForm.fProgress.Width
        If ProgressForm.lblProgress.Width > 15 Then ProgressForm.lblProgress.Width = ProgressForm.lblProgress.Width - 15
    
        DoEvents ' Necessary for Progress form to update
    
        If Left$(s.NameLocal, 21) = "Style 13 pt Bold,Cite" Then
            s.NameLocal = "Style 13 pt Bold,Cite"
        ElseIf Left$(s.NameLocal, 25) = "Style Underline,Underline" Then
            s.NameLocal = "Style Underline,Underline"
        ElseIf Left$(s.NameLocal, 9) = "Emphasis," Then
            s.NameLocal = "Emphasis"
        End If
    Next s

    ProgressForm.lblCaption.Caption = "Converting styles..."

    For Each s In ActiveDocument.Styles
        If ProgressForm.Visible = False Then Exit Sub
        
        i = i + 1
        ProgressForm.lblFile.Caption = "Processing " & s.NameLocal & " (" & i & " of " & StyleCount & ")"
        ProgressForm.lblProgress.Width = (i / StyleCount) * ProgressForm.fProgress.Width
        If ProgressForm.lblProgress.Width > 15 Then ProgressForm.lblProgress.Width = ProgressForm.lblProgress.Width - 15
    
        DoEvents ' Necessary for Progress form to update
  
        ' Ignore headings converted above
        If s.BuiltIn = False _
            And Left$(LCase$(s.NameLocal), 6) <> "pocket" _
            And Left$(LCase$(s.NameLocal), 3) <> "hat" _
            And Left$(LCase$(s.NameLocal), 5) <> "block" _
            And Left$(LCase$(s.NameLocal), 3) <> "tag" _
        Then
            ' Convert cite styles
            If InStr(LCase$(s.NameLocal), "cite") > 0 _
                And Left$(s.NameLocal, 25) <> "Style Underline,Underline" _
                And Left$(LCase$(s.NameLocal), 9) <> "underline" _
                And Left$(LCase$(s.NameLocal), 8) <> "emphasis" _
                And Left$(LCase$(s.NameLocal), 6) <> "normal" _
                And Left$(LCase$(s.NameLocal), 4) <> "card" _
                And Left$(LCase$(s.NameLocal), 8) <> "analytic" _
                And Left$(LCase$(s.NameLocal), 4) <> "body" _
            Then
                With ActiveDocument.Range.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .Text = ""
                    .Replacement.Text = ""
                    .Format = True
                    .Wrap = wdFindContinue
                    .Style = s.NameLocal
                    .Replacement.Style = "Cite"
                    
                    .Execute Replace:=wdReplaceAll
                    
                    .ClearFormatting
                    .Replacement.ClearFormatting
                End With
            
            ' Convert emphasis styles
            ElseIf InStr(LCase$(s.NameLocal), "emphasi") > 0 And Left$(s.NameLocal, 25) <> "Style Underline,Underline" Then
                
                ' Fixes weird linked emphasis styles that don't show up in Find
                If s.Linked Then s.LinkStyle = "Normal"
                
                ' Emphasis with highlighting
                Set r = ActiveDocument.Range
                With r.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .Text = ""
                    .Replacement.Text = ""
                    .Format = True
                    .Wrap = wdFindStop
                    .Style = s.NameLocal
                    .Replacement.Style = "Emphasis"
                    .Highlight = True
                    .Replacement.Highlight = True
                    .Execute
                    
                    Do While .Found = True
                        r.Style = "Emphasis"
                        r.HighlightColorIndex = wdYellow
                        .Execute
                    Loop
                    
                    .ClearFormatting
                    .Replacement.ClearFormatting
                End With
                
                ' Emphasis without highlighting
                Set r = ActiveDocument.Range
                With r.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .Text = ""
                    .Replacement.Text = ""
                    .Format = True
                    .Wrap = wdFindStop
                    .Style = s.NameLocal
                    .Replacement.Style = "Emphasis"
                    .Highlight = False
                    .Replacement.Highlight = False
                    
                    .Execute Replace:=wdReplaceAll
                    
                    .ClearFormatting
                    .Replacement.ClearFormatting
                End With
            
            ' Convert underline styles
            ElseIf InStr(LCase$(s.NameLocal), "underline") > 0 _
                And InStr(LCase$(s.NameLocal), "no underline") = 0 _
                And InStr(LCase$(s.NameLocal), "not underline") = 0 _
                And InStr(LCase$(s.NameLocal), "ununderline") = 0 _
                And InStr(LCase$(s.NameLocal), "un-underline") = 0 _
                And InStr(LCase$(s.NameLocal), "non underline") = 0 _
                And InStr(LCase$(s.NameLocal), "non-underline") = 0 _
            Then
                Set r = ActiveDocument.Range
                With r.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .Text = ""
                    .Replacement.Text = ""
                    .Format = True
                    .Wrap = wdFindStop
                    .Style = s.NameLocal
                    .Replacement.Style = "Style Underline,Underline"
                    .Highlight = True
                    .Replacement.Highlight = True
                    .Execute

                    Do While .Found = True
                        r.Style = "Style Underline,Underline"
                        r.HighlightColorIndex = wdYellow
                        .Execute
                    Loop
                    
                    .ClearFormatting
                    .Replacement.ClearFormatting
                End With
                
                With ActiveDocument.Range.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .Text = ""
                    .Replacement.Text = ""
                    .Format = True
                    .Wrap = wdFindContinue
                    .Style = s.NameLocal
                    .Replacement.Style = "Style Underline,Underline"
                    .Highlight = False
                    .Replacement.Highlight = False
                    
                    .Execute Replace:=wdReplaceAll
                    
                    .ClearFormatting
                    .Replacement.ClearFormatting
                End With
            End If
        End If
    Next s
    
    Application.ScreenUpdating = True
    
    ' Update progress form as complete
    ProgressForm.lblCaption.Caption = "Processing complete."
    ProgressForm.lblFile.Caption = ""
    ProgressForm.lblProgress.Width = ProgressForm.fProgress.Width - 6
    Unload ProgressForm
    Set ProgressForm = Nothing
    
    On Error GoTo 0
End Sub

Public Function LargestHeading() As Long
    LargestHeading = 3
      
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .Style = "Hat"
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute
        
        .ClearFormatting
        .Replacement.ClearFormatting
        
        If .Found Then LargestHeading = 2
    End With
    
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .Style = "Pocket"
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute
        
        .ClearFormatting
        .Replacement.ClearFormatting
        
        If .Found Then LargestHeading = 1
    End With
End Function

Public Sub RemoveNonHighlightedUnderlining()
    Dim r As Range
    
    If Selection.Start <= ActiveDocument.Range.Start And Selection.End = ActiveDocument.Range.Start Then
        If MsgBox("This will remove non-highlighted underlining for all cards in the document. Are you sure?", vbOKCancel) = vbCancel Then Exit Sub
        Selection.WholeStory
    ElseIf Selection.Start = Selection.End Then
        Paperless.SelectCardText
    End If
    
    ' Save a duplicate range to limit the Find to the current selection
    Set r = Selection.Range
    
    ' Have to use Selection instead of a range to have access to ClearFormatting
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .Highlight = False
        .Style = "Underline"
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        Do While .Execute(Forward:=True) = True And Selection.Range.InRange(r)
            ' Directly setting the style to Normal doesn't work, so ClearFormatting instead
            Selection.ClearFormatting
        Loop
                        
        .ClearFormatting
        .Replacement.ClearFormatting
    End With
    
    ' Restore the selection
    Selection.Start = r.Start
    Selection.End = r.End
End Sub

Public Sub FixFormattingGaps()
    Dim SelectionStart As Long
    Dim SelectionEnd As Long
    Dim r As Range
    Dim c As Characters
    Dim i As Long

    SelectionStart = Selection.Start
    SelectionEnd = Selection.End

    If Selection.Start = Selection.End Then Paperless.SelectCardText
    
    If Selection.End - Selection.Start > 5000 Then
        If MsgBox("Repairing card formatting for a long card can take a few minutes. Proceed?", vbYesNo) = vbNo Then Exit Sub
    End If
    
    ' Save a duplicate range to limit the Find to the current selection
    Set r = Selection.Range
    
    ' Find ranges in-between words with spaces/punctuation
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "[0-9A-z][\.\,\;\:\?\(\)\-\! ]{1,}[0-9A-z]"
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        Do While .Execute(Forward:=True) = True And Selection.Range.InRange(r)
            Set c = Selection.Range.Characters
            
            ' Underline or emphasize based on surrounding styles
            If Left$(c.Item(1).Style, 9) = "Underline" And Left$(c.Item(c.Count).Style, 9) = "Underline" Then
                Selection.Style = "Underline"
            ElseIf Left$(c.Item(1).Style, 8) = "Emphasis" And Left$(c.Item(c.Count).Style, 8) = "Emphasis" Then
                Selection.Style = "Emphasis"
            ElseIf Left$(c.Item(1).Style, 9) = "Underline" And Left$(c.Item(c.Count).Style, 8) = "Emphasis" Then
                Selection.Style = "Underline"
            ElseIf Left$(c.Item(1).Style, 8) = "Emphasis" And Left$(c.Item(c.Count).Style, 9) = "Underline" Then
                Selection.Style = "Underline"
            End If
            
            ' Extend highlighting to cover when surrounded, have to apply to each character for Word to display it
            If c.Item(1).HighlightColorIndex > 0 And c.Item(c.Count).HighlightColorIndex > 0 Then
                For i = 2 To c.Count - 1
                    c.Item(i).HighlightColorIndex = c.Item(1).HighlightColorIndex
                Next i
            End If
        Loop
                        
        .ClearFormatting
        .Replacement.ClearFormatting
    End With
    
    ' Restore the selection
    Selection.Start = SelectionStart
    Selection.End = SelectionEnd
    Set r = Nothing
End Sub
