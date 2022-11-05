Attribute VB_Name = "Caselist"
Option Explicit

'*************************************************************************************
'* CITEIFY FUNCTIONS                                                                 *
'*************************************************************************************

Sub CiteRequest()
    Selection.Collapse
    
    ' Make sure cursor is in a card
    If Selection.Paragraphs.outlineLevel <> wdOutlineLevelBodyText Then
        MsgBox "Cursor must be in card text - it appears to be in a heading."
        Exit Sub
    End If
    
    ' If card is longer than 50 words, remove all but the first and last few
    With Selection
        .StartOf Unit:=wdParagraph
        .MoveEnd Unit:=wdParagraph, Count:=1
        If .Range.ComputeStatistics(wdStatisticWords) > 50 Then
            .Range.HighlightColorIndex = wdNoHighlight 'Remove highlighting
            .MoveStart Unit:=wdWord, Count:=15
            .MoveEnd Unit:=wdWord, Count:=-15
            .TypeText vbCrLf & "AND" & vbCrLf
        Else
            MsgBox "Cut longer cards!"
        End If
    
    End With
End Sub

Public Sub CiteRequestAll()
    Dim p
    Dim r As Range
    
    ' Delete blank paragraphs to make processing easier
    For Each p In ActiveDocument.Paragraphs
        If Len(p) = 1 Then p.Range.Delete
    Next p
    
    ' Go to top of document
    Selection.HomeKey Unit:=wdStory
    Selection.Collapse

    ' Find tags
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .ParagraphFormat.outlineLevel = wdOutlineLevel4
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Forward = True
        .Wrap = wdFindContinue
        
        ' Loop all tags
        Do While .Execute And Selection.End <> ActiveDocument.Range.End
            
            ' Select card
            Call Paperless.SelectHeadingAndContent
            
            ' If less than 3 paragraphs (tag, cite, card), something's weird so don't do anything
            If Selection.Paragraphs.Count < 3 Then
                ' Do Nothing
            
            ' If 3 paragraphs, cite request 3rd paragraph, which will almost always be the card text
            ElseIf Selection.Paragraphs.Count = 3 Then
                Set r = Selection.Paragraphs(3).Range
                
            ' If 4 or more paragraphs, non-obvious cite
            Else
                
                ' If 2nd, 3rd or 4th paragraph has a URL, start range with next paragraph
                If InStr(Selection.Paragraphs(2).Range.Text, "http://") > 0 Then
                    Set r = ActiveDocument.Range(Start:=Selection.Paragraphs(3).Range.Start, End:=Selection.Range.End)
                ElseIf InStr(Selection.Paragraphs(3).Range.Text, "http://") > 0 Then
                    Set r = ActiveDocument.Range(Start:=Selection.Paragraphs(4).Range.Start, End:=Selection.Range.End)
                ElseIf InStr(Selection.Paragraphs(4).Range.Text, "http://") > 0 Then
                    Set r = ActiveDocument.Range(Start:=Selection.Paragraphs(5).Range.Start, End:=Selection.Range.End)
                
                ' No URL found, try brackets
                Else
                    
                    ' If starting character of 2nd, 3rd or 4th paragraph is one of (<[, it's likely a cite
                    If Selection.Paragraphs(2).Range.Characters(1) Like "[(<]" Or Selection.Paragraphs(2).Range.Characters(1) Like "[[]" Then
                        Set r = ActiveDocument.Range(Start:=Selection.Paragraphs(3).Range.Start, End:=Selection.Range.End)
                    ElseIf Selection.Paragraphs(3).Range.Characters(1) Like "[(<]" Or Selection.Paragraphs(3).Range.Characters(1) Like "[[]" Then
                        Set r = ActiveDocument.Range(Start:=Selection.Paragraphs(4).Range.Start, End:=Selection.Range.End)
                    ElseIf Selection.Paragraphs(4).Range.Characters(1) Like "[(<]" Or Selection.Paragraphs(4).Range.Characters(1) Like "[[]" Then
                        Set r = ActiveDocument.Range(Start:=Selection.Paragraphs(5).Range.Start, End:=Selection.Range.End)
                    
                    ' No Bracket found, try line-length
                    Else
                        ' If 2nd paragraph is a short line, it's likely to be a 2-line cite, so cite request paragraphs 4+
                        If Selection.Paragraphs(2).Range.Characters.Count < 100 Then
                            Set r = ActiveDocument.Range(Start:=Selection.Paragraphs(4).Range.Start, End:=Selection.Range.End)
                        ' Else it's likely a single line cite, so cite request paragraphs 3+
                        Else
                            Set r = ActiveDocument.Range(Start:=Selection.Paragraphs(3).Range.Start, End:=Selection.Range.End)
                        End If
                    End If
                End If
            End If
            
            ' Cite request the range
            If Not r Is Nothing Then
                If r.Words.Count > 50 Then
                    r.MoveStart Unit:=wdWord, Count:=15
                    r.MoveEnd Unit:=wdWord, Count:=-15
                    r.Text = vbCrLf & "AND" & vbCrLf
                End If
            End If
            
            ' Reset range for next loop
            Set r = Nothing
                
            ' Collapse right so find moves on
            Selection.Collapse wdCollapseEnd
                        
        Loop
    End With
    
    ' Add a newline before each heading to keep plaintext output clean
    For Each p In ActiveDocument.Paragraphs
        If p.outlineLevel < 5 And p.Range.Start <> ActiveDocument.Range.Start Then
            p.Range.InsertBefore vbCrLf
            p.Previous.OutlineDemoteToBody
        End If
    Next p
End Sub

Sub CiteRequestDoc()
    Dim FSO As Scripting.FileSystemObject
    
    On Error GoTo Handler
    
    Set FSO = New Scripting.FileSystemObject

    ' Make sure Debate.dotm exists in template folder
    If FSO.FileExists(Application.NormalTemplate.Path & "\Debate.dotm") = False Then
        MsgBox "Debate.dotm not found in your templates folder - it must be installed to create a cite request doc."
        Exit Sub
    End If
    
    ' Copy everything except header/footer
    ActiveDocument.Content.Select
    Selection.Copy

    ' Add new document based on debate template
    Application.Documents.Add Template:=Application.NormalTemplate.Path & "\Debate.dotm"

    ' Paste into new document
    Selection.Paste
    
    ' Go to top of document and collapse selection
    Selection.HomeKey Unit:=wdStory
    Selection.Collapse

    ' Convert all cites
    Call Caselist.CiteRequestAll
    
    ' Remove highlighting
    ActiveDocument.Content.Select
    Selection.Range.HighlightColorIndex = wdNoHighlight 'Remove highlighting
    Selection.Collapse
    
    Set FSO = Nothing
    
    Exit Sub

Handler:
    Set FSO = Nothing
    MsgBox "Error " & Err.number & ": " & Err.Description
End Sub

'*************************************************************************************
'* WIKIFY FUNCTIONS                                                                  *
'*************************************************************************************

Sub Word2MarkdownCites()

    ' Cite request and wikify doc
    Call Caselist.CiteRequestDoc
    Call Caselist.Word2MarkdownMain
    
    ' Clear all formatting
    ActiveDocument.Content.Select
    Selection.ClearFormatting

    Exit Sub

Handler:
    MsgBox "Error " & Err.number & ": " & Err.Description

End Sub

Public Sub Word2MarkdownMain()
' Based on Word2MediaWiki, modified for Markdown Syntax:
' http://www.mediawiki.org/wiki/Word_macros
' Bold/Italic/Underline text is just set to normal to keep the output clean

    Application.ScreenUpdating = False
    
    On Error Resume Next
       
    WikifyReplaceQuotes
    WikifyReplaceDashes
    Formatting.RemovePilcrows
    WikifyEscapeChars
    WikifyConvertHyperlinks
    WikifyConvertH1
    WikifyConvertH2
    WikifyConvertH3
    WikifyConvertH4
    WikifyConvertH5
    WikifyConvertCites
    WikifyConvertItalic
    WikifyConvertBold
    WikifyConvertUnderline
    WikifyConvertSuperscript
    WikifyConvertSubscript
    WikifyRemoveHighlighting
    WikifyRemoveComments
    
    ' Copy to clipboard
    ActiveDocument.Content.Copy
    Application.ScreenUpdating = True
    
End Sub

Private Function EscapeCharacter(Char As String)
    ReplaceString Char, "\" & Char
End Function

Private Function ReplaceString(findStr As String, replacementStr As String)
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = findStr
        .Replacement.Text = replacementStr
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Function

Private Function ReplaceHeading(outlineLevel As String, headerPrefix As String)
    ActiveDocument.Select
    With Selection.Find
        .ClearFormatting
        .ParagraphFormat.outlineLevel = outlineLevel
        .Text = ""
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Forward = True
        .Wrap = wdFindContinue
        Do While .Execute
            With Selection
                If InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If
              
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    .InsertBefore headerPrefix
                    '.InsertBefore vbCr
                End If
                .Style = ActiveDocument.Styles(wdStyleNormal)
            End With
        Loop
    End With
End Function

Private Sub WikifyReplaceQuotes()
    ' Replace all smart quotes with their dumb equivalents
    Dim Quotes As Boolean
    Quotes = Options.AutoFormatAsYouTypeReplaceQuotes
    Options.AutoFormatAsYouTypeReplaceQuotes = False
    ReplaceString ChrW(8220), """"
    ReplaceString ChrW(8221), """"
    ReplaceString "‘", "'"
    ReplaceString "’", "'"
    ReplaceString "`", "'"
    Options.AutoFormatAsYouTypeReplaceQuotes = Quotes
End Sub

Private Sub WikifyReplaceDashes()
    ReplaceString "--", ChrW(8212)
End Sub

Private Sub WikifyEscapeChars()
    EscapeCharacter "*"
    EscapeCharacter "#"
    EscapeCharacter "_"
    EscapeCharacter "-"
    EscapeCharacter "+"
    EscapeCharacter "{"
    EscapeCharacter "}"
    EscapeCharacter "["
    EscapeCharacter "]"
    'EscapeCharacter "~"
    'EscapeCharacter "^^"
    EscapeCharacter "|"
    'EscapeCharacter "'"
End Sub

Private Sub WikifyConvertHyperlinks()
    Formatting.RemoveHyperlinks
End Sub

Private Sub WikifyConvertH1()
    ReplaceHeading wdOutlineLevel1, "# "
End Sub

Private Sub WikifyConvertH2()
    ReplaceHeading wdOutlineLevel2, "## "
End Sub

Private Sub WikifyConvertH3()
    ReplaceHeading wdOutlineLevel3, "### "
End Sub

Private Sub WikifyConvertH4()
    ReplaceHeading wdOutlineLevel4, "#### "
End Sub

Private Sub WikifyConvertH5()
    ReplaceHeading wdOutlineLevel5, "##### "
End Sub

Private Sub WikifyConvertCites()
    On Error Resume Next
    
    ActiveDocument.Select
    With Selection.Find
        .ClearFormatting
        .Style = "Style Style Bold"
        .Text = ""
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Forward = True
        .Wrap = wdFindContinue
        Do While .Execute
            With Selection
                If Len(.Text) > 1 And InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If
                
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    .InsertBefore "**"
                    .InsertAfter "**"
                End If

                .Style = ActiveDocument.Styles("Default Paragraph Font")
                .Font.Bold = False
            End With
        Loop
    End With
End Sub

Private Sub WikifyConvertItalic()
    ActiveDocument.Select
    With Selection.Find
        .ClearFormatting
        .Font.Italic = True
        .Text = ""
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Forward = True
        .Wrap = wdFindContinue
        Do While .Execute
            With Selection
                If Len(.Text) > 1 And InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If
                
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    '.InsertBefore "//"
                    '.InsertAfter "//"
                End If

                .Style = ActiveDocument.Styles("Default Paragraph Font")
                .Font.Italic = False
            End With
        Loop
    End With
End Sub

Private Sub WikifyConvertBold()
    ActiveDocument.Select
    With Selection.Find
        .ClearFormatting
        .Font.Bold = True
        .Text = ""
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Forward = True
        .Wrap = wdFindContinue
        Do While .Execute
            With Selection
                If Len(.Text) > 1 And InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If
                
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                '    .InsertBefore "**"
                '    .InsertAfter "**"
                End If

                .Style = ActiveDocument.Styles("Default Paragraph Font")
                .Font.Bold = False
            End With
        Loop
    End With
End Sub

Private Sub WikifyConvertUnderline()

    ActiveDocument.Select
    With Selection.Find
        .ClearFormatting
        .Font.Underline = True
        .Text = ""
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Forward = True
        .Wrap = wdFindContinue
        Do While .Execute
            With Selection
                If Len(.Text) > 1 And InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    '.InsertBefore "__"
                    '.InsertAfter "__"
                End If
                .Style = ActiveDocument.Styles("Default Paragraph Font")
                .Font.Underline = False
            End With
        Loop
    End With
End Sub

Private Sub WikifyConvertSuperscript()
    ActiveDocument.Select
    With Selection.Find
        .ClearFormatting
        .Font.Superscript = True
        .Text = ""
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Forward = True
        .Wrap = wdFindContinue
        Do While .Execute
            With Selection
                .Text = Trim(.Text)
                If Len(.Text) > 1 And InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If
             
                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    .InsertBefore ("^")
                    .InsertAfter ("^")
                End If

                .Style = ActiveDocument.Styles("Default Paragraph Font")
                .Font.Superscript = False
            End With
        Loop
    End With
End Sub

Private Sub WikifyConvertSubscript()
    ActiveDocument.Select
    With Selection.Find
        .ClearFormatting
        .Font.Subscript = True
        .Text = ""
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Forward = True
        .Wrap = wdFindContinue
        Do While .Execute
            With Selection
                .Text = Trim(.Text)
                If Len(.Text) > 1 And InStr(1, .Text, vbCr) Then
                    ' Just process the chunk before any newline characters
                    ' We'll pick-up the rest with the next search
                    .Collapse
                    .MoveEndUntil vbCr
                End If

                ' Don't bother to markup newline characters (prevents a loop, as well)
                If Not .Text = vbCr Then
                    .InsertBefore ("~")
                    .InsertAfter ("~")
                End If
                .Style = ActiveDocument.Styles("Default Paragraph Font")
                .Font.Subscript = False
            End With
        Loop
    End With
End Sub

Private Sub WikifyRemoveHighlighting()
    Selection.WholeStory
    Selection.Range.HighlightColorIndex = wdNoHighlight
End Sub

Private Sub WikifyRemoveComments()
    Dim i
    For i = ActiveDocument.Comments.Count To 1 Step -1
        ActiveDocument.Comments(i).Delete
    Next i
End Sub

'*************************************************************************************
'* CASELIST INFO FUNCTIONS                                                           *
'*************************************************************************************

Public Function CheckCaselistToken() As Boolean
    Dim CaselistToken
    Dim CaselistTokenExpires
    
    On Error GoTo Handler
    
    CheckCaselistToken = False
    
    CaselistToken = GetSetting("Verbatim", "Caselist", "CaselistToken", "")
    CaselistTokenExpires = GetSetting("Verbatim", "Caselist", "CaselistTokenExpires", "")
    
    If CaselistToken <> "" And CDate(Now()) < CDate(CaselistTokenExpires) Then
        CheckCaselistToken = True
    End If

    Exit Function

Handler:
    MsgBox "Error " & Err.number & ": " & Err.Description
End Function
