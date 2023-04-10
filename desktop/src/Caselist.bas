Attribute VB_Name = "Caselist"
'@IgnoreModule ProcedureNotUsed
Option Explicit

'*************************************************************************************
'* CITEIFY FUNCTIONS                                                                 *
'*************************************************************************************

Public Sub CiteRequest(Optional ByVal p As Paragraph, Optional ByVal SuppressNotify As Boolean)
    Dim r As Range
    
    If p Is Nothing Then
        ' Make sure cursor is in a card
        If Selection.Paragraphs.OutlineLevel <> wdOutlineLevelBodyText And Selection.Paragraphs.OutlineLevel <> wdOutlineLevel4 Then
            MsgBox "Cursor must be in a card - it appears to be in a larger heading."
            Exit Sub
        End If
        
        ' Use current selection by default
        Set r = Paperless.SelectCardTextRange(Selection.Paragraphs.Item(1))
    Else
        Set r = Paperless.SelectCardTextRange(p)
    End If
    
    ' If card is longer than 50 words, remove all but the first and last few
    If r.ComputeStatistics(wdStatisticWords) > 50 Then
        r.HighlightColorIndex = wdNoHighlight 'Remove highlighting
        r.MoveStart Unit:=wdWord, Count:=15
        r.MoveEnd Unit:=wdWord, Count:=-15
        r.Text = vbCrLf & "AND" & vbCrLf
    Else
        If SuppressNotify <> True Then MsgBox "Cut longer cards!"
    End If
End Sub

Public Sub CiteRequestCard()
    Caselist.CiteRequest
End Sub

Public Sub CiteRequestAll()
    Dim p As Paragraph
    
    ' Delete blank paragraphs to make processing easier
    For Each p In ActiveDocument.Paragraphs
        If Len(p.Range.Text) = 1 Then
            p.Range.Delete
        ElseIf p.OutlineLevel = wdOutlineLevel4 Then
            Caselist.CiteRequest p, True
        End If
    Next p
End Sub

Public Sub CiteRequestDoc(Optional ByVal Wikify As Boolean)
    On Error GoTo Handler
    
    ' Make sure Debate.dotm exists in template folder
    If Filesystem.FileExists(Application.NormalTemplate.Path & Application.PathSeparator & "Debate.dotm") = False Then
        MsgBox "Debate.dotm not found in your templates folder - it must be installed to create a cite request doc."
        Exit Sub
    End If
    
    ' Copy everything except header/footer
    ActiveDocument.Content.Select
    Selection.Copy

    ' Add new document based on debate template
    Paperless.NewDocument

    ' Paste into new document
    Selection.Paste
    
    ' Convert all cites
    Caselist.CiteRequestAll
    
    ' Optionally wikify
    If (Wikify = True) Then Caselist.Word2MarkdownMain
    
    ' Remove highlighting
    ActiveDocument.Content.Select
    Selection.Range.HighlightColorIndex = wdNoHighlight
    Selection.Collapse
    
    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

'*************************************************************************************
'* WIKIFY FUNCTIONS                                                                  *
'*************************************************************************************

Public Sub Word2MarkdownCites()
    On Error GoTo Handler
    
    ' Cite request and wikify doc
    Caselist.CiteRequestDoc Wikify:=True
    
    ' Clear all formatting
    ActiveDocument.Content.Select
    Selection.ClearFormatting

    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Public Sub Word2MarkdownMain()
' Based on Word2MediaWiki, modified for Markdown Syntax:
' http://www.mediawiki.org/wiki/Word_macros
' Bold/Italic/Underline text is just set to normal to keep the output clean

    Application.ScreenUpdating = False
    
    On Error Resume Next
       
    WikifyReplaceQuotes
    WikifyReplaceDashes
    Condense.RemovePilcrows Notify:=False
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
    ' WikifyConvertUnderline
    WikifyConvertSuperscript
    WikifyConvertSubscript
    WikifyRemoveHighlighting
    WikifyRemoveComments
    WikifyReplaceLineBreaks
    
    ' Copy to clipboard
    ActiveDocument.Content.Copy
    Application.ScreenUpdating = True
    
    On Error GoTo 0
End Sub

Private Sub EscapeCharacter(ByVal Char As String)
    ReplaceString Char, "\" & Char
End Sub

Private Sub ReplaceString(ByVal findStr As String, ByVal replacementStr As String)
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
End Sub

Private Sub ReplaceHeading(ByVal OutlineLevel As String, ByVal headerPrefix As String)
    ActiveDocument.Select
    With Selection.Find
        .ClearFormatting
        .ParagraphFormat.OutlineLevel = OutlineLevel
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
                .Style = ActiveDocument.Styles.Item(wdStyleNormal).NameLocal
            End With
        Loop
    End With
End Sub

Private Sub WikifyReplaceQuotes()
    ' Replace all smart quotes with their dumb equivalents
    Dim Quotes As Boolean
    Quotes = Options.AutoFormatAsYouTypeReplaceQuotes
    Options.AutoFormatAsYouTypeReplaceQuotes = False
    ReplaceString ChrW$(8220), """"
    ReplaceString ChrW$(8221), """"
    ReplaceString "‘", "'"
    ReplaceString "’", "'"
    ReplaceString "`", "'"
    Options.AutoFormatAsYouTypeReplaceQuotes = Quotes
End Sub

Private Sub WikifyReplaceDashes()
    ReplaceString "--", ChrW$(8212)
End Sub

Private Sub WikifyReplaceLineBreaks()
    ReplaceString vbCrLf, vbCr
    ReplaceString vbCr, "  " & vbCr
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

                .Style = ActiveDocument.Styles.Item("Default Paragraph Font").NameLocal
                .Font.Bold = False
            End With
        Loop
    End With
    On Error GoTo 0
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
                    .InsertBefore "*"
                    .InsertAfter "*"
                End If

                .Style = ActiveDocument.Styles.Item("Default Paragraph Font").NameLocal
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
                    .InsertBefore "**"
                    .InsertAfter "**"
                End If

                .Style = ActiveDocument.Styles.Item("Default Paragraph Font").NameLocal
                .Font.Bold = False
            End With
        Loop
    End With
End Sub

'@Ignore ProcedureNotUsed
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
                'If Not .Text = vbCr Then
                    '.InsertBefore "__"
                    '.InsertAfter "__"
                'End If
                .Style = ActiveDocument.Styles.Item("Default Paragraph Font").NameLocal
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
                .Text = Trim$(.Text)
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

                .Style = ActiveDocument.Styles.Item("Default Paragraph Font").NameLocal
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
                .Text = Trim$(.Text)
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
                .Style = ActiveDocument.Styles.Item("Default Paragraph Font").NameLocal
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
    Dim i As Long
    For i = ActiveDocument.Comments.Count To 1 Step -1
        ActiveDocument.Comments.Item(i).Delete
    Next i
End Sub

'*************************************************************************************
'* CASELIST INFO FUNCTIONS                                                           *
'*************************************************************************************

Public Function CheckCaselistToken() As Boolean
    Dim CaselistToken As String
    Dim CaselistTokenExpires As String
    
    On Error GoTo Handler
    
    CheckCaselistToken = False
    
    CaselistToken = GetSetting("Verbatim", "Caselist", "CaselistToken", "")
    CaselistTokenExpires = GetSetting("Verbatim", "Caselist", "CaselistTokenExpires", "")
    
    If CaselistToken <> "" And CaselistTokenExpires <> "" Then
        If CDate(Now()) < CDate(CaselistTokenExpires) Then
            CheckCaselistToken = True
        End If
    End If

    Exit Function

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Function
