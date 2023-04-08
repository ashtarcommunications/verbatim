Attribute VB_Name = "Shrink"
Option Explicit

Public Sub ShrinkAllOrCard()
    If Selection.Start <= ActiveDocument.Range.Start And Selection.End = ActiveDocument.Range.Start Then
        If MsgBox("This will shrink all cards in the document. Are you sure?", vbOKCancel) = vbCancel Then Exit Sub
        Shrink.ShrinkAll
    Else
        Shrink.ShrinkText
    End If
End Sub

Public Sub ShrinkText(Optional ByVal ShrinkRange As Range)
' Cycles non-underlined text in the current paragraph down a size at a time from 11-4pt
' Differences in un-underlined font size will be normalized automatically
    Dim r As Range
    Dim r2 As Range
    Dim NewFontSize As Long
    
    Application.ScreenUpdating = False
    
    ' Select the card text if nothing is selected
    If Not ShrinkRange Is Nothing Then
        Set r = Paperless.SelectCardTextRange(ShrinkRange.Paragraphs.Item(1))
    ElseIf Selection.Start = Selection.End Then
        Set r = Paperless.SelectCardTextRange(Selection.Paragraphs.Item(1))
    Else
        If Selection.Type <> wdSelectionNormal Then
            Application.StatusBar = "Can only shrink text, not other document elements"
            Exit Sub
        End If
        If Selection.Paragraphs.OutlineLevel <> wdOutlineLevelBodyText Then
            Application.StatusBar = "Can only shrink card text, not headings"
            Exit Sub
        End If
        Set r = Selection.Range
    End If
    
    ' Duplicate range for limiting find
    Set r2 = r.Duplicate
                              
    ' Find first un-underlined part of card and test font size
    ' have to search for a single character to avoid a runaway find range
    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "[0-9A-Za-z]" ' Ignore punctuation so fixed-size pilcrows don't break the font size
        .Replacement.Text = ""
        .MatchWildcards = True
        .Wrap = wdFindStop
        .Format = True
        .Font.Underline = 0
        .Execute
        
        .Text = ""
        .MatchWildcards = False
    End With
    
    ' Depending on font size, shrink or reset to normal text size
    Select Case r.Font.size
        Case Is = wdUndefined   ' Multiple font sizes, shrink to 8
            NewFontSize = 8
        Case Is > 8
            NewFontSize = 8
        Case Is = 8
            NewFontSize = 7
        Case Is = 7
            NewFontSize = 6
        Case Is = 6
            NewFontSize = 5
        Case Is = 5
            NewFontSize = 4
        Case Is = 4
            NewFontSize = ActiveDocument.Styles.Item("Normal").Font.size
        Case Else   ' Anything weird, go back to normal text size
            NewFontSize = ActiveDocument.Styles.Item("Normal").Font.size
    End Select
                
    ' Reset range
    r.SetRange r2.Start, r2.End
    
    ' If the entire search range is un-underlined, just shrink
    If r.Font.Underline = 0 Then
        r.Font.size = NewFontSize
    ' Otherwise find the un-underlined parts
    Else
        With r.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = ""
            .Replacement.Text = ""
            .MatchWildcards = False
            .IgnorePunct = True
            .IgnoreSpace = True
            .Wrap = wdFindStop
            .Format = True
            .Font.Underline = 0
                        
            Do While r.Find.Execute(Forward:=True) And r.InRange(r2)
                r.Font.size = NewFontSize
            Loop
            
            .MatchWildcards = False
            .ClearFormatting
            .Replacement.ClearFormatting
           End With
    End If
    
    ' Optionally restore bracketed ommissions
    If GetSetting("Verbatim", "Format", "ShrinkOmissions", False) = False Then
        r.SetRange r2.Start, r2.End
        With r.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = "\[*(Omitted)*\]"
            .Replacement.Text = ""
            .Replacement.Font.size = ActiveDocument.Styles.Item("Normal").Font.size
            .MatchWildcards = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Wrap = wdFindStop
            .Format = True
            .Execute Replace:=wdReplaceAll
            
            .Text = "\[\[*(Omitted)*\]\]"
            .Execute Replace:=wdReplaceAll
            
            .Text = "\<*(Omitted)*\>"
            .Execute Replace:=wdReplaceAll
            
            .ClearFormatting
            .Replacement.ClearFormatting
            .MatchWildcards = False
        End With
    End If
    
    ' Shrink pilcrows too, just in case they've been underlined
    r.SetRange r2.Start, r2.End
    Shrink.ShrinkPilcrows r
       
    Application.ScreenUpdating = True
    
    Set r = Nothing
    Set r2 = Nothing
End Sub

Public Sub ShrinkAll()
    Dim p As Paragraph
    
    For Each p In ActiveDocument.Paragraphs
        If p.OutlineLevel = 4 Then
            Shrink.ShrinkText p.Range
        End If
    Next p
End Sub

Public Sub UnshrinkAll()
    With ActiveDocument.Range.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .ParagraphFormat.OutlineLevel = wdOutlineLevelBodyText
        .Font.Underline = wdUnderlineNone
        .Font.Bold = False
        .Replacement.Font.size = ActiveDocument.Styles.Item("Normal").Font.size
        .Format = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .Replacement.ClearFormatting
    End With
End Sub

Public Sub ShrinkPilcrows(Optional ByVal ShrinkRange As Range)
' Shrinks, un-underlines and unbolds all pilcrows in current paragraph to 6pt
' If run with the insertion point at the very beginning of the document, shrinks all pilcrows
    Dim r As Range
    Dim PilcrowCode As Long
    
    #If Mac Then
        PilcrowCode = 166
    #Else
        PilcrowCode = 182
    #End If
    
    Application.ScreenUpdating = False
    
    If Not ShrinkRange Is Nothing Then
        Set r = ShrinkRange
    ElseIf Selection.Start <= ActiveDocument.Range.Start And Selection.Start = Selection.End Then
        Set r = ActiveDocument.Range
    ElseIf Selection.Start = Selection.End Then
        Set r = Paperless.SelectHeadingAndContentRange(Selection.Paragraphs.Item(1))
    Else
        If Selection.Type <> wdSelectionNormal Then
            Application.StatusBar = "Can only shrink text, not other document elements"
            Exit Sub
        End If
        Set r = Selection.Range
    End If
    
    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = Chr$(PilcrowCode)
        .Replacement.Text = Chr$(PilcrowCode)
        .Replacement.Font.size = 6
        .Replacement.Font.Underline = wdUnderlineNone
        .Replacement.Font.Bold = 0
        .Format = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
        
        .ClearFormatting
        .Replacement.ClearFormatting
    End With
            
    Application.ScreenUpdating = True
    
    Set r = Nothing
End Sub
