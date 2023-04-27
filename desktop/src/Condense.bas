Attribute VB_Name = "Condense"
'@IgnoreModule EmptyDoWhileBlock, ObsoleteWhileWendStatement
Option Explicit

'@Ignore ProcedureNotUsed
Public Sub CondenseAllOrCard()
    Dim p As Paragraph
    If Selection.Start <= ActiveDocument.Range.Start And Selection.End = ActiveDocument.Range.Start Then
        If MsgBox("This will condense all cards in the document. Are you sure?", vbOKCancel) = vbCancel Then Exit Sub
        For Each p In ActiveDocument.Paragraphs
            If p.OutlineLevel = 4 Then
                Condense.CondenseCard p.Range
            End If
        Next p
    Else
        Condense.CondenseCard
    End If
End Sub

Public Sub CondenseCard(Optional ByVal r As Range)
' Removes white-space from selection and optionally retains paragraph integrity

    Dim CondenseRange As Range
    Dim r2 As Range
    Dim PilcrowCode As Long
    
    #If Mac Then
        PilcrowCode = 166
    #Else
        PilcrowCode = 182
    #End If
    
    Application.ScreenUpdating = False
    
    ' Default to condensing the provided range, or the selection, otherwise select the current card
    If Not r Is Nothing Then
        Set CondenseRange = Paperless.SelectCardTextRange(r.Paragraphs.Item(1))
    ElseIf Selection.Start = Selection.End Then
        Set CondenseRange = Paperless.SelectCardTextRange(Selection.Paragraphs.Item(1))
    Else
        If Selection.Type <> wdSelectionNormal Then
            Application.StatusBar = "Can only condense text, not other document elements"
            Exit Sub
        End If
        Set CondenseRange = Selection.Range
    End If
    
    ' If selection is too short, exit
    If Len(CondenseRange.Text) < 2 Then Exit Sub
        
    ' If end of range is a line break, shorten it
    If CondenseRange.Characters.Last.Text = vbCr Or CondenseRange.Characters.Last.Text = vbCrLf Then CondenseRange.MoveEnd , -1
       
    ' Condense everything except hard returns
    With CondenseRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Wrap = wdFindStop
    
        .Text = "^m"                    ' page breaks
        .Replacement.Text = " "
        .Execute Replace:=wdReplaceAll
        
        .Text = "^t"                    ' tabs
        .Replacement.Text = " "
        .Execute Replace:=wdReplaceAll
        
        .Text = "^s"                    ' non-breaking space
        .Replacement.Text = " "
        .Execute Replace:=wdReplaceAll
        
        .Text = "^b"                    ' section break
        .Replacement.Text = " "
        .Execute Replace:=wdReplaceAll
        
        .Text = "^l"                    ' new line
        .Replacement.Text = " "
        .Execute Replace:=wdReplaceAll
        
        .Text = "^n"                    ' column break
        .Replacement.Text = " "
        .Execute Replace:=wdReplaceAll
    End With
    
    ' If paragraph integrity is off, just condense
    If GetSetting("Verbatim", "Format", "ParagraphIntegrity", False) = False Then
        With CondenseRange.Find
            .Text = "^p"
            .Replacement.Text = " "
            .Execute Replace:=wdReplaceAll
        
            .Text = "  "
            .Replacement.Text = " "
            
            While InStr(CondenseRange.Text, "  ")
                .Execute Replace:=wdReplaceAll
            Wend
            
            If CondenseRange.Characters.Item(1).Text = " " And _
            CondenseRange.Paragraphs.Item(1).Range.Start = CondenseRange.Start Then _
            CondenseRange.Characters.Item(1).Delete
        End With
    
    Else
        ' If paragraph integrity and Pilcrows are on, replace paragraph breaks with Pilcrow sign
        If GetSetting("Verbatim", "Format", "UsePilcrows", False) = True Then
            With CondenseRange.Find
                .Text = "^p"
                .Replacement.Text = Chr$(PilcrowCode) & " " ' Pilcrow sign
                .Replacement.Font.size = 6
                .Execute Replace:=wdReplaceAll
                
                .Text = Chr$(PilcrowCode) & " " & Chr$(PilcrowCode)
                .Replacement.Text = Chr$(PilcrowCode)
                
                While InStr(CondenseRange.Text, Chr$(PilcrowCode) & " " & Chr$(PilcrowCode))
                    .Execute Replace:=wdReplaceAll
                Wend
                
                .Text = "  "
                .Replacement.ClearFormatting
                .Replacement.Text = " "
                
                While InStr(CondenseRange, "  ")
                    .Execute Replace:=wdReplaceAll
                Wend
                
                If CondenseRange.Characters.Item(1).Text = " " And _
                CondenseRange.Paragraphs.Item(1).Range.Start = CondenseRange.Start Then _
                CondenseRange.Characters.Item(1).Delete
                
                ' Remove trailing pilcrows
                If CondenseRange.Characters.Last.Previous.Text = Chr$(PilcrowCode) Then CondenseRange.Characters.Last.Previous.Delete
            End With
    
        Else ' Else, paragraph integrity is off and Pilcrows are off, leave single paragraph marks
            ' Use duplicate range to prevent runaway find bug
            Set r2 = CondenseRange
            
            With CondenseRange.Find
                .Text = "^p^w"
                .Replacement.Text = "^p"
                
                Do While .Execute(Forward:=True, Replace:=wdReplaceAll) And CondenseRange.InRange(r2)
                Loop
                
                .Text = "^p^p"
                .Replacement.Text = "^p"
                
                Do While .Execute(Forward:=True, Replace:=wdReplaceAll) And CondenseRange.InRange(r2)
                Loop
                
                .Text = "  "
                .Replacement.Text = " "
                .Execute Replace:=wdReplaceAll
                
                If CondenseRange.Characters.Item(1).Text = " " And _
                CondenseRange.Paragraphs.Item(1).Range.Start = CondenseRange.Start Then _
                CondenseRange.Characters.Item(1).Delete
            End With
            
            Set r2 = Nothing
        End If
    End If
    
    CondenseRange.Find.ClearFormatting
    CondenseRange.Find.Replacement.ClearFormatting
    
    Application.ScreenUpdating = True
End Sub

Public Sub CondenseNoPilcrows()
' Easier to override saved settings temporarily because of poor VBA handling of optional boolean parameters
    Dim ParagraphIntegrity As Boolean
    Dim UsePilcrows As Boolean
    
    ParagraphIntegrity = GetSetting("Verbatim", "Format", "ParagraphIntegrity", False)
    UsePilcrows = GetSetting("Verbatim", "Format", "UsePilcrows", False)
    
    SaveSetting "Verbatim", "Format", "ParagraphIntegrity", False
    SaveSetting "Verbatim", "Format", "UsePilcrows", False
    
    Condense.CondenseCard
    
    SaveSetting "Verbatim", "Format", "ParagraphIntegrity", ParagraphIntegrity
    SaveSetting "Verbatim", "Format", "UsePilcrows", UsePilcrows
End Sub

Public Sub CondenseWithPilcrows()
    Dim ParagraphIntegrity As Boolean
    Dim UsePilcrows As Boolean
    
    ParagraphIntegrity = GetSetting("Verbatim", "Format", "ParagraphIntegrity", False)
    UsePilcrows = GetSetting("Verbatim", "Format", "UsePilcrows", False)
    
    SaveSetting "Verbatim", "Format", "ParagraphIntegrity", True
    SaveSetting "Verbatim", "Format", "UsePilcrows", True
    
    Condense.CondenseCard
    
    SaveSetting "Verbatim", "Format", "ParagraphIntegrity", ParagraphIntegrity
    SaveSetting "Verbatim", "Format", "UsePilcrows", UsePilcrows
End Sub

Public Sub Uncondense()
' Replaces pilcrows with paragraph breaks
    Dim r As Range
    Dim PilcrowCode As Long
    
    #If Mac Then
        PilcrowCode = 166
    #Else
        PilcrowCode = 182
    #End If
    
    Application.ScreenUpdating = False
    
    If Selection.Start <= ActiveDocument.Range.Start And Selection.Start = Selection.End Then
        If MsgBox("This will uncondense all cards in the document. Are you sure?", vbOKCancel) = vbCancel Then Exit Sub
        Set r = ActiveDocument.Range
    ElseIf Selection.Start = Selection.End Then
        Set r = Paperless.SelectHeadingAndContentRange(Selection.Paragraphs.Item(1))
    Else
        If Selection.Type <> wdSelectionNormal Then
            Application.StatusBar = "Can only uncondense text, not other document elements"
            Exit Sub
        End If
        Set r = Selection.Range
    End If
    
    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = Chr$(PilcrowCode) ' Pilcrow
        .Replacement.Text = "^p"
        .Execute Replace:=wdReplaceAll
        
        .Text = ""
        .Replacement.Text = ""
    End With
    
    Set r = Nothing
    
    Application.ScreenUpdating = True
End Sub

Public Sub RemovePilcrows(Optional ByVal Notify As Boolean)
    Dim r As Range
    Dim PilcrowCode As Long
    
    #If Mac Then
        PilcrowCode = 166
    #Else
        PilcrowCode = 182
    #End If
    
    Application.ScreenUpdating = False
    
    If Selection.Start <= ActiveDocument.Range.Start And Selection.Start = Selection.End Then
        If Notify = True Then
            If MsgBox("This will remove all pilcrows in the document. Are you sure?", vbOKCancel) = vbCancel Then Exit Sub
        End If
        Set r = ActiveDocument.Range
    ElseIf Selection.Start = Selection.End Then
        Set r = Paperless.SelectHeadingAndContentRange(Selection.Paragraphs.Item(1))
    Else
        If Selection.Type <> wdSelectionNormal Then
            Application.StatusBar = "Can only remove pilcrows from text, not other document elements"
            Exit Sub
        End If
        Set r = Selection.Range
    End If

    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = Chr$(PilcrowCode)
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
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
    
    Application.ScreenUpdating = True
End Sub

'@Ignore ProcedureNotUsed
'@Ignore ParameterNotUsed
Public Sub ToggleParagraphIntegrity(ByVal c As IRibbonControl, ByVal pressed As Boolean)
' Toggle setting for Paragraph Integrity
    If GetSetting("Verbatim", "Format", "ParagraphIntegrity", True) = True Then
        SaveSetting "Verbatim", "Format", "ParagraphIntegrity", False
        SaveSetting "Verbatim", "Format", "UsePilcrows", False
        Globals.ParagraphIntegrityToggle = False
        Globals.UsePilcrowsToggle = False
    Else
        SaveSetting "Verbatim", "Format", "ParagraphIntegrity", True
        Globals.ParagraphIntegrityToggle = True
    End If

    Ribbon.RefreshRibbon
End Sub

'@Ignore ProcedureNotUsed
'@Ignore ParameterNotUsed
Public Sub ToggleUsePilcrows(ByVal c As IRibbonControl, ByVal pressed As Boolean)
' Toggle setting for Use Pilcrows
    If GetSetting("Verbatim", "Format", "UsePilcrows", True) = True Then
        SaveSetting "Verbatim", "Format", "UsePilcrows", False
        Globals.UsePilcrowsToggle = False
    Else
        SaveSetting "Verbatim", "Format", "UsePilcrows", True
        Globals.UsePilcrowsToggle = True
    End If

    Ribbon.RefreshRibbon
End Sub


