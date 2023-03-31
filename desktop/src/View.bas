Attribute VB_Name = "View"
'@IgnoreModule ProcedureNotUsed
Option Explicit

Public Sub ArrangeWindows()
    Dim CurrentWindow As Window
    Dim w As Window
    Dim MaxLeft As Long
    Dim MaxTop As Long
    Dim LeftSplitPct As Single
    Dim RightSplitPct As Single
    
    On Error GoTo Handler
        
    ' Save current window
    Set CurrentWindow = ActiveWindow
    
    ' Find largest usable window size
    #If Mac Then
        Dim MaxSize() As String
        Dim MaxWidth As Long
        
        ' Mac has issues maximizing windows with the new "full-screen" mode so use an AppleScript to get the usable bounds
        MaxSize = Split(AppleScriptTask("Verbatim.scpt", "GetHorizontalWindowSize", ""), ",")
        MaxWidth = CInt(MaxSize(0))
        MaxLeft = CInt(MaxSize(1))
    #Else
        ' PC can just maximize the current window to get the largest possible size
        ActiveWindow.WindowState = wdWindowStateMaximize
        MaxLeft = ActiveWindow.Left
        MaxTop = ActiveWindow.Top
    #End If

    ' Set to zero if maximized window returns a negative number
    If MaxLeft < 0 Then MaxLeft = 0
    If MaxTop < 0 Then MaxTop = 0
         
    ' Get split percentages from settings, default to 50/50
    LeftSplitPct = GetSetting("Verbatim", "View", "DocsPct", 50) / 100
    RightSplitPct = GetSetting("Verbatim", "View", "SpeechPct", 50) / 100
        
    ' Loop through open windows and organize
    For Each w In Application.Windows
        ' Windows must not be minimized/maximized to assign properties
        w.WindowState = wdWindowStateNormal
        
        ' Mac Word has a bug that treats full-size windows as maximized even in normal mode
        ' and won't let you change window state directly, so fake it out with a small fixed size
        #If Mac Then
            w.Width = 100
            w.Height = 100
        #End If
        
        ' If an ActiveSpeechDoc, put only that doc on the right, otherwise any doc with "Speech" in the name
        If ( _
            (Globals.ActiveSpeechDoc <> "" And w.Document.Name = Globals.ActiveSpeechDoc) _
            Or (Globals.ActiveSpeechDoc = "" And InStr(LCase$(w.Document.Name), "speech") > 0) _
        ) Then
            ' Mac has trouble with Application.UsableWidth, so use values from above to take dock into account
            #If Mac Then
                w.Width = MaxWidth * RightSplitPct
                w.Left = MaxWidth - w.Width + MaxLeft
            #Else
                w.Width = Application.UsableWidth * RightSplitPct
                w.Left = MaxLeft + (Application.UsableWidth * LeftSplitPct)
            #End If
            
        ' Else put doc on the left
        Else
            #If Mac Then
                w.Width = MaxWidth * LeftSplitPct
                w.Left = MaxLeft
            #Else
                w.Width = Application.UsableWidth * LeftSplitPct
                w.Left = MaxLeft
            #End If
        End If
        
        ' Always set 100% height
        #If Mac Then
            ' On Mac, Application.UsableHeight automatically positions vertically.
            ' Need extra check and no Top setting to avoid maximized window bug.
            If w.WindowState <> wdWindowStateMaximize Then w.Height = Application.UsableHeight
        #Else
            w.Height = Application.UsableHeight
            w.Top = MaxTop
        #End If
    Next w
    
    ' Activate original window and clean up
    CurrentWindow.Activate
    Set CurrentWindow = Nothing

    Exit Sub
    
Handler:
    Set CurrentWindow = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Public Sub SwitchWindows()
    Dim i As Long
    
    ' Cycle through all open windows
    For i = 1 To Documents.Count
        If Documents.Item(i).Name = ActiveDocument.Name Then Exit For
    Next i
    
    If i = 1 Then i = Documents.Count + 1
    Documents.Item(i - 1).Activate
End Sub

Public Sub ToggleReadingView()
    #If Mac Then
        MsgBox "Reading View is not supported on Mac. Try the ""Focus Mode"" button at the bottom of the window."
        Exit Sub
    #End If
   
    If ActiveWindow.ActivePane.View.Type <> wdReadingView Then
        ActiveWindow.ActivePane.View.Type = wdReadingView
    Else
        View.DefaultView
    End If
End Sub

Public Sub DefaultView()
    If GetSetting("Verbatim", "View", "DefaultView", Globals.DefaultView) = "Web" Then
        ActiveWindow.ActivePane.View.Type = wdWebView
    Else
        ActiveWindow.ActivePane.View.Type = wdNormalView
    End If
End Sub

Public Sub ReadingView()
    #If Mac Then
        ActiveWindow.ActivePane.View.Type = wdWebView
    #Else
        ActiveWindow.ActivePane.View.Type = wdReadingView
    #End If
End Sub

Public Sub SetZoom()
    ActiveWindow.ActivePane.View.Zoom.Percentage = GetSetting("Verbatim", "View", "ZoomPct", "100")
End Sub

'@Ignore ParameterNotUsed
Public Sub InvisibilityMode(ByVal c As IRibbonControl, ByVal pressed As Boolean)
    On Error Resume Next
    
    If pressed Then
        InvisibilityToggle = True
        View.InvisibilityOn
        MsgBox "Invisibility Mode On. Press the button again to turn it off."
    Else
        InvisibilityToggle = False
        View.InvisibilityOff
        MsgBox "Invisibility Mode Off"
    End If
    
    Ribbon.RefreshRibbon

    On Error GoTo 0
End Sub

Public Sub InvisibilityOn()
    Dim p As Paragraph
    Dim pCount As Long
 
    pCount = 0
    
    ' Make sure status bar is visible for progress indicator
    Application.StatusBar = True
 
    ' Loop each paragraph
    For Each p In ActiveDocument.Paragraphs
        pCount = pCount + 1
        Application.StatusBar = "Processing paragraph " & pCount & " of " & ActiveDocument.Paragraphs.Count
        
        ' Select each non-blank body text paragraph
        If p.OutlineLevel = wdOutlineLevelBodyText And Len(p.Range.Text) > 1 Then
            p.Range.Select
            
            ' Highlight the cites so they don't disappear
            With Selection.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = ""
                .Wrap = wdFindStop
                .Replacement.Text = ""
                .Format = True
                .Style = "Cite"
                .Execute
                
                ' Skip the paragraph if cite is found
                If .Found = True Then GoTo Skip
            End With
            
            ' Select the paragraph, shorten to keep line breaks
            p.Range.Select
            Selection.MoveEndWhile Cset:=vbCrLf, Count:=-1
            Selection.MoveEndWhile Cset:=" ", Count:=-1
            Selection.MoveStartWhile Cset:=vbCrLf, Count:=1
            Selection.MoveStartWhile Cset:=" ", Count:=1
            
            ' Hide all non-highlighted text
            With Selection.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = "[! ]"
                .Wrap = wdFindStop
                .MatchWildcards = True
                .Format = True
                .Highlight = False
                .ParagraphFormat.OutlineLevel = wdOutlineLevelBodyText
                .Replacement.Font.Hidden = True
                .Execute Replace:=wdReplaceAll
            End With
            
        End If
Skip:
    Next p

    ' Clean up and supress errors
    Selection.Find.ClearFormatting
    Selection.Find.MatchWildcards = False
    Selection.Find.Replacement.ClearFormatting
                    
    ActiveDocument.ShowGrammaticalErrors = False
    ActiveDocument.ShowSpellingErrors = False
End Sub

Public Sub InvisibilityOff()
    ' Set the whole doc visible
    ActiveDocument.Range.Font.Hidden = False
    
    ' Turn error checking back on but set it to checked
    ActiveDocument.ShowGrammaticalErrors = False
    ActiveDocument.ShowSpellingErrors = False
    ActiveDocument.GrammarChecked = True
    ActiveDocument.SpellingChecked = True
    ActiveDocument.ShowGrammaticalErrors = True
    ActiveDocument.ShowSpellingErrors = True
End Sub


