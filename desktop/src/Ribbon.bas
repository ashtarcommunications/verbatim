Attribute VB_Name = "Ribbon"
'@IgnoreModule ProcedureCanBeWrittenAsFunction, ProcedureNotUsed
Option Explicit

Public Sub OnLoad(ByVal Ribbon As IRibbonUI)
    Dim SavedState As Boolean
    Set Globals.DebateRibbon = Ribbon
    
    ' Save a pointer to the Ribbon in case it gets lost
    SavedState = ActiveDocument.Saved
    ActiveDocument.Variables.Item("RibbonPointer").Value = ObjPtr(Ribbon)
    ActiveDocument.Saved = SavedState
End Sub

Public Function GetRibbon(ByVal lRibbonPointer As LongPtr) As Object
    Dim objRibbon As Object
    #If Mac Then
        CopyMemory_byVar objRibbon, lRibbonPointer, LenB(lRibbonPointer)
    #Else
        '@Ignore ImplicitUnboundDefaultMemberAccess
        CopyMemory objRibbon, lRibbonPointer, LenB(lRibbonPointer)
    #End If
    Set GetRibbon = objRibbon
    Set objRibbon = Nothing
End Function

Public Sub RefreshRibbon()
    If Globals.DebateRibbon Is Nothing Then
        Set Globals.DebateRibbon = GetRibbon(ActiveDocument.Variables.Item("RibbonPointer").Value)
        Globals.DebateRibbon.Invalidate
    Else
        Globals.DebateRibbon.Invalidate
    End If
End Sub

Public Sub RibbonMain(ByVal c As IRibbonControl)
    ' Set Customization context so FindKey returns correct shortcuts
    '@Ignore ImplicitUnboundDefaultMemberAccess
    If Application.CustomizationContext <> "Debate.dotm" Then Application.CustomizationContext = ActiveDocument.AttachedTemplate

    Select Case c.ID
   
    ' Speech group
    Case Is = "SendToSpeech"
        Paperless.SendToSpeech
    Case Is = "SendToSpeech2"
        Paperless.SendToSpeech
    Case Is = "SendToSpeechEnd"
        Paperless.SendToSpeechEnd
    Case Is = "SendToFlowCell"
        Flow.SendToFlowCell
    Case Is = "SendToFlowColumn"
        Flow.SendToFlowColumn
    Case Is = "SendHeadingsToFlowCell"
        Flow.SendHeadingsToFlowCell
    Case Is = "SendHeadingsToFlowColumn"
        Flow.SendHeadingsToFlowColumn
    Case Is = "SelectHeadingAndContent"
        Paperless.SelectHeadingAndContent
    Case Is = "MoveDown"
        Paperless.MoveDown
    Case Is = "MoveUp"
        Paperless.MoveUp
    Case Is = "MoveToBottom"
        Paperless.MoveToBottom
    
    Case Is = "QuickCardSettings"
        UI.ShowForm "QuickCards"
        
    ' Organize group
    Case Is = "F2Button"
        FindKey(wdKeyF2).Execute
    Case Is = "F3Button"
        FindKey(wdKeyF3).Execute
    Case Is = "F4Button"
        FindKey(wdKeyF4).Execute
    Case Is = "F5Button"
        FindKey(wdKeyF5).Execute
    Case Is = "F6Button"
        FindKey(wdKeyF6).Execute
    Case Is = "F7Button"
        FindKey(wdKeyF7).Execute
    Case Is = "F8Button"
        FindKey(wdKeyF8).Execute
    Case Is = "F9Button"
        FindKey(wdKeyF9).Execute
    Case Is = "F10Button"
        FindKey(wdKeyF10).Execute
    Case Is = "F11Button"
        FindKey(wdKeyF11).Execute
    Case Is = "F12Button"
        FindKey(wdKeyF12).Execute

    ' Format Group
    Case Is = "ShrinkText", "ShrinkText2"
        Shrink.ShrinkAllOrCard
    Case Is = "ShrinkAll"
        Shrink.ShrinkAll
    Case Is = "ShrinkPilcrows"
        Shrink.ShrinkPilcrows
    Case Is = "UnshrinkAll"
        Shrink.UnshrinkAll
    
    Case Is = "FixFakeTags"
        Formatting.FixFakeTags
    Case Is = "ConvertAnalyticsToTags"
        Formatting.ConvertAnalyticsToTags
    Case Is = "FixFormattingGaps"
        Formatting.FixFormattingGaps
    Case Is = "ConvertToDefaultStyles"
        Formatting.ConvertToDefaultStyles
    Case Is = "RemoveExtraStyles"
        Formatting.RemoveExtraStyles
    Case Is = "AutoNumberTags"
        Formatting.AutoNumberTags
    Case Is = "DeNumberTags"
        Formatting.DeNumberTags
    Case Is = "InsertHeader"
        Formatting.InsertHeader
        
    Case Is = "RemoveEmphasis"
        Formatting.RemoveEmphasis
    Case Is = "RemoveNonHighlightedUnderlining"
        Formatting.RemoveNonHighlightedUnderlining
    Case Is = "RemoveBlanks"
        Formatting.RemoveBlanks
    Case Is = "RemovePilcrows"
        Condense.RemovePilcrows Notify:=True
    Case Is = "RemoveHyperlinks"
        Formatting.RemoveHyperlinks
    Case Is = "RemoveBookmarks"
        VirtualTub.RemoveBookmarks
        
    Case Is = "UpdateStyles"
        Formatting.UpdateStyles
    Case Is = "SelectSimilar"
        Formatting.SelectSimilar
    
    Case Is = "CondenseNoPilcrows"
        Condense.CondenseNoPilcrows
    Case Is = "CondenseWithPilcrows"
        Condense.CondenseWithPilcrows
    Case Is = "Uncondense"
        Condense.Uncondense

    Case Is = "UniHighlight"
        Formatting.UniHighlight
    Case Is = "UniHighlightWithException"
        Formatting.UniHighlightWithException
            
    Case Is = "AutoEmphasizeFirst"
        Formatting.AutoEmphasizeFirst
    Case Is = "AutoUnderline"
        Formatting.AutoUnderline

    Case Is = "DuplicateCite"
        Formatting.CopyPreviousCite
    Case Is = "AutoFormatCite"
        Formatting.AutoFormatCite
    Case Is = "ReformatAllCites"
        Formatting.ReformatAllCites
    Case Is = "GetFromCiteCreator"
        Plugins.GetFromCiteCreator
    
    ' Paperless group
    Case Is = "NewSpeech"
        Paperless.NewSpeech
    Case Is = "ChooseSpeechDoc"
        UI.ShowForm "ChooseSpeechDoc"
    Case Is = "NewDocument", "NewDocument1"
        Paperless.NewDocument
    Case Is = "CreateFlow"
        Flow.CreateFlow
    Case Is = "CombineDocs"
        UI.ShowForm "CombineDocs"
    Case Is = "CopyToUSB"
        Paperless.CopyToUSB
    Case Is = "TabroomShare"
        UI.ShowForm "Share"
     
    ' Tools Group
    Case Is = "StartTimer"
        Plugins.StartTimer
    Case Is = "PasteOCR"
        OCR.PasteOCR
    Case Is = "DocumentStats"
        UI.ShowForm "Stats"

    Case Is = "NavPaneCycle"
        Plugins.NavPaneCycle
    
    Case Is = "NewWarrant", "NewWarrant1"
        Paperless.NewWarrant
    Case Is = "DeleteAllWarrants"
        Paperless.DeleteAllWarrants
        
    ' View Group
    Case Is = "ReadingView"
        View.ReadingView
    Case Is = "DefaultView"
        View.DefaultView
    
    Case Is = "WindowArranger"
        View.ArrangeWindows
            
    ' Caselist Group
    Case Is = "CaselistWizard"
        UI.ShowForm "Caselist"
    Case Is = "ConvertToWiki"
        Caselist.Word2MarkdownCites
    Case Is = "CiteRequestDoc"
        Caselist.CiteRequestDoc
    Case Is = "CiteRequest"
        Caselist.CiteRequest
    
    ' Settings Group
    Case Is = "LaunchWebsite"
        Settings.LaunchWebsite Globals.PAPERLESSDEBATE_URL
    Case Is = "VerbatimHelp"
        UI.ShowForm "Help"
    Case Is = "CheatSheet"
        UI.ShowForm "CheatSheet"
    Case Is = "VerbatimSettings", "VerbatimSettings1", "VerbatimSettings2"
        UI.ShowForm "Settings"
        
    Case Else
        ' Do Nothing

    End Select

    ' Reset Customization Context
    '@Ignore ValueRequired
    Application.CustomizationContext = ThisDocument
End Sub

Public Sub GetRibbonLabels(ByVal c As IRibbonControl, ByRef label As Variant)
' Assign labels to F key controls from registry

    Select Case c.ID
    
    Case Is = "F2Button"
        label = "F2 " & GetSetting("Verbatim", "Keyboard", "F2Shortcut", "Paste")
    Case Is = "F3Button"
        label = "F3 " & GetSetting("Verbatim", "Keyboard", "F3Shortcut", "Condense")
    Case Is = "F4Button"
        label = "F4 " & GetSetting("Verbatim", "Keyboard", "F4Shortcut", "Pocket")
    Case Is = "F5Button"
        label = "F5 " & GetSetting("Verbatim", "Keyboard", "F5Shortcut", "Hat")
    Case Is = "F6Button"
        label = "F6 " & GetSetting("Verbatim", "Keyboard", "F6Shortcut", "Block")
    Case Is = "F7Button"
        label = "F7 " & GetSetting("Verbatim", "Keyboard", "F7Shortcut", "Tag")
    Case Is = "F8Button"
        label = "F8 " & GetSetting("Verbatim", "Keyboard", "F8Shortcut", "Cite")
    Case Is = "F9Button"
        label = "F9 " & GetSetting("Verbatim", "Keyboard", "F9Shortcut", "Underline")
    Case Is = "F10Button"
        label = "F10 " & GetSetting("Verbatim", "Keyboard", "F10Shortcut", "Emphasis")
    Case Is = "F11Button"
        label = "F11 " & GetSetting("Verbatim", "Keyboard", "F11Shortcut", "Highlight")
    Case Is = "F12Button"
        label = "F12 " & GetSetting("Verbatim", "Keyboard", "F12Shortcut", "Clear")
    
    Case Is = "DefaultView"
        If GetSetting("Verbatim", "View", "DefaultView", Globals.DefaultView) = "Web" Then
            label = "Web"
        Else
            label = "Draft"
        End If
        
    Case Else
        label = "Undefined"
    
    End Select
End Sub

Public Sub GetRibbonImages(ByVal c As IRibbonControl, ByRef returnedBitmap As Variant)
' Get image for Default View
    Select Case c.ID
        Case Is = "DefaultView"
            If GetSetting("Verbatim", "View", "DefaultView", Globals.DefaultView) = "Web" Then
                returnedBitmap = "ViewWebLayoutView"
            Else
                returnedBitmap = "ViewDraftView"
            End If
        Case Is = "ReadingView"
            #If Mac Then
                returnedBitmap = "ViewDraftView"
            #Else
                returnedBitmap = "ViewFullScreenReadingView"
            #End If
        
        Case Is = "SendToFlowCell", "SendToFlowColumn"
            #If Mac Then
                returnedBitmap = "ChartShowDataContextualMenu"
            #Else
                returnedBitmap = "ExportExcel"
            #End If
            
        Case Is = "SendHeadingsToFlowCell", "SendHeadingsToFlowColumn"
            returnedBitmap = "ViewOutlineView"
            
        Case Is = "CaselistWizard"
            #If Mac Then
                returnedBitmap = "WebPagePreview"
            #Else
                returnedBitmap = "UpgradeDocument"
            #End If
        Case Else
            returnedBitmap = ""
    
    End Select
End Sub

Public Sub GetRibbonToggles(ByVal c As IRibbonControl, ByRef state As Variant)
    Select Case c.ID
        
    Case Is = "AutoOpenFolder"
        state = Globals.AutoOpenFolderToggle
        
    Case Is = "RecordAudio"
        state = Globals.RecordAudioToggle
        
    Case Is = "InvisibilityMode"
        state = Globals.InvisibilityToggle
        
    Case Is = "AutoUnderline"
        state = Globals.UnderlineModeToggle
    
    Case Is = "ParagraphIntegrity"
        state = Globals.ParagraphIntegrityToggle
    
    Case Is = "UsePilcrows"
        state = Globals.UsePilcrowsToggle
        
    Case Else
        state = False
        
    End Select
End Sub

Public Sub GetRibbonVisibility(ByVal c As IRibbonControl, ByRef Visible As Variant)
' Get visibility of ribbon groups
    
    ' Default to true
    Visible = True
    
    Select Case c.ID
        Case "Speech", "Organize", "Format", "Paperless", "Tools", "View", "Caselist", "Settings"
            If GetSetting("Verbatim", "View", "RibbonDisable" & c.ID, False) = True Then Visible = False
        Case Else
            Visible = True
    End Select
End Sub
