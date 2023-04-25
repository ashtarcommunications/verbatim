VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCheatSheet 
   Caption         =   "Cheat Sheet"
   ClientHeight    =   9420
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6360
   OleObjectBlob   =   "frmCheatSheet.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCheatSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Activate()
    Dim k As KeyBinding
    Dim Shortcuts As Dictionary
    Set Shortcuts = New Dictionary
    
    On Error GoTo Handler

    '@Ignore ImplicitUnboundDefaultMemberAccess
    If Application.CustomizationContext <> "Debate.dotm" Then Application.CustomizationContext = ActiveDocument.AttachedTemplate
    
    ' Convert keybindings to a dictionary for easier lookup
    ' Mac converts backticks to !, so skip those
    On Error Resume Next
    For Each k In KeyBindings
        If Shortcuts.Exists(k.Command) Then
            Shortcuts.Item(k.Command) = Shortcuts.Item(k.Command) & " / " & Replace(k.KeyString, "!", "`")
        Else
            Shortcuts.Add k.Command, Replace(k.KeyString, "!", "`")
        End If
    Next k
    On Error GoTo Handler
    
    Me.lboxShortcuts.AddItem "----------Speech----------"
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Send To Speech/Mark Card"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Paperless.SendToSpeechCursor")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Send To Speech End"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Paperless.SendToSpeechEnd")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Send To Flow (Cell)"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Flow.SendToFlowCell")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Send To Flow (Column)"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Flow.SendToFlowColumn")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Send Headings To Flow (Cell)"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Flow.SendHeadingsToFlowCell")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Send Headings To Flow (Column)"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Flow.SendHeadingsToFlowColumn")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Insert Quick Card"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.QuickCards.InsertCurrentQuickCard")
    
    
    Me.lboxShortcuts.AddItem ""
    Me.lboxShortcuts.AddItem "----------Organize----------"
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Verbatim Help"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.UI.ShowFormHelp")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Paste"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Formatting.PasteText")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Condense"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Condense.CondenseAllOrCard")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Pocket"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Pocket")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Hat"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Hat")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Block"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Block")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Tag"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Tag")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Cite"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Cite")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Underline"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Formatting.ToggleUnderline")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Emphasis"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Emphasis")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Highlight"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Formatting.Highlight")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Clear"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Formatting.ClearToNormal")
    
    
    Me.lboxShortcuts.AddItem ""
    Me.lboxShortcuts.AddItem "----------Format----------"

    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Shrink"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Shrink.ShrinkAllOrCard")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Condense With Pilcrows"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Condense.CondenseWithPilcrows")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Condense No Pilcrows"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Condense.CondenseNoPilcrows")
        
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Uncondense"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Condense.Uncondense")
        
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Auto Format Cite"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Formatting.AutoFormatCite")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Copy Previous Cite"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Formatting.CopyPreviousCite")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Auto Underline"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Formatting.AutoUnderline")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Auto Emphasize First"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Formatting.AutoEmphasizeFirst")
        
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Update Styles"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Formatting.UpdateStyles")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Select Similar"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Formatting.SelectSimilar")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Get From CiteCreator"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Plugins.GetFromCiteCreator")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Auto Number Tags"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Formatting.AutoNumberTags")
    
    
    Me.lboxShortcuts.AddItem ""
    Me.lboxShortcuts.AddItem "----------Paperless----------"
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Move Up"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Paperless.MoveUp")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Move Down"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Paperless.MoveDown")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Move To Bottom"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Paperless.MoveToBottom")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Select Heading"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Paperless.SelectHeadingAndContent")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Delete Heading"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Paperless.DeleteHeading")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "New Speech"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Paperless.NewSpeech")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Copy To USB"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Paperless.CopyToUSB")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Share To Tabroom"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.UI.ShowFormShare")
    
    
    Me.lboxShortcuts.AddItem ""
    Me.lboxShortcuts.AddItem "----------Tools----------"
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Start Timer"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Plugins.StartTimer")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Document Stats"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.UI.ShowFormStats")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Run NavPaneCycle"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Plugins.NavPaneCycle")
    
    
    Me.lboxShortcuts.AddItem ""
    Me.lboxShortcuts.AddItem "----------View----------"
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Arrange Windows"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.View.ArrangeWindows")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Cycle Windows"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.View.SwitchWindows")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Invisibility Off"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.View.InvisibilityOff")
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Toggle Reading View"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.View.ToggleReadingView")
    
        
    Me.lboxShortcuts.AddItem ""
    Me.lboxShortcuts.AddItem "----------Caselist----------"
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Cite Request Card"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.Caselist.CiteRequestCard")
    
    
    Me.lboxShortcuts.AddItem ""
    Me.lboxShortcuts.AddItem "----------Settings----------"
    
    Me.lboxShortcuts.AddItem
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Verbatim Settings"
    Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = Shortcuts.Item("Verbatim.UI.ShowFormSettings")
    
    '@Ignore ValueRequired
    Application.CustomizationContext = ThisDocument
    
    Exit Sub
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub
