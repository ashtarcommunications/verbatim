VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCheatSheet 
   Caption         =   "Cheat Sheet"
   ClientHeight    =   8985.001
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4020
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
    Dim CommandArray
    
    On Error GoTo Handler

    CustomizationContext = ActiveDocument.AttachedTemplate
    
    For Each k In KeyBindings
        Me.lboxShortcuts.AddItem
        Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 1) = k.KeyString
        
        Select Case k.Command
            Case Is = "Verbatim.Formatting.PasteText"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Paste"
            Case Is = "Verbatim.Settings.ShowVerbatimHelp"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Verbatim Help"
            Case Is = "Verbatim.Formatting.Condense"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Condense"
            Case Is = "Verbatim.Formatting.ToggleUnderline"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Underline"
            Case Is = "Verbatim.Formatting.Highlight"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Highlight"
            Case Is = "Verbatim.Formatting.ClearToNormal"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Clear"
            Case Is = "Verbatim.View.SwitchWindows"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Cycle Windows"
            Case Is = "Verbatim.Settings.ShowSettingsForm"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Verbatim Settings"
            Case Is = "Verbatim.Formatting.GetFromCiteMaker"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Get From CiteMaker"
            Case Is = "Verbatim.Formatting.SelectSimilar"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Select Similar"
            Case Is = "Verbatim.Formatting.CondenseNoPilcrows"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Condense (No Pilcrows)"
            Case Is = "Verbatim.Formatting.ShrinkText"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Shrink"
            Case Is = "Verbatim.Formatting.AutoFormatCite"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Auto Format Cite"
            Case Is = "Verbatim.Formatting.CopyPreviousCite"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Copy Previous Cite"
            Case Is = "Verbatim.Formatting.AutoUnderline"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Auto Underline"
            Case Is = "Verbatim.Formatting.RemoveEmphasis"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Remove Emphasis"
            Case Is = "Verbatim.Formatting.UpdateStyles"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Update Styles"
            Case Is = "Verbatim.Formatting.AutoNumberTags"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Auto Number Tags"
            Case Is = "Verbatim.Paperless.MoveUp"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Move Up"
            Case Is = "Verbatim.Paperless.MoveDown"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Move Down"
            Case Is = "Verbatim.Paperless.DeleteHeading"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Delete Heading"
            Case Is = "Verbatim.Email.ShowEmailForm"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Email"
            Case Is = "Verbatim.Paperless.CopyToUSB"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Copy To USB"
            Case Is = "Verbatim.PaDS.PaDSPublic"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Public PaDS"
            Case Is = "Verbatim.PaDS.UploadToPaDSDummy"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Upload To PaDS"
            Case Is = "Verbatim.PaDS.OpenFromPaDSDummy"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Open From PaDS"
            Case Is = "Verbatim.View.ArrangeWindows"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Arrange Windows"
            Case Is = "Verbatim.Paperless.StartTimer"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Start Timer"
            Case Is = "Verbatim.Caselist.CiteRequest"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Cite Request"
            Case Is = "Verbatim.Stats.ShowStatsForm"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Document Stats"
            Case Is = "Verbatim.View.InvisibilityOff"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "No Invisibility"
            Case Is = "Verbatim.Paperless.NewSpeech"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "New Speech"
            Case Is = "Verbatim.View.ToggleReadingView"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Toggle View"
            Case Is = "Verbatim.Paperless.SendToSpeech"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Send To Speech\Mark"
            Case Is = "Verbatim.Paperless.ShowChooseSpeechDoc"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "Choose Doc"
            Case Is = "Verbatim.View.NavPaneCycle"
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = "NavPaneCycle"
            Case Else
                CommandArray = Split(k.Command, ".")
                Me.lboxShortcuts.List(Me.lboxShortcuts.ListCount - 1, 0) = CommandArray(UBound(CommandArray))
        End Select
        
    Next k

    CustomizationContext = ThisDocument

    Exit Sub
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub
