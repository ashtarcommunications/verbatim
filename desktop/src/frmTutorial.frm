VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTutorial 
   Caption         =   "Verbatim Tutorial"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   OleObjectBlob   =   "frmTutorial.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTutorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'@Ignore EncapsulatePublicField
Public TutorialStep As Long
 
Private Sub UserForm_Initialize()
    Globals.InitializeGlobals
    
    #If Mac Then
        ' Don't use UI.ResizeUserForm for this form to make positioning easier
        Me.btnExit.ForeColor = Globals.RED
        Me.btnNext.ForeColor = Globals.BLUE
    #End If
    
    ' Reset tutorial step counter
    TutorialStep = 0
End Sub
 
Public Sub btnExit_Click()
    If Globals.InvisibilityToggle = True Then View.InvisibilityOff
    Unload Me
End Sub

Public Sub btnNext_Click()
    On Error GoTo Handler

    ' Increment step counter
    TutorialStep = TutorialStep + 1
    Me.Caption = "Verbatim Tutorial (" & TutorialStep & " / 20)"

    ' Make sure window is still maximized and reset doc
    ActiveWindow.WindowState = wdWindowStateMaximize
    If TutorialStep = 20 Then ActiveWindow.View.Type = wdWebView
    ClearTutorialDoc
    
    Select Case TutorialStep
        ' Introduction
        Case Is = 1
            
            ' Make sure ribbon is visible
            If CommandBars.Item("Ribbon").Controls.Item(1).Height < 100 Then ActiveWindow.ToggleRibbon
            
            Me.lblMessage.Caption = "First, let's get acquainted with the Debate ribbon - it contains buttons for almost every feature. " _
                & "Many features also have keyboard shortcuts."
    
            Selection.Style = "Tag"
            Selection.TypeText "Welcome to the interactive Verbatim tutorial! You can use this document to experiment and follow along." & vbCrLf
            Selection.TypeText "Use the Next button to step through the tutorial." & vbCrLf
            Selection.ClearFormatting
    
        ' F keys
        Case Is = 2
            Me.lblMessage.Caption = "The ""Organize"" section of the ribbon shows basic formatting functions for things like Blocks and Tags, and their corresponding F-key shortcuts. " _
            & "You can configure these shortcuts in the Verbatim settings."
            
            ShowImage "Organize"
            
            Selection.Style = "Tag"
            Selection.TypeText "Try using some of the F-key shortcuts to paste or format text:" & vbCrLf
            Selection.ClearFormatting
            Selection.TypeText vbCrLf & "For example, if you" & vbCrLf & vbCrLf & "select these four paragraphs" & vbCrLf & vbCrLf & "and press F3, they will be condensed" & vbCrLf & vbCrLf & "to a single paragraph." & vbCrLf & vbCrLf
            Selection.Style = "Tag"
            Selection.TypeText "Remember to always use the Paste function instead of Ctrl-V when pasting from the internet to remove weird formatting!" & vbCrLf
            Selection.ClearFormatting
    
        ' Heading levels
        Case Is = 3
            Me.lblMessage.Caption = "Think of each Word document like an expando - with Pockets, Hats, Blocks and Tags, you have 4 levels available for organizing your files. " _
                & "Note how these levels show up in a hierarchy in the Navigation Pane on the left, and can be dragged up and down."
                
            Selection.Style = "Pocket"
            Selection.TypeText "Pocket" & vbCrLf
            Selection.Style = "Hat"
            Selection.TypeText "Hat" & vbCrLf
            Selection.Style = "Block"
            Selection.TypeText "Block" & vbCrLf
            Selection.Style = "Tag"
            Selection.TypeText "Tag" & vbCrLf
            Selection.ClearFormatting
    
        ' Other formatting options
        Case Is = 4
            Me.lblMessage.Caption = "There are many other formatting functions and useful features - check the Verbatim manual for more information."
            
            ShowImage "Format"
            
            Selection.Style = "Tag"
            Selection.TypeText "Some of the other formatting features include:" & vbCrLf
            Selection.ClearFormatting
            Selection.TypeText "* Shrink un-underlined text to 8pt and smaller" & vbCrLf _
                & "* Condense paragraphs into one" & vbCrLf _
                & "* Automatically underline a card" & vbCrLf _
                & "* Automatically fix various formatting errors" & vbCrLf _
                & "* Auto-format citations" & vbCrLf _
                & "* And many more..." & vbCrLf
    
        ' Send To Speech button
        Case Is = 5
            Me.lblMessage.Caption = "The speech section contains functions useful while constructing a speech document. " _
                & "The arrow sends the current Pocket, Hat, Block, or Card (or the selected text) to the active Speech document (the menu has additional options). " _
                & "You can also press the `\~ key instead (next to the number 1 key)."
        
            ShowImage "Speech"
            #If Mac Then
                HighlightControl 175, 120, 24, 50
            #Else
                HighlightControl 180, 132, 22, 38
            #End If
        
        Case Is = 6
            Me.lblMessage.Caption = "Try it out! Click on a heading and click the arrow button to send the heading to the speech document. " _
                & "When you're done experimenting, click Next. The temporary speech doc will be closed automatically."
            
            Selection.Style = "Block"
            Selection.TypeText "Block Title" & vbCrLf
            Selection.Style = "Tag"
            Selection.TypeText "This is a sample tag" & vbCrLf
            Selection.ClearFormatting
            Selection.TypeText "Sample card text" & vbCrLf
            
            Selection.Style = "Block"
            Selection.TypeText "Block 1" & vbCrLf
            Selection.ClearFormatting
            Selection.TypeText "You can also try entering a card marker with the `/~ key in the speech doc." & vbCrLf
            Selection.Style = "Block"
            Selection.TypeText "Block 2" & vbCrLf & "Block 3" & vbCrLf & "Block 4" & vbCrLf
            
            Dim w As Window
            Set w = ActiveWindow
            Dim TempSpeechDoc As String
            TempSpeechDoc = Documents.Add(Template:=ActiveDocument.AttachedTemplate.FullName).Name
            Globals.ActiveSpeechDoc = TempSpeechDoc
            View.ArrangeWindows
            w.Activate
            Set w = Nothing
                    
        ' VTub/Quick Cards
        Case Is = 7
            Dim d As Document
            For Each d In Documents
                If d.Name = ActiveSpeechDoc Then Documents.Item(ActiveSpeechDoc).Close wdDoNotSaveChanges
            Next d
            
            ActiveWindow.WindowState = wdWindowStateMaximize
            ClearTutorialDoc
            
            Me.lblMessage.Caption = "These menus open your ""Quick Cards"" and ""Virtual Tub,"" which let you quickly insert cards or blocks without needing to actually open the files. " _
            & "It must be configured in the Verbatim Settings before use."
        
            #If Mac Then
                HighlightControl 178, 146, 46, 42
            #Else
                HighlightControl 180, 149, 39, 38
            #End If
        
            Selection.Style = "Tag"
            Selection.TypeText "Tip: The Virtual Tub is designed to be used with a relatively small number of files that are very well organized - " _
                & "it's not meant for your entire Tub. Make sure to read the manual if you're having trouble." & vbCrLf
            
        ' New speech
        Case Is = 8
            Me.lblMessage.Caption = "This button creates a new ""Speech"" document. If you use the drop-down menu instead, it will let you select from a list of pre-selected speech names, " _
                & "including ones auto-detected from the tournament you're at."
            
            ShowImage "Paperless"

            #If Mac Then
                HighlightControl 120, 120, 22, 90
            #Else
                HighlightControl 135, 132, 18, 72
            #End If
            
            Selection.Style = "Tag"
            Selection.TypeText "Tip: For the auto-naming functions to work, you must have a Tabroom.com account, and the tournament you are attending must be run on Tabroom." & vbCrLf

        ' Choose Speech Doc
        Case Is = 9
            Me.lblMessage.Caption = "By default, any document with the name ""Speech"" in the name will be your speech document for sending things to. " _
                & "This button lets you choose any document you want as the current speech document."
                        
            #If Mac Then
                HighlightControl 120, 142, 22, 90
            #Else
                HighlightControl 135, 149, 18, 72
            #End If

        ' Share
        Case Is = 10
            Me.lblMessage.Caption = "These buttons let you quickly share a speech document via USB or share.tabroom.com"
            
            #If Mac Then
                HighlightControl 210, 142, 50, 75
            #Else
                HighlightControl 206, 149, 36, 58
            #End If
            
            Selection.Style = "Tag"
            Selection.TypeText "Tip: share.tabroom.com lets you share your document without sharing your email address, and can automatically email a copy to everyone in your round." & vbCrLf

        ' Tools
        Case Is = 11
            Me.lblMessage.Caption = "This section contains useful tools - a built-in speech Timer, OCR, a ""Stats"" window that will " _
                & "estimate how long your speech doc would take to read, an audio recorder for recording your speeches, and more."
            
            Me.lblHighlight.Visible = False
            ShowImage "Tools"
            
            Selection.Style = "Tag"
            Selection.TypeText "The default directory for saving recorded audio can be configured in the Verbatim settings. You can also configure your words-per-minute " _
                & "count for a more accurate time estimate in the Stats form." & vbCrLf

        ' Search
        Case Is = 12
            Me.lblMessage.Caption = "You can type a search term into the box and press Enter - the dropdown menu will contain a list of documents on your computer " _
            & "which contain that phrase, which you can open just by clicking. It also integrates with ""Everything Search"" on the PC for even more advanced searching."
                        
            Selection.Style = "Tag"
            Selection.TypeText "Tip: By default, the Search box will search everything under your home folder. You can set a more specific search location in the Verbatim settings." & vbCrLf

        ' Windows arranger
        Case Is = 13
            Me.lblMessage.Caption = "These functions let you quickly adjust your view for greater efficiency, like toggling or cycling the Navigation Pane or switching between Web and Read view." _
            & "The highlighted button arranges your docs split-screen with your Speech on the right (like in an earlier step)."
                        
            ShowImage "View"
            
            #If Mac Then
                HighlightControl 212, 142, 22, 26
            #Else
                HighlightControl 210, 149, 16, 24
            #End If
            
            Selection.Style = "Tag"
            Selection.TypeText "Tip: You can configure your default view and the layout of the automatic windows arranger in the Verbatim Settings." & vbCrLf

        ' Invisibility
        Case Is = 14
            Me.lblMessage.Caption = "This button toggles ""Invisibility Mode,"" which temporarily hides all non-highlighted card text for easier reading or judging. " _
                & "Click Next to see it in action."
            
            #If Mac Then
                HighlightControl 212, 120, 22, 26
            #Else
                HighlightControl 210, 132, 16, 24
            #End If
            
            SampleCard
            
        Case Is = 15
            Me.lblMessage.Caption = "Click Next to turn invisibility mode back off."
            
            SampleCard
            
            View.InvisibilityOn
            
        Case Is = 16
            View.InvisibilityOff
            
            Me.lblMessage.Caption = "Invisibility mode is off! Click Next to move on."
            
            SampleCard

        ' Caselist
        Case Is = 17
            Me.lblMessage.Caption = "These buttons let you automatically upload cites or open source documents to opencaselist.com, or convert your docs to cites and/or wiki syntax for manual posting."
                        
            Me.lblHighlight.Visible = False
            ShowImage "Caselist"
            
            Selection.Style = "Tag"
            Selection.TypeText "Caselist functions can be configured in the Verbatim settings, and require a tabroom.com account." & vbCrLf
            
        ' Settings
        Case Is = 18
            Me.lblMessage.Caption = "Use these buttons to get more help from paperlessdebate.com or the Verbatim settings."
            
            ShowImage "Settings"
            
            Selection.Style = "Tag"
            Selection.TypeText "Tip: You can also open the Verbatim help at any time by pressing F1." & vbCrLf
            
        ' Cheat Sheet
        Case Is = 19
            Me.lblMessage.Caption = "This button opens a handy cheat sheet of all the Verbatim keyboard shortcuts."
            
            #If Mac Then
                HighlightControl 142, 168, 22, 24
            #Else
                HighlightControl 152, 168, 16, 24
            #End If
            
        ' Finish
        Case Is = 20
            Me.lblHighlight.Visible = False
            ShowImage "None"
            Me.btnNext.Caption = "Exit"
            Me.btnExit.Visible = False
            Me.lblMessage.Caption = "That's it! For more information read the manual on paperlessdebate.com"
            
        Case Is = 21
            Me.Hide
            Unload Me
            
        Case Else
            Exit Sub
    End Select
    
    Exit Sub
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub ClearTutorialDoc()
    Selection.WholeStory
    Selection.Delete
    Selection.ClearFormatting
End Sub

Private Sub ShowImage(ByVal Image As String)
    Dim c As Object
    Dim i As String
    i = "img" & Image

    Select Case Image
        Case Is = "Speech"
            Me.Controls(i).Left = 173
            Me.Controls(i).Top = 120
            Me.Controls(i).Width = 54
            Me.Controls(i).Height = 95
            Me.Controls(i).Visible = True
        Case Is = "Organize"
            Me.Controls(i).Left = 68
            Me.Controls(i).Top = 120
            Me.Controls(i).Width = 263
            Me.Controls(i).Height = 92
            Me.Controls(i).Visible = True
        Case Is = "Format"
            Me.Controls(i).Left = 130
            Me.Controls(i).Top = 120
            Me.Controls(i).Width = 140
            Me.Controls(i).Height = 93
            Me.Controls(i).Visible = True
        Case Is = "Paperless"
            Me.Controls(i).Left = 116
            Me.Controls(i).Top = 120
            Me.Controls(i).Width = 168
            Me.Controls(i).Height = 93
            Me.Controls(i).Visible = True
        Case Is = "Tools"
            Me.Controls(i).Left = 109
            Me.Controls(i).Top = 120
            Me.Controls(i).Width = 181
            Me.Controls(i).Height = 93
            Me.Controls(i).Visible = True
        Case Is = "View"
            Me.Controls(i).Left = 142
            Me.Controls(i).Top = 120
            Me.Controls(i).Width = 115
            Me.Controls(i).Height = 92
            Me.Controls(i).Visible = True
        Case Is = "Caselist"
            Me.Controls(i).Left = 155
            Me.Controls(i).Top = 120
            Me.Controls(i).Width = 89
            Me.Controls(i).Height = 93
            Me.Controls(i).Visible = True
        Case Is = "Settings"
            Me.Controls(i).Left = 135
            Me.Controls(i).Top = 120
            Me.Controls(i).Width = 129
            Me.Controls(i).Height = 91
            Me.Controls(i).Visible = True
        Case Else
            ' Do Nothing
    End Select
    
    For Each c In Me.Controls
        If Left$(c.Name, 3) = "img" And c.Name <> i Then
            c.Visible = False
        End If
    Next c
End Sub

Private Sub HighlightControl(ByVal Left As Long, ByVal Top As Long, ByVal Height As Long, ByVal Width As Long)
    Me.lblHighlight.Left = Left
    Me.lblHighlight.Top = Top
    Me.lblHighlight.Height = Height
    Me.lblHighlight.Width = Width
    Me.lblHighlight.Visible = True
End Sub

Private Sub SampleCard()
    Selection.Style = "Tag"
    Selection.TypeText "Sample tag" & vbCrLf
    Selection.Style = "Normal/Card"
    Selection.TypeText "Jean-Luc Picard, Captain, 2364" & vbCrLf
    Selection.TypeText "Space, the final frontier. These are the voyages of the Starship Enterprise. "
    Selection.TypeText "Its continuing mission: to explore strange new worlds, to seek out new life and new civilizations, "
    Selection.TypeText "to boldly go where no one has gone before."
    
    ActiveDocument.Range(20, 26).Style = "Cite"
    ActiveDocument.Range(37, 41).Style = "Cite"
    
    ActiveDocument.Range(146, 172).Style = "Underline"
    ActiveDocument.Range(146, 172).HighlightColorIndex = wdYellow
End Sub
