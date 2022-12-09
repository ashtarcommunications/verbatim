VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTutorial 
   Caption         =   "Verbatim Tutorial"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7530
   OleObjectBlob   =   "frmTutorial.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTutorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Windows API declarations for making form transparent
#If Mac Then
    ' Do Nothing
#Else
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
#End If

Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2

Public TutorialStep As Long
Public TutorialDoc As String
 
Private Sub UserForm_Initialize()
    
    'Reset tutorial step counter
    TutorialStep = 0
      
End Sub
 
Private Sub GoTransparent()
    Dim formhandle As Long
    Dim Style As Long
    Dim Menu As Long

    'Find Form window handle and make it transparent
    formhandle = FindWindow(vbNullString, Me.Caption)
    SetWindowLong formhandle, GWL_EXSTYLE, GetWindowLong(formhandle, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes formhandle, vbCyan, 0&, LWA_COLORKEY
    Me.BackColor = vbCyan

    formhandle = FindWindow("ThunderDFrame", Me.Caption)
    Style = GetWindowLong(formhandle, &HFFF0)
    Style = Style And Not &HC00000
    SetWindowLong formhandle, &HFFF0, Style
    DrawMenuBar formhandle
   
End Sub

Sub btnCancel_Click()
    Unload Me
End Sub

Sub btnNext_Click()

    On Error GoTo Handler

    'Increment step counter
    TutorialStep = TutorialStep + 1

    'Make sure window is still maximized and reset doc
    ActiveWindow.WindowState = wdWindowStateMaximize
    If TutorialStep = 20 Then ActiveWindow.View.Type = wdWebView
    ClearTutorialDoc
    
    Select Case TutorialStep
    
        'Introduction
        Case Is = 1
                   
            'Make sure ribbon is visible
            If CommandBars("Ribbon").Controls(1).Height < 100 Then ActiveWindow.ToggleRibbon
            
            'Make form transparent and cover ribbon
            Call GoTransparent
            Me.Left = ActiveWindow.Left
            Me.Top = ActiveWindow.Top
            Me.Width = ActiveWindow.Width
            Me.Height = 115
        
            'Change labels and buttons for overlay
            Me.lblMessage.BackColor = vbBlack
            Me.lblMessage.ForeColor = vbWhite
            Me.btnCancel.BackColor = vbRed
            Me.btnNext.BackColor = vbGreen
            Me.btnNext.Caption = "Next"
            
            Call ChangeMessage("This is the Verbatim ribbon - it contains buttons for almost every feature. Many features also have keyboard shortcuts.", 45, 330, 250, 40)
    
            Selection.Style = "Tag"
            Selection.TypeText "Welcome to the interactive Verbatim tutorial! You can use this document to experiment and follow along." & vbCrLf
            Selection.TypeText "Use the Next button above to step through the tutorial." & vbCrLf
            Selection.TypeText "Features in each step will be highlighted with a red box." & vbCrLf
            Selection.ClearFormatting
    
        'F keys
        Case Is = 2
            Call ChangeMessage("This section of the ribbon shows basic formatting functions for things like Blocks and Tags, and their corresponding F-key shortcuts. You can configure these shortcuts in the Verbatim settings.", 45, 8, 375, 40)
            If Application.Version < "15.0" Then
                Call ChangeRedBox(43, 383, 270, 33)
            Else
                Call ChangeRedBox(45, 399, 270, 33)
            End If
            Selection.Style = "Tag"
            Selection.TypeText "Try using some of the F-key shortcuts to paste or format text:" & vbCrLf
            Selection.ClearFormatting
            Selection.TypeText vbCrLf & "For example, if you" & vbCrLf & vbCrLf & "select these four paragraphs" & vbCrLf & vbCrLf & "and press F3, they will be condensed" & vbCrLf & vbCrLf & "to a single paragraph." & vbCrLf & vbCrLf
            Selection.Style = "Tag"
            Selection.TypeText "Remember to always use the Paste function instead of Ctrl-V when pasting from the internet to remove weird formatting!" & vbCrLf
            Selection.ClearFormatting
    
        'Heading levels
        Case Is = 3
            Call ChangeMessage("Think of each Word document like an expando - with Pockets, Hats, Blocks and Tags, you have 4 levels available for organizing your files. Note how these levels show up in a hierarchy in the Navigation Pane on the left, and can be dragged up and down.", 45, 8, 375, 40)
            Selection.Style = "Pocket"
            Selection.TypeText "Pocket" & vbCrLf
            Selection.Style = "Hat"
            Selection.TypeText "Hat" & vbCrLf
            Selection.Style = "Block"
            Selection.TypeText "Block" & vbCrLf
            Selection.Style = "Tag"
            Selection.TypeText "Tag" & vbCrLf
            Selection.ClearFormatting
    
        'Other formatting options
        Case Is = 4
            Call ChangeMessage("There are many other formatting functions and useful features - check the Verbatim help for more information. Next up, features for in-round debating.", 45, 125, 250, 40)
            If Application.Version < "15.0" Then
                Call ChangeRedBox(75, 382, 275, 16)
            Else
                Call ChangeRedBox(78, 399, 285, 16)
            End If
            Selection.Style = "Tag"
            Selection.TypeText "Some of the other formatting features include:" & vbCrLf
            Selection.ClearFormatting
            Selection.TypeText "*Shrink un-underlined text to 8pt and smaller" & vbCrLf & "*Convert your backfiles" & vbCrLf & "*Automatically underline a card" & vbCrLf & "*Automatically underline selected text" & vbCrLf & "*Duplicate the previous cite" & vbCrLf & "*Auto-format your cite by bolding last name and date" & vbCrLf & "*And many more..." & vbCrLf
    
        'Send To Speech button
        Case Is = 5
            Call ChangeMessage("The arrow to the left sends the current Pocket, Hat, Block, or Card (or the selected text) to the active Speech document. You can also press the `\~ key instead (next to the number 1 key).", 45, 32, 300, 40)
            If Application.Version < "15.0" Then
                Call ChangeRedBox(42, 8, 15, 15)
            Else
                Call ChangeRedBox(45, 8, 15, 15)
            End If
        
        Case Is = 6
            Call ChangeMessage("Try it out! Click in the sample text below and try sending to the new speech doc. When you're done experimenting, click Next. The temporary speech doc will be closed automatically.", 45, 32, 300, 40)
            
            Selection.Style = "Block"
            Selection.TypeText "Block Title" & vbCrLf
            Selection.Style = "Tag"
            Selection.TypeText "This is a sample tag" & vbCrLf
            Selection.ClearFormatting
            Selection.TypeText "Sample card text"
            
            Dim TempSpeechDoc As String
            TempSpeechDoc = Documents.Add(Template:=ActiveDocument.AttachedTemplate.FullName)
            ActiveSpeechDoc = TempSpeechDoc
            View.ArrangeWindows
            If Application.Version < "15.0" Then
                Call ChangeRedBox(44, 11, 15, 15)
            Else
                Call ChangeRedBox(47, 8, 15, 15)
            End If
            
        'Speech doc chooser
        Case Is = 7
            Dim d
            For Each d In Documents
                If d = ActiveSpeechDoc Then Documents(ActiveSpeechDoc).Close wdDoNotSaveChanges
            Next d
            
            ActiveWindow.WindowState = wdWindowStateMaximize
            ClearTutorialDoc
            Call ChangeMessage("By default, any document with the name ""Speech"" in the name will be your speech document for sending things to. This button lets you choose any document you want as the current speech document.", 45, 100, 300, 40)
            
            If Application.Version < "15.0" Then
                Call ChangeRedBox(42, 30, 68, 16)
            Else
                Call ChangeRedBox(45, 30, 68, 16)
            End If
            
        'Windows arranger
        Case Is = 8
            Call ChangeMessage("These buttons help arrange your docs for greater efficiency. The left menu shows a list of open docs, the right button arranges them split-screen with your Speech on the right (like in the last step).", 45, 110, 300, 40)
            
            If Application.Version < "15.0" Then
                Call ChangeRedBox(58, 7, 86, 18)
            Else
                Call ChangeRedBox(60, 7, 86, 18)
            End If
            
            Selection.Style = "Tag"
            Selection.TypeText "Tip: You can configure the layout of the automatic windows arranger in the Verbatim Settings." & vbCrLf
            
            'Shift focus to document so SendKeys works
            'SetCursorPos 500, 500
            'Tutorial.SingleClick
            WordBasic.SendKeys "%d%w"
            
        'VTub
        Case Is = 9
            Call ChangeMessage("This menu opens your ""Virtual Tub,"" which lets you insert sections of documents without needing to actually open them. It must be configured in the Verbatim Settings before use.", 45, 32, 270, 40)
            
            If Application.Version < "15.0" Then
                Call ChangeRedBox(76, 7, 22, 16)
            Else
                Call ChangeRedBox(78, 7, 22, 16)
            End If
            
            Selection.Style = "Tag"
            Selection.TypeText "Tip: The Virtual Tub is designed to be used with a relatively small number of files that are very well organized - it's not meant for your entire Tub. The VTub can be tricky to set up - make sure to read the manual if you're having trouble." & vbCrLf
                    
        'Search
        Case Is = 10
            Call ChangeMessage("You can type a search term into this box and press Enter - the dropdown menu will contain a list of documents on your computer which contain that phrase, which you can open just by clicking.", 45, 125, 330, 40)
            
            If Application.Version < "15.0" Then
                Call ChangeRedBox(72, 36, 72, 23)
            Else
                Call ChangeRedBox(76, 38, 75, 23)
            End If
            
            Selection.Style = "Tag"
            Selection.TypeText "Tip: By default, the Search box will search everything under your Users folder. You can set a more specific search location in the Verbatim settings." & vbCrLf
                    
        'New Speech
        Case Is = 11
            Call ChangeMessage("This button creates a new ""Speech"" document. If you use the drop-down menu instead, it will let you select from a list of pre-selected speech names, including ones auto-detected from the tournament you're at.", 45, 200, 340, 40)
            
            If Application.Version < "15.0" Then
                Call ChangeRedBox(42, 113, 65, 16)
            Else
                Call ChangeRedBox(45, 117, 68, 16)
            End If
            
            Selection.Style = "Tag"
            Selection.TypeText "Tip: For the auto-naming functions to work, you must enter your tabroom.com username and password in the Verbatim settings, and the tournament you are attending must be run on Tabroom." & vbCrLf
                    
        'Doc Combiner
        Case Is = 12
            Call ChangeMessage("The button on the left creates a new blank Verbatim document. The button on the right starts a wizard that lets you quickly combine documents, for example to combine speech docs into one post-round document for the judge.", 45, 200, 340, 40)
            
            If Application.Version < "15.0" Then
                Call ChangeRedBox(58, 114, 65, 18)
            Else
                Call ChangeRedBox(60, 117, 68, 18)
            End If
                    
        'Coauthoring
        Case Is = 13
            Call ChangeMessage("If you have a PaDS account, this menu will let you quickly upload or open documents for coauthoring with other people.", 45, 200, 200, 40)
            
            If Application.Version < "15.0" Then
                Call ChangeRedBox(75, 112, 68, 18)
            Else
                Call ChangeRedBox(78, 115, 70, 18)
            End If
            
            Selection.Style = "Tag"
            Selection.TypeText """Coauthoring"" is when multiple people edit the same Word document simultaneously, for example partners prepping the same speech document." & vbCrLf
            Selection.TypeText "You can configure the default locations for using PaDS in the Verbatim settings." & vbCrLf
            Selection.TypeText "For more info on PaDS, check out: http://paperlessdebate.com/pads" & vbCrLf
            
        'Misc Paperless
        Case Is = 14
            Call ChangeMessage("This section includes several paperless functions, including an ""Auto Open"" folder, adding warrant boxes to cards and turning on automatic coauthoring updates.", 45, 230, 250, 40)
            
            If Application.Version < "15.0" Then
                Call ChangeRedBox(42, 183, 26, 51)
            Else
                Call ChangeRedBox(44, 187, 26, 51)
            End If
            
            Selection.Style = "Tag"
            Selection.TypeText "You can configure the ""Auto Open"" folder in the Verbatim settings. It will watch any folder you choose (e.g. a PaDS or Dropbox folder) and automatically open any new document which appears there." & vbCrLf
            
        'Share
        Case Is = 15
            Call ChangeMessage("These buttons let you quickly share a speech document via USB, Email, or a public PaDS folder.", 45, 285, 180, 40)
            
            If Application.Version < "15.0" Then
                Call ChangeRedBox(42, 214, 44, 52)
            Else
                Call ChangeRedBox(43, 220, 47, 52)
            End If
            
            Selection.Style = "Tag"
            Selection.TypeText "Email can be configured in the Verbatim settings, and works easily with Gmail. It will also try to automatically look up your opponents email addresses from tabroom.com - make sure you've entered your Tabroom username and password." & vbCrLf
            
        'Tools
        Case Is = 16
            Call ChangeMessage("This section contains useful tools - a built-in speech Timer, an audio recorder for quickly capturing speeches, and a ""Stats"" window that will estimate how long your speech doc would take to read.", 45, 332, 300, 40)
            
            If Application.Version < "15.0" Then
                Call ChangeRedBox(42, 261, 50, 52)
            Else
                Call ChangeRedBox(43, 272, 50, 52)
            End If
            
            Selection.Style = "Tag"
            Selection.TypeText "The default directory for saving recorded audio can be configured in the Verbatim settings. You can also configure your words-per-minute count for a more accurate time estimate in the Stats form." & vbCrLf
            
        'View
        Case Is = 17
            Call ChangeMessage("These functions let you quickly change your view, like toggling or cycling the Navigation Pane or switching between Web and Read view.", 45, 404, 220, 40)
            
            If Application.Version < "15.0" Then
                Call ChangeRedBox(43, 313, 69, 52)
            Else
                Call ChangeRedBox(43, 327, 69, 52)
            End If
            
            Selection.Style = "Tag"
            Selection.TypeText "Tip: If you prefer ""Draft"" view to Web View, you can configure your default view in the Verbatim settings." & vbCrLf
                     
        'Reading Mode
        Case Is = 18
            Call ChangeMessage("Click Next to give Reading view a try.", 45, 404, 80, 40)
            
            If Application.Version < "15.0" Then
                Call ChangeRedBox(60, 313, 40, 18)
            Else
                Call ChangeRedBox(60, 327, 40, 18)
            End If
            
        Case Is = 19
            Call ChangeMessage("While in Reading View, you can advance pages using the arrow keys or mousewheel. To mark a card, click where you stopped reading and press the ""`/~"" key.", 5, 250, 700, 16)
            Call ChangeRedBox(0, 0, 0, 0)
            Me.Height = 50
            Selection.Style = "Block"
            Selection.TypeText "Block 1" & vbCrLf
            Selection.ClearFormatting
            Selection.TypeText "Try inserting a speech marker in the middle of this sentence." & vbCrLf
            Selection.Style = "Block"
            Selection.TypeText "Block 2" & vbCrLf & "Block 3" & vbCrLf & "Block 4" & vbCrLf
            ActiveWindow.View.Type = wdReadingView
            
        'Invisibility
        Case Is = 20
            Me.Height = 115
            Call ChangeMessage("This button toggles ""Invisibility Mode,"" which temporarily hides all non-highlighted card text for easier reading or judging. Click Next to see it in action.", 45, 404, 250, 40)
            
            If Application.Version < "15.0" Then
                Call ChangeRedBox(60, 360, 16, 16)
            Else
                Call ChangeRedBox(60, 376, 16, 16)
            End If
            
            ActiveDocument.AttachedTemplate.BuildingBlockTypes(wdTypeCustomQuickParts).Categories("General").BuildingBlocks("VSCVerbatimSampleCard").Insert Selection.Range
            ActiveDocument.AttachedTemplate.BuildingBlockTypes(wdTypeCustomQuickParts).Categories("General").BuildingBlocks("VSCVerbatimSampleCard").Insert Selection.Range
            ActiveDocument.AttachedTemplate.BuildingBlockTypes(wdTypeCustomQuickParts).Categories("General").BuildingBlocks("VSCVerbatimSampleCard").Insert Selection.Range
            
        Case Is = 21
            Call ChangeMessage("Click Next to turn invisibility mode back off.", 45, 404, 80, 40)
            ActiveDocument.AttachedTemplate.BuildingBlockTypes(wdTypeCustomQuickParts).Categories("General").BuildingBlocks("VSCVerbatimSampleCard").Insert Selection.Range
            ActiveDocument.AttachedTemplate.BuildingBlockTypes(wdTypeCustomQuickParts).Categories("General").BuildingBlocks("VSCVerbatimSampleCard").Insert Selection.Range
            ActiveDocument.AttachedTemplate.BuildingBlockTypes(wdTypeCustomQuickParts).Categories("General").BuildingBlocks("VSCVerbatimSampleCard").Insert Selection.Range
            View.InvisibilityOn
            
        Case Is = 22
            View.InvisibilityOff
            Call ChangeMessage("Visibility mode back on! Click Next to move on.", 45, 404, 80, 40)
            ActiveDocument.AttachedTemplate.BuildingBlockTypes(wdTypeCustomQuickParts).Categories("General").BuildingBlocks("VSCVerbatimSampleCard").Insert Selection.Range
            ActiveDocument.AttachedTemplate.BuildingBlockTypes(wdTypeCustomQuickParts).Categories("General").BuildingBlocks("VSCVerbatimSampleCard").Insert Selection.Range
            ActiveDocument.AttachedTemplate.BuildingBlockTypes(wdTypeCustomQuickParts).Categories("General").BuildingBlocks("VSCVerbatimSampleCard").Insert Selection.Range
        
        'Automatic Underliner
        Case Is = 23
            Call ChangeMessage("This button attempts to automatically underline a card based on the tag - make sure to check the results!", 45, 592, 180, 40)
            
            If Application.Version < "15.0" Then
                Call ChangeRedBox(74, 521, 16, 16)
            Else
                Call ChangeRedBox(78, 544, 16, 16)
            End If
            
            Selection.Style = "Tag"
            Selection.TypeText "Automatic Underliner Instructions: To use the automatic underliner, your cursor must be on the tag. The better and more specific your tag, the better it will work." & vbCrLf
        
        'OCR
        Case Is = 24
            Call ChangeMessage("This button lets you OCR a section of your screen and paste the result directly into your doc.", 45, 632, 180, 40)
            
            If Application.Version < "15.0" Then
                Call ChangeRedBox(74, 564, 37, 16)
            Else
                Call ChangeRedBox(78, 587, 37, 16)
            End If
            
            Selection.Style = "Tag"
            Selection.TypeText "OCR Instructions: The OCR function works by letting you draw a box around a portion of the screen - you'll click once in the upper left corner of the capture area to start the box, then click a second time in the lower right to finish. Results will paste at your cursor." & vbCrLf
            
        'Caselist
        Case Is = 25
            Call ChangeMessage("These buttons let you automatically upload cites or open source documents to the caselist of your choice, or convert your docs to cites and/or wiki syntax for manual posting.", 45, 350, 300, 40)
            
            If Application.Version < "15.0" Then
                Call ChangeRedBox(41, 662, 50, 53)
            Else
                Call ChangeRedBox(43, 691, 50, 53)
            End If
            
            Selection.Style = "Tag"
            Selection.TypeText "Caselist functions can be configured in the Verbatim settings, and require you to set up a tabroom.com account." & vbCrLf
            
        'Link and Help
        Case Is = 26
            Call ChangeMessage("Use these buttons to get more help from paperlessdebate.com or the built-in Verbatim help.", 45, 508, 200, 40)
            
            If Application.Version < "15.0" Then
                Call ChangeRedBox(41, 718, 87, 35)
            Else
                Call ChangeRedBox(43, 750, 87, 35)
            End If
            
            Selection.Style = "Tag"
            Selection.TypeText "Tip: You can also open the Verbatim help at any time by pressing F1." & vbCrLf
            
        'Cheat Sheet
        Case Is = 27
            Call ChangeMessage("This button opens a handy cheat sheet of all the Verbatim keyboard shortcuts.", 45, 508, 150, 40)
            
            If Application.Version < "15.0" Then
                Call ChangeRedBox(75, 718, 17, 17)
            Else
                Call ChangeRedBox(78, 750, 17, 17)
            End If
            
        
        'Settings
        Case Is = 28
            Call ChangeMessage("This button opens the main Verbatim Settings, where you can configure all of the features you've just seen.", 45, 505, 200, 40)
            
            If Application.Version < "15.0" Then
                Call ChangeRedBox(76, 740, 61, 17)
            Else
                Call ChangeRedBox(78, 772, 61, 17)
            End If
            
        'Finish
        Case Is = 29
            Me.lblRedBox.visible = False
            Me.btnNext.Caption = "Exit"
            Call ChangeMessage("That's it! For more information read the built-in help or the manual on paperlessdebate.com", 45, 332, 180, 40)
            
        Case Is = 30
            Unload Me
            
        Case Else
            Exit Sub
    End Select
    
    Exit Sub
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Sub ChangeMessage(Message As String, Top As Long, Left As Long, Width As Long, Height As Long)

    'Change message label
    Me.lblMessage.Caption = Message
    Me.lblMessage.Top = Top
    Me.lblMessage.Left = Left
    Me.lblMessage.Width = Width
    Me.lblMessage.Height = Height

    'Reposition buttons
    Me.btnCancel.Top = Me.lblMessage.Top + Me.lblMessage.Height
    Me.btnCancel.Left = Me.lblMessage.Left
    Me.btnCancel.Width = 30
    Me.btnCancel.Height = 20
    
    Me.btnNext.Top = Me.lblMessage.Top + Me.lblMessage.Height
    Me.btnNext.Left = Me.lblMessage.Left + Me.lblMessage.Width - 30
    Me.btnNext.Width = 30
    Me.btnNext.Height = 20
    Me.btnNext.SetFocus

End Sub

Sub ChangeRedBox(Top As Long, Left As Long, Width As Long, Height As Long)

    'Resize and position box
    Me.lblRedBox.visible = True
    Me.lblRedBox.Top = Top
    Me.lblRedBox.Left = Left
    Me.lblRedBox.Width = Width
    Me.lblRedBox.Height = Height
    
End Sub

Private Sub ClearTutorialDoc()
    Selection.WholeStory
    Selection.Delete
    Selection.ClearFormatting
End Sub


Sub lblRedBox_Click()

    Select Case TutorialStep
        Case Is = 6
            Call Paperless.SendToSpeech
        Case Else
            'Do nothing
    End Select

End Sub


Private Sub LaunchTutorial()

    Dim TutorialDoc As String
    Dim d As Document
    
    'If more than one non-empty doc is open, prompt to close
    If Documents.Count > 1 Or ActiveDocument.Words.Count > 1 Then
        If MsgBox("Tutorial can only be run while a single blank document is open. Open a new blank doc and close everything else?", vbYesNo) = vbYes Then
            TutorialDoc = Documents.Add(ActiveDocument.AttachedTemplate.FullName)
            
            For Each d In Documents
                If d <> TutorialDoc Then d.Close wdPromptToSaveChanges
            Next d
        Else
            Exit Sub
        End If
    End If
    
    'Make sure debate tab is active on ribbon
    WordBasic.SendKeys "%d%"
    
    UI.ShowForm "Tutorial"
    
End Sub


