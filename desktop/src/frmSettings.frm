VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   Caption         =   "Verbatim Settings"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11925
   OleObjectBlob   =   "frmSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SetPage(MenuTab As String)
 
    Me.lblTabMain.BackColor = RGB(255, 255, 255)
    Me.lblTabMain.ForeColor = RGB(40, 40, 40)
    Me.lblTabAdmin.BackColor = RGB(255, 255, 255)
    Me.lblTabAdmin.ForeColor = RGB(40, 40, 40)
    Me.lblTabView.BackColor = RGB(255, 255, 255)
    Me.lblTabView.ForeColor = RGB(40, 40, 40)
    Me.lblTabFormat.BackColor = RGB(255, 255, 255)
    Me.lblTabFormat.ForeColor = RGB(40, 40, 40)
    Me.lblTabPaperless.BackColor = RGB(255, 255, 255)
    Me.lblTabPaperless.ForeColor = RGB(40, 40, 40)
    
    ' TODO - refactor globals
    ' TODO - add extra pages and fix order
    ' TODO - add hover state
    Select Case MenuTab
        Case "Main"
            Me.mpgSettings.Value = 0
            Me.lblTabMain.BackColor = Globals.BLUE_BUTTON_NORMAL
            Me.lblTabMain.ForeColor = RGB(255, 255, 255)
        Case "Admin"
            Me.mpgSettings.Value = 1
            Me.lblTabAdmin.BackColor = Globals.BLUE_BUTTON_NORMAL
            Me.lblTabAdmin.ForeColor = RGB(255, 255, 255)
        Case "View"
            Me.mpgSettings.Value = 2
            Me.lblTabView.BackColor = Globals.BLUE_BUTTON_NORMAL
            Me.lblTabView.ForeColor = RGB(255, 255, 255)
        Case "Format"
            Me.mpgSettings.Value = 3
            Me.lblTabFormat.BackColor = Globals.BLUE_BUTTON_NORMAL
            Me.lblTabFormat.ForeColor = RGB(255, 255, 255)
        Case "Paperless"
            Me.mpgSettings.Value = 4
            Me.lblTabPaperless.BackColor = Globals.BLUE_BUTTON_NORMAL
            Me.lblTabPaperless.ForeColor = RGB(255, 255, 255)
        Case Else
            Me.mpgSettings.Value = 0
            Me.lblTabMain.BackColor = Globals.BLUE_BUTTON_NORMAL
            Me.lblTabMain.ForeColor = RGB(255, 255, 255)
    End Select
    
End Sub



Private Sub btnTabroomLogin_Click()
    Me.Hide
    UI.ShowForm "Login"
    Me.Show
End Sub


Private Sub btnTabroomLogout_Click()
    SaveSetting "Verbatim", "Caselist", "CaselistToken", ""
    SaveSetting "Verbatim", "Caselist", "CaselistTokenExpires", ""
    Me.lblTabroomLoggedIn.Caption = "You are logged out of Tabroom"
End Sub

Private Sub lblTabMain_Click()
    SetPage "Main"
End Sub
Private Sub lblTabAdmin_Click()
    SetPage "Admin"
End Sub
Private Sub lblTabPaperless_Click()
    SetPage "Paperless"
End Sub
Private Sub lblTabView_Click()
    SetPage "View"
End Sub
Private Sub lblTabFormat_Click()
    SetPage "Format"
End Sub

Private Sub UserForm_Activate()
    If GetSetting("Verbatim", "Caselist", "CaselistToken", "") <> "" And Caselist.CheckCaselistToken = True Then
        Me.lblTabroomLoggedIn.Caption = "You are logged in to Tabroom"
    Else
        Me.lblTabroomLoggedIn.Caption = "You are logged out of Tabroom"
    End If
End Sub



Private Sub UserForm_Initialize()

    Dim FontSize As Integer
    Dim f
    Dim MacroArray
    
    'Turn on Error handling
    On Error GoTo Handler
    
    'Get Settings from the registry to populate the settings boxes
    
    'Main Tab
    Me.txtSchoolName.Value = GetSetting("Verbatim", "Main", "SchoolName")
    Me.txtName.Value = GetSetting("Verbatim", "Main", "Name")
    
    If GetSetting("Verbatim", "Main", "CollegeHS", "College") = "College" Then
        Me.optCollege.Value = True
    Else
        Me.optHS.Value = True
    End If
    
    Me.txtWPM.Value = GetSetting("Verbatim", "Main", "WPM", 350)
    
    
    Me.chkAutomaticUpdates.Value = GetSetting("Verbatim", "Main", "AutomaticUpdates", True)
    Me.lblLastUpdateCheck.Caption = "Last Update Check:" & vbCrLf & _
        Format(GetSetting("Verbatim", "Main", "LastUpdateCheck", ""), "mm-dd-yy hh:mm")
    
    'Admin Tab
    Me.chkAlwaysOn.Value = GetSetting("Verbatim", "Admin", "AlwaysOn", True)
    Me.chkAutoUpdateStyles.Value = GetSetting("Verbatim", "Admin", "AutoUpdateStyles", True)
    Me.chkSuppressInstallChecks.Value = GetSetting("Verbatim", "Admin", "SuppressInstallChecks", False)
    Me.chkSuppressDocCheck.Value = GetSetting("Verbatim", "Admin", "SuppressDocCheck", False)
    Me.chkFirstRun = GetSetting("Verbatim", "Admin", "FirstRun", False)
    
    'View Tab
    If GetSetting("Verbatim", "View", "DefaultView", "Web") = "Web" Then
        Me.optWebView.Value = True
    Else
        Me.optDraftView.Value = True
    End If
    
    Me.chkNPCStartup.Value = GetSetting("Verbatim", "View", "NPCStartup", False)
    
    Me.spnDocs.Value = GetSetting("Verbatim", "View", "DocsPct", 50)
    Me.spnSpeech.Value = GetSetting("Verbatim", "View", "SpeechPct", 50)
    
    'Paperless Tab
    Me.chkAutoSaveSpeech.Value = GetSetting("Verbatim", "Paperless", "AutoSaveSpeech", False)
    Me.cboAutoSaveDir.Value = GetSetting("Verbatim", "Paperless", "AutoSaveDir")
    Me.chkStripSpeech.Value = GetSetting("Verbatim", "Paperless", "StripSpeech", True)
    Me.cboSearchDir.Value = GetSetting("Verbatim", "Paperless", "SearchDir")
    Me.cboAutoOpenDir.Value = GetSetting("Verbatim", "Paperless", "AutoOpenDir")
    Me.cboAudioDir.Value = GetSetting("Verbatim", "Paperless", "AudioDir")
      
    'Populate Format Tab Comboboxes - Allow 8pt-32pt
    FontSize = 8
    Do While FontSize < 33
        Me.cboNormalSize.AddItem FontSize
        Me.cboPocketSize.AddItem FontSize
        Me.cboHatSize.AddItem FontSize
        Me.cboBlockSize.AddItem FontSize
        Me.cboTagSize.AddItem FontSize
        Me.cboCiteSize.AddItem FontSize
        Me.cboUnderlineSize.AddItem FontSize
        Me.cboEmphasisSize.AddItem FontSize
        FontSize = FontSize + 1
    Loop
    
    'Populate Format Tab Normal Font Combobox
    For Each f In Application.FontNames
        Me.cboNormalFont.AddItem f
    Next f
    
    'Populate Format Tab Emphasis box size combobox
    Me.cboEmphasisBoxSize.AddItem "1pt"
    Me.cboEmphasisBoxSize.AddItem "1.5pt"
    Me.cboEmphasisBoxSize.AddItem "2.25pt"
    Me.cboEmphasisBoxSize.AddItem "3pt"
    
    'Format Tab
    Me.cboNormalSize.Value = GetSetting("Verbatim", "Format", "NormalSize", 11)
    Me.cboNormalFont.Value = GetSetting("Verbatim", "Format", "NormalFont", "Calibri")
    
    If GetSetting("Verbatim", "Format", "Spacing", "Wide") = "Wide" Then
        Me.optSpacingWide.Value = True
    Else
        Me.optSpacingNarrow.Value = True
    End If
    
    Me.cboPocketSize.Value = GetSetting("Verbatim", "Format", "PocketSize", 26)
    Me.cboHatSize.Value = GetSetting("Verbatim", "Format", "HatSize", 22)
    Me.cboBlockSize.Value = GetSetting("Verbatim", "Format", "BlockSize", 16)
    Me.cboTagSize.Value = GetSetting("Verbatim", "Format", "TagSize", 13)
    
    Me.cboCiteSize.Value = GetSetting("Verbatim", "Format", "CiteSize", 13)
    Me.chkUnderlineCite.Value = GetSetting("Verbatim", "Format", "UnderlineCite", False)
    
    Me.cboUnderlineSize.Value = GetSetting("Verbatim", "Format", "UnderlineSize", 11)
    Me.chkBoldUnderline.Value = GetSetting("Verbatim", "Format", "BoldUnderline", False)
    
    Me.cboEmphasisSize.Value = GetSetting("Verbatim", "Format", "EmphasisSize", 11)
    Me.chkEmphasisBold.Value = GetSetting("Verbatim", "Format", "EmphasisBold", True)
    Me.chkEmphasisItalic.Value = GetSetting("Verbatim", "Format", "EmphasisItalic", False)
    Me.chkEmphasisBox.Value = GetSetting("Verbatim", "Format", "EmphasisBox", False)
    Me.cboEmphasisBoxSize.Value = GetSetting("Verbatim", "Format", "EmphasisBoxSize", "1pt")
    
    Me.chkParagraphIntegrity.Value = GetSetting("Verbatim", "Format", "ParagraphIntegrity", False)
    Me.chkUsePilcrows.Value = GetSetting("Verbatim", "Format", "UsePilcrows", False)
    
    If GetSetting("Verbatim", "Format", "ShrinkMode", "Paragraph") = "Paragraph" Then
        Me.optParagraph.Value = True
    Else
        Me.optSelected.Value = True
    End If
    
    Me.chkAutoUnderlineEmphasis.Value = GetSetting("Verbatim", "Format", "AutoUnderlineEmphasis", False)
            
    'Populate Keyboard Tab Comboboxes
    MacroArray = Array("Paste", "Condense", "Pocket", "Hat", "Block", "Tag", "Cite", "Underline", "Emphasis", "Highlight", "Clear", "Shrink Text", "Select Similar")
    
    Me.cboF2Shortcut.List = MacroArray
    Me.cboF3Shortcut.List = MacroArray
    Me.cboF4Shortcut.List = MacroArray
    Me.cboF5Shortcut.List = MacroArray
    Me.cboF6Shortcut.List = MacroArray
    Me.cboF7Shortcut.List = MacroArray
    Me.cboF8Shortcut.List = MacroArray
    Me.cboF9Shortcut.List = MacroArray
    Me.cboF10Shortcut.List = MacroArray
    Me.cboF11Shortcut.List = MacroArray
    Me.cboF12Shortcut.List = MacroArray
    
    'Keyboard Tab
    Me.cboF2Shortcut.Value = GetSetting("Verbatim", "Keyboard", "F2Shortcut", "Paste")
    Me.cboF3Shortcut.Value = GetSetting("Verbatim", "Keyboard", "F3Shortcut", "Condense")
    Me.cboF4Shortcut.Value = GetSetting("Verbatim", "Keyboard", "F4Shortcut", "Pocket")
    Me.cboF5Shortcut.Value = GetSetting("Verbatim", "Keyboard", "F5Shortcut", "Hat")
    Me.cboF6Shortcut.Value = GetSetting("Verbatim", "Keyboard", "F6Shortcut", "Block")
    Me.cboF7Shortcut.Value = GetSetting("Verbatim", "Keyboard", "F7Shortcut", "Tag")
    Me.cboF8Shortcut.Value = GetSetting("Verbatim", "Keyboard", "F8Shortcut", "Cite")
    Me.cboF9Shortcut.Value = GetSetting("Verbatim", "Keyboard", "F9Shortcut", "Underline")
    Me.cboF10Shortcut.Value = GetSetting("Verbatim", "Keyboard", "F10Shortcut", "Emphasis")
    Me.cboF11Shortcut.Value = GetSetting("Verbatim", "Keyboard", "F11Shortcut", "Highlight")
    Me.cboF12Shortcut.Value = GetSetting("Verbatim", "Keyboard", "F12Shortcut", "Clear")
    
    'VTub Tab
    Me.cboVTubPath.Value = GetSetting("Verbatim", "VTub", "VTubPath")
    Me.chkVTubRefreshPrompt.Value = GetSetting("Verbatim", "VTub", "VTubRefreshPrompt", True)
       
    'Caselist Tab
    Me.cboCaselistSchoolName.Value = GetSetting("Verbatim", "Caselist", "CaselistSchoolName")
    Me.cboCaselistTeamName.Value = GetSetting("Verbatim", "Caselist", "CaselistTeamName")
        
    'About Tab
    Me.lblAbout2.Caption = "Verbatim v. " & Settings.GetVersion
    
    Exit Sub
        
Handler:
    MsgBox "Error " & Err.number & ": " & Err.Description

End Sub

Private Sub btnResetAllSettings_Click()
'Resets all settings to the default

    On Error GoTo Handler
    
    'Prompt for confirmation
    If MsgBox("This will reset all settings to their default values - changes will not be committed until you click Save. Are you sure?", vbOKCancel) = vbCancel Then Exit Sub
    
    'Main Tab
    Me.txtSchoolName.Value = ""
    Me.txtName.Value = ""
    Me.optCollege.Value = True
    Me.txtWPM.Value = 350
    Me.txtTabroomUsername.Value = ""
    Me.txtTabroomPassword.Value = ""
    Me.chkAutomaticUpdates.Value = True
    
    'Admin Tab
    Me.chkAlwaysOn.Value = True
    Me.chkAutoUpdateStyles.Value = True
    Me.chkSuppressInstallChecks.Value = False
    Me.chkSuppressDocCheck.Value = False
    Me.chkFirstRun.Value = False
    
    'View Tab
    Me.optWebView.Value = True
    Me.chkNPCStartup.Value = False
    Me.spnDocs.Value = 50
    Me.spnSpeech.Value = 50
    
    'Paperless Tab
    Me.chkAutoSaveSpeech.Value = False
    Me.cboAutoSaveDir.Value = ""
    Me.chkStripSpeech.Value = True
    Me.cboSearchDir.Value = ""
    Me.cboAutoOpenDir.Value = ""
    Me.cboAudioDir.Value = ""
    
    'Format Tab
    Me.cboNormalSize.Value = 11
    Me.cboNormalFont.Value = "Calibri"
    Me.optSpacingWide.Value = True
    Me.cboPocketSize.Value = 26
    Me.cboHatSize.Value = 22
    Me.cboBlockSize.Value = 16
    Me.cboTagSize.Value = 13
    Me.cboCiteSize.Value = 13
    Me.chkUnderlineCite.Value = False
    Me.cboUnderlineSize.Value = 11
    Me.chkBoldUnderline.Value = False
    Me.cboEmphasisSize.Value = 11
    Me.chkEmphasisBold.Value = True
    Me.chkEmphasisItalic.Value = False
    Me.chkEmphasisBox.Value = False
    Me.cboEmphasisBoxSize.Value = "1pt"
    
    Me.chkParagraphIntegrity.Value = False
    Me.chkUsePilcrows.Value = False
    Me.optParagraph.Value = True
    Me.chkAutoUnderlineEmphasis.Value = False
    
    'Keyboard Tab
    Me.cboF2Shortcut.Value = "Paste"
    Me.cboF3Shortcut.Value = "Condense"
    Me.cboF4Shortcut.Value = "Pocket"
    Me.cboF5Shortcut.Value = "Hat"
    Me.cboF6Shortcut.Value = "Block"
    Me.cboF7Shortcut.Value = "Tag"
    Me.cboF8Shortcut.Value = "Cite"
    Me.cboF9Shortcut.Value = "Underline"
    Me.cboF10Shortcut.Value = "Emphasis"
    Me.cboF11Shortcut.Value = "Highlight"
    Me.cboF12Shortcut.Value = "Clear"
    
    'VTub Tab
    Me.cboVTubPath.Value = ""
    Me.chkVTubRefreshPrompt.Value = True
    
    'PaDS Tab
    Me.txtPaDSSiteName.Value = ""
    Me.chkManualPaDSFolders.Value = False
    Me.cboCoauthoringFolder.Value = ""
    Me.cboPublicFolder.Value = ""
    Me.chkAutoSavePaDS.Value = False
    
    'Caselist Tab
    Me.optOpenCaselist.Value = True
    Me.cboCaselistSchoolName.Value = ""
    Me.cboCaselistTeamName.Value = ""
    Me.txtCustomPrefixes.Value = ""

    'Email Tab
    Me.optGmail.Value = True
    Me.txtGmailUsername.Value = ""
    Me.txtGmailPassword.Value = ""
    Me.txtEmailUsername.Value = ""
    Me.txtEmailPassword.Value = ""
    Me.txtSMTPServer.Value = ""
    Me.txtSMTPPort.Value = ""
    Me.chkUseSSL.Value = False
    
    'About Tab
    
    Exit Sub
        
Handler:
    MsgBox "Error " & Err.number & ": " & Err.Description

End Sub

Private Sub btnSave_Click()
'Save Settings to Registry

    Dim DebateTemplate As Document
    Dim CloseDebateTemplate As Boolean
    
    'Turn on Error handling
    On Error GoTo Handler
    
    'Main Tab
    SaveSetting "Verbatim", "Main", "SchoolName", Me.txtSchoolName.Value
    SaveSetting "Verbatim", "Main", "Name", Me.txtName.Value
    
    If Me.optCollege.Value = True Then
        SaveSetting "Verbatim", "Main", "CollegeHS", "College"
    Else
        SaveSetting "Verbatim", "Main", "CollegeHS", "HS"
    End If
        
    SaveSetting "Verbatim", "Main", "WPM", Me.txtWPM.Value
    SaveSetting "Verbatim", "Main", "TabroomUsername", Trim(Me.txtTabroomUsername.Value)
    If Me.txtTabroomPassword.Value <> "" Then
        SaveSetting "Verbatim", "Main", "TabroomPassword", XOREncryption(Me.txtTabroomPassword.Value)
    End If
    SaveSetting "Verbatim", "Main", "AutomaticUpdates", Me.chkAutomaticUpdates.Value
    
    'Admin Tab
    SaveSetting "Verbatim", "Admin", "AlwaysOn", Me.chkAlwaysOn.Value
    SaveSetting "Verbatim", "Admin", "AutoUpdateStyles", Me.chkAutoUpdateStyles.Value
    SaveSetting "Verbatim", "Admin", "SuppressInstallChecks", Me.chkSuppressInstallChecks.Value
    SaveSetting "Verbatim", "Admin", "SuppressDocCheck", Me.chkSuppressDocCheck.Value
    SaveSetting "Verbatim", "Admin", "FirstRun", Me.chkFirstRun.Value
    
    'View Tab
    If Me.optWebView.Value = True Then
        SaveSetting "Verbatim", "View", "DefaultView", "Web"
    Else
        SaveSetting "Verbatim", "View", "DefaultView", "Draft"
    End If

    SaveSetting "Verbatim", "View", "NPCStartup", Me.chkNPCStartup.Value
    SaveSetting "Verbatim", "View", "DocsPct", Me.spnDocs.Value
    SaveSetting "Verbatim", "View", "SpeechPct", Me.spnSpeech.Value
    
    'Paperless Tab
    SaveSetting "Verbatim", "Paperless", "AutoSaveSpeech", Me.chkAutoSaveSpeech.Value
    SaveSetting "Verbatim", "Paperless", "AutoSaveDir", Me.cboAutoSaveDir.Value
    SaveSetting "Verbatim", "Paperless", "StripSpeech", Me.chkStripSpeech.Value
    SaveSetting "Verbatim", "Paperless", "SearchDir", Me.cboSearchDir.Value
    SaveSetting "Verbatim", "Paperless", "AutoOpenDir", Me.cboAutoOpenDir.Value
    SaveSetting "Verbatim", "Paperless", "AudioDir", Me.cboAudioDir.Value
    
    'Format Tab
    SaveSetting "Verbatim", "Format", "NormalSize", Me.cboNormalSize.Value
    SaveSetting "Verbatim", "Format", "NormalFont", Me.cboNormalFont.Value
    
    If Me.optSpacingWide.Value = True Then
        SaveSetting "Verbatim", "Format", "Spacing", "Wide"
    Else
        SaveSetting "Verbatim", "Format", "Spacing", "Narrow"
    End If
    
    SaveSetting "Verbatim", "Format", "PocketSize", Me.cboPocketSize.Value
    SaveSetting "Verbatim", "Format", "HatSize", Me.cboHatSize.Value
    SaveSetting "Verbatim", "Format", "BlockSize", Me.cboBlockSize.Value
    SaveSetting "Verbatim", "Format", "TagSize", Me.cboTagSize.Value
    SaveSetting "Verbatim", "Format", "CiteSize", Me.cboCiteSize.Value
    SaveSetting "Verbatim", "Format", "UnderlineCite", Me.chkUnderlineCite.Value
    SaveSetting "Verbatim", "Format", "UnderlineSize", Me.cboUnderlineSize.Value
    SaveSetting "Verbatim", "Format", "BoldUnderline", Me.chkBoldUnderline.Value
    SaveSetting "Verbatim", "Format", "EmphasisSize", Me.cboEmphasisSize.Value
    SaveSetting "Verbatim", "Format", "EmphasisBold", Me.chkEmphasisBold.Value
    SaveSetting "Verbatim", "Format", "EmphasisItalic", Me.chkEmphasisItalic.Value
    SaveSetting "Verbatim", "Format", "EmphasisBox", Me.chkEmphasisBox.Value
    SaveSetting "Verbatim", "Format", "EmphasisBoxSize", Me.cboEmphasisBoxSize.Value
    SaveSetting "Verbatim", "Format", "ParagraphIntegrity", Me.chkParagraphIntegrity.Value
    SaveSetting "Verbatim", "Format", "UsePilcrows", Me.chkUsePilcrows.Value
    
    If Me.optParagraph.Value = True Then
        SaveSetting "Verbatim", "Format", "ShrinkMode", "Paragraph"
    Else
        SaveSetting "Verbatim", "Format", "ShrinkMode", "Selected"
    End If
    
    SaveSetting "Verbatim", "Format", "AutoUnderlineEmphasis", Me.chkAutoUnderlineEmphasis.Value
    
    'Check if Template itself is open, or open it as a Document
    If ActiveDocument.FullName = ActiveDocument.AttachedTemplate.FullName Then
        Set DebateTemplate = ActiveDocument
        CloseDebateTemplate = False
    Else
        Set DebateTemplate = ActiveDocument.AttachedTemplate.OpenAsDocument
        CloseDebateTemplate = True
    End If
    
    'Update template styles based on Format settings
    DebateTemplate.Styles("Normal").Font.Size = Me.cboNormalSize.Value
    DebateTemplate.Styles("Normal").Font.Name = Me.cboNormalFont.Value
    
    If Me.optSpacingWide.Value = True Then
        DebateTemplate.Styles("Normal").ParagraphFormat.SpaceBefore = 0
        DebateTemplate.Styles("Normal").ParagraphFormat.SpaceAfter = 8
        DebateTemplate.Styles("Normal").ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple
        DebateTemplate.Styles("Normal").ParagraphFormat.LineSpacing = LinesToPoints(1.08)
        DebateTemplate.Styles("Pocket").ParagraphFormat.SpaceBefore = 12
        DebateTemplate.Styles("Pocket").ParagraphFormat.SpaceAfter = 0
        DebateTemplate.Styles("Hat").ParagraphFormat.SpaceBefore = 2
        DebateTemplate.Styles("Hat").ParagraphFormat.SpaceAfter = 0
        DebateTemplate.Styles("Block").ParagraphFormat.SpaceBefore = 2
        DebateTemplate.Styles("Block").ParagraphFormat.SpaceAfter = 0
        DebateTemplate.Styles("Tag").ParagraphFormat.SpaceBefore = 2
        DebateTemplate.Styles("Tag").ParagraphFormat.SpaceAfter = 0
    Else
        DebateTemplate.Styles("Normal").ParagraphFormat.SpaceBefore = 0
        DebateTemplate.Styles("Normal").ParagraphFormat.SpaceAfter = 0
        DebateTemplate.Styles("Normal").ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
        DebateTemplate.Styles("Pocket").ParagraphFormat.SpaceBefore = 24
        DebateTemplate.Styles("Pocket").ParagraphFormat.SpaceAfter = 0
        DebateTemplate.Styles("Hat").ParagraphFormat.SpaceBefore = 24
        DebateTemplate.Styles("Hat").ParagraphFormat.SpaceAfter = 0
        DebateTemplate.Styles("Block").ParagraphFormat.SpaceBefore = 10
        DebateTemplate.Styles("Block").ParagraphFormat.SpaceAfter = 0
        DebateTemplate.Styles("Tag").ParagraphFormat.SpaceBefore = 10
        DebateTemplate.Styles("Tag").ParagraphFormat.SpaceAfter = 0
    End If
    
    DebateTemplate.Styles("Pocket").Font.Size = Me.cboPocketSize.Value
    DebateTemplate.Styles("Hat").Font.Size = Me.cboHatSize.Value
    DebateTemplate.Styles("Block").Font.Size = Me.cboBlockSize.Value
    DebateTemplate.Styles("Tag").Font.Size = Me.cboTagSize.Value
    DebateTemplate.Styles("Cite").Font.Size = Me.cboCiteSize.Value
    If Me.chkUnderlineCite.Value = True Then
        DebateTemplate.Styles("Cite").Font.Underline = wdUnderlineSingle
    Else
        DebateTemplate.Styles("Cite").Font.Underline = wdUnderlineNone
    End If
    DebateTemplate.Styles("Underline").Font.Size = Me.cboUnderlineSize.Value
    If Me.chkBoldUnderline.Value = True Then
        DebateTemplate.Styles("Underline").Font.Bold = True
    Else
        DebateTemplate.Styles("Underline").Font.Bold = False
    End If
    DebateTemplate.Styles("Emphasis").Font.Size = Me.cboEmphasisSize.Value
    DebateTemplate.Styles("Emphasis").Font.Name = Me.cboNormalFont.Value
    DebateTemplate.Styles("Emphasis").Font.Bold = Me.chkEmphasisBold.Value
    DebateTemplate.Styles("Emphasis").Font.Italic = Me.chkEmphasisItalic.Value
    
    If Me.chkEmphasisBox.Value = True Then
        DebateTemplate.Styles("Emphasis").Font.Borders(1).LineStyle = wdLineStyleSingle
    
        Select Case Me.cboEmphasisBoxSize.Value
            Case Is = "1pt"
                DebateTemplate.Styles("Emphasis").Font.Borders(1).LineWidth = wdLineWidth100pt
            Case Is = "1.5pt"
                DebateTemplate.Styles("Emphasis").Font.Borders(1).LineWidth = wdLineWidth150pt
            Case Is = "2.25pt"
                DebateTemplate.Styles("Emphasis").Font.Borders(1).LineWidth = wdLineWidth225pt
            Case Is = "3pt"
                DebateTemplate.Styles("Emphasis").Font.Borders(1).LineWidth = wdLineWidth300pt
            Case Else
                DebateTemplate.Styles("Emphasis").Font.Borders(1).LineWidth = wdLineWidth100pt
        End Select
    Else
        DebateTemplate.Styles("Emphasis").Font.Borders(1).LineStyle = wdLineStyleNone
    End If
    
    'Keyboard Tab
    SaveSetting "Verbatim", "Keyboard", "F2Shortcut", Me.cboF2Shortcut.Value
    SaveSetting "Verbatim", "Keyboard", "F3Shortcut", Me.cboF3Shortcut.Value
    SaveSetting "Verbatim", "Keyboard", "F4Shortcut", Me.cboF4Shortcut.Value
    SaveSetting "Verbatim", "Keyboard", "F5Shortcut", Me.cboF5Shortcut.Value
    SaveSetting "Verbatim", "Keyboard", "F6Shortcut", Me.cboF6Shortcut.Value
    SaveSetting "Verbatim", "Keyboard", "F7Shortcut", Me.cboF7Shortcut.Value
    SaveSetting "Verbatim", "Keyboard", "F8Shortcut", Me.cboF8Shortcut.Value
    SaveSetting "Verbatim", "Keyboard", "F9Shortcut", Me.cboF9Shortcut.Value
    SaveSetting "Verbatim", "Keyboard", "F10Shortcut", Me.cboF10Shortcut.Value
    SaveSetting "Verbatim", "Keyboard", "F11Shortcut", Me.cboF11Shortcut.Value
    SaveSetting "Verbatim", "Keyboard", "F12Shortcut", Me.cboF12Shortcut.Value
    
    'Update template keyboard shortcuts based on keyboard settings
    Call Settings.ChangeKeyboardShortcut(wdKeyF2, Me.cboF2Shortcut.Value)
    Call Settings.ChangeKeyboardShortcut(wdKeyF3, Me.cboF3Shortcut.Value)
    Call Settings.ChangeKeyboardShortcut(wdKeyF4, Me.cboF4Shortcut.Value)
    Call Settings.ChangeKeyboardShortcut(wdKeyF5, Me.cboF5Shortcut.Value)
    Call Settings.ChangeKeyboardShortcut(wdKeyF6, Me.cboF6Shortcut.Value)
    Call Settings.ChangeKeyboardShortcut(wdKeyF7, Me.cboF7Shortcut.Value)
    Call Settings.ChangeKeyboardShortcut(wdKeyF8, Me.cboF8Shortcut.Value)
    Call Settings.ChangeKeyboardShortcut(wdKeyF9, Me.cboF9Shortcut.Value)
    Call Settings.ChangeKeyboardShortcut(wdKeyF10, Me.cboF10Shortcut.Value)
    Call Settings.ChangeKeyboardShortcut(wdKeyF11, Me.cboF11Shortcut.Value)
    Call Settings.ChangeKeyboardShortcut(wdKeyF12, Me.cboF12Shortcut.Value)
    
    'Close template if opened separately
    If CloseDebateTemplate = True Then
        DebateTemplate.Close SaveChanges:=wdSaveChanges
    End If
    
    ActiveDocument.UpdateStyles
    
    'VTub Tab
    SaveSetting "Verbatim", "VTub", "VTubPath", Me.cboVTubPath.Value
    SaveSetting "Verbatim", "VTub", "VTubRefreshPrompt", chkVTubRefreshPrompt.Value
    
    'PaDS Tab
    SaveSetting "Verbatim", "PaDS", "PaDSSiteName", Me.txtPaDSSiteName.Value
    SaveSetting "Verbatim", "PaDS", "ManualPaDSFolders", Me.chkManualPaDSFolders.Value
    SaveSetting "Verbatim", "PaDS", "CoauthoringFolder", Me.cboCoauthoringFolder.Value
    SaveSetting "Verbatim", "PaDS", "PublicFolder", Me.cboPublicFolder.Value
    SaveSetting "Verbatim", "PaDS", "AutoSavePaDS", Me.chkAutoSavePaDS.Value
    
    'Caselist Tab
    If Me.optOpenCaselist.Value = True Then SaveSetting "Verbatim", "Caselist", "DefaultWiki", "openCaselist"
    If Me.optNDCAPolicy.Value = True Then SaveSetting "Verbatim", "Caselist", "DefaultWiki", "NDCAPolicy"
    If Me.optNDCALD.Value = True Then SaveSetting "Verbatim", "Caselist", "DefaultWiki", "NDCALD"
    
    SaveSetting "Verbatim", "Caselist", "CaselistSchoolName", Me.cboCaselistSchoolName.Value
    If Me.cboCaselistTeamName.Value <> "No teams found." Then SaveSetting "Verbatim", "Caselist", "CaselistTeamName", Me.cboCaselistTeamName.Value
    SaveSetting "Verbatim", "Caselist", "CustomPrefixes", Me.txtCustomPrefixes.Value
        
    'Email Tab
    If Me.optGmail.Value = True Then
        SaveSetting "Verbatim", "Email", "UseGmail", True
    Else
        SaveSetting "Verbatim", "Email", "UseGmail", False
    End If
    
    SaveSetting "Verbatim", "Email", "GmailUsername", Me.txtGmailUsername.Value
    If Me.txtGmailPassword.Value <> "" Then
        SaveSetting "Verbatim", "Email", "GmailPassword", XOREncryption(Me.txtGmailPassword.Value)
    End If
    SaveSetting "Verbatim", "Email", "EmailUsername", Me.txtEmailUsername.Value
    If Me.txtEmailPassword.Value <> "" Then
        SaveSetting "Verbatim", "Email", "EmailPassword", XOREncryption(Me.txtEmailPassword.Value)
    End If
    SaveSetting "Verbatim", "Email", "SMTPServer", Me.txtSMTPServer.Value
    SaveSetting "Verbatim", "Email", "SMTPPort", Me.txtSMTPPort.Value
    SaveSetting "Verbatim", "Email", "UseSSL", Me.chkUseSSL.Value
    
    'About Tab
    SaveSetting "Verbatim", "Main", "Version", Settings.GetVersion
    
    'Refresh ribbon in case keyboard shortcuts changed
    Ribbon.RefreshRibbon
    
    'Unload the form
    Unload Me
    Exit Sub
        
Handler:
    MsgBox "Error " & Err.number & ": " & Err.Description

End Sub

Private Sub btnCancel_Click()
  Unload Me
End Sub

'*************************************************************************************
'* MAIN TAB                                                                          *
'*************************************************************************************

Private Sub lblWPMLink_Click()
    Settings.LaunchWebsite ("http://www.readingsoft.com/")
End Sub
Private Sub lblTabroomRegister_Click()
    Settings.LaunchWebsite ("https://www.tabroom.com/user/login/new_user.mhtml")
End Sub

Private Sub btnUpdateCheck_Click()
    Settings.UpdateCheck (True)
End Sub

'*************************************************************************************
'* ADMIN TAB                                                                         *
'*************************************************************************************

Private Sub btnVerbatimizeNormal_Click()
    Settings.VerbatimizeNormal
End Sub

Private Sub btnUnverbatimizeNormal_Click()
    Settings.UnverbatimizeNormal
End Sub

Private Sub btnTemplatesFolder_Click()
    Settings.OpenTemplatesFolder
End Sub

Private Sub btnTutorial_Click()
    Unload Me
    Tutorial.LaunchTutorial
End Sub

Private Sub btnSetupWizard_Click()
    Unload Me
    UI.ShowForm "Setup"
End Sub

Private Sub btnTroubleshooter_Click()
    Unload Me
    UI.ShowForm "Troubleshooter"
End Sub

Private Sub btnImportSettings_Click()

    Dim SettingsFileName As String

    'Turn on Error Handling
    On Error GoTo Handler

    'Show the built-in file picker, only allow picking 1 file at a time
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    Application.FileDialog(msoFileDialogOpen).Filters.Clear
    Application.FileDialog(msoFileDialogOpen).Filters.Add "Verbatim Settings", "*.ini"
    Application.FileDialog(msoFileDialogOpen).Title = "Select Verbatim Settings file to import..."
    Application.FileDialog(msoFileDialogOpen).ButtonName = "Import"
    If Application.FileDialog(msoFileDialogOpen).Show = 0 Then 'Error trap cancel button
        Call Settings.ResetFileDialog(msoFileDialogOpen)
        Exit Sub
    End If
    
    SettingsFileName = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
    Settings.ResetFileDialog (msoFileDialogOpen)

    'Exit if trying to import an old settings file
    If System.PrivateProfileString(SettingsFileName, "Main", "Version") = "" Or System.PrivateProfileString(SettingsFileName, "Main", "Version") < 5 Then
        MsgBox "Outdated settings file. You must use a Verbatim settings file exported from v5.0 or newer."
        Exit Sub
    End If
    
    'Import settings - Main
    If System.PrivateProfileString(SettingsFileName, "Main", "CollegeHS") = "College" Then
        Me.optCollege.Value = True
    Else
        Me.optHS.Value = True
    End If
    
    Me.chkAutomaticUpdates.Value = System.PrivateProfileString(SettingsFileName, "Main", "AutomaticUpdates")
    
    'Import settings - Admin
    Me.chkAlwaysOn.Value = System.PrivateProfileString(SettingsFileName, "Admin", "AlwaysOn")
    Me.chkAutoUpdateStyles.Value = System.PrivateProfileString(SettingsFileName, "Admin", "AutoUpdateStyles")
    Me.chkSuppressInstallChecks.Value = System.PrivateProfileString(SettingsFileName, "Admin", "SuppressInstallChecks")
    Me.chkSuppressDocCheck.Value = System.PrivateProfileString(SettingsFileName, "Admin", "SuppressDocCheck")
    
    'Import settings - View
    If System.PrivateProfileString(SettingsFileName, "View", "DefaultView") = "Web" Then
        Me.optWebView.Value = True
    Else
        Me.optDraftView.Value = True
    End If
    
    Me.chkNPCStartup.Value = System.PrivateProfileString(SettingsFileName, "View", "NPCStartup")
    Me.spnDocs.Value = System.PrivateProfileString(SettingsFileName, "View", "DocsPct")
    Me.spnSpeech.Value = System.PrivateProfileString(SettingsFileName, "View", "SpeechPct")
    
    'Import settings - Paperless
    Me.chkAutoSaveSpeech.Value = System.PrivateProfileString(SettingsFileName, "Paperless", "AutoSaveSpeech")
    Me.cboAutoSaveDir.Value = System.PrivateProfileString(SettingsFileName, "Paperless", "AutoSaveDir")
    Me.chkStripSpeech.Value = System.PrivateProfileString(SettingsFileName, "Paperless", "StripSpeech")
    Me.cboSearchDir.Value = System.PrivateProfileString(SettingsFileName, "Paperless", "SearchDir")
    Me.cboAutoOpenDir.Value = System.PrivateProfileString(SettingsFileName, "Paperless", "AutoOpenDir")
    Me.cboAudioDir.Value = System.PrivateProfileString(SettingsFileName, "Paperless", "AudioDir")
    
    'Import settings - Format
    Me.cboNormalSize.Value = System.PrivateProfileString(SettingsFileName, "Format", "NormalSize")
    Me.cboNormalFont.Value = System.PrivateProfileString(SettingsFileName, "Format", "NormalFont")
    
    If System.PrivateProfileString(SettingsFileName, "Format", "Spacing") = "Wide" Then
        Me.optSpacingWide.Value = True
    Else
        Me.optSpacingNarrow.Value = True
    End If
    
    Me.cboPocketSize.Value = System.PrivateProfileString(SettingsFileName, "Format", "PocketSize")
    Me.cboHatSize.Value = System.PrivateProfileString(SettingsFileName, "Format", "HatSize")
    Me.cboBlockSize.Value = System.PrivateProfileString(SettingsFileName, "Format", "BlockSize")
    Me.cboTagSize.Value = System.PrivateProfileString(SettingsFileName, "Format", "TagSize")
    Me.cboCiteSize.Value = System.PrivateProfileString(SettingsFileName, "Format", "CiteSize")
    Me.chkUnderlineCite.Value = System.PrivateProfileString(SettingsFileName, "Format", "UnderlineCite")
    Me.cboUnderlineSize.Value = System.PrivateProfileString(SettingsFileName, "Format", "UnderlineSize")
    Me.chkBoldUnderline.Value = System.PrivateProfileString(SettingsFileName, "Format", "BoldUnderline")
    Me.cboEmphasisSize.Value = System.PrivateProfileString(SettingsFileName, "Format", "EmphasisSize")
    Me.chkEmphasisBold.Value = System.PrivateProfileString(SettingsFileName, "Format", "EmphasisBold")
    Me.chkEmphasisItalic.Value = System.PrivateProfileString(SettingsFileName, "Format", "EmphasisItalic")
    Me.chkEmphasisBox.Value = System.PrivateProfileString(SettingsFileName, "Format", "EmphasisBox")
    Me.cboEmphasisBoxSize.Value = System.PrivateProfileString(SettingsFileName, "Format", "EmphasisBoxSize")
    Me.chkParagraphIntegrity.Value = System.PrivateProfileString(SettingsFileName, "Format", "ParagraphIntegrity")
    Me.chkUsePilcrows.Value = System.PrivateProfileString(SettingsFileName, "Format", "UsePilcrows")

    If System.PrivateProfileString(SettingsFileName, "Format", "ShrinkMode") = "Paragraph" Then
        Me.optParagraph.Value = True
    Else
        Me.optSelected.Value = True
    End If

    Me.chkAutoUnderlineEmphasis.Value = System.PrivateProfileString(SettingsFileName, "Format", "AutoUnderlineEmphasis")

    'Import settings - Keyboard
    Me.cboF2Shortcut.Value = System.PrivateProfileString(SettingsFileName, "Keyboard", "F2Shortcut")
    Me.cboF3Shortcut.Value = System.PrivateProfileString(SettingsFileName, "Keyboard", "F3Shortcut")
    Me.cboF4Shortcut.Value = System.PrivateProfileString(SettingsFileName, "Keyboard", "F4Shortcut")
    Me.cboF5Shortcut.Value = System.PrivateProfileString(SettingsFileName, "Keyboard", "F5Shortcut")
    Me.cboF6Shortcut.Value = System.PrivateProfileString(SettingsFileName, "Keyboard", "F6Shortcut")
    Me.cboF7Shortcut.Value = System.PrivateProfileString(SettingsFileName, "Keyboard", "F7Shortcut")
    Me.cboF8Shortcut.Value = System.PrivateProfileString(SettingsFileName, "Keyboard", "F8Shortcut")
    Me.cboF9Shortcut.Value = System.PrivateProfileString(SettingsFileName, "Keyboard", "F9Shortcut")
    Me.cboF10Shortcut.Value = System.PrivateProfileString(SettingsFileName, "Keyboard", "F10Shortcut")
    Me.cboF11Shortcut.Value = System.PrivateProfileString(SettingsFileName, "Keyboard", "F11Shortcut")
    Me.cboF12Shortcut.Value = System.PrivateProfileString(SettingsFileName, "Keyboard", "F12Shortcut")
    
    'Import settings - VTub
    Me.cboVTubPath.Value = System.PrivateProfileString(SettingsFileName, "VTub", "VTubPath")
    Me.chkVTubRefreshPrompt.Value = System.PrivateProfileString(SettingsFileName, "VTub", "VTubRefreshPrompt")

    'Import settings - PaDS
    Me.txtPaDSSiteName.Value = System.PrivateProfileString(SettingsFileName, "PaDS", "PaDSSiteName")
    Me.chkManualPaDSFolders.Value = System.PrivateProfileString(SettingsFileName, "PaDS", "ManualPaDSFolders")
    Me.cboCoauthoringFolder.Value = System.PrivateProfileString(SettingsFileName, "PaDS", "CoauthoringFolder")
    Me.cboPublicFolder.Value = System.PrivateProfileString(SettingsFileName, "PaDS", "PublicFolder")
    Me.chkAutoSavePaDS.Value = System.PrivateProfileString(SettingsFileName, "PaDS", "AutoSavePaDS")
    
    'Import settings - Caselist
    Select Case System.PrivateProfileString(SettingsFileName, "Caselist", "DefaultWiki")
        Case Is = "openCaselist"
            Me.optOpenCaselist.Value = True
        Case Is = "NDCAPolicy"
            Me.optNDCAPolicy.Value = True
        Case Is = "NDCALD"
            Me.optNDCALD.Value = True
        Case Else
            Me.optOpenCaselist.Value = True
    End Select
    
    Me.cboCaselistSchoolName.Value = System.PrivateProfileString(SettingsFileName, "Caselist", "CaselistSchoolName")
    Me.cboCaselistTeamName.Value = System.PrivateProfileString(SettingsFileName, "Caselist", "CaselistTeamName")
    Me.txtCustomPrefixes.Value = System.PrivateProfileString(SettingsFileName, "Caselist", "CustomPrefixes")

    'Import settings - Email
    Me.txtSMTPServer.Value = System.PrivateProfileString(SettingsFileName, "Email", "SMTPServer")
    Me.txtSMTPPort.Value = System.PrivateProfileString(SettingsFileName, "Email", "SMTPPort")
    Me.chkUseSSL.Value = System.PrivateProfileString(SettingsFileName, "Email", "UseSSL")
    
    'Report success
    MsgBox "Settings successfully imported. They will not be committed until you click Save."
    
    Exit Sub

Handler:
    MsgBox "Error " & Err.number & ": " & Err.Description

End Sub

Private Sub btnExportSettings_Click()

    Dim SettingsFileName As String
    Dim ExportPath As String

    'Turn on Error Handling
    On Error GoTo Handler

    'Create SettingsFile name
    SettingsFileName = "VerbatimSettings"
    If Me.txtSchoolName.Value <> "" Then
        SettingsFileName = SettingsFileName & " - " & Me.txtSchoolName.Value
    End If
    If Me.txtName.Value <> "" Then
        SettingsFileName = SettingsFileName & " - " & Me.txtName.Value
    End If
    SettingsFileName = SettingsFileName & ".ini"

    'Show the built-in folder picker, only allow picking 1 folder at a time
    Application.FileDialog(msoFileDialogFolderPicker).AllowMultiSelect = False
    Application.FileDialog(msoFileDialogFolderPicker).Title = "Choose folder for export..."
    Application.FileDialog(msoFileDialogFolderPicker).ButtonName = "Export"
    If Application.FileDialog(msoFileDialogFolderPicker).Show = 0 Then 'Error trap cancel button
        Settings.ResetFileDialog (msoFileDialogFolderPicker)
        Exit Sub
    End If
    
    ExportPath = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
    SettingsFileName = ExportPath & "\" & SettingsFileName
    Settings.ResetFileDialog (msoFileDialogFolderPicker)

    'Set settings file version
    System.PrivateProfileString(SettingsFileName, "Main", "Version") = Settings.GetVersion

    'Export settings - Main
    If Me.optCollege.Value = True Then
        System.PrivateProfileString(SettingsFileName, "Main", "CollegeHS") = "College"
    Else
        System.PrivateProfileString(SettingsFileName, "Main", "CollegeHS") = "HS"
    End If
    
    System.PrivateProfileString(SettingsFileName, "Main", "AutomaticUpdates") = Me.chkAutomaticUpdates.Value
    
    'Export settings - Admin
    System.PrivateProfileString(SettingsFileName, "Admin", "AlwaysOn") = Me.chkAlwaysOn.Value
    System.PrivateProfileString(SettingsFileName, "Admin", "AutoUpdateStyles") = Me.chkAutoUpdateStyles.Value
    System.PrivateProfileString(SettingsFileName, "Admin", "SuppressInstallChecks") = Me.chkSuppressInstallChecks.Value
    System.PrivateProfileString(SettingsFileName, "Admin", "SuppressDocCheck") = Me.chkSuppressDocCheck.Value
    
    'Export settings - View
    If Me.optWebView.Value = True Then
        System.PrivateProfileString(SettingsFileName, "View", "DefaultView") = "Web"
    Else
        System.PrivateProfileString(SettingsFileName, "View", "DefaultView") = "Draft"
    End If
        
    System.PrivateProfileString(SettingsFileName, "View", "NPCStartup") = Me.chkNPCStartup.Value
    System.PrivateProfileString(SettingsFileName, "View", "DocsPct") = Me.spnDocs.Value
    System.PrivateProfileString(SettingsFileName, "View", "SpeechPct") = Me.spnDocs.Value
    
    'Export settings - Paperless
    System.PrivateProfileString(SettingsFileName, "Paperless", "AutoSaveSpeech") = Me.chkAutoSaveSpeech.Value
    System.PrivateProfileString(SettingsFileName, "Paperless", "AutoSaveDir") = Me.cboAutoSaveDir.Value
    System.PrivateProfileString(SettingsFileName, "Paperless", "StripSpeech") = Me.chkStripSpeech.Value
    System.PrivateProfileString(SettingsFileName, "Paperless", "SearchDir") = Me.cboSearchDir.Value
    System.PrivateProfileString(SettingsFileName, "Paperless", "AutoOpenDir") = Me.cboAutoOpenDir.Value
    System.PrivateProfileString(SettingsFileName, "Paperless", "AudioDir") = Me.cboAudioDir.Value

    'Export settings - Format
    System.PrivateProfileString(SettingsFileName, "Format", "NormalSize") = Me.cboNormalSize.Value
    System.PrivateProfileString(SettingsFileName, "Format", "NormalFont") = Me.cboNormalFont.Value
    
    If Me.optSpacingWide.Value = True Then
        System.PrivateProfileString(SettingsFileName, "Format", "Spacing") = "Wide"
    Else
        System.PrivateProfileString(SettingsFileName, "Format", "Spacing") = "Narrow"
    End If
    
    System.PrivateProfileString(SettingsFileName, "Format", "PocketSize") = Me.cboPocketSize.Value
    System.PrivateProfileString(SettingsFileName, "Format", "HatSize") = Me.cboHatSize.Value
    System.PrivateProfileString(SettingsFileName, "Format", "BlockSize") = Me.cboBlockSize.Value
    System.PrivateProfileString(SettingsFileName, "Format", "TagSize") = Me.cboTagSize.Value
    System.PrivateProfileString(SettingsFileName, "Format", "CiteSize") = Me.cboCiteSize.Value
    System.PrivateProfileString(SettingsFileName, "Format", "UnderlineCite") = Me.chkUnderlineCite.Value
    System.PrivateProfileString(SettingsFileName, "Format", "UnderlineSize") = Me.cboUnderlineSize.Value
    System.PrivateProfileString(SettingsFileName, "Format", "BoldUnderline") = Me.chkBoldUnderline.Value
    System.PrivateProfileString(SettingsFileName, "Format", "EmphasisSize") = Me.cboEmphasisSize.Value
    System.PrivateProfileString(SettingsFileName, "Format", "EmphasisBold") = Me.chkEmphasisBold.Value
    System.PrivateProfileString(SettingsFileName, "Format", "EmphasisItalic") = Me.chkEmphasisItalic.Value
    System.PrivateProfileString(SettingsFileName, "Format", "EmphasisBox") = Me.chkEmphasisBox.Value
    System.PrivateProfileString(SettingsFileName, "Format", "EmphasisBoxSize") = Me.cboEmphasisBoxSize.Value
    System.PrivateProfileString(SettingsFileName, "Format", "ParagraphIntegrity") = Me.chkParagraphIntegrity.Value
    System.PrivateProfileString(SettingsFileName, "Format", "UsePilcrows") = Me.chkUsePilcrows.Value
    
    If Me.optParagraph.Value = True Then
        System.PrivateProfileString(SettingsFileName, "Format", "ShrinkMode") = "Paragraph"
    Else
        System.PrivateProfileString(SettingsFileName, "Format", "ShrinkMode") = "Selected"
    End If

    System.PrivateProfileString(SettingsFileName, "Format", "AutoUnderlineEmphasis") = Me.chkAutoUnderlineEmphasis.Value
    
    'Export settings - Keyboard
    System.PrivateProfileString(SettingsFileName, "Keyboard", "F2Shortcut") = Me.cboF2Shortcut.Value
    System.PrivateProfileString(SettingsFileName, "Keyboard", "F3Shortcut") = Me.cboF3Shortcut.Value
    System.PrivateProfileString(SettingsFileName, "Keyboard", "F4Shortcut") = Me.cboF4Shortcut.Value
    System.PrivateProfileString(SettingsFileName, "Keyboard", "F5Shortcut") = Me.cboF5Shortcut.Value
    System.PrivateProfileString(SettingsFileName, "Keyboard", "F6Shortcut") = Me.cboF6Shortcut.Value
    System.PrivateProfileString(SettingsFileName, "Keyboard", "F7Shortcut") = Me.cboF7Shortcut.Value
    System.PrivateProfileString(SettingsFileName, "Keyboard", "F8Shortcut") = Me.cboF8Shortcut.Value
    System.PrivateProfileString(SettingsFileName, "Keyboard", "F9Shortcut") = Me.cboF9Shortcut.Value
    System.PrivateProfileString(SettingsFileName, "Keyboard", "F10Shortcut") = Me.cboF10Shortcut.Value
    System.PrivateProfileString(SettingsFileName, "Keyboard", "F11Shortcut") = Me.cboF11Shortcut.Value
    System.PrivateProfileString(SettingsFileName, "Keyboard", "F12Shortcut") = Me.cboF12Shortcut.Value

    'Export settings - VTub
    System.PrivateProfileString(SettingsFileName, "VTub", "VTubPath") = Me.cboVTubPath.Value
    System.PrivateProfileString(SettingsFileName, "VTub", "VTubRefreshPrompt") = Me.chkVTubRefreshPrompt.Value

    'Export settings - PaDS
    System.PrivateProfileString(SettingsFileName, "PaDS", "PaDSSiteName") = Me.txtPaDSSiteName.Value
    System.PrivateProfileString(SettingsFileName, "PaDS", "ManualPaDSFolders") = Me.chkManualPaDSFolders.Value
    System.PrivateProfileString(SettingsFileName, "PaDS", "CoauthoringFolder") = Me.cboCoauthoringFolder.Value
    System.PrivateProfileString(SettingsFileName, "PaDS", "PublicFolder") = Me.cboPublicFolder.Value
    System.PrivateProfileString(SettingsFileName, "PaDS", "AutoSavePaDS") = Me.chkAutoSavePaDS.Value

    'Export settings - Caselist
    If Me.optOpenCaselist.Value = True Then System.PrivateProfileString(SettingsFileName, "Caselist", "DefaultWiki") = "openCaselist"
    If Me.optNDCAPolicy.Value = True Then System.PrivateProfileString(SettingsFileName, "Caselist", "DefaultWiki") = "NDCAPolicy"
    If Me.optNDCALD.Value = True Then System.PrivateProfileString(SettingsFileName, "Caselist", "DefaultWiki") = "NDCALD"
    
    System.PrivateProfileString(SettingsFileName, "Caselist", "CaselistSchoolName") = Me.cboCaselistSchoolName.Value
    If Me.cboCaselistTeamName.Value <> "No teams found." Then System.PrivateProfileString(SettingsFileName, "Caselist", "CaselistTeamName") = Me.cboCaselistTeamName.Value
    System.PrivateProfileString(SettingsFileName, "Caselist", "CustomPrefixes") = Me.txtCustomPrefixes.Value

    'Export settings - Email
    System.PrivateProfileString(SettingsFileName, "Email", "SMTPServer") = Me.txtSMTPServer.Value
    System.PrivateProfileString(SettingsFileName, "Email", "SMTPPort") = Me.txtSMTPPort.Value
    System.PrivateProfileString(SettingsFileName, "Email", "UseSSL") = Me.chkUseSSL.Value

    'Report success
    MsgBox "Settings successfully exported as:" & vbCrLf & SettingsFileName
    
    Exit Sub

Handler:
    MsgBox "Error " & Err.number & ": " & Err.Description

End Sub

Private Sub btnImportCustomCode_Click()
    Call Settings.ImportCustomCode(Notify:=True)
End Sub

Private Sub btnExportCustomCode_Click()
    Call Settings.ExportCustomCode(Notify:=True)
End Sub

'*************************************************************************************
'* VIEW TAB                                                                          *
'*************************************************************************************

Private Sub spnDocs_Change()
    Me.txtDocPct.Value = Me.spnDocs.Value
    Me.lblDocs.Width = 200 * Me.spnDocs.Value / 100
    Me.lblSpeech.Width = (200 * Me.spnSpeech.Value / 100)
    Me.lblSpeech.Left = 200 - Me.lblSpeech.Width
End Sub

Private Sub spnSpeech_Change()
    Me.txtSpeechPct.Value = Me.spnSpeech.Value
    Me.lblDocs.Width = 200 * Me.spnDocs.Value / 100
    Me.lblSpeech.Width = (200 * Me.spnSpeech.Value / 100)
    Me.lblSpeech.Left = 200 - Me.lblSpeech.Width
End Sub

Private Sub btnResetView_Click()

    If MsgBox("This will reset view settings to their default values. Changes will not be committed until you click Save. Are you sure?", vbOKCancel) = vbCancel Then Exit Sub
    
    Me.optWebView.Value = True
    Me.chkNPCStartup.Value = False
    Me.spnDocs.Value = 50
    Me.spnSpeech.Value = 50
End Sub

'*************************************************************************************
'* PAPERLESS TAB                                                                     *
'*************************************************************************************

Private Sub chkAutoSaveSpeech_Change()
    If Me.chkAutoSaveSpeech.Value = True Then
        Me.cboAutoSaveDir.Enabled = True
        Me.lblAutoSaveDir.Enabled = True
    Else
        Me.cboAutoSaveDir.Enabled = False
        Me.lblAutoSaveDir.Enabled = False
    End If
End Sub

Private Sub cboAutoSaveDir_DropButtonClick()

    On Error Resume Next
    
    'Show the built-in folder picker, only allow picking 1 folder at a time
    Application.FileDialog(msoFileDialogFolderPicker).AllowMultiSelect = False
    If Application.FileDialog(msoFileDialogFolderPicker).Show = 0 Then 'Error trap cancel button
        Me.btnCancel.SetFocus 'Have to switch focus to avoid dropdown getting stuck
        Settings.ResetFileDialog (msoFileDialogFolderPicker)
        Exit Sub
    End If
    
    'Populate the combobox with the current directory, set by the folder dialog
    Me.cboAutoSaveDir.Value = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
    Me.btnCancel.SetFocus 'Have to switch focus to avoid dropdown getting stuck
    Settings.ResetFileDialog (msoFileDialogFolderPicker)

End Sub
Private Sub cboSearchDir_DropButtonClick()

    On Error Resume Next
    
    'Show the built-in folder picker, only allow picking 1 folder at a time
    Application.FileDialog(msoFileDialogFolderPicker).AllowMultiSelect = False
    If Application.FileDialog(msoFileDialogFolderPicker).Show = 0 Then 'Error trap cancel button
        Me.btnCancel.SetFocus 'Have to switch focus to avoid dropdown getting stuck
        Settings.ResetFileDialog (msoFileDialogFolderPicker)
        Exit Sub
    End If
    
    'Populate the combobox with the current directory, set by the folder dialog
    Me.cboSearchDir.Value = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
    Me.btnCancel.SetFocus 'Have to switch focus to avoid dropdown getting stuck
    Settings.ResetFileDialog (msoFileDialogFolderPicker)
    
End Sub
Private Sub cboAutoOpenDir_DropButtonClick()

    On Error Resume Next
    
    'Show the built-in folder picker, only allow picking 1 folder at a time
    Application.FileDialog(msoFileDialogFolderPicker).AllowMultiSelect = False
    If Application.FileDialog(msoFileDialogFolderPicker).Show = 0 Then 'Error trap cancel button
        Me.btnCancel.SetFocus 'Have to switch focus to avoid dropdown getting stuck
        Settings.ResetFileDialog (msoFileDialogFolderPicker)
        Exit Sub
    End If
    
    'Populate the combobox with the current directory, set by the folder dialog
    Me.cboAutoOpenDir.Value = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
    Me.btnCancel.SetFocus 'Have to switch focus to avoid dropdown getting stuck
    Settings.ResetFileDialog (msoFileDialogFolderPicker)
    
End Sub

Private Sub cboAudioDir_DropButtonClick()

    On Error Resume Next
    
    'Show the built-in folder picker, only allow picking 1 folder at a time
    Application.FileDialog(msoFileDialogFolderPicker).AllowMultiSelect = False
    If Application.FileDialog(msoFileDialogFolderPicker).Show = 0 Then 'Error trap cancel button
        Me.btnCancel.SetFocus 'Have to switch focus to avoid dropdown getting stuck
        Settings.ResetFileDialog (msoFileDialogFolderPicker)
        Exit Sub
    End If
    
    'Populate the combobox with the current directory, set by the folder dialog
    Me.cboAudioDir.Value = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
    Me.btnCancel.SetFocus 'Have to switch focus to avoid dropdown getting stuck
    Settings.ResetFileDialog (msoFileDialogFolderPicker)
End Sub

'*************************************************************************************
'* FORMAT TAB                                                                        *
'*************************************************************************************

Private Sub cboNormalFont_Change()
    'Changes the font sample
    Me.lblFontSample2.Font.Name = Me.cboNormalFont.Value
End Sub

Private Sub chkEmphasisBox_Change()
    If Me.chkEmphasisBox.Value = True Then
        Me.cboEmphasisBoxSize.Enabled = True
    Else
        Me.cboEmphasisBoxSize.Enabled = False
    End If

End Sub

Private Sub chkParagraphIntegrity_Change()
    'Disable Pilcrows button if unchecked
    If Me.chkParagraphIntegrity.Value = False Then
        Me.chkUsePilcrows.Enabled = False
    Else
        Me.chkUsePilcrows.Enabled = True
    End If

End Sub

Private Sub btnResetFormatting_Click()
'Resets formatting settings to the default
    
    On Error GoTo Handler
    
    'Prompt for confirmation
    If MsgBox("This will reset formatting settings to their default values. Changes will not be committed until you click Save. Are you sure?", vbOKCancel) = vbCancel Then Exit Sub
    
    'Format Tab
    Me.cboNormalSize.Value = 11
    Me.cboNormalFont.Value = "Calibri"
    Me.optSpacingWide.Value = True
    Me.cboPocketSize.Value = 26
    Me.cboHatSize.Value = 22
    Me.cboBlockSize.Value = 16
    Me.cboTagSize.Value = 13
    Me.cboCiteSize.Value = 13
    Me.chkUnderlineCite.Value = False
    Me.cboUnderlineSize.Value = 11
    Me.chkBoldUnderline.Value = False
    Me.cboEmphasisSize.Value = 11
    Me.chkEmphasisBold.Value = True
    Me.chkEmphasisItalic.Value = False
    Me.chkEmphasisBox.Value = False
    Me.cboEmphasisBoxSize.Value = "1pt"
    
    Me.chkParagraphIntegrity.Value = False
    Me.chkUsePilcrows.Value = False
    Me.optParagraph.Value = True
    
    Exit Sub
        
Handler:
    MsgBox "Error " & Err.number & ": " & Err.Description

End Sub

'*************************************************************************************
'* KEYBOARD TAB                                                                      *
'*************************************************************************************

Private Sub btnOtherKeyboardShortcuts_Click()
    'Shows the Customize Keyboard dialogue
    Dialogs(wdDialogToolsCustomizeKeyboard).Show
End Sub

Private Sub btnResetKeyboard_Click()
'Resets keyboard settings to the default
    
    On Error GoTo Handler
    
    'Prompt for confirmation
    If MsgBox("This will reset keyboard settings to their default values. Changes will not be committed until you click Save. Are you sure?", vbOKCancel) = vbCancel Then Exit Sub
    
    'Keyboard Tab
    Me.cboF2Shortcut.Value = "Paste"
    Me.cboF3Shortcut.Value = "Condense"
    Me.cboF4Shortcut.Value = "Pocket"
    Me.cboF5Shortcut.Value = "Hat"
    Me.cboF6Shortcut.Value = "Block"
    Me.cboF7Shortcut.Value = "Tag"
    Me.cboF8Shortcut.Value = "Cite"
    Me.cboF9Shortcut.Value = "Underline"
    Me.cboF10Shortcut.Value = "Emphasis"
    Me.cboF11Shortcut.Value = "Highlight"
    Me.cboF12Shortcut.Value = "Clear"
    
    Exit Sub
    
Handler:
    MsgBox "Error " & Err.number & ": " & Err.Description

End Sub

Private Sub btnFixTilde_Click()
    Call Troubleshooting.FixTilde
    MsgBox "Tilde shortcuts fixed!"
End Sub

'*************************************************************************************
'* VTUB TAB                                                                          *
'*************************************************************************************

Private Sub cboVTubPath_DropButtonClick()

    On Error Resume Next
    
    'Show the built-in folder picker, only allow picking 1 folder at a time
    Application.FileDialog(msoFileDialogFolderPicker).AllowMultiSelect = False
    If Application.FileDialog(msoFileDialogFolderPicker).Show = 0 Then 'Error trap cancel button
        Me.btnCancel.SetFocus 'Have to switch focus to avoid dropdown getting stuck
        Settings.ResetFileDialog (msoFileDialogFolderPicker)
        Exit Sub
    End If
    
    'Populate the combobox with the current directory, set by the folder dialog
    Me.cboVTubPath.Value = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
    Me.btnCancel.SetFocus 'Have to switch focus to avoid dropdown getting stuck
    Settings.ResetFileDialog (msoFileDialogFolderPicker)
    
    'Save immediately so the Create button can find a path
    SaveSetting "Verbatim", "VTub", "VTubPath", Me.cboVTubPath.Value
End Sub

Private Sub btnCreateVTub_Click()
    If Me.cboVTubPath.Value = "" Then
        MsgBox "You must select a path for the VTub first."
        Exit Sub
    Else
        Me.Hide
        Call VirtualTub.VTubCreate
        Me.Show
    End If
End Sub

'*************************************************************************************
'* PADS TAB                                                                          *
'*************************************************************************************

Private Sub txtPaDSSiteName_Change()
    If Me.chkManualPaDSFolders.Value = False Then
        Me.cboCoauthoringFolder.Value = "http://" & Me.txtPaDSSiteName.Value & ".paperlessdebate.com/Team Tubs/"
        Me.cboPublicFolder.Value = "http://" & Me.txtPaDSSiteName.Value & ".paperlessdebate.com/Public/"
    End If
End Sub

Private Sub chkManualPaDSFolders_Click()
    If Me.chkManualPaDSFolders.Value = True Then
        Me.lblCoauthoringFolder.Enabled = True
        Me.cboCoauthoringFolder.Enabled = True
        Me.lblPublicFolder.Enabled = True
        Me.cboPublicFolder.Enabled = True
        Me.txtPaDSSiteName.Enabled = False
    Else
        Me.lblCoauthoringFolder.Enabled = False
        Me.cboCoauthoringFolder.Enabled = False
        Me.lblPublicFolder.Enabled = False
        Me.cboPublicFolder.Enabled = False
        Me.txtPaDSSiteName.Enabled = True
    End If
End Sub

Private Sub cboCoauthoringFolder_DropButtonClick()
    
    Call GetPaDSFolder(Me.cboCoauthoringFolder)

End Sub

Private Sub cboPublicFolder_DropButtonClick()
    
    Call GetPaDSFolder(Me.cboPublicFolder)
    
End Sub

Private Sub GetPaDSFolder(c As control)
    
    On Error GoTo Handler
    
    'Check tabroom username and password are entered
    If Me.txtTabroomUsername.Value = "" Then
        MsgBox "Cannot connect to PaDS unless you have entered a tabroom username and password on the Main tab."
        Exit Sub
    End If
    
    'Send SOAP authorization to log-in to PaDS
    Call PaDS.SOAPLogin(Me.txtPaDSSiteName.Value, Me.txtTabroomUsername.Value, XORDecryption(GetSetting("Verbatim", "Main", "TabroomPassword")))
    Call PaDS.PaDSIELogin(Me.txtPaDSSiteName.Value, Me.txtTabroomUsername.Value, XORDecryption(GetSetting("Verbatim", "Main", "TabroomPassword")))
    
    'Set default browse location
    If c.Value <> "" Then
        If Right(c.Value, 1) <> "/" Then c.Value = c.Value & "/" 'Add a trailing slash to avoid breaking URL check
        If PaDS.SharepointURLExists(c.Value, True) = False Then
            MsgBox "The PaDS location you have entered doesn't exist."
            Exit Sub
        End If
        Application.FileDialog(msoFileDialogFolderPicker).InitialFileName = c.Value
    ElseIf Me.txtPaDSSiteName.Value <> "" Then
        Application.FileDialog(msoFileDialogFolderPicker).InitialFileName = "http://" & Me.txtPaDSSiteName & ".paperlessdebate.com"
    Else
        MsgBox "You must enter a PaDS Site Name before you can browse for a folder"
        Exit Sub
    End If
    
    'Show the built-in folder picker, only allow picking 1 folder at a time
    Application.FileDialog(msoFileDialogFolderPicker).AllowMultiSelect = False
    If Application.FileDialog(msoFileDialogFolderPicker).Show = 0 Then 'Error trap cancel button
        Me.btnCancel.SetFocus 'Have to switch focus to avoid dropdown getting stuck
        Settings.ResetFileDialog (msoFileDialogFolderPicker)
        Exit Sub
    End If
    
    'Populate the box with the current directory, set by the folder dialog
    c.Value = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
    Me.btnCancel.SetFocus 'Have to switch focus to avoid dropdown getting stuck
    Settings.ResetFileDialog (msoFileDialogFolderPicker)
    
    Exit Sub
    
Handler:
    MsgBox "Error " & Err.number & ": " & Err.Description
    
End Sub

Private Sub btnPaDSSetupCheck_Click()
    Unload Me
    Call Settings.ShowPaDSSetupCheck
End Sub

Private Sub btnNukeODC_Click()
    If MsgBox("Clearing the Office Document Cache requires shutting down Word. It is also recommended that you manually create an offline copy of your Sharepoint folder before clearing the ODC, just to be safe." & vbCrLf & vbCrLf & "You will see a pop-up that will prompt you to continue when Word is finished closing. Continue?", vbYesNo) = vbNo Then Exit Sub
    Call PaDS.NukeODC
End Sub

Private Sub btnNukeOneDrive_Click()
    If MsgBox("Before proceeding, you MUST ensure that you have stopped syncing any folders in OneDrive for Business. Failure to do so may result in deleting files from the server. All OneDrive for Business folder configuration will be lost." & vbCrLf & vbCrLf & "Are you sure?", vbYesNo) = vbNo Then Exit Sub
    If MsgBox("No, seriously. Are you 100% sure you have stopped syncing any folders in OneDrive for Business?" & vbCrLf & vbCrLf & "This will attempt to create a backup of your Sharepoint folder in your User folder, but you should manually back it up as well, just to be safe." & vbCrLf & vbCrLf & "This requires shutting down Word. You will see a pop-up that will prompt you to continue when Word is finished closing. Continue?", vbYesNo) = vbNo Then Exit Sub
    Call PaDS.NukeODC(True)
End Sub

Private Sub lblPaDSLink_Click()
    Settings.LaunchWebsite ("http://paperlessdebate.com/pads/")
End Sub

'*************************************************************************************
'* CASELIST TAB                                                                      *
'*************************************************************************************

Private Sub optOpenCaselist_Change()
    Me.cboCaselistSchoolName.Value = ""
    Me.cboCaselistSchoolName.Clear
    Me.cboCaselistTeamName.Value = ""
    Me.cboCaselistTeamName.Clear
End Sub

Private Sub optNDCAPolicy_Change()
    Me.cboCaselistSchoolName.Value = ""
    Me.cboCaselistSchoolName.Clear
    Me.cboCaselistTeamName.Value = ""
    Me.cboCaselistTeamName.Clear
End Sub

Private Sub optNDCALD_Change()
    Me.cboCaselistSchoolName.Value = ""
    Me.cboCaselistSchoolName.Clear
    Me.cboCaselistTeamName.Value = ""
    Me.cboCaselistTeamName.Clear
End Sub

Private Sub cboCaselistSchoolName_Change()
    Me.cboCaselistTeamName.Value = ""
    Me.cboCaselistTeamName.Clear
End Sub

Private Sub cboCaselistSchoolName_DropButtonClick()
'Populates the SchoolName combo box with schools from the caselist
        
    'If the list is already populated, exit
    If Me.cboCaselistSchoolName.ListCount > 0 Then Exit Sub
        
    'Clear ComboBoxes - clear TeamName too, so there's not a mismatch when changing
    Me.cboCaselistSchoolName.Value = ""
    Me.cboCaselistTeamName.Value = ""
    Me.cboCaselistSchoolName.Clear
    Me.cboCaselistTeamName.Clear
        
    'Populate box
    If Me.optOpenCaselist.Value = True Then Call GetCaselistSchoolNames("openCaselist", Me.cboCaselistSchoolName)
    If Me.optNDCAPolicy.Value = True Then Call GetCaselistSchoolNames("NDCAPolicy", Me.cboCaselistSchoolName)
    If Me.optNDCALD.Value = True Then Call GetCaselistSchoolNames("NDCALD", Me.cboCaselistSchoolName)
    
    Exit Sub

Handler:
    MsgBox "Error " & Err.number & ": " & Err.Description

End Sub

Private Sub cboCaselistTeamName_DropButtonClick()
'Populates the TeamName combo box with pages from the school's space

    'If the list is already populated, exit
    If Me.cboCaselistTeamName.ListCount > 0 Then Exit Sub
    
    'Check CaselistSchoolName has a value
    If Me.cboCaselistSchoolName.Value = "" Then
        Me.cboCaselistTeamName.Value = "Please choose a school first"
        Me.cboCaselistTeamName.Clear
        Exit Sub
    End If
    
    'Clear ComboBox
    Me.cboCaselistTeamName.Value = ""
    Me.cboCaselistTeamName.Clear
    
    'Turn on error checking
    On Error GoTo Handler
  
    If Me.optOpenCaselist.Value = True Then Call GetCaselistTeamNames("openCaselist", Me.cboCaselistSchoolName.Value, Me.cboCaselistTeamName)
    If Me.optNDCAPolicy.Value = True Then Call GetCaselistTeamNames("NDCAPolicy", Me.cboCaselistSchoolName.Value, Me.cboCaselistTeamName)
    If Me.optNDCALD.Value = True Then Call GetCaselistTeamNames("NDCALD", Me.cboCaselistSchoolName.Value, Me.cboCaselistTeamName)
    
    Exit Sub

Handler:
    MsgBox "Error " & Err.number & ": " & Err.Description

End Sub

'*************************************************************************************
'* EMAIL TAB                                                                         *
'*************************************************************************************

Private Sub optGmail_Change()
    If Me.optGmail.Value = True Then
        Me.lblGmailUsername.Enabled = True
        Me.txtGmailUsername.Enabled = True
        Me.lblGmailPassword.Enabled = True
        Me.txtGmailPassword.Enabled = True

        Me.lblEmailUsername.Enabled = False
        Me.txtEmailUsername.Enabled = False
        Me.lblEmailPassword.Enabled = False
        Me.txtEmailPassword.Enabled = False
        Me.lblSMTPServer.Enabled = False
        Me.txtSMTPServer.Enabled = False
        Me.lblSMTPPort.Enabled = False
        Me.txtSMTPPort.Enabled = False
        Me.chkUseSSL.Enabled = False
    End If
End Sub

Private Sub optManualEmail_Change()
    If Me.optManualEmail.Value = True Then
        Me.lblGmailUsername.Enabled = False
        Me.txtGmailUsername.Enabled = False
        Me.lblGmailPassword.Enabled = False
        Me.txtGmailPassword.Enabled = False

        Me.lblEmailUsername.Enabled = True
        Me.txtEmailUsername.Enabled = True
        Me.lblEmailPassword.Enabled = True
        Me.txtEmailPassword.Enabled = True
        Me.lblSMTPServer.Enabled = True
        Me.txtSMTPServer.Enabled = True
        Me.lblSMTPPort.Enabled = True
        Me.txtSMTPPort.Enabled = True
        Me.chkUseSSL.Enabled = True
    End If
End Sub

'*************************************************************************************
'* ABOUT TAB                                                                         *
'*************************************************************************************

Private Sub lblAbout5_Click()
    Settings.LaunchWebsite ("https://paperlessdebate.com/")
End Sub

Private Sub btnVerbatimHelp_Click()
    UI.ShowForm "Help"
End Sub
Private Sub btnEasterEgg_Click()
    MsgBox "You have ascended to the Ashtar Command!"
End Sub

Sub btnSave_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnSave.BackColor = Globals.BLUE_BUTTON_HOVER
End Sub
Sub btnResetAllSettings_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnResetAllSettings.BackColor = Globals.GREEN_BUTTON_HOVER
End Sub
Sub btnCancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnCancel.BackColor = Globals.RED_BUTTON_HOVER
End Sub
Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnSave.BackColor = Globals.BLUE_BUTTON_NORMAL
    btnResetAllSettings.BackColor = Globals.GREEN_BUTTON_NORMAL
    btnCancel.BackColor = Globals.RED_BUTTON_NORMAL
End Sub
