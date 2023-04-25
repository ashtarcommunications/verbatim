VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   Caption         =   "Verbatim Settings"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9705
   OleObjectBlob   =   "frmSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*************************************************************************************
'* FORM UI                                                                           *
'*************************************************************************************
Private Sub SetPage(ByVal MenuTab As String)
    ' Reset all tabs
    Me.lblTabProfile.BackColor = Globals.WHITE
    Me.lblTabProfile.ForeColor = Globals.BLACK
    Me.lblTabAdmin.BackColor = Globals.WHITE
    Me.lblTabAdmin.ForeColor = Globals.BLACK
    Me.lblTabView.BackColor = Globals.WHITE
    Me.lblTabView.ForeColor = Globals.BLACK
    Me.lblTabPaperless.BackColor = Globals.WHITE
    Me.lblTabPaperless.ForeColor = Globals.BLACK
    Me.lblTabStyles.BackColor = Globals.WHITE
    Me.lblTabStyles.ForeColor = Globals.BLACK
    Me.lblTabFormat.BackColor = Globals.WHITE
    Me.lblTabFormat.ForeColor = Globals.BLACK
    Me.lblTabKeyboard.BackColor = Globals.WHITE
    Me.lblTabKeyboard.ForeColor = Globals.BLACK
    Me.lblTabVTub.BackColor = Globals.WHITE
    Me.lblTabVTub.ForeColor = Globals.BLACK
    Me.lblTabCaselist.BackColor = Globals.WHITE
    Me.lblTabCaselist.ForeColor = Globals.BLACK
    Me.lblTabPlugins.BackColor = Globals.WHITE
    Me.lblTabPlugins.ForeColor = Globals.BLACK
    Me.lblTabAbout.BackColor = Globals.WHITE
    Me.lblTabAbout.ForeColor = Globals.BLACK
    
    Select Case MenuTab
        Case "Profile"
            Me.mpgSettings.Value = 0
            Me.lblTabProfile.BackColor = Globals.BLUE
            Me.lblTabProfile.ForeColor = Globals.WHITE
        Case "Admin"
            Me.mpgSettings.Value = 1
            Me.lblTabAdmin.BackColor = Globals.BLUE
            Me.lblTabAdmin.ForeColor = Globals.WHITE
        Case "View"
            Me.mpgSettings.Value = 2
            Me.lblTabView.BackColor = Globals.BLUE
            Me.lblTabView.ForeColor = Globals.WHITE
        Case "Paperless"
            Me.mpgSettings.Value = 3
            Me.lblTabPaperless.BackColor = Globals.BLUE
            Me.lblTabPaperless.ForeColor = Globals.WHITE
        Case "Styles"
            Me.mpgSettings.Value = 4
            Me.lblTabStyles.BackColor = Globals.BLUE
            Me.lblTabStyles.ForeColor = Globals.WHITE
        Case "Format"
            Me.mpgSettings.Value = 5
            Me.lblTabFormat.BackColor = Globals.BLUE
            Me.lblTabFormat.ForeColor = Globals.WHITE
        Case "Keyboard"
            Me.mpgSettings.Value = 6
            Me.lblTabKeyboard.BackColor = Globals.BLUE
            Me.lblTabKeyboard.ForeColor = Globals.WHITE
        Case "VTub"
            Me.mpgSettings.Value = 7
            Me.lblTabVTub.BackColor = Globals.BLUE
            Me.lblTabVTub.ForeColor = Globals.WHITE
        Case "Caselist"
            Me.mpgSettings.Value = 8
            Me.lblTabCaselist.BackColor = Globals.BLUE
            Me.lblTabCaselist.ForeColor = Globals.WHITE
        Case "Plugins"
            Me.mpgSettings.Value = 9
            Me.lblTabPlugins.BackColor = Globals.BLUE
            Me.lblTabPlugins.ForeColor = Globals.WHITE
        Case "About"
            Me.mpgSettings.Value = 10
            Me.lblTabAbout.BackColor = Globals.BLUE
            Me.lblTabAbout.ForeColor = Globals.WHITE
        Case Else
            Me.mpgSettings.Value = 0
            Me.lblTabProfile.BackColor = Globals.BLUE
            Me.lblTabProfile.ForeColor = Globals.WHITE
    End Select
End Sub

Private Sub lblTabProfile_Click()
    SetPage "Profile"
End Sub

Private Sub lblTabAdmin_Click()
    SetPage "Admin"
End Sub

Private Sub lblTabView_Click()
    SetPage "View"
End Sub

Private Sub lblTabPaperless_Click()
    SetPage "Paperless"
End Sub

Private Sub lblTabStyles_Click()
    SetPage "Styles"
End Sub

Private Sub lblTabFormat_Click()
    SetPage "Format"
End Sub

Private Sub lblTabKeyboard_Click()
    SetPage "Keyboard"
End Sub

Private Sub lblTabVTub_Click()
    SetPage "VTub"
End Sub

Private Sub lblTabCaselist_Click()
    SetPage "Caselist"
End Sub

Private Sub lblTabPlugins_Click()
    SetPage "Plugins"
End Sub

Private Sub lblTabAbout_Click()
    SetPage "About"
End Sub

'*************************************************************************************
'* GENERAL FUNCTIONS                                                                 *
'*************************************************************************************
Private Sub UserForm_Initialize()
    Dim e As String
    Dim FontSize As Long
    Dim f As Variant
    Dim MacroArray() As Variant
    
    ' Turn on Error handling
    On Error GoTo Handler
    
    Globals.InitializeGlobals
    
    #If Mac Then
        UI.ResizeUserForm Me
        Me.btnCancel.ForeColor = Globals.RED
        Me.btnResetAllSettings.ForeColor = Globals.ORANGE
        Me.btnSave.ForeColor = Globals.BLUE
        Me.btnTabroomLogout.ForeColor = Globals.RED
        Me.btnTabroomLogin.ForeColor = Globals.BLUE
        Me.btnCreateVTub.ForeColor = Globals.BLUE
        Me.btnUpdateCheck.ForeColor = Globals.BLUE
    #End If
    
    ' Get Settings from the registry to populate the settings boxes
    
    ' Profile Tab
    Me.txtName.Value = GetSetting("Verbatim", "Profile", "Name", "")
    Me.txtSchoolName.Value = GetSetting("Verbatim", "Profile", "SchoolName", "")
    
    If GetSetting("Verbatim", "Profile", "CollegeHS", "K12") = "College" Then
        Me.optCollege.Value = True
    Else
        Me.optK12.Value = True
    End If
    
    e = GetSetting("Verbatim", "Profile", "Event", "CX")
    If e = "LD" Then
        Me.optLD.Value = True
    ElseIf e = "PF" Then
        Me.optPF.Value = True
    Else
        Me.optCX.Value = True
    End If
    
    Me.txtWPM.Value = GetSetting("Verbatim", "Profile", "WPM", 250)
    
    Me.chkDisableTabroom.Value = GetSetting("Verbatim", "Profile", "DisableTabroom", False)
    
    ' Admin Tab
    Me.chkAlwaysOn.Value = GetSetting("Verbatim", "Admin", "AlwaysOn", True)
    Me.chkAutoUpdateStyles.Value = GetSetting("Verbatim", "Admin", "AutoUpdateStyles", True)
    Me.chkSuppressStyleChecks.Value = GetSetting("Verbatim", "Admin", "SuppressStyleChecks", True)
    Me.chkSuppressInstallChecks.Value = GetSetting("Verbatim", "Admin", "SuppressInstallChecks", False)
    Me.chkSuppressDocCheck.Value = GetSetting("Verbatim", "Admin", "SuppressDocCheck", False)
    Me.chkFirstRun.Value = GetSetting("Verbatim", "Admin", "FirstRun", False)
    
    ' View Tab
    If GetSetting("Verbatim", "View", "DefaultView", Globals.DefaultView) = "Web" Then
        Me.optWebView.Value = True
    Else
        Me.optDraftView.Value = True
    End If
    
    Me.chkNPCStartup.Value = GetSetting("Verbatim", "View", "NPCStartup", False)
    
    Me.spnDocs.Value = GetSetting("Verbatim", "View", "DocsPct", 50)
    Me.spnSpeech.Value = GetSetting("Verbatim", "View", "SpeechPct", 50)
    
    Me.chkRibbonDisableSpeech.Value = GetSetting("Verbatim", "View", "RibbonDisableSpeech", False)
    Me.chkRibbonDisableOrganize.Value = GetSetting("Verbatim", "View", "RibbonDisableOrganize", False)
    Me.chkRibbonDisableFormat.Value = GetSetting("Verbatim", "View", "RibbonDisableFormat", False)
    Me.chkRibbonDisablePaperless.Value = GetSetting("Verbatim", "View", "RibbonDisablePaperless", False)
    Me.chkRibbonDisableTools.Value = GetSetting("Verbatim", "View", "RibbonDisableTools", False)
    Me.chkRibbonDisableView.Value = GetSetting("Verbatim", "View", "RibbonDisableView", False)
    Me.chkRibbonDisableCaselist.Value = GetSetting("Verbatim", "View", "RibbonDisableCaselist", False)
    Me.chkRibbonDisableSettings.Value = GetSetting("Verbatim", "View", "RibbonDisableSettings", False)
        
    ' Paperless Tab
    Me.chkAutoSaveSpeech.Value = GetSetting("Verbatim", "Paperless", "AutoSaveSpeech", False)
    Me.cboAutoSaveDir.Value = GetSetting("Verbatim", "Paperless", "AutoSaveDir", "")
    Me.chkStripSpeech.Value = GetSetting("Verbatim", "Paperless", "StripSpeech", True)
    Me.cboSearchDir.Value = GetSetting("Verbatim", "Paperless", "SearchDir", "")
    Me.cboAutoOpenDir.Value = GetSetting("Verbatim", "Paperless", "AutoOpenDir", "")
    Me.cboAudioDir.Value = GetSetting("Verbatim", "Paperless", "AudioDir", "")
      
    ' Populate Styles Tab Comboboxes - Allow 8pt-32pt
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
    
    ' Populate Styles Tab Normal Font Combobox
    For Each f In Application.FontNames
        Me.cboNormalFont.AddItem f
    Next f
    
    ' Populate Styles Tab Emphasis box size combobox
    Me.cboEmphasisBoxSize.AddItem "1pt"
    Me.cboEmphasisBoxSize.AddItem "1.5pt"
    Me.cboEmphasisBoxSize.AddItem "2.25pt"
    Me.cboEmphasisBoxSize.AddItem "3pt"
    
    ' Styles Tab
    Me.cboNormalSize.Value = GetSetting("Verbatim", "Styles", "NormalSize", 11)
    Me.cboNormalFont.Value = GetSetting("Verbatim", "Styles", "NormalFont", "Calibri")
    
    Me.cboPocketSize.Value = GetSetting("Verbatim", "Styles", "PocketSize", 26)
    Me.cboHatSize.Value = GetSetting("Verbatim", "Styles", "HatSize", 22)
    Me.cboBlockSize.Value = GetSetting("Verbatim", "Styles", "BlockSize", 16)
    Me.cboTagSize.Value = GetSetting("Verbatim", "Styles", "TagSize", 13)
    
    Me.cboCiteSize.Value = GetSetting("Verbatim", "Styles", "CiteSize", 13)
    Me.chkUnderlineCite.Value = GetSetting("Verbatim", "Styles", "UnderlineCite", False)
    
    Me.cboUnderlineSize.Value = GetSetting("Verbatim", "Styles", "UnderlineSize", 11)
    Me.chkBoldUnderline.Value = GetSetting("Verbatim", "Styles", "BoldUnderline", False)
    
    Me.cboEmphasisSize.Value = GetSetting("Verbatim", "Styles", "EmphasisSize", 11)
    Me.chkEmphasisBold.Value = GetSetting("Verbatim", "Styles", "EmphasisBold", True)
    Me.chkEmphasisItalic.Value = GetSetting("Verbatim", "Styles", "EmphasisItalic", False)
    Me.chkEmphasisBox.Value = GetSetting("Verbatim", "Styles", "EmphasisBox", False)
    Me.cboEmphasisBoxSize.Value = GetSetting("Verbatim", "Styles", "EmphasisBoxSize", "1pt")
    
    ' Format Tab
    If GetSetting("Verbatim", "Format", "Spacing", "Wide") = "Wide" Then
        Me.optSpacingWide.Value = True
    Else
        Me.optSpacingNarrow.Value = True
    End If
        
    Me.chkShrinkOmissions.Value = GetSetting("Verbatim", "Format", "ShrinkOmissions", False)
    
    Me.chkParagraphIntegrity.Value = GetSetting("Verbatim", "Format", "ParagraphIntegrity", False)
    Me.chkUsePilcrows.Value = GetSetting("Verbatim", "Format", "UsePilcrows", False)
    Me.chkCondenseOnPaste.Value = GetSetting("Verbatim", "Format", "CondenseOnPaste", False)
    
    Me.chkAutoUnderlineEmphasis.Value = GetSetting("Verbatim", "Format", "AutoUnderlineEmphasis", False)
            
    ' Populate highlighting exception dropdown
    Me.cboHighlightingException.AddItem "None"
    Me.cboHighlightingException.AddItem "Black"
    Me.cboHighlightingException.AddItem "Blue"
    Me.cboHighlightingException.AddItem "Bright Green"
    Me.cboHighlightingException.AddItem "Dark Blue"
    Me.cboHighlightingException.AddItem "Dark Red"
    Me.cboHighlightingException.AddItem "Dark Yellow"
    Me.cboHighlightingException.AddItem "Light Gray"
    Me.cboHighlightingException.AddItem "Dark Gray"
    Me.cboHighlightingException.AddItem "Green"
    Me.cboHighlightingException.AddItem "Pink"
    Me.cboHighlightingException.AddItem "Red"
    Me.cboHighlightingException.AddItem "Teal"
    Me.cboHighlightingException.AddItem "Turquoise"
    Me.cboHighlightingException.AddItem "Violet"
    Me.cboHighlightingException.AddItem "White"
    Me.cboHighlightingException.AddItem "Yellow"
    
    Me.cboHighlightingException.Value = GetSetting("Verbatim", "Format", "HighlightingException", "None")
    
    ' Populate Keyboard Tab Comboboxes
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
    
    ' Keyboard Tab
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
        
    ' VTub Tab
    Me.cboVTubPath.Value = GetSetting("Verbatim", "VTub", "VTubPath", "")
    Me.chkVTubRefreshPrompt.Value = GetSetting("Verbatim", "VTub", "VTubRefreshPrompt", True)
       
    ' Caselist Tab
    Me.chkOpenSource.Value = GetSetting("Verbatim", "Caselist", "OpenSource", True)
    Me.chkProcessCites.Value = GetSetting("Verbatim", "Caselist", "ProcessCites", True)
        
    ' Plugins Tab
    Me.cboTimer.Value = GetSetting("Verbatim", "Plugins", "TimerPath", "")
    Me.cboOCR.Value = GetSetting("Verbatim", "Plugins", "OCRPath", "")
    Me.cboSearch.Value = GetSetting("Verbatim", "Plugins", "SearchPath", "")
    
    ' About Tab
    Me.lblVersion.Caption = "Verbatim v. " & Settings.GetVersion
    Me.chkAutomaticUpdates.Value = GetSetting("Verbatim", "Profile", "AutomaticUpdates", True)
    Me.lblLastUpdateCheck.Caption = "Last Update Check:" & vbCrLf & _
        Format$(GetSetting("Verbatim", "Profile", "LastUpdateCheck", ""), "mm-dd-yy hh:mm")
    Exit Sub
        
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub UserForm_Activate()
    ' Ensure correct button colors don't get lost
    Globals.InitializeGlobals
    
    ' Set Tabroom logged in state
    If GetSetting("Verbatim", "Caselist", "CaselistToken", "") <> "" And Caselist.CheckCaselistToken = True Then
        Me.lblTabroomLoggedIn.Caption = "You are logged in to Tabroom"
        Me.btnTabroomLogout.Visible = True
        Me.btnTabroomLogin.Visible = False
    Else
        Me.lblTabroomLoggedIn.Caption = "You are logged out of Tabroom"
        Me.btnTabroomLogout.Visible = False
        Me.btnTabroomLogin.Visible = True
    End If
End Sub

Private Sub btnResetAllSettings_Click()
'Resets all settings to the default
    On Error GoTo Handler
    
    If MsgBox("This will reset all settings to their default values - changes will not be committed until you click Save. Are you sure?", vbOKCancel) = vbCancel Then Exit Sub
    
    ' Profile Tab
    Me.txtName.Value = ""
    Me.txtSchoolName.Value = ""
    Me.optK12.Value = True
    Me.optCX.Value = True
    Me.txtWPM.Value = 250
    Me.chkDisableTabroom.Value = False
        
    ' Admin Tab
    Me.chkAlwaysOn.Value = True
    Me.chkAutoUpdateStyles.Value = True
    Me.chkSuppressStyleChecks.Value = False
    Me.chkSuppressInstallChecks.Value = False
    Me.chkSuppressDocCheck.Value = False
    Me.chkFirstRun.Value = False
    
    ' View Tab
    If Globals.DefaultView = "Web" Then
        Me.optWebView.Value = True
    Else
        Me.optDraftView.Value = True
    End If
    Me.chkNPCStartup.Value = False
    Me.spnDocs.Value = 50
    Me.spnSpeech.Value = 50
    
    Me.chkRibbonDisableSpeech.Value = False
    Me.chkRibbonDisableOrganize.Value = False
    Me.chkRibbonDisableFormat.Value = False
    Me.chkRibbonDisablePaperless.Value = False
    Me.chkRibbonDisableTools.Value = False
    Me.chkRibbonDisableView.Value = False
    Me.chkRibbonDisableCaselist.Value = False
    Me.chkRibbonDisableSettings.Value = False
    
    ' Paperless Tab
    Me.chkAutoSaveSpeech.Value = False
    Me.cboAutoSaveDir.Value = ""
    Me.chkStripSpeech.Value = True
    Me.cboSearchDir.Value = ""
    Me.cboAutoOpenDir.Value = ""
    Me.cboAudioDir.Value = ""
    
    ' Styles Tab
    Me.cboNormalSize.Value = 11
    Me.cboNormalFont.Value = "Calibri"
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
    
    ' Format Tab
    Me.optSpacingWide.Value = True
    Me.chkShrinkOmissions.Value = False
    
    Me.chkParagraphIntegrity.Value = False
    Me.chkUsePilcrows.Value = False
    Me.chkCondenseOnPaste.Value = False
    
    Me.chkAutoUnderlineEmphasis.Value = False
    
    Me.cboHighlightingException.Value = "None"
    
    ' Keyboard Tab
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
    
    ' VTub Tab
    Me.cboVTubPath.Value = ""
    Me.chkVTubRefreshPrompt.Value = True
    
    ' Caselist Tab
    Me.chkOpenSource.Value = True
    Me.chkProcessCites.Value = True
       
    ' Plugins Tab
    Me.cboTimer.Value = ""
    Me.cboOCR.Value = ""
    Me.cboSearch.Value = ""
    
    ' About Tab
    Me.chkAutomaticUpdates.Value = True
    
    Exit Sub
        
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Private Sub btnSave_Click()
    Dim DebateTemplate As Document
    Dim CloseDebateTemplate As Boolean
    
    On Error GoTo Handler
    
    ' Profile Tab
    SaveSetting "Verbatim", "Profile", "Name", Me.txtName.Value
    SaveSetting "Verbatim", "Profile", "SchoolName", Me.txtSchoolName.Value
    
    If Me.optCollege.Value = True Then
        SaveSetting "Verbatim", "Profile", "CollegeHS", "College"
    Else
        SaveSetting "Verbatim", "Profile", "CollegeHS", "K12"
    End If
    
    If Me.optLD.Value = True Then
        SaveSetting "Verbatim", "Profile", "Event", "LD"
    ElseIf Me.optPF.Value = True Then
        SaveSetting "Verbatim", "Profile", "Event", "PF"
    Else
        SaveSetting "Verbatim", "Profile", "Event", "CX"
    End If
        
    SaveSetting "Verbatim", "Profile", "WPM", Me.txtWPM.Value
    
    SaveSetting "Verbatim", "Profile", "DisableTabroom", Me.chkDisableTabroom.Value
    
    ' Admin Tab
    SaveSetting "Verbatim", "Admin", "AlwaysOn", Me.chkAlwaysOn.Value
    SaveSetting "Verbatim", "Admin", "AutoUpdateStyles", Me.chkAutoUpdateStyles.Value
    SaveSetting "Verbatim", "Admin", "SuppressStyleChecks", Me.chkSuppressStyleChecks.Value
    SaveSetting "Verbatim", "Admin", "SuppressInstallChecks", Me.chkSuppressInstallChecks.Value
    SaveSetting "Verbatim", "Admin", "SuppressDocCheck", Me.chkSuppressDocCheck.Value
    SaveSetting "Verbatim", "Admin", "FirstRun", Me.chkFirstRun.Value
    
    ' View Tab
    If Me.optWebView.Value = True Then
        SaveSetting "Verbatim", "View", "DefaultView", "Web"
    Else
        SaveSetting "Verbatim", "View", "DefaultView", "Draft"
    End If

    SaveSetting "Verbatim", "View", "NPCStartup", Me.chkNPCStartup.Value
    SaveSetting "Verbatim", "View", "DocsPct", Me.spnDocs.Value
    SaveSetting "Verbatim", "View", "SpeechPct", Me.spnSpeech.Value
    
    SaveSetting "Verbatim", "View", "RibbonDisableSpeech", Me.chkRibbonDisableSpeech.Value
    SaveSetting "Verbatim", "View", "RibbonDisableOrganize", Me.chkRibbonDisableOrganize.Value
    SaveSetting "Verbatim", "View", "RibbonDisableFormat", Me.chkRibbonDisableFormat.Value
    SaveSetting "Verbatim", "View", "RibbonDisablePaperless", Me.chkRibbonDisablePaperless.Value
    SaveSetting "Verbatim", "View", "RibbonDisableTools", Me.chkRibbonDisableTools.Value
    SaveSetting "Verbatim", "View", "RibbonDisableView", Me.chkRibbonDisableView.Value
    SaveSetting "Verbatim", "View", "RibbonDisableCaselist", Me.chkRibbonDisableCaselist.Value
    SaveSetting "Verbatim", "View", "RibbonDisableSettings", Me.chkRibbonDisableSettings.Value
    
    ' Paperless Tab
    SaveSetting "Verbatim", "Paperless", "AutoSaveSpeech", Me.chkAutoSaveSpeech.Value
    SaveSetting "Verbatim", "Paperless", "AutoSaveDir", Me.cboAutoSaveDir.Value
    SaveSetting "Verbatim", "Paperless", "StripSpeech", Me.chkStripSpeech.Value
    SaveSetting "Verbatim", "Paperless", "SearchDir", Me.cboSearchDir.Value
    SaveSetting "Verbatim", "Paperless", "AutoOpenDir", Me.cboAutoOpenDir.Value
    SaveSetting "Verbatim", "Paperless", "AudioDir", Me.cboAudioDir.Value
    
    ' Styles Tab
    SaveSetting "Verbatim", "Styles", "NormalSize", Me.cboNormalSize.Value
    SaveSetting "Verbatim", "Styles", "NormalFont", Me.cboNormalFont.Value
    SaveSetting "Verbatim", "Styles", "PocketSize", Me.cboPocketSize.Value
    SaveSetting "Verbatim", "Styles", "HatSize", Me.cboHatSize.Value
    SaveSetting "Verbatim", "Styles", "BlockSize", Me.cboBlockSize.Value
    SaveSetting "Verbatim", "Styles", "TagSize", Me.cboTagSize.Value
    SaveSetting "Verbatim", "Styles", "CiteSize", Me.cboCiteSize.Value
    SaveSetting "Verbatim", "Styles", "UnderlineCite", Me.chkUnderlineCite.Value
    SaveSetting "Verbatim", "Styles", "UnderlineSize", Me.cboUnderlineSize.Value
    SaveSetting "Verbatim", "Styles", "BoldUnderline", Me.chkBoldUnderline.Value
    SaveSetting "Verbatim", "Styles", "EmphasisSize", Me.cboEmphasisSize.Value
    SaveSetting "Verbatim", "Styles", "EmphasisBold", Me.chkEmphasisBold.Value
    SaveSetting "Verbatim", "Styles", "EmphasisItalic", Me.chkEmphasisItalic.Value
    SaveSetting "Verbatim", "Styles", "EmphasisBox", Me.chkEmphasisBox.Value
    SaveSetting "Verbatim", "Styles", "EmphasisBoxSize", Me.cboEmphasisBoxSize.Value
    
    ' Format Tab
    If Me.optSpacingWide.Value = True Then
        SaveSetting "Verbatim", "Format", "Spacing", "Wide"
    Else
        SaveSetting "Verbatim", "Format", "Spacing", "Narrow"
    End If
   
    SaveSetting "Verbatim", "Format", "ShrinkOmissions", Me.chkShrinkOmissions.Value
    
    SaveSetting "Verbatim", "Format", "ParagraphIntegrity", Me.chkParagraphIntegrity.Value
    SaveSetting "Verbatim", "Format", "UsePilcrows", Me.chkUsePilcrows.Value
    SaveSetting "Verbatim", "Format", "CondenseOnPaste", Me.chkCondenseOnPaste.Value
    
    SaveSetting "Verbatim", "Format", "AutoUnderlineEmphasis", Me.chkAutoUnderlineEmphasis.Value
    
    SaveSetting "Verbatim", "Format", "HighlightingException", Me.cboHighlightingException.Value
    
    ' Check if Template itself is open, or open it as a Document
    On Error Resume Next
    If ActiveDocument.FullName = ActiveDocument.AttachedTemplate.FullName Then
        Set DebateTemplate = ActiveDocument
        CloseDebateTemplate = False
    Else
        Set DebateTemplate = ActiveDocument.AttachedTemplate.OpenAsDocument
        CloseDebateTemplate = True
    End If
    On Error GoTo Handler
    
    ' Update template styles based on Styles settings
    DebateTemplate.Styles.Item("Normal").Font.size = Me.cboNormalSize.Value
    DebateTemplate.Styles.Item("Normal").Font.Name = Me.cboNormalFont.Value
    DebateTemplate.Styles.Item("Pocket").Font.size = Me.cboPocketSize.Value
    DebateTemplate.Styles.Item("Hat").Font.size = Me.cboHatSize.Value
    DebateTemplate.Styles.Item("Block").Font.size = Me.cboBlockSize.Value
    DebateTemplate.Styles.Item("Tag").Font.size = Me.cboTagSize.Value
    DebateTemplate.Styles.Item("Cite").Font.size = Me.cboCiteSize.Value
    If Me.chkUnderlineCite.Value = True Then
        DebateTemplate.Styles.Item("Cite").Font.Underline = wdUnderlineSingle
    Else
        DebateTemplate.Styles.Item("Cite").Font.Underline = wdUnderlineNone
    End If
    DebateTemplate.Styles.Item("Underline").Font.size = Me.cboUnderlineSize.Value
    DebateTemplate.Styles.Item("Underline").Font.Bold = Me.chkBoldUnderline.Value
    DebateTemplate.Styles.Item("Emphasis").Font.size = Me.cboEmphasisSize.Value
    DebateTemplate.Styles.Item("Emphasis").Font.Name = Me.cboNormalFont.Value
    DebateTemplate.Styles.Item("Emphasis").Font.Bold = Me.chkEmphasisBold.Value
    DebateTemplate.Styles.Item("Emphasis").Font.Italic = Me.chkEmphasisItalic.Value
    
    If Me.chkEmphasisBox.Value = True Then
        DebateTemplate.Styles.Item("Emphasis").Font.Borders.Item(1).LineStyle = wdLineStyleSingle
    
        Select Case Me.cboEmphasisBoxSize.Value
            Case Is = "1pt"
                DebateTemplate.Styles.Item("Emphasis").Font.Borders.Item(1).LineWidth = wdLineWidth100pt
            Case Is = "1.5pt"
                DebateTemplate.Styles.Item("Emphasis").Font.Borders.Item(1).LineWidth = wdLineWidth150pt
            Case Is = "2.25pt"
                DebateTemplate.Styles.Item("Emphasis").Font.Borders.Item(1).LineWidth = wdLineWidth225pt
            Case Is = "3pt"
                DebateTemplate.Styles.Item("Emphasis").Font.Borders.Item(1).LineWidth = wdLineWidth300pt
            Case Else
                DebateTemplate.Styles.Item("Emphasis").Font.Borders.Item(1).LineWidth = wdLineWidth100pt
        End Select
    Else
        DebateTemplate.Styles.Item("Emphasis").Font.Borders.Item(1).LineStyle = wdLineStyleNone
    End If
    
    If Me.optSpacingWide.Value = True Then
        DebateTemplate.Styles.Item("Normal").ParagraphFormat.SpaceBefore = 0
        DebateTemplate.Styles.Item("Normal").ParagraphFormat.SpaceAfter = 8
        DebateTemplate.Styles.Item("Normal").ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple
        DebateTemplate.Styles.Item("Normal").ParagraphFormat.LineSpacing = LinesToPoints(1.08)
        DebateTemplate.Styles.Item("Pocket").ParagraphFormat.SpaceBefore = 12
        DebateTemplate.Styles.Item("Pocket").ParagraphFormat.SpaceAfter = 0
        DebateTemplate.Styles.Item("Hat").ParagraphFormat.SpaceBefore = 2
        DebateTemplate.Styles.Item("Hat").ParagraphFormat.SpaceAfter = 0
        DebateTemplate.Styles.Item("Block").ParagraphFormat.SpaceBefore = 2
        DebateTemplate.Styles.Item("Block").ParagraphFormat.SpaceAfter = 0
        DebateTemplate.Styles.Item("Tag").ParagraphFormat.SpaceBefore = 2
        DebateTemplate.Styles.Item("Tag").ParagraphFormat.SpaceAfter = 0
    Else
        DebateTemplate.Styles.Item("Normal").ParagraphFormat.SpaceBefore = 0
        DebateTemplate.Styles.Item("Normal").ParagraphFormat.SpaceAfter = 0
        DebateTemplate.Styles.Item("Normal").ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
        DebateTemplate.Styles.Item("Pocket").ParagraphFormat.SpaceBefore = 24
        DebateTemplate.Styles.Item("Pocket").ParagraphFormat.SpaceAfter = 0
        DebateTemplate.Styles.Item("Hat").ParagraphFormat.SpaceBefore = 24
        DebateTemplate.Styles.Item("Hat").ParagraphFormat.SpaceAfter = 0
        DebateTemplate.Styles.Item("Block").ParagraphFormat.SpaceBefore = 10
        DebateTemplate.Styles.Item("Block").ParagraphFormat.SpaceAfter = 0
        DebateTemplate.Styles.Item("Tag").ParagraphFormat.SpaceBefore = 10
        DebateTemplate.Styles.Item("Tag").ParagraphFormat.SpaceAfter = 0
    End If
    
    ' Keyboard Tab
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
    
    ' Update template keyboard shortcuts based on keyboard settings
    Settings.ChangeKeyboardShortcut wdKeyF2, Me.cboF2Shortcut.Value
    Settings.ChangeKeyboardShortcut wdKeyF3, Me.cboF3Shortcut.Value
    Settings.ChangeKeyboardShortcut wdKeyF4, Me.cboF4Shortcut.Value
    Settings.ChangeKeyboardShortcut wdKeyF5, Me.cboF5Shortcut.Value
    Settings.ChangeKeyboardShortcut wdKeyF6, Me.cboF6Shortcut.Value
    Settings.ChangeKeyboardShortcut wdKeyF7, Me.cboF7Shortcut.Value
    Settings.ChangeKeyboardShortcut wdKeyF8, Me.cboF8Shortcut.Value
    Settings.ChangeKeyboardShortcut wdKeyF9, Me.cboF9Shortcut.Value
    Settings.ChangeKeyboardShortcut wdKeyF10, Me.cboF10Shortcut.Value
    Settings.ChangeKeyboardShortcut wdKeyF11, Me.cboF11Shortcut.Value
    Settings.ChangeKeyboardShortcut wdKeyF12, Me.cboF12Shortcut.Value
    
    ' Close template if opened separately
    If CloseDebateTemplate = True Then
        '@Ignore MemberNotOnInterface
        DebateTemplate.Close SaveChanges:=wdSaveChanges
    End If
    
    On Error Resume Next
    ActiveDocument.UpdateStyles
    On Error GoTo Handler
    
    ' VTub Tab
    SaveSetting "Verbatim", "VTub", "VTubPath", Me.cboVTubPath.Value
    SaveSetting "Verbatim", "VTub", "VTubRefreshPrompt", chkVTubRefreshPrompt.Value
      
    ' Caselist Tab
    SaveSetting "Verbatim", "Caselist", "OpenSource", Me.chkOpenSource.Value
    SaveSetting "Verbatim", "Caselist", "ProcessCites", Me.chkProcessCites.Value
        
    ' Plugins Tab
    SaveSetting "Verbatim", "Plugins", "TimerPath", Me.cboTimer.Value
    SaveSetting "Verbatim", "Plugins", "OCRPath", Me.cboOCR.Value
    SaveSetting "Verbatim", "Plugins", "SearchPath", Me.cboSearch.Value
    
    ' About Tab
    SaveSetting "Verbatim", "Profile", "Version", Settings.GetVersion
    SaveSetting "Verbatim", "Profile", "AutomaticUpdates", Me.chkAutomaticUpdates.Value
    
    ' Refresh ribbon in case keyboard shortcuts changed
    Ribbon.RefreshRibbon
    
    ' Unload the form
    Unload Me
    Exit Sub
        
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub btnCancel_Click()
  Unload Me
End Sub

'*************************************************************************************
'* PROFILE TAB                                                                       *
'*************************************************************************************
Private Sub lblWPMLink_Click()
    Settings.LaunchWebsite (Globals.WPM_URL)
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
    Me.btnTabroomLogout.Visible = False
    Me.btnTabroomLogin.Visible = True
End Sub

Private Sub lblTabroomRegister_Click()
    Settings.LaunchWebsite (Globals.TABROOM_REGISTER_URL)
End Sub

'*************************************************************************************
'* ADMIN TAB                                                                         *
'*************************************************************************************

Private Sub btnSetupWizard_Click()
    Unload Me
    UI.ShowForm "Setup"
End Sub

Private Sub btnTroubleshooter_Click()
    Unload Me
    UI.ShowForm "Troubleshooter"
End Sub

Private Sub btnTemplatesFolder_Click()
    Settings.OpenTemplatesFolder
End Sub

Private Sub btnTutorial_Click()
    Unload Me
    UI.LaunchTutorial
End Sub

Private Sub btnUnverbatimizeNormal_Click()
    Settings.UnverbatimizeNormal Notify:=True
End Sub

Private Sub btnImportSettings_Click()

    Dim SettingsFileName As String

    On Error GoTo Handler
    
    #If Mac Then
        MsgBox "Importing settings isn't available for Mac. See the online manual for more information on workarounds.", vbOKOnly
        Exit Sub
    #End If

    SettingsFileName = UI.GetFileFromDialog("Verbatim Settings", "*.ini", "Select Verbatim Settings file to import...", "Import")

    ' Exit if trying to import an old settings file
    If System.PrivateProfileString(SettingsFileName, "Profile", "Version") = "" Or _
        CLng(Left$(System.PrivateProfileString(SettingsFileName, "Profile", "Version"), 1)) < 6 Then
        MsgBox "Outdated settings file. You must use a Verbatim settings file exported from v6.0 or newer."
        Exit Sub
    End If
    
    ' Import settings - Profile
    If System.PrivateProfileString(SettingsFileName, "Profile", "CollegeHS") = "College" Then
        Me.optCollege.Value = True
    Else
        Me.optK12.Value = True
    End If
    
    If System.PrivateProfileString(SettingsFileName, "Profile", "Event") = "LD" Then
        Me.optLD.Value = True
    ElseIf System.PrivateProfileString(SettingsFileName, "Profile", "Event") = "PF" Then
        Me.optPF.Value = True
    Else
        Me.optCX.Value = True
    End If
    
    Me.chkDisableTabroom.Value = System.PrivateProfileString(SettingsFileName, "Profile", "DisableTabroom")
    
    Me.chkAutomaticUpdates.Value = System.PrivateProfileString(SettingsFileName, "Profile", "AutomaticUpdates")
    
    ' Import settings - Admin
    Me.chkAlwaysOn.Value = System.PrivateProfileString(SettingsFileName, "Admin", "AlwaysOn")
    Me.chkAutoUpdateStyles.Value = System.PrivateProfileString(SettingsFileName, "Admin", "AutoUpdateStyles")
    Me.chkSuppressStyleChecks.Value = System.PrivateProfileString(SettingsFileName, "Admin", "SuppressStyleChecks")
    Me.chkSuppressInstallChecks.Value = System.PrivateProfileString(SettingsFileName, "Admin", "SuppressInstallChecks")
    Me.chkSuppressDocCheck.Value = System.PrivateProfileString(SettingsFileName, "Admin", "SuppressDocCheck")
    
    ' Import settings - View
    If System.PrivateProfileString(SettingsFileName, "View", "DefaultView") = "Web" Then
        Me.optWebView.Value = True
    Else
        Me.optDraftView.Value = True
    End If
       
    Me.chkNPCStartup.Value = System.PrivateProfileString(SettingsFileName, "View", "NPCStartup")
    Me.spnDocs.Value = System.PrivateProfileString(SettingsFileName, "View", "DocsPct")
    Me.spnSpeech.Value = System.PrivateProfileString(SettingsFileName, "View", "SpeechPct")
    
    ' Import settings - Paperless
    Me.chkAutoSaveSpeech.Value = System.PrivateProfileString(SettingsFileName, "Paperless", "AutoSaveSpeech")
    Me.cboAutoSaveDir.Value = System.PrivateProfileString(SettingsFileName, "Paperless", "AutoSaveDir")
    Me.chkStripSpeech.Value = System.PrivateProfileString(SettingsFileName, "Paperless", "StripSpeech")
    Me.cboSearchDir.Value = System.PrivateProfileString(SettingsFileName, "Paperless", "SearchDir")
    Me.cboAutoOpenDir.Value = System.PrivateProfileString(SettingsFileName, "Paperless", "AutoOpenDir")
    Me.cboAudioDir.Value = System.PrivateProfileString(SettingsFileName, "Paperless", "AudioDir")
    
    ' Import settings - Styles
    Me.cboNormalSize.Value = System.PrivateProfileString(SettingsFileName, "Styles", "NormalSize")
    Me.cboNormalFont.Value = System.PrivateProfileString(SettingsFileName, "Styles", "NormalFont")
    Me.cboPocketSize.Value = System.PrivateProfileString(SettingsFileName, "Styles", "PocketSize")
    Me.cboHatSize.Value = System.PrivateProfileString(SettingsFileName, "Styles", "HatSize")
    Me.cboBlockSize.Value = System.PrivateProfileString(SettingsFileName, "Styles", "BlockSize")
    Me.cboTagSize.Value = System.PrivateProfileString(SettingsFileName, "Styles", "TagSize")
    Me.cboCiteSize.Value = System.PrivateProfileString(SettingsFileName, "Styles", "CiteSize")
    Me.chkUnderlineCite.Value = System.PrivateProfileString(SettingsFileName, "Styles", "UnderlineCite")
    Me.cboUnderlineSize.Value = System.PrivateProfileString(SettingsFileName, "Styles", "UnderlineSize")
    Me.chkBoldUnderline.Value = System.PrivateProfileString(SettingsFileName, "Styles", "BoldUnderline")
    Me.cboEmphasisSize.Value = System.PrivateProfileString(SettingsFileName, "Styles", "EmphasisSize")
    Me.chkEmphasisBold.Value = System.PrivateProfileString(SettingsFileName, "Styles", "EmphasisBold")
    Me.chkEmphasisItalic.Value = System.PrivateProfileString(SettingsFileName, "Styles", "EmphasisItalic")
    Me.chkEmphasisBox.Value = System.PrivateProfileString(SettingsFileName, "Styles", "EmphasisBox")
    Me.cboEmphasisBoxSize.Value = System.PrivateProfileString(SettingsFileName, "Styles", "EmphasisBoxSize")
    
    ' Import settings - Format
    If System.PrivateProfileString(SettingsFileName, "Format", "Spacing") = "Wide" Then
        Me.optSpacingWide.Value = True
    Else
        Me.optSpacingNarrow.Value = True
    End If

    Me.chkShrinkOmissions.Value = System.PrivateProfileString(SettingsFileName, "Format", "ShrinkOmissions")

    Me.chkParagraphIntegrity.Value = System.PrivateProfileString(SettingsFileName, "Format", "ParagraphIntegrity")
    Me.chkUsePilcrows.Value = System.PrivateProfileString(SettingsFileName, "Format", "UsePilcrows")
    Me.chkCondenseOnPaste.Value = System.PrivateProfileString(SettingsFileName, "Format", "CondenseOnPaste")

    Me.chkAutoUnderlineEmphasis.Value = System.PrivateProfileString(SettingsFileName, "Format", "AutoUnderlineEmphasis")
    
    Me.cboHighlightingException.Value = System.PrivateProfileString(SettingsFileName, "Format", "HighlightingException")
    
    ' Import settings - Keyboard
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
    
    ' Import settings - VTub
    Me.cboVTubPath.Value = System.PrivateProfileString(SettingsFileName, "VTub", "VTubPath")
    Me.chkVTubRefreshPrompt.Value = System.PrivateProfileString(SettingsFileName, "VTub", "VTubRefreshPrompt")
  
    ' Import settings - Caselist
    Me.chkOpenSource.Value = System.PrivateProfileString(SettingsFileName, "Caselist", "OpenSource")
    Me.chkProcessCites.Value = System.PrivateProfileString(SettingsFileName, "Caselist", "ProcessCites")
    
    ' Import settings - Plugins
    Me.cboTimer.Value = System.PrivateProfileString(SettingsFileName, "Plugins", "TimerPath")
    Me.cboOCR.Value = System.PrivateProfileString(SettingsFileName, "Plugins", "OCRPath")
    Me.cboSearch.Value = System.PrivateProfileString(SettingsFileName, "Plugins", "SearchPath")
    
    ' Report success
    MsgBox "Settings successfully imported. They will not be committed until you click Save."
    
    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub btnExportSettings_Click()
    Dim SettingsFileName As String
    Dim ExportPath As String

    On Error GoTo Handler
    
    #If Mac Then
        MsgBox "Exporting settings isn't available for Mac. See the online manual for more information on workarounds.", vbOKOnly
        Exit Sub
    #End If

    ' Create SettingsFile name
    SettingsFileName = "VerbatimSettings"
    If Me.txtSchoolName.Value <> "" Then
        SettingsFileName = SettingsFileName & " - " & Me.txtSchoolName.Value
    End If
    If Me.txtName.Value <> "" Then
        SettingsFileName = SettingsFileName & " - " & Me.txtName.Value
    End If
    SettingsFileName = SettingsFileName & ".ini"

    ExportPath = UI.GetFolderFromDialog("Choose folder for export...", "Export")
    SettingsFileName = ExportPath & "\" & SettingsFileName

    ' Set settings file version
    System.PrivateProfileString(SettingsFileName, "Profile", "Version") = Settings.GetVersion

    ' Export settings - Profile
    If Me.optCollege.Value = True Then
        System.PrivateProfileString(SettingsFileName, "Profile", "CollegeHS") = "College"
    Else
        System.PrivateProfileString(SettingsFileName, "Profile", "CollegeHS") = "K12"
    End If
    
    If Me.optLD.Value = True Then
        System.PrivateProfileString(SettingsFileName, "Profile", "Event") = "LD"
    ElseIf Me.optPF.Value = True Then
        System.PrivateProfileString(SettingsFileName, "Profile", "Event") = "PF"
    Else
        System.PrivateProfileString(SettingsFileName, "Profile", "Event") = "CX"
    End If
    
    System.PrivateProfileString(SettingsFileName, "Profile", "DisableTabroom") = Me.chkDisableTabroom.Value
    
    System.PrivateProfileString(SettingsFileName, "Profile", "AutomaticUpdates") = Me.chkAutomaticUpdates.Value
    
    ' Export settings - Admin
    System.PrivateProfileString(SettingsFileName, "Admin", "AlwaysOn") = Me.chkAlwaysOn.Value
    System.PrivateProfileString(SettingsFileName, "Admin", "AutoUpdateStyles") = Me.chkAutoUpdateStyles.Value
    System.PrivateProfileString(SettingsFileName, "Admin", "SuppressStyleChecks") = Me.chkSuppressStyleChecks.Value
    System.PrivateProfileString(SettingsFileName, "Admin", "SuppressInstallChecks") = Me.chkSuppressInstallChecks.Value
    System.PrivateProfileString(SettingsFileName, "Admin", "SuppressDocCheck") = Me.chkSuppressDocCheck.Value
    
    ' Export settings - View
    If Me.optWebView.Value = True Then
        System.PrivateProfileString(SettingsFileName, "View", "DefaultView") = "Web"
    Else
        System.PrivateProfileString(SettingsFileName, "View", "DefaultView") = "Draft"
    End If
        
    System.PrivateProfileString(SettingsFileName, "View", "NPCStartup") = Me.chkNPCStartup.Value
    System.PrivateProfileString(SettingsFileName, "View", "DocsPct") = Me.spnDocs.Value
    System.PrivateProfileString(SettingsFileName, "View", "SpeechPct") = Me.spnDocs.Value
    
    ' Export settings - Paperless
    System.PrivateProfileString(SettingsFileName, "Paperless", "AutoSaveSpeech") = Me.chkAutoSaveSpeech.Value
    System.PrivateProfileString(SettingsFileName, "Paperless", "AutoSaveDir") = Me.cboAutoSaveDir.Value
    System.PrivateProfileString(SettingsFileName, "Paperless", "StripSpeech") = Me.chkStripSpeech.Value
    System.PrivateProfileString(SettingsFileName, "Paperless", "SearchDir") = Me.cboSearchDir.Value
    System.PrivateProfileString(SettingsFileName, "Paperless", "AutoOpenDir") = Me.cboAutoOpenDir.Value
    System.PrivateProfileString(SettingsFileName, "Paperless", "AudioDir") = Me.cboAudioDir.Value

    ' Export settings - Styles
    System.PrivateProfileString(SettingsFileName, "Styles", "NormalSize") = Me.cboNormalSize.Value
    System.PrivateProfileString(SettingsFileName, "Styles", "NormalFont") = Me.cboNormalFont.Value
    System.PrivateProfileString(SettingsFileName, "Styles", "PocketSize") = Me.cboPocketSize.Value
    System.PrivateProfileString(SettingsFileName, "Styles", "HatSize") = Me.cboHatSize.Value
    System.PrivateProfileString(SettingsFileName, "Styles", "BlockSize") = Me.cboBlockSize.Value
    System.PrivateProfileString(SettingsFileName, "Styles", "TagSize") = Me.cboTagSize.Value
    System.PrivateProfileString(SettingsFileName, "Styles", "CiteSize") = Me.cboCiteSize.Value
    System.PrivateProfileString(SettingsFileName, "Styles", "UnderlineCite") = Me.chkUnderlineCite.Value
    System.PrivateProfileString(SettingsFileName, "Styles", "UnderlineSize") = Me.cboUnderlineSize.Value
    System.PrivateProfileString(SettingsFileName, "Styles", "BoldUnderline") = Me.chkBoldUnderline.Value
    System.PrivateProfileString(SettingsFileName, "Styles", "EmphasisSize") = Me.cboEmphasisSize.Value
    System.PrivateProfileString(SettingsFileName, "Styles", "EmphasisBold") = Me.chkEmphasisBold.Value
    System.PrivateProfileString(SettingsFileName, "Styles", "EmphasisItalic") = Me.chkEmphasisItalic.Value
    System.PrivateProfileString(SettingsFileName, "Styles", "EmphasisBox") = Me.chkEmphasisBox.Value
    System.PrivateProfileString(SettingsFileName, "Styles", "EmphasisBoxSize") = Me.cboEmphasisBoxSize.Value
    
    ' Export settings - Format
    If Me.optSpacingWide.Value = True Then
        System.PrivateProfileString(SettingsFileName, "Format", "Spacing") = "Wide"
    Else
        System.PrivateProfileString(SettingsFileName, "Format", "Spacing") = "Narrow"
    End If
    
    System.PrivateProfileString(SettingsFileName, "Format", "ShrinkOmissions") = Me.chkShrinkOmissions.Value

    System.PrivateProfileString(SettingsFileName, "Format", "ParagraphIntegrity") = Me.chkParagraphIntegrity.Value
    System.PrivateProfileString(SettingsFileName, "Format", "UsePilcrows") = Me.chkUsePilcrows.Value
    System.PrivateProfileString(SettingsFileName, "Format", "CondenseOnPaste") = Me.chkCondenseOnPaste.Value

    System.PrivateProfileString(SettingsFileName, "Format", "AutoUnderlineEmphasis") = Me.chkAutoUnderlineEmphasis.Value
        
    System.PrivateProfileString(SettingsFileName, "Format", "HighlightingException") = Me.cboHighlightingException.Value

    ' Export settings - Keyboard
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

    ' Export settings - VTub
    System.PrivateProfileString(SettingsFileName, "VTub", "VTubPath") = Me.cboVTubPath.Value
    System.PrivateProfileString(SettingsFileName, "VTub", "VTubRefreshPrompt") = Me.chkVTubRefreshPrompt.Value

    ' Export settings - Caselist
    System.PrivateProfileString(SettingsFileName, "Caselist", "OpenSource") = Me.chkOpenSource.Value
    System.PrivateProfileString(SettingsFileName, "Caselist", "ProcessCites") = Me.chkProcessCites.Value
    
    ' Export settings - Plugins
    System.PrivateProfileString(SettingsFileName, "Plugins", "TimerPath") = Me.cboTimer.Value
    System.PrivateProfileString(SettingsFileName, "Plugins", "OCRPath") = Me.cboOCR.Value
    System.PrivateProfileString(SettingsFileName, "Plugins", "SearchPath") = Me.cboSearch.Value

    ' Report success
    MsgBox "Settings successfully exported as:" & vbCrLf & SettingsFileName
    
    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub btnImportCustomCode_Click()
    Settings.ImportCustomCode True
End Sub

Private Sub btnExportCustomCode_Click()
    Settings.ExportCustomCode True
End Sub

'*************************************************************************************
'* VIEW TAB                                                                          *
'*************************************************************************************

Private Sub spnDocs_Change()
    Me.txtDocPct.Value = Me.spnDocs.Value
    Me.lblDocs.Width = 250 * Me.spnDocs.Value / 100
    Me.lblSpeech.Width = (250 * Me.spnSpeech.Value / 100)
    Me.lblSpeech.Left = 250 - Me.lblSpeech.Width
End Sub

Private Sub spnSpeech_Change()
    Me.txtSpeechPct.Value = Me.spnSpeech.Value
    Me.lblDocs.Width = 250 * Me.spnDocs.Value / 100
    Me.lblSpeech.Width = (250 * Me.spnSpeech.Value / 100)
    Me.lblSpeech.Left = 250 - Me.lblSpeech.Width
End Sub

Private Sub btnResetView_Click()
    If MsgBox("This will reset view settings to their default values. Changes will not be committed until you click Save. Are you sure?", vbOKCancel) = vbCancel Then Exit Sub
    
    If Globals.DefaultView = "Web" Then
        Me.optWebView.Value = True
    Else
        Me.optDraftView.Value = True
    End If
    Me.chkNPCStartup.Value = False
    Me.spnDocs.Value = 50
    Me.spnSpeech.Value = 50
    
    Me.chkRibbonDisableSpeech.Value = False
    Me.chkRibbonDisableOrganize.Value = False
    Me.chkRibbonDisableFormat.Value = False
    Me.chkRibbonDisablePaperless.Value = False
    Me.chkRibbonDisableTools.Value = False
    Me.chkRibbonDisableView.Value = False
    Me.chkRibbonDisableCaselist.Value = False
    Me.chkRibbonDisableSettings.Value = False
    
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
    
    Me.cboAutoSaveDir.Value = UI.GetFolderFromDialog("Choose an AutoSave Folder", "Select")
    Me.btnCancel.SetFocus ' Have to switch focus to avoid dropdown getting stuck
    
    On Error GoTo 0
End Sub

Private Sub cboSearchDir_DropButtonClick()
    On Error Resume Next
    
    Me.cboSearchDir.Value = UI.GetFolderFromDialog("Choose a Search Folder", "Select")
    Me.btnCancel.SetFocus 'Have to switch focus to avoid dropdown getting stuck
    
    On Error GoTo 0
End Sub

Private Sub cboAutoOpenDir_DropButtonClick()
    On Error Resume Next
    
    Me.cboAutoOpenDir.Value = UI.GetFolderFromDialog("Choose an AutoOpen Folder", "Select")
    Me.btnCancel.SetFocus 'Have to switch focus to avoid dropdown getting stuck
    
    On Error GoTo 0
End Sub

Private Sub cboAudioDir_DropButtonClick()
    On Error Resume Next
    
    Me.cboAudioDir.Value = UI.GetFolderFromDialog("Choose an Audio Recordings Folder", "Select")
    Me.btnCancel.SetFocus 'Have to switch focus to avoid dropdown getting stuck
    
    On Error GoTo 0
End Sub

'*************************************************************************************
'* STYLES TAB                                                                        *
'*************************************************************************************

Private Sub cboNormalFont_Change()
    ' Changes the font sample
    Me.lblFontSample.Font.Name = Me.cboNormalFont.Value
End Sub

Private Sub chkEmphasisBox_Change()
    Me.cboEmphasisBoxSize.Enabled = Me.chkEmphasisBox.Value
End Sub

Private Sub btnResetStyles_Click()
' Resets style settings to the default
    
    On Error GoTo Handler
    
    ' Prompt for confirmation
    If MsgBox("This will reset styles settings to their default values. Changes will not be committed until you click Save. Are you sure?", vbOKCancel) = vbCancel Then Exit Sub
    
    Me.cboNormalSize.Value = 11
    Me.cboNormalFont.Value = "Calibri"
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
    
    Exit Sub
        
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

'*************************************************************************************
'* FORMAT TAB                                                                        *
'*************************************************************************************

Private Sub chkParagraphIntegrity_Change()
    ' Disable Pilcrows button if unchecked
    Me.chkUsePilcrows.Enabled = Me.chkParagraphIntegrity.Value
End Sub

Private Sub btnResetFormatting_Click()
' Resets formatting settings to the default
    
    On Error GoTo Handler
    
    ' Prompt for confirmation
    If MsgBox("This will reset formatting settings to their default values. Changes will not be committed until you click Save. Are you sure?", vbOKCancel) = vbCancel Then Exit Sub
    
    Me.chkAutoUnderlineEmphasis.Value = False
    Me.chkParagraphIntegrity.Value = False
    Me.chkUsePilcrows.Value = False
    Me.chkCondenseOnPaste.Value = False
    Me.optSpacingWide.Value = True
    
    Me.cboHighlightingException.Value = "None"
    
    Exit Sub
        
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

'*************************************************************************************
'* KEYBOARD TAB                                                                      *
'*************************************************************************************

Private Sub btnOtherKeyboardShortcuts_Click()
    ' Shows the Customize Keyboard dialogue
    Dialogs.Item(wdDialogToolsCustomizeKeyboard).Show
End Sub

Private Sub btnResetKeyboard_Click()
' Resets keyboard settings to the default
    
    On Error GoTo Handler
    
    ' Prompt for confirmation
    If MsgBox("This will reset keyboard settings to their default values. Changes will not be committed until you click Save. Are you sure?", vbOKCancel) = vbCancel Then Exit Sub
    
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
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub btnFixTilde_Click()
    Troubleshooting.FixTilde
    MsgBox "Tilde shortcuts fixed!"
End Sub

'*************************************************************************************
'* VTUB TAB                                                                          *
'*************************************************************************************

Private Sub cboVTubPath_DropButtonClick()
    On Error Resume Next
    
    Me.cboVTubPath.Value = UI.GetFolderFromDialog("Choose a VTub Folder...", "Select")
    Me.btnCancel.SetFocus 'Have to switch focus to avoid dropdown getting stuck
    
    ' Save immediately so the Create button can find a path
    SaveSetting "Verbatim", "VTub", "VTubPath", Me.cboVTubPath.Value
    
    On Error GoTo 0
End Sub

Private Sub btnCreateVTub_Click()
    If Me.cboVTubPath.Value = "" Then
        MsgBox "You must select a path for the VTub first."
        Exit Sub
    Else
        Me.Hide
        If MsgBox("Are you sure you want to create the VTub now?", vbYesNo) = vbYes Then
            VirtualTub.VTubCreate
        End If
        Me.Show
    End If
End Sub

'*************************************************************************************
'* PLUGINS TAB                                                                         *
'*************************************************************************************
Private Sub cboTimer_DropButtonClick()
    On Error Resume Next
    
    Me.cboTimer.Value = UI.GetFileFromDialog("Timer Application", "*.*", "Choose a Timer application...", "Select")
    Me.btnCancel.SetFocus 'Have to switch focus to avoid dropdown getting stuck
    
    On Error GoTo 0
End Sub

Private Sub cboOCR_DropButtonClick()
    On Error Resume Next
    
    Me.cboOCR.Value = UI.GetFileFromDialog("OCR Application", "*.*", "Choose an OCR application...", "Select")
    Me.btnCancel.SetFocus 'Have to switch focus to avoid dropdown getting stuck
    
    On Error GoTo 0
End Sub

Private Sub cboSearch_DropButtonClick()
    On Error Resume Next
    
    Me.cboSearch.Value = UI.GetFileFromDialog("Search Application", "*.*", "Choose a Search application...", "Select")
    Me.btnCancel.SetFocus 'Have to switch focus to avoid dropdown getting stuck
    
    On Error GoTo 0
End Sub

'*************************************************************************************
'* ABOUT TAB                                                                         *
'*************************************************************************************
Private Sub btnUpdateCheck_Click()
    Settings.UpdateCheck True
End Sub

Private Sub lblWebsite_Click()
    Settings.LaunchWebsite Globals.PAPERLESSDEBATE_URL
End Sub

Private Sub btnVerbatimHelp_Click()
    UI.ShowForm "Help"
End Sub

Private Sub btnEasterEgg_Click()
    MsgBox "You have ascended to the Ashtar Command!"
End Sub

