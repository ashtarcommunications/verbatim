VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   Caption         =   "Settings"
   ClientHeight    =   7710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6360
   OleObjectBlob   =   "frmSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim FontSize As Long
    
    Globals.InitializeGlobals
    
    #If Mac Then
        UI.ResizeUserForm Me
        Me.btnCancel.ForeColor = Globals.RED
        Me.btnReset.ForeColor = Globals.ORANGE
        Me.btnSave.ForeColor = Globals.BLUE
    #End If
    
    Me.chkInsertMode.Value = GetSetting("Verbatim", "Flow", "InsertMode", False)
    Me.chkExtendWithArrow.Value = GetSetting("Verbatim", "Flow", "ExtendWithArrow", False)
    Me.chkAutoLabelFlows.Value = GetSetting("Verbatim", "Flow", "AutoLabelFlows", True)
    Me.chkDisableSheetPopup.Value = GetSetting("Verbatim", "Flow", "DisableSheetPopup", False)
    Me.chkFreezeSpeechNames.Value = GetSetting("Verbatim", "Flow", "FreezeSpeechNames", True)
    Me.chkAlternateTildeCode.Value = GetSetting("Verbatim", "Flow", "AlternateTildeCode", False)
    
    Me.cboSpeechNames.AddItem "1AC,1NC,2AC,Block,1AR,2NR,2AR"
    Me.cboSpeechNames.AddItem "AC,NC,1AR,NR,2AR"
    Me.cboSpeechNames.AddItem "AC,NC,AR,NR,AS,NS,AF,NF"
    
    Me.cboSpeechNames.Value = GetSetting("Verbatim", "Flow", "SpeechNames", "1AC,1NC,2AC,Block,1AR,2NR,2AR")
    
    FontSize = 4
    Do While FontSize < 16
        Me.cboFontSize.AddItem FontSize
        FontSize = FontSize + 1
    Loop

    Me.cboFontSize.Value = GetSetting("Verbatim", "Flow", "FontSize", 8)
    Me.spnRowHeight.Value = GetSetting("Verbatim", "Flow", "RowHeight", 12)
    Me.spnColumnWidth.Value = GetSetting("Verbatim", "Flow", "ColumnWidth", 36)
End Sub

Private Sub spnRowHeight_Change()
    Me.txtRowHeight.Value = Me.spnRowHeight.Value
End Sub

Private Sub spnColumnWidth_Change()
    Me.txtColumnWidth.Value = Me.spnColumnWidth.Value
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnReset_Click()
    If MsgBox("This will reset all settings, but will not make changes until you press Save. Proceed?", vbYesNo) = vbNo Then Exit Sub
    
    Me.chkInsertMode.Value = False
    Me.chkAutoLabelFlows.Value = True
    Me.chkDisableSheetPopup.Value = False
    Me.chkFreezeSpeechNames.Value = True
    Me.chkAlternateTildeCode.Value = False
    
    Me.cboSpeechNames.Value = "1AC,1NC,2AC,Block,1AR,2NR,2AR"
    
    Me.cboFontSize.Value = 8
    Me.spnRowHeight.Value = 12
    Me.spnColumnWidth.Value = 36
End Sub

Private Sub btnSave_Click()
    SaveSetting "Verbatim", "Flow", "InsertMode", Me.chkInsertMode.Value
    SaveSetting "Verbatim", "Flow", "ExtendWithArrow", Me.chkExtendWithArrow.Value
    SaveSetting "Verbatim", "Flow", "AutoLabelFlows", Me.chkAutoLabelFlows.Value
    SaveSetting "Verbatim", "Flow", "DisableSpeechPopup", Me.chkDisableSheetPopup.Value
    SaveSetting "Verbatim", "Flow", "FreezeSpeechNames", Me.chkFreezeSpeechNames.Value
    SaveSetting "Verbatim", "Flow", "AlternateTildeCode", Me.chkAlternateTildeCode.Value
    
    SaveSetting "Verbatim", "Flow", "SpeechNames", Strings.OnlyCSV(Me.cboSpeechNames.Value)
    
    SaveSetting "Verbatim", "Flow", "FontSize", Me.cboFontSize.Value
    SaveSetting "Verbatim", "Flow", "RowHeight", Me.spnRowHeight.Value
    SaveSetting "Verbatim", "Flow", "ColumnWidth", Me.spnColumnWidth.Value

    Me.Hide
    Unload Me
End Sub
