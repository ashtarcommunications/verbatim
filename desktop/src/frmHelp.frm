VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHelp 
   Caption         =   "Verbatim Help"
   ClientHeight    =   1770
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4770
   OleObjectBlob   =   "frmHelp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnManual_Click()
    Unload Me
    Settings.LaunchWebsite ("https://paperlessdebate.com/verbatim")
End Sub

Private Sub btnTutorial_Click()
    Unload Me
    Tutorial.LaunchTutorial
End Sub

Private Sub btnTroubleshooter_Click()
    Unload Me
    UI.ShowForm "Troubleshooting"
End Sub

Private Sub btnSettings_Click()
    Unload Me
    UI.ShowForm "Settings"
End Sub

Private Sub btnOfficeHelp_Click()
    Unload Me
    Settings.OpenWordHelp
End Sub
