VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHelp 
   Caption         =   "Verbatim Help"
   ClientHeight    =   1770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5520
   OleObjectBlob   =   "frmHelp.frx":0000
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'@Ignore EmptyMethod
Private Sub UserForm_Initialize()
    #If Mac Then
        UI.ResizeUserForm Me
    #End If
End Sub

Private Sub btnManual_Click()
    Unload Me
    Settings.LaunchWebsite Globals.PAPERLESSDEBATE_URL
End Sub

Private Sub btnTutorial_Click()
    Unload Me
    UI.LaunchTutorial
End Sub

Private Sub btnTroubleshooter_Click()
    Unload Me
    UI.ShowForm "Troubleshooter"
End Sub

Private Sub btnSettings_Click()
    Unload Me
    UI.ShowForm "Settings"
End Sub

Private Sub btnOfficeHelp_Click()
    Unload Me
    Settings.OpenWordHelp
End Sub

