VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSetupWizard 
   Caption         =   "Verbatim Setup Wizard"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7515
   OleObjectBlob   =   "frmSetupWizard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSetupWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    On Error GoTo Handler
    
    Dim TemplateFolder As String
    
    Globals.InitializeGlobals
    
    #If Mac Then
        UI.ResizeUserForm Me
        Me.btnCancel.ForeColor = Globals.RED
        Me.btnOK.ForeColor = Globals.BLUE
    #End If
    
    ' Run install checks
    If GetSetting("Verbatim", "Admin", "SuppressInstallChecks", False) = False Then
        If Troubleshooting.InstallCheckTemplateName = False Then
            MsgBox "WARNING - Verbatim appears to be installed incorrectly as " _
                & ActiveDocument.AttachedTemplate.Name _
                & " - Verbatim should always be named ""Debate.dotm"" or many features will not work correctly. " _
                & "It is strongly recommended you change the file name back." _
                & vbCrLf & vbCrLf & "Please close Verbatim and install correctly before proceeding."
            Unload Me
            Exit Sub
        ElseIf Troubleshooting.InstallCheckTemplateLocation = False Then
            #If Mac Then
                TemplateFolder = "/Users/" & Environ("USER") & "/Library/Group Containers/UBF8T34G9.Office/User Content/Templates"
            #Else
                TemplateFolder = "c:\Users\<yourname>\AppData\Roaming\Microsoft\Templates"
            #End If
            MsgBox "WARNING - Verbatim appears to be installed in the wrong location. " _
                   & "The Verbatim template file (Debate.dotm) should be located in your Word Templates folder, usually located at: " _
                   & TemplateFolder _
                   & ". Using it from a different location will break many features." _
                   & vbCrLf & vbCrLf & "Please close Verbatim and install correctly before proceeding."
            Unload Me
            Exit Sub
        End If
    End If
    
    ' Set defaults
    Me.chkAlwaysOn.Value = GetSetting("Verbatim", "Admin", "AlwaysOn", True)
    
    If GetSetting("Verbatim", "Profile", "CollegeHS", "College") = "College" Then
        Me.optCollege.Value = True
    Else
        Me.optK12.Value = True
    End If
    
    Me.optCX.Value = True
    
    Me.chkTutorial.Value = True
    
    Exit Sub
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub btnOK_Click()
    On Error GoTo Handler
    
    SaveSetting "Verbatim", "Admin", "AlwaysOn", Me.chkAlwaysOn.Value
        
    If Me.optCollege.Value = True Then
        SaveSetting "Verbatim", "Profile", "CollegeHS", "College"
    Else
        SaveSetting "Verbatim", "Profile", "CollegeHS", "K12"
    End If
    
    If Me.optCX.Value = True Then
        SaveSetting "Verbatim", "Profile", "Event", "CX"
    ElseIf Me.optLD.Value = True Then
        SaveSetting "Verbatim", "Profile", "Event", "LD"
    ElseIf Me.optPF.Value = True Then
        SaveSetting "Verbatim", "Profile", "Event", "PF"
    Else
        SaveSetting "Verbatim", "Profile", "Event", "CX"
    End If
    
    Unload Me
        
    If Me.chkTutorial.Value = True Then UI.LaunchTutorial
           
    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub btnCancel_Click()
    If MsgBox("Are you sure you want to exit without completing the Setup Wizard?", vbYesNo) = vbYes Then Unload Me
End Sub

#If Mac Then
    ' Do Nothing
#Else
Public Sub btnOK_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnOK.BackColor = Globals.LIGHT_BLUE
End Sub

Public Sub btnCancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnCancel.BackColor = Globals.LIGHT_RED
End Sub

Public Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnOK.BackColor = Globals.BLUE
    btnCancel.BackColor = Globals.RED
End Sub
#End If


