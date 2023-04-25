VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTroubleshooter 
   Caption         =   "Verbatim Troubleshooter"
   ClientHeight    =   5775
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9045
   OleObjectBlob   =   "frmTroubleshooter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTroubleshooter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Activate()

    On Error GoTo Handler

    Globals.InitializeGlobals

    #If Mac Then
        UI.ResizeUserForm Me
        Me.btnFix.ForeColor = Globals.BLUE
        Me.btnCancel.ForeColor = Globals.RED
    #End If

    Dim TemplateLocation As String

    Me.lblWarnings.Caption = ""

    ' Run install checks and exit early if installed incorrectly
    If Troubleshooting.InstallCheckTemplateName = False Then
        Me.lblWarnings.ForeColor = Globals.RED
        Me.lblWarnings.Caption = Me.lblWarnings.Caption _
            & "Verbatim appears to be installed incorrectly as " _
            & ActiveDocument.AttachedTemplate.Name _
            & " - Verbatim should always be named ""Debate.dotm"" or many features will not work correctly. " _
            & "It is strongly recommended you change the file name back." _
            & vbCrLf & vbCrLf & "Please close Verbatim and install correctly before proceeding."
        Exit Sub
    ElseIf Troubleshooting.InstallCheckTemplateLocation = False Then
        Me.lblWarnings.ForeColor = Globals.RED

        #If Mac Then
            TemplateLocation = "/Users/<yourusername>/Library/Group Containers/UBF8T346G9.Office/User Content/Templates"
        #Else
            TemplateLocation = "c:\Users\<yourname>\AppData\Roaming\Microsoft\Templates"
        #End If
        
        Me.lblWarnings.Caption = Me.lblWarnings.Caption _
            & "Verbatim appears to be installed in the wrong location. " _
            & "The Verbatim template file (Debate.dotm) should be located in your Word Templates folder, usually located at:" _
            & vbCrLf & TemplateLocation & vbCrLf _
            & "Using it from a different location will break many features." _
            & vbCrLf & vbCrLf & "Please close Verbatim and install correctly before proceeding."
        Exit Sub
    Else
        Me.lblWarnings.ForeColor = Globals.BLACK
        Me.lblWarnings.Caption = "Verbatim appears to be installed correctly." & vbCrLf
    End If
    
    #If Mac Then
        If Troubleshooting.InstallCheckScptFileExists = False Then
            Me.lblWarnings.ForeColor = Globals.RED
            Me.lblWarnings.Caption = Me.lblWarnings.Caption _
                & "Your Verbatim.scpt file is not found at ~/Library/Application Scripts/com.Microsoft.Word/Verbatim.scpt " _
                & "- this will cause many Verbatim features to break. You should rerun the Verbatim installer or install the file manually before proceeding."
            Exit Sub
        End If
    #End If
    
    ' Run rest of checks
    If Troubleshooting.CheckDuplicateTemplates = False Then
        Me.btnFix.Visible = True
        Me.lblWarnings.ForeColor = Globals.RED
        Me.lblWarnings.Caption = Me.lblWarnings.Caption & vbCrLf _
            & "Multiple versions of the Verbatim template detected on your computer (likely in your Downloads or Desktop folder). " _
            & "This can cause difficulties with file interoperability." _
            & vbCrLf
    End If
    
    If Troubleshooting.CheckAddins = False Then
        Me.btnFix.Visible = True
        Me.lblWarnings.ForeColor = Globals.RED
        Me.lblWarnings.Caption = Me.lblWarnings.Caption & vbCrLf _
            & "A buggy non-Verbatim Word Add-in has been detected. This can cause annoying prompts to save changes to Debate.dotm. " _
            & "You can manually disable Word Add-Ins in Word Options - Add-Ins." _
            & vbCrLf
    End If
    
    If Troubleshooting.CheckSaveFormat = False Then
        Me.btnFix.Visible = True
        Me.lblWarnings.ForeColor = Globals.RED
        Me.lblWarnings.Caption = Me.lblWarnings.Caption & vbCrLf _
            & "Your default file format is not set to .docx - this will cause problems with interoperability and many Verbatim functions." _
            & vbCrLf
    End If
    
    Exit Sub
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnFix_Click()
    Troubleshooting.DeleteDuplicateTemplates
    Troubleshooting.DisableAddins
    Troubleshooting.SetDefaultSave
    Me.Hide
    Me.Show
End Sub

#If Mac Then
#Else
Public Sub btnFix_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnFix.BackColor = Globals.LIGHT_BLUE
End Sub

Public Sub btnCancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnCancel.BackColor = Globals.LIGHT_RED
End Sub

Public Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnFix.BackColor = Globals.BLUE
    btnCancel.BackColor = Globals.RED
End Sub
#End If
