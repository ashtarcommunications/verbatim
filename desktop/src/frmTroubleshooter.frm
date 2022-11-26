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

    #If Mac Then
        UI.ResizeUserForm Me
    #End If

    Dim TemplateLocation As String

    Me.lblWarnings.Caption = vbNullString

    ' Run install checks and exit early if installed incorrectly
    If Troubleshooting.InstallCheckTemplateName = True Then
        Me.lblWarnings.ForeColor = Globals.RED
        Me.lblWarnings.Caption = Me.lblWarnings.Caption & "Verbatim appears to be installed incorrectly as " & ActiveDocument.AttachedTemplate.Name & " - Verbatim should always be named ""Debate.dotm"" or many features will not work correctly. It is strongly recommended you change the file name back." & vbCrLf & vbCrLf & "Please close Verbatim and install correctly before proceeding."
        Exit Sub
    ElseIf Troubleshooting.InstallCheckTemplateLocation = True Then
        Me.lblWarnings.ForeColor = Globals.RED

        #If Mac Then
            ' TODO get right location for Mac
            TemplateLocation = "User:whatever"
        #Else
            TemplateLocation = "c:\Users\<yourname>\AppData\Roaming\Microsoft\Templates"
        #End If
        
        Me.lblWarnings.Caption = Me.lblWarnings.Caption & "Verbatim appears to be installed in the wrong location. The Verbatim template file (Debate.dotm) should be located in your Word Templates folder, usually located at:" & vbCrLf & TemplateLocation & vbCrLf & "Using it from a different location will break many features." & vbCrLf & vbCrLf & "Please close Verbatim and install correctly before proceeding."
        Exit Sub
    Else
        Me.lblWarnings.ForeColor = Globals.BLACK
        Me.lblWarnings.Caption = "Verbatim appears to be installed correctly."
    End If
    
    'Run rest of checks
    If Troubleshooting.CheckDuplicateTemplates = False Then
        Me.btnFix.Visible = True
        Me.lblWarnings.ForeColor = Globals.RED
        Me.lblWarnings.Caption = Me.lblWarnings.Caption & vbCrLf & "Multiple versions of the Verbatim template detected on your computer (likely in your Downloads or Desktop folder). This can cause difficulties with file interoperability."
    End If
    
    If Troubleshooting.CheckAddins = False Then
        Me.btnFix.Visible = True
        Me.lblWarnings.ForeColor = Globals.RED
        Me.lblWarnings.Caption = Me.lblWarnings.Caption & vbCrLf & "A buggy non-Verbatim Word Add-in has been detected. This can cause annoying prompts to save changes to Debate.dotm. You can manually disable Word Add-Ins in Word Options - Add-Ins.."
    End If
    
    If Troubleshooting.CheckDefaultSave = False Then
        Me.btnFix.Visible = True
        Me.lblWarnings.ForeColor = Globals.RED
        Me.lblWarnings.Caption = Me.lblWarnings.Caption & vbCrLf & "Your default file format is not set to .docx - this will cause problems with interoperability and many Verbatim functions."
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

Sub btnFix_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnFix.BackColor = Globals.LIGHT_BLUE
End Sub

Sub btnCancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnCancel.BackColor = Globals.LIGHT_RED
End Sub

Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnFix.BackColor = Globals.BLUE
    btnCancel.BackColor = Globals.RED
End Sub


