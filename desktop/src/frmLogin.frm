VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "Login"
   ClientHeight    =   3975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4860
   OleObjectBlob   =   "frmLogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Have to use global to store password because password masking doesn't work on Mac
#If Mac Then
    Dim Password As String
#End If

Private Sub UserForm_Initialize()
    #If Mac Then
        Password = ""
    #End If
    
    Globals.InitializeGlobals
    
    If GetSetting("Verbatim", "Profile", "DisableTabroom", False) = True Then
        MsgBox "Caselist functions are disabled in the Verbatim settings. Please enable to use this feature."
        Me.Hide
        Unload Me
    End If
    
    #If Mac Then
        UI.ResizeUserForm Me
        Me.btnCancel.ForeColor = Globals.RED
        Me.btnLogin.ForeColor = Globals.BLUE
    #End If
End Sub

#If Mac Then
#Else
Sub btnLogin_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnLogin.BackColor = Globals.LIGHT_BLUE
End Sub

Sub btnCancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnCancel.BackColor = Globals.LIGHT_RED
End Sub

Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnLogin.BackColor = Globals.BLUE
    btnCancel.BackColor = Globals.RED
End Sub
#End If

Private Sub btnLogin_Click()
    Dim Body As Dictionary
    Set Body = New Dictionary
    Body.Add "username", Me.txtUsername.Value
    #If Mac Then
        Body.Add "password", Password
    #Else
        Body.Add "password", Me.txtPassword.Value
    #End If
       
    Dim Response
    Set Response = HTTP.PostReq(Globals.CASELIST_URL & "/login", Body)
    
    Dim status
    status = Response("status")
    
    Dim b
    Dim token As String
    Dim expires As String

    Select Case status
        Case "401"
            MsgBox "Invalid username or password."
            Exit Sub
        Case "201"
            Set b = Response("body")
            token = b("token")
            expires = b("expires")
            SaveSetting "Verbatim", "Caselist", "CaselistToken", token
            SaveSetting "Verbatim", "Caselist", "CaselistTokenExpires", JSONTools.ParseIso(expires)
            MsgBox "Successfully logged in, you can now use integrated caselist features!"
            Me.Hide
            Unload Me
            Exit Sub
        Case Else
            If InStr(status, "Error") > 0 Then
                MsgBox status
            Else
                MsgBox Response("body")("message")
            End If
    End Select
    
    Exit Sub
    
Handler:
    Set Response = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub
 
#If Mac Then
Private Sub txtPassword_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyBack
            If Len(Password) <= 1 Then
                Password = ""
            Else
                Password = Left(Password, Len(Password) - 1)
            End If
        Case Else
            Password = Password & Chr(KeyCode)
            KeyCode = 0
    End Select
    Me.txtPassword.Value = String(Len(Password), "*")
End Sub
#End If
