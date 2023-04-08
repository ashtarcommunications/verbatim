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

Private Sub UserForm_Initialize()
    Globals.InitializeGlobals
    
    #If Mac Then
        Me.lblPassword.Caption = "Password (careful, shown while typing)"
    #End If
    
    If GetSetting("Verbatim", "Profile", "DisableTabroom", False) = True Then
        MsgBox "Tabroom/Caselist functions are disabled in the Verbatim settings. Please enable to use this feature."
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
Public Sub btnLogin_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnLogin.BackColor = Globals.LIGHT_BLUE
End Sub

Public Sub btnCancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnCancel.BackColor = Globals.LIGHT_RED
End Sub

Public Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnLogin.BackColor = Globals.BLUE
    btnCancel.BackColor = Globals.RED
End Sub
#End If

Private Sub btnLogin_Click()
    Dim Body As Dictionary
    Set Body = New Dictionary
    
    On Error GoTo Handler
    
    Body.Add "username", Me.txtUsername.Value
    Body.Add "password", Me.txtPassword.Value
       
    Dim Response As Variant
    Set Response = HTTP.PostReq(Globals.CASELIST_URL & "/login", Body)
    
    Dim status As String
    status = Response("status")
    
    Dim b As Dictionary
    Dim token As String
    Dim expires As String

    Select Case status
        Case "401"
            MsgBox "Invalid username or password."
            Exit Sub
        Case "201"
            Set b = Response("body")
            token = b.Item("token")
            expires = b.Item("expires")
            SaveSetting "Verbatim", "Caselist", "CaselistToken", token
            SaveSetting "Verbatim", "Caselist", "CaselistTokenExpires", JSONTools.ParseIso(expires)
            Ribbon.RefreshRibbon
            MsgBox "Successfully logged in, you can now use integrated Tabroom/Caselist features!"
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
