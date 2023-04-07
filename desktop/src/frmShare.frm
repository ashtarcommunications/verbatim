VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShare 
   Caption         =   "Share on share.tabroom.com"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8445
   OleObjectBlob   =   "frmShare.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmShare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    On Error GoTo Handler
    
    Globals.InitializeGlobals
    
    #If Mac Then
        UI.ResizeUserForm Me
        Me.btnCancel.ForeColor = Globals.RED
        Me.btnShare.ForeColor = Globals.BLUE
        Me.btnBrowser.ForeColor = Globals.BLUE
    #End If
    
    Dim Phrase As String
    Phrase = RandomPhrase
    
    If GetSetting("Verbatim", "Profile", "DisableTabroom", False) = False Then
        Dim Response As Dictionary
        Set Response = HTTP.GetReq(Globals.CASELIST_URL & "/tabroom/rounds?current=true")
        
        Select Case Response.Item("status")
        Case Is = "200"
            Me.lblError.Visible = False
            
            Dim Round As Variant
            Dim RoundName As String
            Dim Side As String
            Dim SideName As String
            
            For Each Round In Response.Item("body")
                If (Round("share") <> "") Then
                    RoundName = Strings.RoundName(Round("round"))
                    Side = Round("side")
                    SideName = Strings.DisplaySide(Side)
                    Me.lboxRounds.AddItem
                    Me.lboxRounds.List(Me.lboxRounds.ListCount - 1, 0) = Round("tournament") & " " & RoundName & " " & SideName & " vs " & Round("opponent") & " (" & Round("share") & ")"
                    Me.lboxRounds.List(Me.lboxRounds.ListCount - 1, 1) = Round("share")
                End If
            Next Round
        
        Case Is = "401"
            Me.lblLogin1.Visible = True
            Me.lblLoginLink.Visible = True
        Case Else
            Me.lblLogin1.Visible = True
            Me.lblLoginLink.Visible = True
        End Select

    End If

    Me.lboxRounds.AddItem
    Me.lboxRounds.List(Me.lboxRounds.ListCount - 1, 0) = "Random New Room: " & Phrase
    Me.lboxRounds.List(Me.lboxRounds.ListCount - 1, 1) = Phrase

    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    
End Sub

Private Sub lboxRounds_Change()
    Me.txtRoom.Value = ""
    '@Ignore FunctionReturnValueDiscarded
    ValidateForm
End Sub

Private Sub txtRoom_Change()
    Me.lboxRounds.Value = ""
    Me.txtRoom.Value = Strings.OnlyAlphaNumericChars(Me.txtRoom.Value)
    '@Ignore FunctionReturnValueDiscarded
    ValidateForm
End Sub

Private Function ValidateForm() As Boolean
    If Me.txtRoom.Value = "" And (Me.lboxRounds.Value = "" Or IsNull(Me.lboxRounds.Value)) Then
        Me.btnShare.Enabled = False
        Me.btnBrowser.Enabled = False
        ValidateForm = False
    ElseIf (Me.lboxRounds.Value = "" Or IsNull(Me.lboxRounds.Value)) And Me.txtRoom.TextLength < 8 Then
        Me.btnShare.Enabled = False
        Me.btnBrowser.Enabled = False
        ValidateForm = False
    ElseIf (Me.lboxRounds.Value = "" Or IsNull(Me.lboxRounds.Value)) And Len(Strings.OnlyAlphaNumericChars(Me.txtRoom.Value)) < 8 Then
        Me.btnShare.Enabled = False
        Me.btnBrowser.Enabled = False
        ValidateForm = False
    Else
        Me.btnShare.Enabled = True
        Me.btnBrowser.Enabled = True
        ValidateForm = True
    End If
End Function

#If Mac Then
#Else
Public Sub btnShare_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.btnShare.BackColor = Globals.LIGHT_BLUE
End Sub

Public Sub btnBrowser_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.btnBrowser.BackColor = Globals.LIGHT_BLUE
End Sub

Public Sub btnCancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.btnCancel.BackColor = Globals.LIGHT_RED
End Sub

Public Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.btnShare.BackColor = Globals.BLUE
    Me.btnBrowser.BackColor = Globals.BLUE
    Me.btnCancel.BackColor = Globals.RED
End Sub
#End If

Private Sub lblURL_Click()
    Settings.LaunchWebsite Globals.SHARE_URL
End Sub

Public Sub btnBrowser_Click()
    Dim Room As String
    If Me.lboxRounds.Value <> "" Then Room = Me.lboxRounds.Value
    If Me.txtRoom.Value <> "" Then Room = Strings.OnlyAlphaNumericChars(Me.txtRoom.Value)
    
    If Room = "" Then
        MsgBox "You must select or enter a round name first!"
        Exit Sub
    End If

    Settings.LaunchWebsite Globals.SHARE_URL & "/" & Room
End Sub

Public Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub lblLoginLink_Click()
    Me.Hide
    UI.ShowForm "Login"
    Unload Me
    Exit Sub
End Sub

Private Sub btnShare_Click()
    UploadToShare
End Sub

Private Function RandomPhrase() As String
    Dim Adjectives() As Variant
    Dim Animals() As Variant
    
    Dim Adjective As String
    Dim Animal As String
    Dim RandomNumber As Long
    
    Adjectives = Array( _
        "ancient", "average", "best", "better", "big", "blue", "brave", "breezy", "bright", "brown", "calm", _
        "chatty", "chilly", "clear", "clever", "cold", "curly", "dry", "early", "easy", "empty", "fast", "fluffy", _
        "free", "fresh", "friendly", "funny", "fuzzy", "gentle", "giant", "good", "great", "green", "happy", _
        "honest", "hot", "huge", "hungry", "jolly", "kind", "large", "late", "light", "little", "local", "long", _
        "loud", "lovely", "lucky", "massive", "mighty", "modern", "natural", "neat", "new", "nice", "odd", "old", _
        "open", "orange", "ordinary", "perfect", "pink", "plastic", "polite", "popular", "pretty", "proud", "public", _
        "purple", "quick", "quiet", "rare", "real", "red", "serious", "shaggy", "sharp", "short", "shy", "silent", _
        "silly", "simple", "single", "small", "smart", "smooth", "social", "soft", "sour", "spicy", "splendid", _
        "spotty", "stale", "strange", "strong", "sweet", "swift", "tall", "tame", "tender", "thin", "tidy", "tiny", _
        "tough", "weak", "wet", "wise", "witty", "wonderful", "yellow", "young" _
    )
    
    Animals = Array( _
        "badger", "bat", "bear", "bird", "bobcat", "bulldog", "cat", "catfish", "cheetah", "chicken", _
        "chipmunk", "cobra", "cougar", "cow", "crab", "deer", "dog", "dolphin", "dragon", "duck", _
        "eagle", "eel", "elephant", "emu", "falcon", "fish", "fly", "fox", "frog", "gecko", "goat", _
        "goose", "grasshopper", "horse", "hound", "insect", "jellyfish", "kangaroo", "ladybug", "lion", _
        "lizard", "monkey", "moose", "moth", "mouse", "mule", "newt", "octopus", "otter", "owl", _
        "panda", "panther", "parrot", "penguin", "pig", "puma", "rabbit", "rat", "robin", "seahorse", _
        "sheep", "shrimp", "skunk", "sloth", "snail", "snake", "squid", "swan", "tiger", "turkey", _
        "turtle", "walrus", "wasp", "wombat", "yak", "zebra" _
    )

    Adjective = StrConv(Adjectives(CInt(Rnd() * UBound(Adjectives))), vbProperCase)
    Animal = StrConv(Animals(CInt(Rnd() * UBound(Animals))), vbProperCase)
    RandomNumber = CLng(Rnd() * (999 - 1) + 1)
    If (RandomNumber = 69 Or RandomNumber = 420 Or RandomNumber = 666) Then RandomNumber = RandomNumber + 1

    RandomPhrase = Adjective + Animal + CStr(RandomNumber)
End Function

Private Sub UploadToShare()
    On Error GoTo Handler
    
    Dim Filename As String
    
    Application.ScreenUpdating = False
    System.Cursor = wdCursorWait
    
    If ValidateForm = False Then Exit Sub

    If ActiveDocument.Saved = False Then ActiveDocument.Save
    
    ' Strip "Speech" if option set
    If GetSetting("Verbatim", "Paperless", "StripSpeech", True) = True And Len(ActiveDocument.Name) > 11 Then
        Filename = Trim$(Replace(ActiveDocument.Name, "speech", "", 1, -1, vbTextCompare))
    Else
        Filename = Trim$(ActiveDocument.Name)
    End If
    
    Dim Body As Dictionary
    Set Body = New Dictionary
       
    Dim Base64 As String
    Base64 = Filesystem.GetFileAsBase64(ActiveDocument.FullName)
    Body.Add "file", Base64
    Body.Add "filename", Filename
           
    Dim Room As String
    If Me.lboxRounds.Value <> "" Then Room = Me.lboxRounds.Value
    If Me.txtRoom.Value <> "" Then Room = Strings.OnlyAlphaNumericChars(Me.txtRoom.Value)
    
    Dim Request As Dictionary
    Set Request = HTTP.PostReq(Globals.SHARE_URL & "/" & Room, Body)
    
    Select Case Request.Item("status")
    Case Is = "200" ' Success
        MsgBox "File successfully shared to https://share.tabroom.com/" & Room & " - anyone linked to your round on Tabroom has been emailed!"
        Me.Hide
        Unload Me
        
    Case Is = "400" ' Bad file
        Me.lblError.Caption = "Something appears to be wrong with your file, please try again"
        Me.lblError.Visible = True

    Case Is = "500" ' Server error
        Me.lblError.Caption = "Failed to upload file, please try again"
        Me.lblError.Visible = True
    
    Case Is = "504" ' Gateway timeout
        Me.lblError.Caption = "Connection time out, please try again"
        Me.lblError.Visible = True

    Case Else
        Me.lblError.Caption = "Error uploading file, please try again. " & Request.Item("status")
        Me.lblError.Visible = True
    End Select
    
    Set Body = Nothing
    Set Request = Nothing
    
    Application.ScreenUpdating = True
    System.Cursor = wdCursorNormal
    Exit Sub
    
Handler:
    Set Body = Nothing
    Set Request = Nothing
    Application.ScreenUpdating = True
    System.Cursor = wdCursorNormal
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub
