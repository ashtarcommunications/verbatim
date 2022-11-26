VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShare 
   Caption         =   "Share on share.tabroom.com"
   ClientHeight    =   6975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8955
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
    
    #If Mac Then
        UI.ResizeUserForm Me
    #End If
    
    Dim Phrase
    Phrase = RandomPhrase
    
    If GetSetting("Verbatim", "Profile", "DisableTabroom", False) = False Then
        Dim Response As Dictionary
        'Set Response = HTTP.GetReq(Globals.CASELIST_URL & "/tabroom/rounds?current=true")
        Set Response = HTTP.GetReq(Globals.MOCK_ROUNDS)
        
        Select Case Response("status")
        Case Is = "200"
            Me.lblError.Visible = False
        Case Is = "401"
            Me.lblLogin1.Visible = True
            Me.lblLogin2.Visible = True
            Exit Sub
        Case Else
            Me.lblError.Caption = Response("body")("message")
            Me.lblError.Visible = True
            Exit Sub
        End Select
                
        Dim Round
        Dim RoundName
        Dim Side As String
        Dim SideName
        
        For Each Round In Response("body")
            If (Round("share") <> vbNullString) Then
                RoundName = Strings.RoundName(Round("round"))
                Side = Round("side")
                SideName = Strings.DisplaySide(Side)
                Me.lboxRounds.AddItem
                Me.lboxRounds.List(Me.lboxRounds.ListCount - 1, 0) = Round("tournament") & " " & RoundName & " " & SideName & " vs " & Round("opponent") & " (" & Round("share") & ")"
                Me.lboxRounds.List(Me.lboxRounds.ListCount - 1, 1) = Round("share")
            End If
        Next Round

    End If

    Me.lboxRounds.AddItem
    Me.lboxRounds.List(Me.lboxRounds.ListCount - 1, 0) = "Random New Room: " & Phrase
    Me.lboxRounds.List(Me.lboxRounds.ListCount - 1, 1) = Phrase

    Exit Sub

Handler:
    Set Request = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description
    
End Sub

Private Sub lboxRounds_Change()
    Me.txtRoom.Value = ""
    ValidateForm
End Sub

Private Sub txtRoom_Change()
    Me.lboxRounds.Value = vbNullString
    Me.txtRoom.Value = Strings.OnlyAlphaNumericChars(Me.txtRoom.Value)
    ValidateForm
End Sub

Private Sub ValidateForm()
    If Me.txtRoom.Value = vbNullString And Me.lboxRounds.Value = vbNullString Then
        Me.btnShare.Enabled = False
        Me.btnBrowser.Enabled = False
    ElseIf Me.lboxRounds.Value = vbNullString And Me.txtRoom.TextLength < 8 Then
        Me.btnShare.Enabled = False
        Me.btnBrowser.Enabled = False
    ElseIf Me.lboxRounds.Value = vbNullString And Len(Strings.OnlyAlphaNumericChars(Me.txtRoom.Value)) < 8 Then
        Me.btnShare.Enabled = False
        Me.btnBrowser.Enabled = False
    Else
        Me.btnShare.Enabled = True
        Me.btnBrowser.Enabled = True
    End If
End Sub

Sub btnShare_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.btnShare.BackColor = Globals.LIGHT_BLUE
End Sub

Sub btnBrowser_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.btnBrowser.BackColor = Globals.LIGHT_BLUE
End Sub

Sub btnCancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.btnCancel.BackColor = Globals.LIGHT_RED
End Sub

Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.btnShare.BackColor = Globals.BLUE
    Me.btnBrowser.BackColor = Globals.BLUE
    Me.btnCancel.BackColor = Globals.RED
End Sub

Private Sub lblURL_Click()
    Settings.LaunchWebsite Globals.SHARE_URL
End Sub

Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub lblLoginLink_Click()
    Me.Hide
    UI.ShowForm "Login"
    Unload Me
    Exit Sub
End Sub

Private Sub btnShare_Click()
    UploadToShare lboxRounds.List(lboxRounds.ListIndex, 1)
End Sub

Private Function RandomPhrase() As String
    Dim Adjectives
    Dim Animals
    
    Dim Adjective
    Dim Animal
    Dim RandomNumber
    
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
    RandomNumber = CInt(Rnd() * (999 - 1) + 1)
    If (RandomNumber = 69 Or RandomNumber = 420 Or RandomNumber = 666) Then RandomNumber = RandomNumber + 1

    RandomPhrase = Adjective + Animal + CStr(RandomNumber)
End Function

Private Sub UploadToShare()
    On Error GoTo Handler
    
    Application.ScreenUpdating = False
    System.Cursor = wdCursorWait
    
    If ValidateForm = False Then Exit Sub

    If ActiveDocument.Saved = False Then ActiveDocument.Save
    
    Dim Body As Dictionary
    Set Body = New Dictionary
       
    Dim Base64
    Base64 = Filesystem.GetFileAsBase64(ActiveDocument.FullName)
    Body.Add "file", Base64
    Body.Add "filename", ActiveDocument.Name
           
    Dim Room As String
    If Me.lboxRounds.Value <> vbNullString Then Room = Me.lboxRounds.Value
    If Me.txtRoom.Value <> vbNullString Then Room = Strings.OnlyAlphaNumericChars(Me.txtRoom.Value)
    
    Dim Request
    Set Request = HTTP.PostReq(Globals.SHARE_URL & "/" & Room, Body)
    
    Select Case Request("status")
    Case Is = "200"                              ' Success
        MsgBox "File successfully shared to https://share.tabroom.com/" & Room & " - anyone linked to your round on Tabroom has been emailed!"
        Me.Hide
        Unload Me
        
    Case Is = "400"                              ' Bad file
        Me.lblError.Caption = "Something appears to be wrong with your file, please try again"
        Me.lblError.Visible = True

    Case Is = "500"                              ' Server error
        Me.lblError.Caption = "Failed to upload file, please try again"
        Me.lblError.Visible = True
    
    Case Else
        Me.lblError.Caption = Request("body")("message")
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


