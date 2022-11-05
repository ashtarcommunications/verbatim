VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShare 
   Caption         =   "Share"
   ClientHeight    =   7035
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "frmShare.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmShare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnShare_Click()
    Debug.Print lboxRounds.List(lboxRounds.ListIndex, 1)
    ' UploadToShare lboxRounds.List(lboxRounds.ListIndex, 1)
End Sub

Private Sub UserForm_Initialize()
    #If Mac Then
        UI.ResizeUserForm Me
    #End If
    
    Dim Phrase
    Phrase = RandomPhrase
    
    Dim Rounds As Object
    Set Rounds = Caselist.GetRounds
    
    If Rounds Is Nothing Then
        ' Me.lblNoRounds.Visible = True
    Else
        Dim Item
        For Each Item In Rounds
            Dim Round
            If Item("share") <> "" Then
                Round = Item("tournament") + " Round " + Item("round") + " (" + Item("share") + ")"
                Me.lboxRounds.AddItem Round
                Me.lboxRounds.List(Me.lboxRounds.ListCount - 1, 1) = Item("share")
            End If
        Next Item
     
    End If
    Me.lboxRounds.AddItem "Random New Room: " + Phrase
    
End Sub

Private Sub lboxRounds_Click()
    Me.txtRoom.Value = ""
    ValidateForm
End Sub

Private Sub txtRoom_Change()
    Me.lboxRounds.Selected(0) = False
    ValidateForm
End Sub

Private Sub ValidateForm()

    If Me.txtRoom.Value = "" And Me.lboxRounds.Value = "" Then
        Me.btnShare.Enabled = False
        Me.btnBrowser.Enabled = False
    ElseIf Me.lboxRounds.Value = "" And Me.txtRoom.TextLength < 8 Then
        Me.btnShare.Enabled = False
        Me.btnBrowser.Enabled = False
    ' ElseIf Me.lboxRounds.Value = "" And Not Strings.IsAlphanumeric(Me.txtRoom.Value) Then
    
    Else
    
        Me.btnShare.Enabled = True
        Me.btnBrowser.Enabled = True
    End If

End Sub

Sub btnSubmit_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    ' btnSubmit.BackColor = RGB(114, 142, 171)
End Sub
Sub btnCancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    ' btnCancel.BackColor = RGB(241, 136, 136)
End Sub
Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    ' btnSubmit.BackColor = RGB(64, 92, 121)
    ' btnCancel.BackColor = RGB(191, 86, 86)
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
