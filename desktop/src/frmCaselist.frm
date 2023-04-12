VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCaselist 
   Caption         =   "Upload to openCaselist.com"
   ClientHeight    =   10515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14760
   OleObjectBlob   =   "frmCaselist.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCaselist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule MemberNotOnInterface
Option Explicit

#If Mac Then
#Else
Public Sub btnSubmit_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnSubmit.BackColor = Globals.LIGHT_BLUE
End Sub

Public Sub btnCancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnCancel.BackColor = Globals.LIGHT_RED
End Sub

Public Sub btnAddCite_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnAddCite.BackColor = Globals.LIGHT_GREEN
End Sub

Public Sub btnDeleteCite_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnDeleteCite.BackColor = Globals.LIGHT_RED
End Sub

Public Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnSubmit.BackColor = Globals.BLUE
    btnCancel.BackColor = Globals.RED
    btnAddCite.BackColor = Globals.GREEN
    btnDeleteCite.BackColor = Globals.RED
End Sub
#End If

Private Sub UserForm_Initialize()
    On Error GoTo Handler
    
    Globals.InitializeGlobals
    
    If GetSetting("Verbatim", "Profile", "DisableTabroom", False) = True Then
        MsgBox "Caselist functions are disabled in the Verbatim settings. Please enable to use this feature."
        Me.Hide
        Unload Me
    End If

    #If Mac Then
        UI.ResizeUserForm Me
        Me.btnCancel.ForeColor = Globals.RED
        Me.btnSubmit.ForeColor = Globals.BLUE
        Me.btnAddCite.ForeColor = Globals.GREEN
        Me.btnDeleteCite.ForeColor = Globals.RED
    #End If

    Application.StatusBar = "Checking whether logged into openCaselist..."
    If Caselist.CheckCaselistToken = False Then
        Me.Hide
        UI.ShowForm "Login"
        Unload Me
        Exit Sub
    End If
    
    ' Get rounds from Tabroom
    Application.StatusBar = "Retrieving rounds from openCaselist..."
    Dim Response As Dictionary
    Set Response = HTTP.GetReq(Globals.CASELIST_URL & "/tabroom/rounds")
    Application.StatusBar = "Retrieved rounds from openCaselist"
    
    If Response.Item("status") = 401 Then
        Me.Hide
        UI.ShowForm "Login"
        Unload Me
        Exit Sub
    End If
    
    Me.cboTournament.AddItem ""
    Me.cboTournament.AddItem "All Tournaments / General Disclosure"
    
    Dim Round As Variant
    Dim RoundName As String
    Dim Side As String
    Dim SideName As String

    Me.cboSelectRound.AddItem
    Me.cboSelectRound.List(Me.cboSelectRound.ListCount - 1, 1) = ""
    Me.cboSelectRound.List(Me.cboSelectRound.ListCount - 1, 2) = ""
    Me.cboSelectRound.List(Me.cboSelectRound.ListCount - 1, 3) = ""
    Me.cboSelectRound.List(Me.cboSelectRound.ListCount - 1, 4) = ""
    Me.cboSelectRound.List(Me.cboSelectRound.ListCount - 1, 5) = ""
    
    If Response.Item("body").Count = 0 Then
        Me.cboSelectRound.List(Me.cboSelectRound.ListCount - 1, 0) = "No rounds found on Tabroom"
    Else
        Me.cboSelectRound.List(Me.cboSelectRound.ListCount - 1, 0) = "Select a Round"
    End If

    Me.cboSelectRound.ListIndex = 0

    For Each Round In Response.Item("body")
        RoundName = Strings.RoundName(Round("round"))
        Side = Round("side")
        SideName = Strings.DisplaySide(Side)
        
        Me.cboSelectRound.AddItem
        
        Me.cboSelectRound.List(Me.cboSelectRound.ListCount - 1, 0) = Round("tournament") & " " & RoundName & " " & SideName & " vs " & Round("opponent")
        Me.cboSelectRound.List(Me.cboSelectRound.ListCount - 1, 1) = Round("tournament")
        Me.cboSelectRound.List(Me.cboSelectRound.ListCount - 1, 2) = Round("round")
        Me.cboSelectRound.List(Me.cboSelectRound.ListCount - 1, 3) = Strings.NormalizeSide(Side)
        Me.cboSelectRound.List(Me.cboSelectRound.ListCount - 1, 4) = Round("opponent")
        Me.cboSelectRound.List(Me.cboSelectRound.ListCount - 1, 5) = Round("judge")
    Next Round
    
    ' Add Round and side options to dropdowns
    Me.cboRound.AddItem
    Me.cboRound.List(Me.cboRound.ListCount - 1, 0) = ""
    Me.cboRound.List(Me.cboRound.ListCount - 1, 1) = ""
    Me.cboRound.AddItem
    Me.cboRound.List(Me.cboRound.ListCount - 1, 0) = "All"
    Me.cboRound.List(Me.cboRound.ListCount - 1, 1) = "All"
    Me.cboRound.AddItem
    Me.cboRound.List(Me.cboRound.ListCount - 1, 0) = "Round 1"
    Me.cboRound.List(Me.cboRound.ListCount - 1, 1) = "1"
    Me.cboRound.AddItem
    Me.cboRound.List(Me.cboRound.ListCount - 1, 0) = "Round 2"
    Me.cboRound.List(Me.cboRound.ListCount - 1, 1) = "2"
    Me.cboRound.AddItem
    Me.cboRound.List(Me.cboRound.ListCount - 1, 0) = "Round 3"
    Me.cboRound.List(Me.cboRound.ListCount - 1, 1) = "3"
    Me.cboRound.AddItem
    Me.cboRound.List(Me.cboRound.ListCount - 1, 0) = "Round 4"
    Me.cboRound.List(Me.cboRound.ListCount - 1, 1) = "4"
    Me.cboRound.AddItem
    Me.cboRound.List(Me.cboRound.ListCount - 1, 0) = "Round 5"
    Me.cboRound.List(Me.cboRound.ListCount - 1, 1) = "5"
    Me.cboRound.AddItem
    Me.cboRound.List(Me.cboRound.ListCount - 1, 0) = "Round 6"
    Me.cboRound.List(Me.cboRound.ListCount - 1, 1) = "6"
    Me.cboRound.AddItem
    Me.cboRound.List(Me.cboRound.ListCount - 1, 0) = "Round 7"
    Me.cboRound.List(Me.cboRound.ListCount - 1, 1) = "7"
    Me.cboRound.AddItem
    Me.cboRound.List(Me.cboRound.ListCount - 1, 0) = "Round 8"
    Me.cboRound.List(Me.cboRound.ListCount - 1, 1) = "8"
    Me.cboRound.AddItem
    Me.cboRound.List(Me.cboRound.ListCount - 1, 0) = "Round 9"
    Me.cboRound.List(Me.cboRound.ListCount - 1, 1) = "9"
    Me.cboRound.AddItem
    Me.cboRound.List(Me.cboRound.ListCount - 1, 0) = "Quads"
    Me.cboRound.List(Me.cboRound.ListCount - 1, 1) = "Quads"
    Me.cboRound.AddItem
    Me.cboRound.List(Me.cboRound.ListCount - 1, 0) = "Triples"
    Me.cboRound.List(Me.cboRound.ListCount - 1, 1) = "Triples"
    Me.cboRound.AddItem
    Me.cboRound.List(Me.cboRound.ListCount - 1, 0) = "Doubles"
    Me.cboRound.List(Me.cboRound.ListCount - 1, 1) = "Doubles"
    Me.cboRound.AddItem
    Me.cboRound.List(Me.cboRound.ListCount - 1, 0) = "Octas"
    Me.cboRound.List(Me.cboRound.ListCount - 1, 1) = "Octas"
    Me.cboRound.AddItem
    Me.cboRound.List(Me.cboRound.ListCount - 1, 0) = "Quarters"
    Me.cboRound.List(Me.cboRound.ListCount - 1, 1) = "Quarters"
    Me.cboRound.AddItem
    Me.cboRound.List(Me.cboRound.ListCount - 1, 0) = "Semis"
    Me.cboRound.List(Me.cboRound.ListCount - 1, 1) = "Semis"
    Me.cboRound.AddItem
    Me.cboRound.List(Me.cboRound.ListCount - 1, 0) = "Finals"
    Me.cboRound.List(Me.cboRound.ListCount - 1, 1) = "Finals"
    
    Me.cboSide.AddItem
    Me.cboSide.List(Me.cboSide.ListCount - 1, 0) = ""
    Me.cboSide.List(Me.cboSide.ListCount - 1, 1) = ""
    Me.cboSide.AddItem
    Me.cboSide.List(Me.cboSide.ListCount - 1, 0) = "Aff"
    Me.cboSide.List(Me.cboSide.ListCount - 1, 1) = "A"
    Me.cboSide.AddItem
    Me.cboSide.List(Me.cboSide.ListCount - 1, 0) = "Neg"
    Me.cboSide.List(Me.cboSide.ListCount - 1, 1) = "N"
        
    If GetSetting("Verbatim", "Caselist", "OpenSource", False) = True Then
        Me.chkOpenSource.Value = True
    End If
    
    ' Blank options for caselist page selectors
    Me.cboCaselists.AddItem
    Me.cboCaselists.List(Me.cboCaselists.ListCount - 1, 0) = ""
    Me.cboCaselists.List(Me.cboCaselists.ListCount - 1, 1) = ""
    Me.cboCaselists.Value = ""
    
    Me.cboCaselistSchoolName.AddItem
    Me.cboCaselistSchoolName.List(Me.cboCaselistSchoolName.ListCount - 1, 0) = ""
    Me.cboCaselistSchoolName.List(Me.cboCaselistSchoolName.ListCount - 1, 1) = ""
    Me.cboCaselistSchoolName.Value = ""
    
    Me.cboCaselistTeamName.AddItem
    Me.cboCaselistTeamName.List(Me.cboCaselistTeamName.ListCount - 1, 0) = ""
    Me.cboCaselistTeamName.List(Me.cboCaselistTeamName.ListCount - 1, 1) = ""
    Me.cboCaselistTeamName.Value = ""
    
    ' Use default caselist page
    Dim DefaultCaselist As String
    DefaultCaselist = GetSetting("Verbatim", "Caselist", "DefaultCaselist", "")
    If DefaultCaselist <> "" Then
        Me.cboCaselists.AddItem
        Me.cboCaselists.List(Me.cboCaselists.ListCount - 1, 0) = Split(DefaultCaselist, "|")(0)
        Me.cboCaselists.List(Me.cboCaselists.ListCount - 1, 1) = Split(DefaultCaselist, "|")(1)
        Me.cboCaselists.Value = Split(DefaultCaselist, "|")(1)
    End If
    
    Dim DefaultCaselistSchool As String
    DefaultCaselistSchool = GetSetting("Verbatim", "Caselist", "DefaultCaselistSchool", "")
    If DefaultCaselistSchool <> "" Then
        Me.cboCaselistSchoolName.AddItem
        Me.cboCaselistSchoolName.List(Me.cboCaselistSchoolName.ListCount - 1, 0) = Split(DefaultCaselistSchool, "|")(0)
        Me.cboCaselistSchoolName.List(Me.cboCaselistSchoolName.ListCount - 1, 1) = Split(DefaultCaselistSchool, "|")(1)
        Me.cboCaselistSchoolName.Value = Split(DefaultCaselistSchool, "|")(1)
    End If
    
    Dim DefaultCaselistTeam As String
    DefaultCaselistTeam = GetSetting("Verbatim", "Caselist", "DefaultCaselistTeam", "")
    If DefaultCaselistTeam <> "" Then
        Me.cboCaselistTeamName.AddItem
        Me.cboCaselistTeamName.List(Me.cboCaselistTeamName.ListCount - 1, 0) = Split(DefaultCaselistTeam, "|")(0)
        Me.cboCaselistTeamName.List(Me.cboCaselistTeamName.ListCount - 1, 1) = Split(DefaultCaselistTeam, "|")(1)
        Me.cboCaselistTeamName.Value = Split(DefaultCaselistTeam, "|")(1)
    End If
    
    If GetSetting("Verbatim", "Caselist", "ProcessCites", True) = True Then
        ProcessCiteEntries
    Else
        AddCiteEntry "", ""
    End If
    
    Exit Sub
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub lblCaselistLink_Click()
    Settings.LaunchWebsite ("https://opencaselist.com")
End Sub

Private Sub cboSelectRound_Change()
    If Me.cboSelectRound.ListIndex > -1 Then
        If Not IsNull(Me.cboSelectRound.List(Me.cboSelectRound.ListIndex, 1)) _
            And Me.cboSelectRound.List(Me.cboSelectRound.ListIndex, 1) <> "" _
        Then
            Me.cboTournament.AddItem Me.cboSelectRound.List(Me.cboSelectRound.ListIndex, 1)
            Me.cboTournament.Value = Me.cboSelectRound.List(Me.cboSelectRound.ListIndex, 1)
        End If
        If Not IsNull(Me.cboSelectRound.List(Me.cboSelectRound.ListIndex, 2)) Then Me.cboRound.Value = Me.cboSelectRound.List(Me.cboSelectRound.ListIndex, 2)
        If Not IsNull(Me.cboSelectRound.List(Me.cboSelectRound.ListIndex, 3)) Then Me.cboSide.Value = Me.cboSelectRound.List(Me.cboSelectRound.ListIndex, 3)
        If Not IsNull(Me.cboSelectRound.List(Me.cboSelectRound.ListIndex, 4)) Then Me.txtOpponent.Value = Me.cboSelectRound.List(Me.cboSelectRound.ListIndex, 4)
        If Not IsNull(Me.cboSelectRound.List(Me.cboSelectRound.ListIndex, 5)) Then Me.txtJudge.Value = Me.cboSelectRound.List(Me.cboSelectRound.ListIndex, 5)
    Else
        Me.cboRound.Enabled = True
        Me.txtOpponent.Enabled = True
        Me.txtJudge.Enabled = True
        Me.txtRoundReport.Enabled = True
        
        Me.cboTournament.Value = ""
        Me.cboSide.Value = ""
        Me.cboRound.Value = ""
        Me.txtOpponent.Value = ""
        Me.txtJudge.Value = ""
    End If
End Sub

Private Sub cboTournament_Change()
    If Me.cboTournament.Value = "All Tournaments / General Disclosure" Then
        Me.cboRound.Value = "All"
        Me.cboRound.Enabled = False
        Me.txtOpponent.Value = ""
        Me.txtOpponent.Enabled = False
        Me.txtJudge.Value = ""
        Me.txtJudge.Enabled = False
        Me.txtRoundReport.Value = ""
        Me.txtRoundReport.Enabled = False
    Else
        Me.cboRound.Enabled = True
        Me.txtOpponent.Enabled = True
        Me.txtJudge.Enabled = True
        Me.txtRoundReport.Enabled = True
    End If
End Sub

Private Sub ProcessCiteEntries()
    Dim LargestHeading As Long
    LargestHeading = Formatting.LargestHeading
       
    Dim CiteEntries As Dictionary
    Set CiteEntries = New Dictionary
    Dim Entry As Dictionary
    
    Dim i As Long
    i = 1
    
    Dim p As Paragraph
    Dim key As Variant
    
    For Each p In ActiveDocument.Paragraphs
        If p.OutlineLevel = LargestHeading Then
            ' Limit submission to 5 cite entries
            i = i + 1
            If i <= 5 Then
                Set Entry = New Dictionary
    
                Selection.Start = p.Range.Start
                Paperless.SelectHeadingAndContent
                
                Entry.Add "Title", Strings.HeadingToTitle(p.Range.Text)
                Selection.MoveStart wdParagraph, 1
                Entry.Add "Content", WikifySelection
                CiteEntries.Add p.Range.Text, Entry
                Selection.Collapse
            End If
        End If
    Next p
    
    For Each key In CiteEntries.Keys
        Set Entry = CiteEntries.Item(key)
        AddCiteEntry Trim$(Entry.Item("Title")), Trim$(Entry.Item("Content"))
    Next key
End Sub

Private Sub AddCiteEntry(ByRef Title As String, ByRef Content As String)
    Dim TitleLabel As Object
    Dim TitleBox As Object
    Dim EntryLabel As Object
    Dim EntryText As Object
    Dim RuleLabel As Object
    
    Dim NumEntries As Long
    
    ' Each Cite Entry is comprised of five controls
    If Me.fCites.Controls.Count > 0 Then
        NumEntries = Me.fCites.Controls.Count / 5
    Else
        NumEntries = 0
    End If
    NumEntries = NumEntries + 1
            
    ' Create Title Label - All other controls positioning keyed off this
    Set TitleLabel = Me.fCites.Controls.Add("Forms.Label.1", "lblEntryTitle" & NumEntries)
    TitleLabel.Caption = "Title " & NumEntries
    TitleLabel.Height = 12
    TitleLabel.Width = 65
    TitleLabel.Left = 5
    TitleLabel.Top = fCites.ScrollHeight + 10
    
    ' Create Title Box
    Set TitleBox = Me.fCites.Controls.Add("Forms.TextBox.1", "txtEntryTitle" & NumEntries)
    TitleBox.Height = 18
    TitleBox.Width = fCites.Width - 60
    TitleBox.Left = 5
    TitleBox.Top = TitleLabel.Top + TitleLabel.Height + 5
    TitleBox.Value = Trim$(Title)
        
    ' Create Entry Label
    Set EntryLabel = Me.fCites.Controls.Add("Forms.Label.1", "lblEntryContent" & NumEntries)
    EntryLabel.Caption = "Entry " & NumEntries
    EntryLabel.Height = 12
    EntryLabel.Width = 65
    EntryLabel.Left = 5
    EntryLabel.Top = TitleBox.Top + TitleBox.Height + 5
    
    ' Create Entry Box
    Set EntryText = Me.fCites.Controls.Add("Forms.TextBox.1", "txtEntryContent" & NumEntries)
    EntryText.Height = 100
    EntryText.Width = fCites.Width - 40
    EntryText.Left = 5
    EntryText.Top = EntryLabel.Top + EntryLabel.Height + 5
    EntryText.MultiLine = True
    EntryText.EnterKeyBehavior = True
    EntryText.ScrollBars = 2
    EntryText.Font.size = 8
    EntryText.Value = Trim$(Content)
    
    Set RuleLabel = Me.fCites.Controls.Add("Forms.Label.1", "lblEntryRule" & NumEntries)
    RuleLabel.Height = 1
    RuleLabel.Width = fCites.Width - 40
    RuleLabel.Left = 5
    RuleLabel.Top = EntryText.Top + EntryText.Height + 10
    RuleLabel.Caption = ""
    RuleLabel.BorderStyle = 1
    
    ' Add ScrollHeight and scroll to bottom
    Me.fCites.ScrollHeight = Me.fCites.ScrollHeight + 180
    Me.fCites.ScrollTop = Me.fCites.ScrollHeight
End Sub

Private Function WikifySelection() As String
    Dim d As Document
    On Error GoTo Handler
    
    Application.ScreenUpdating = False
        
    ' Copy selection
    Selection.Copy
    
    ' Add new document based on debate template
    Set d = Application.Documents.Add(Template:=ActiveDocument.AttachedTemplate.FullName, Visible:=False)
    d.Activate

    ' Paste into new document
    Selection.Paste
    
    ' Go to top of document and collapse selection
    Selection.HomeKey Unit:=wdStory
    Selection.Collapse

    ' Convert to cites
    Caselist.CiteRequestAll

    ' Wikify and clear formatting
    Caselist.Word2MarkdownMain
    ActiveDocument.Content.Select
    Selection.ClearFormatting
    
    WikifySelection = Selection.Text
    d.Close wdDoNotSaveChanges
       
    Application.ScreenUpdating = True

    Set d = Nothing
    
    Exit Function
    
Handler:
    Set d = Nothing
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Function

Private Sub btnAddCite_Click()
    AddCiteEntry "", ""
End Sub

Private Sub btnDeleteCite_Click()
    Dim NumEntries As Long
    If Me.fCites.Controls.Count > 0 Then
        NumEntries = Me.fCites.Controls.Count / 5

        ' Delete last entry
        Me.fCites.Controls.Remove ("lblEntryTitle" & NumEntries)
        Me.fCites.Controls.Remove ("txtEntryTitle" & NumEntries)
        Me.fCites.Controls.Remove ("lblEntryContent" & NumEntries)
        Me.fCites.Controls.Remove ("txtEntryContent" & NumEntries)
        Me.fCites.Controls.Remove ("lblEntryRule" & NumEntries)
           
        ' Remove excess ScrollHeight
        Me.fCites.ScrollHeight = Me.fCites.ScrollHeight - 180
    End If
End Sub

Private Sub cboCaselists_DropButtonClick()
    On Error GoTo Handler

    Dim DefaultCaselist As String
    DefaultCaselist = GetSetting("Verbatim", "Caselist", "DefaultCaselist", "")
    
    ' Only fetch the list once per form load, if we have more than default entry we've already fetched
    If (DefaultCaselist <> "" And Me.cboCaselists.ListCount > 2) Or (DefaultCaselist = "" And Me.cboCaselists.ListCount > 1) Then Exit Sub
    
    Me.cboCaselists.Clear
    Me.cboCaselists.AddItem
    Me.cboCaselists.List(Me.cboCaselists.ListCount - 1, 0) = ""
    Me.cboCaselists.List(Me.cboCaselists.ListCount - 1, 1) = ""
    Me.cboCaselists.Value = ""
            
    UI.PopulateComboBoxFromJSON Globals.CASELIST_URL & "/caselists", "display_name", "name", Me.cboCaselists
    
    If DefaultCaselist <> "" Then
        Dim i As Long
        For i = 0 To Me.cboCaselists.ListCount - 1
            If Me.cboCaselists.List(i, 1) = Split(DefaultCaselist, "|")(1) Then
                Me.cboCaselists.Value = Split(DefaultCaselist, "|")(1)
            End If
        Next i
    End If
    
    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub cboCaselists_Change()
    ' Clear ComboBoxes - clear TeamName too, so there's not a mismatch when changing
    Me.cboCaselistSchoolName.Clear
    Me.cboCaselistTeamName.Clear
    
    Me.cboCaselistSchoolName.AddItem
    Me.cboCaselistSchoolName.List(Me.cboCaselistSchoolName.ListCount - 1, 0) = ""
    Me.cboCaselistSchoolName.List(Me.cboCaselistSchoolName.ListCount - 1, 1) = ""
    Me.cboCaselistSchoolName.Value = ""
    
    Me.cboCaselistTeamName.AddItem
    Me.cboCaselistTeamName.List(Me.cboCaselistTeamName.ListCount - 1, 0) = ""
    Me.cboCaselistTeamName.List(Me.cboCaselistTeamName.ListCount - 1, 1) = ""
    Me.cboCaselistTeamName.Value = ""
End Sub

Private Sub cboCaselistSchoolName_DropButtonClick()
    Dim DefaultCaselistSchool As String
    
    On Error GoTo Handler

    DefaultCaselistSchool = GetSetting("Verbatim", "Caselist", "DefaultCaselistSchool", "")
    
    ' If the list is already populated, exit
    If (DefaultCaselistSchool <> "" And Me.cboCaselistSchoolName.ListCount > 2) Or (DefaultCaselistSchool = "" And Me.cboCaselistSchoolName.ListCount > 1) Then Exit Sub
    
    If Me.cboCaselists.Value = "" Or IsNull(Me.cboCaselists.Value) = True Then Exit Sub
               
    Me.cboCaselistSchoolName.Clear
    Me.cboCaselistSchoolName.AddItem
    Me.cboCaselistSchoolName.List(Me.cboCaselistSchoolName.ListCount - 1, 0) = ""
    Me.cboCaselistSchoolName.List(Me.cboCaselistSchoolName.ListCount - 1, 1) = ""
    Me.cboCaselistSchoolName.Value = ""
               
    Dim URL As String
    URL = Globals.CASELIST_URL & "/caselists/" + Me.cboCaselists.Value & "/schools"
    UI.PopulateComboBoxFromJSON URL, "display_name", "name", Me.cboCaselistSchoolName
        
    If DefaultCaselistSchool <> "" Then
        Dim i As Long
        For i = 0 To Me.cboCaselistSchoolName.ListCount - 1
            If Me.cboCaselistSchoolName.List(i, 1) = Split(DefaultCaselistSchool, "|")(1) Then
                Me.cboCaselistSchoolName.Value = Split(DefaultCaselistSchool, "|")(1)
            End If
        Next i
    End If
    
    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub cboCaselistSchoolName_Change()
    ' Clear TeamName too, so there's not a mismatch when changing
    Me.cboCaselistTeamName.Clear
    
    Me.cboCaselistTeamName.AddItem
    Me.cboCaselistTeamName.List(Me.cboCaselistTeamName.ListCount - 1, 0) = ""
    Me.cboCaselistTeamName.List(Me.cboCaselistTeamName.ListCount - 1, 1) = ""
    Me.cboCaselistTeamName.Value = ""

End Sub

Private Sub cboCaselistTeamName_DropButtonClick()
    Dim DefaultCaselistTeam As String
    
    On Error GoTo Handler
    
    DefaultCaselistTeam = GetSetting("Verbatim", "Caselist", "DefaultCaselistTeam", "")
    
    ' If the list is already populated, exit
    If (DefaultCaselistTeam <> "" And Me.cboCaselistTeamName.ListCount > 2) Or (DefaultCaselistTeam = "" And Me.cboCaselistTeamName.ListCount > 1) Then Exit Sub
    
    If Me.cboCaselists.Value = "" Or Me.cboCaselistSchoolName.Value = "" Or IsNull(Me.cboCaselists.Value) = True Or IsNull(Me.cboCaselistSchoolName.Value) = True Then Exit Sub
               
    Me.cboCaselistTeamName.Clear
    Me.cboCaselistTeamName.AddItem
    Me.cboCaselistTeamName.List(Me.cboCaselistTeamName.ListCount - 1, 0) = ""
    Me.cboCaselistTeamName.List(Me.cboCaselistTeamName.ListCount - 1, 1) = ""
    Me.cboCaselistTeamName.Value = ""
    
    Dim URL As String
    URL = Globals.CASELIST_URL & "/caselists/" + Me.cboCaselists.Value & "/schools/" & Me.cboCaselistSchoolName.Value & "/teams"
    UI.PopulateComboBoxFromJSON URL, "display_name", "name", Me.cboCaselistTeamName
    
    If DefaultCaselistTeam <> "" Then
        Dim i As Long
        For i = 0 To Me.cboCaselistTeamName.ListCount - 1
            If Me.cboCaselistTeamName.List(i, 1) = Split(DefaultCaselistTeam, "|")(1) Then
                Me.cboCaselistTeamName.Value = Split(DefaultCaselistTeam, "|")(1)
            End If
        Next i
    End If
    
    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub btnSubmit_Click()
    Me.UploadToCaselist
End Sub

Public Function ValidateForm() As Boolean

    Me.cboTournament.BorderColor = Globals.DARK_GRAY
    Me.cboSide.BorderColor = Globals.DARK_GRAY
    Me.cboRound.BorderColor = Globals.DARK_GRAY
    
    Me.lblError.Visible = False
    Me.lblError.Caption = "No Errors"
    ValidateForm = True

    If Trim$(Me.cboTournament.Value) = "" Then
        Me.cboTournament.BorderColor = Globals.RED
        Me.lblError.Caption = "Tournament, side, and round are required!"
        Me.lblError.Visible = True
        ValidateForm = False
        Exit Function
    End If
    
    If Me.cboSide.Value = "" Then
        Me.cboSide.BorderColor = Globals.RED
        Me.lblError.Caption = "Tournament, side, and round are required!"
        Me.lblError.Visible = True
        ValidateForm = False
        Exit Function
    End If
    
    If Me.cboRound.Value = "" Then
        Me.cboRound.BorderColor = Globals.RED
        Me.lblError.Caption = "Tournament, side, and round are required!"
        Me.lblError.Visible = True
        ValidateForm = False
        Exit Function
    End If
    
    If Me.chkOpenSource.Value = False And Me.fCites.Controls.Count = 0 Then
        Me.lblError.Caption = "Nothing to upload! You must either include cites or upload as open source."
        Me.lblError.Visible = True
        ValidateForm = False
        Exit Function
    End If
    
    If Me.fCites.Controls.Count > 0 Then
        If Trim$(Me.fCites.Controls.Item("txtEntryTitle1").Value) = "" Or Trim$(Me.fCites.Controls.Item("txtEntryContent1").Value) = "" Then
            Me.lblError.Caption = "Cite entries require a title and content!"
            Me.lblError.Visible = True
            ValidateForm = False
            Exit Function
        End If
    End If
    
    If Me.cboCaselists.Value = "" _
        Or Me.cboCaselistSchoolName.Value = "" _
        Or Me.cboCaselistTeamName.Value = "" _
        Or IsNull(Me.cboCaselists.Value) = True _
        Or IsNull(Me.cboCaselistSchoolName.Value) = True _
        Or IsNull(Me.cboCaselistTeamName.Value) = True _
    Then
        Me.lblError.Caption = "You must select a caselist, school, and team!"
        Me.lblError.Visible = True
        ValidateForm = False
        Exit Function
    End If
End Function

Public Sub UploadToCaselist()

    On Error GoTo Handler
    
    Application.ScreenUpdating = False
    System.Cursor = wdCursorWait
    
    If ValidateForm = False Then Exit Sub

    If ActiveDocument.Saved = False Then ActiveDocument.Save
    
    Dim Body As Dictionary
    Set Body = New Dictionary
    
    Dim Cites As Collection
    Set Cites = New Collection
    
    Dim Cite As Dictionary
    
    If Me.cboTournament.Value = "All Tournaments / General Disclosure" Then
        Body.Add "tournament", "All Tournaments"
        Body.Add "side", Me.cboSide.Value
        Body.Add "round", "All"
    Else
        Body.Add "tournament", Trim$(Me.cboTournament.Value)
        Body.Add "side", Me.cboSide.Value
        Body.Add "round", Me.cboRound.Value
        Body.Add "opponent", Trim$(Me.txtOpponent.Value)
        Body.Add "judge", Trim$(Me.txtJudge.Value)
        Body.Add "report", Trim$(Me.txtRoundReport.Value)
    End If
    
    If Me.chkOpenSource.Value = True Then
        Application.StatusBar = "Preparing file for upload to openCaselist..."
        Dim Base64 As String
        Base64 = Filesystem.GetFileAsBase64(ActiveDocument.FullName)
        Body.Add "opensource", Base64
        Body.Add "filename", ActiveDocument.Name
    End If
    
    Dim NumCites As Long
    NumCites = Me.fCites.Controls.Count / 5
    
    If NumCites > 0 Then
        Dim i As Long
        For i = 1 To NumCites
            Set Cite = New Dictionary
            Cite.Add "title", Trim$(Me.fCites.Controls.Item("txtEntryTitle" & i).Value)
            Cite.Add "cites", Trim$(Me.fCites.Controls.Item("txtEntryContent" & i).Value)
            Cites.Add Cite
        Next
        
        Body.Add "cites", Cites
    End If
       
    Dim URL As String
    URL = Globals.CASELIST_URL
    URL = URL & "/caselists/" & Me.cboCaselists.Value
    URL = URL & "/schools/" & Me.cboCaselistSchoolName.Value
    URL = URL & "/teams/" & Me.cboCaselistTeamName.Value
    URL = URL & "/rounds"
    
    Application.StatusBar = "Uploading to openCaselist..."
    
    Dim Request As Dictionary
    Set Request = HTTP.PostReq(URL, Body)
    
    Select Case Request.Item("status")
        Case Is = "201" ' Created
            MsgBox "Round successfully created on openCaselist"
            Application.StatusBar = "Round successfully created on openCaselist"
            Me.Hide
            Unload Me
        Case Is = "401" ' Unauthorized
            Me.lblError.Caption = "Unauthorized, log in first and try again"
            Me.lblError.Visible = True
            Me.Hide
            UI.ShowForm "Login"
            Me.Show
        Case Is = "404" ' Not Found
            Me.lblError.Caption = "Unauthorized, log in first"
            Me.lblError.Visible = True
        Case Else
            Me.lblError.Caption = Request.Item("status")
            Me.lblError.Visible = True
    End Select
    
    If Me.chkSaveDefault.Value = True Then
        SaveSetting "Verbatim", "Caselist", "DefaultCaselist", Me.cboCaselists.Text & "|" & Me.cboCaselists.Value
        SaveSetting "Verbatim", "Caselist", "DefaultCaselistSchool", Me.cboCaselistSchoolName.Text & "|" & Me.cboCaselistSchoolName.Value
        SaveSetting "Verbatim", "Caselist", "DefaultCaselistTeam", Me.cboCaselistTeamName.Text & "|" & Me.cboCaselistTeamName.Value
    End If
    
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

Private Sub btnCancel_Click()
    Unload Me
End Sub

