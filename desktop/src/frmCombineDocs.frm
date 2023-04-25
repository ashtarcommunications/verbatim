VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCombineDocs 
   Caption         =   "Combine Docs"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8220
   OleObjectBlob   =   "frmCombineDocs.frx":0000
End
Attribute VB_Name = "frmCombineDocs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim rFile As RecentFile
           
    On Error GoTo Handler
    
    Globals.InitializeGlobals
    
    #If Mac Then
        UI.ResizeUserForm Me
        Me.btnCancel.ForeColor = Globals.RED
        Me.btnCombine.ForeColor = Globals.BLUE
        Me.btnManualAdd.ForeColor = Globals.GREEN
    #End If
    
    ' Add all recent files to the box
    For Each rFile In Application.RecentFiles
        Me.lboxDocs.AddItem
        Me.lboxDocs.List(Me.lboxDocs.ListCount - 1, 0) = rFile.Name
        Me.lboxDocs.List(Me.lboxDocs.ListCount - 1, 1) = rFile.Path & "\" & rFile.Name
    Next rFile
       
    ' Reset AutoName box and add a blank item
    Me.cboAutoName.Clear
    Me.cboAutoName.AddItem
    Me.cboAutoName.List(Me.cboAutoName.ListCount - 1) = ""
    
    ' Get rounds from Tabroom
    If GetSetting("Verbatim", "Profile", "DisableTabroom", False) = True Then
        Exit Sub
    End If

    If Caselist.CheckCaselistToken = False Then
        Me.Hide
        UI.ShowForm "Login"
        Unload Me
        Exit Sub
    End If
    
    Dim Response As Dictionary
    Set Response = HTTP.GetReq(Globals.CASELIST_URL & "/tabroom/rounds")
    
    If Response.Item("status") = 401 Then
        Me.Hide
        UI.ShowForm "Login"
        Unload Me
        Exit Sub
    End If
    
    Me.cboAutoName.AddItem
    Me.cboAutoName.List(Me.cboAutoName.ListCount - 1, 0) = "Select a Round"
    Me.cboAutoName.List(Me.cboAutoName.ListCount - 1, 1) = ""
    Me.cboAutoName.ListIndex = 0
       
    Dim Round As Variant
    Dim RoundName As String
    Dim Side As String
    Dim SideName As String

    If Response.Item("body").Count = 0 Then
        Me.cboAutoName.List(0, 0) = "No rounds found on Tabroom"
        Me.cboAutoName.ListIndex = 0
        Me.cboAutoName.Enabled = False
    End If

    For Each Round In Response.Item("body")
        RoundName = Strings.RoundName(Round("round"))
        Side = Round("side")
        SideName = Strings.DisplaySide(Side)
        
        Me.lboxDocs.AddItem
        
        Me.cboAutoName.List(Me.cboAutoName.ListCount - 1, 0) = Round("tournament") & " " & RoundName & " " & SideName & " vs " & Round("opponent")
        Me.cboAutoName.List(Me.cboAutoName.ListCount - 1, 1) = Round("tournament") & " " & RoundName & " " & SideName & " vs " & Round("opponent")
        
    Next Round
    
    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

#If Mac Then
#Else
Public Sub btnCombine_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnCombine.BackColor = Globals.LIGHT_BLUE
End Sub

Public Sub btnCancel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnCancel.BackColor = Globals.LIGHT_RED
End Sub

Public Sub btnManualAdd_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnCombine.BackColor = Globals.LIGHT_GREEN
End Sub

Public Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnCombine.BackColor = Globals.BLUE
    btnCancel.BackColor = Globals.RED
    btnManualAdd.BackColor = Globals.GREEN
End Sub
#End If

Private Sub btnManualAdd_Click()
    Dim FilePath As String
    
    On Error Resume Next
    
    FilePath = UI.GetFileFromDialog("Select document", "*.doc*", "Choose a document...", "Select")
    
    Me.lboxDocs.AddItem , 0
    Me.lboxDocs.List(0, 0) = Right$(FilePath, Len(FilePath) - InStrRev(FilePath, "\"))
    Me.lboxDocs.List(0, 1) = FilePath
    Me.lboxDocs.Selected(0) = True
    
    On Error GoTo 0
End Sub

Private Sub btnCombine_Click()
    Dim i As Long
    Dim FileCount As Long
    Dim CombinedName As String
    
    On Error GoTo Handler
    
    ' Make sure only docx, doc and rtf files are selected
    For i = 0 To Me.lboxDocs.ListCount - 1
        If Me.lboxDocs.Selected(i) = True Then
            If Right$(Me.lboxDocs.List(i, 1), Len(Me.lboxDocs.List(i, 1)) - InStrRev(Me.lboxDocs.List(i, 1), ".")) <> "docx" And _
            Right$(Me.lboxDocs.List(i, 1), Len(Me.lboxDocs.List(i, 1)) - InStrRev(Me.lboxDocs.List(i, 1), ".")) <> "doc" And _
            Right$(Me.lboxDocs.List(i, 1), Len(Me.lboxDocs.List(i, 1)) - InStrRev(Me.lboxDocs.List(i, 1), ".")) <> "rtf" Then
                MsgBox "You can only combine .docx, .doc, and .rtf files - please deselect other file formats before proceeding."
                Exit Sub
            End If
            FileCount = FileCount + 1
        End If
    Next i
    
    ' Make sure at least 2 files are selected
    If FileCount < 2 Then
        MsgBox "You must select at least 2 files to combine."
        Exit Sub
    End If
        
    ' Add a new blank document
    Paperless.NewDocument
  
    ' Insert selected files in new pockets
    For i = 0 To Me.lboxDocs.ListCount - 1
        If Me.lboxDocs.Selected(i) = True Then
            Selection.TypeText Left$(Me.lboxDocs.List(i, 0), InStrRev(Me.lboxDocs.List(i, 0), ".") - 1)
            Selection.Style = "Pocket"
            Selection.TypeParagraph
            Selection.InsertFile Me.lboxDocs.List(i, 1)
        End If
    Next i
   
    If Me.cboAutoName.Value <> "" And Me.cboAutoName.Value <> "No rounds found on Tabroom" Then
        CombinedName = Me.cboAutoName.Value
    Else
        CombinedName = "Combined Doc"
    End If
    
    ' Save file
    If GetSetting("Verbatim", "Paperless", "AutoSaveDir", "") <> "" And Me.cboAutoName.Value <> "" Then
        If Right$(GetSetting("Verbatim", "Paperless", "AutoSaveDir", ""), 1) = Application.PathSeparator Then
            ActiveDocument.SaveAs Filename:=GetSetting("Verbatim", "Paperless", "AutoSaveDir") & CombinedName, FileFormat:=wdFormatXMLDocument
        Else
            ActiveDocument.SaveAs Filename:=GetSetting("Verbatim", "Paperless", "AutoSaveDir") & Application.PathSeparator & CombinedName, FileFormat:=wdFormatXMLDocument
        End If
    Else
        With Application.Dialogs.Item(wdDialogFileSaveAs)
            .Name = CombinedName
            .Show
        End With
    End If
    
    Unload Me
    
    Exit Sub
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

