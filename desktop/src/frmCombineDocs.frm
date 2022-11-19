VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCombineDocs 
   Caption         =   "Combine Docs"
   ClientHeight    =   5505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8160
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
           
    'Turn on error checking
    On Error GoTo Handler
    
    'Add all recent files to the box
    For Each rFile In Application.RecentFiles
        Me.lboxDocs.AddItem
        Me.lboxDocs.List(Me.lboxDocs.ListCount - 1, 0) = rFile.Name
        Me.lboxDocs.List(Me.lboxDocs.ListCount - 1, 1) = rFile.Path & "\" & rFile.Name
    Next rFile
       
    'Reset AutoName box and add a blank item
    Me.cboAutoName.Clear
    Me.cboAutoName.AddItem
    Me.cboAutoName.List(Me.cboAutoName.ListCount - 1) = ""
    
    'Get rounds from Tabroom
    Caselist.PopulateComboBox Globals.MOCK_ROUNDS, "tournament", Me.cboAutoName
    
    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Sub btnCombine_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnCombine.BackColor = Globals.SUBMIT_BUTTON_HOVER
End Sub
Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    btnCombine.BackColor = Globals.SUBMIT_BUTTON_NORMAL
End Sub

Private Sub btnManualAdd_Click()
    
    Dim FilePath As String
    
    On Error Resume Next
    
    'Show the built-in file picker, only allow picking 1 file at a time
    Application.FileDialog(msoFileDialogFilePicker).AllowMultiSelect = False
    If Application.FileDialog(msoFileDialogFilePicker).Show = 0 Then 'Error trap cancel button
        Exit Sub
    End If
    
    'Add selected file to the box
    FilePath = Application.FileDialog(msoFileDialogFilePicker).SelectedItems(1)
    Me.lboxDocs.AddItem , 0
    Me.lboxDocs.List(0, 0) = Right(FilePath, Len(FilePath) - InStrRev(FilePath, "\"))
    Me.lboxDocs.List(0, 1) = FilePath
    Me.lboxDocs.Selected(0) = True
        
End Sub

Private Sub btnCombine_Click()

    Dim i As Integer
    Dim FileCount As Integer
    
    On Error GoTo Handler
    
    'Make sure only docx, doc and rtf files are selected
    For i = 0 To Me.lboxDocs.ListCount - 1
        If Me.lboxDocs.Selected(i) = True Then
            If Right(Me.lboxDocs.List(i, 1), Len(Me.lboxDocs.List(i, 1)) - InStrRev(Me.lboxDocs.List(i, 1), ".")) <> "docx" And _
            Right(Me.lboxDocs.List(i, 1), Len(Me.lboxDocs.List(i, 1)) - InStrRev(Me.lboxDocs.List(i, 1), ".")) <> "doc" And _
            Right(Me.lboxDocs.List(i, 1), Len(Me.lboxDocs.List(i, 1)) - InStrRev(Me.lboxDocs.List(i, 1), ".")) <> "rtf" Then
                MsgBox "You can only combine .docx, .doc, and .rtf files - please deselect other file formats before proceeding."
                Exit Sub
            End If
            FileCount = FileCount + 1
        End If
    Next i
    
    'Make sure at least 2 files are selected
    If FileCount < 2 Then
        MsgBox "You must select at least 2 files to combine."
        Exit Sub
    End If
        
    'Add a new blank document
    Paperless.NewDocument
  
    'Insert selected files in new pockets
    For i = 0 To Me.lboxDocs.ListCount - 1
        If Me.lboxDocs.Selected(i) = True Then
            Selection.TypeText Left(Me.lboxDocs.List(i, 0), InStrRev(Me.lboxDocs.List(i, 0), ".") - 1)
            Selection.Style = "Pocket"
            Selection.TypeParagraph
            Selection.InsertFile Me.lboxDocs.List(i, 1)
        End If
    Next i
   
    'Save file
    If GetSetting("Verbatim", "Paperless", "AutoSaveDir") <> "" And Me.cboAutoName.Value <> "" Then
        If Right(GetSetting("Verbatim", "Paperless", "AutoSaveDir"), 1) = "\" Then
            ActiveDocument.SaveAs FileName:=GetSetting("Verbatim", "Paperless", "AutoSaveDir") & Me.cboAutoName.Value, FileFormat:=wdFormatXMLDocument
        Else
            ActiveDocument.SaveAs FileName:=GetSetting("Verbatim", "Paperless", "AutoSaveDir") & "\" & Me.cboAutoName.Value, FileFormat:=wdFormatXMLDocument
        End If
    Else
        With Application.Dialogs(wdDialogFileSaveAs)
            If Me.cboAutoName.Value <> "" Then
                .Name = Me.cboAutoName.Value
            Else
                .Name = "Combined Doc"
            End If
            .Show
        End With
    End If
    
    Unload Me
    
    Exit Sub
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub
