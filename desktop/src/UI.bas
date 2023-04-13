Attribute VB_Name = "UI"
'@IgnoreModule ProcedureNotUsed
Option Explicit

Public Sub ShowForm(ByVal FormName As String)
    Dim Form As Object
    
    On Error Resume Next
    
    Select Case FormName
        Case "Caselist"
            Set Form = New frmCaselist
        Case "CheatSheet"
            Set Form = New frmCheatSheet
        Case "ChooseSpeechDoc"
            Set Form = New frmChooseSpeechDoc
        Case "CombineDocs"
            Set Form = New frmCombineDocs
        Case "Help"
            Set Form = New frmHelp
        Case "Login"
            Set Form = New frmLogin
        Case "Progress"
            Set Form = New frmProgress
        Case "Settings"
            Set Form = New frmSettings
        Case "Setup"
            Set Form = New frmSetupWizard
        Case "Share"
            Set Form = New frmShare
        Case "Stats"
            If Globals.InvisibilityToggle = True Then
                MsgBox "Stats form cannot be opened while in Invisibility Mode. Please turn off Invisibility Mode and try again."
                Exit Sub
            End If
            Set Form = New frmStats
        Case "QuickCards"
            Set Form = New frmQuickCards
        Case "Troubleshooter"
            Set Form = New frmTroubleshooter
        Case "Tutorial"
            Set Form = New frmTutorial
        Case Else
            ' Do nothing
            Exit Sub
    End Select

    Form.Show
    
    On Error GoTo 0
End Sub

' Functions for assigning keyboard shortcuts to forms
Public Sub ShowFormHelp()
    UI.ShowForm "Help"
End Sub

Public Sub ShowFormSettings()
    UI.ShowForm "Settings"
End Sub

Public Sub ShowFormShare()
    UI.ShowForm "Share"
End Sub

Public Sub ShowFormStats()
    UI.ShowForm "Stats"
End Sub

Public Sub ShowFormCaselist()
    UI.ShowForm "Caselist"
End Sub

Public Sub ShowFormChooseSpeechDoc()
    UI.ShowForm "ChooseSpeechDoc"
End Sub

Public Sub LaunchTutorial()
    Dim d As Document
    Dim TutorialDoc As String
    
    ' If more than one non-empty doc is open, prompt to close
    If Documents.Count > 1 Or ActiveDocument.Words.Count > 1 Then
        If MsgBox("Tutorial can only be run while a single blank document is open. Open a new blank doc and close everything else?", vbYesNo) = vbYes Then
            TutorialDoc = Documents.Add(ActiveDocument.AttachedTemplate.FullName).Name
            
            For Each d In Documents
                '@Ignore MemberNotOnInterface
                If d.Name <> TutorialDoc Then d.Close wdPromptToSaveChanges
            Next d
        Else
            Exit Sub
        End If
    End If
    
    ' Make sure debate tab is active on ribbon
    #If Mac Then
    #Else
        WordBasic.SendKeys "%d%"
    #End If
    
    UI.ShowForm "Tutorial"
End Sub

'@Ignore ProcedureCanBeWrittenAsFunction
Public Sub ResizeUserForm(ByVal frm As Object, Optional ByRef dResizeFactor As Double = 0#)
' From https://peltiertech.com/userforms-for-mac-and-windows/
    Dim ctrl As Object
    Dim sColWidths As String
    Dim vColWidths As Variant
    Dim iCol As Long

    If dResizeFactor = 0 Then dResizeFactor = USER_FORM_RESIZE_FACTOR
    With frm
        .Height = .Height * dResizeFactor
        .Width = .Width * dResizeFactor

        For Each ctrl In frm.Controls
            With ctrl
                .Height = .Height * dResizeFactor
                .Width = .Width * dResizeFactor
                .Left = .Left * dResizeFactor
                .Top = .Top * dResizeFactor
                On Error Resume Next
                .Font.size = .Font.size * dResizeFactor
                On Error GoTo 0

                ' Multi column listboxes, comboboxes
                Select Case TypeName(ctrl)
                    Case "ListBox", "ComboBox"
                        If ctrl.ColumnCount > 1 Then
                            sColWidths = ctrl.ColumnWidths
                            vColWidths = Split(sColWidths, ";")
                            For iCol = LBound(vColWidths) To UBound(vColWidths)
                                vColWidths(iCol) = Val(vColWidths(iCol)) * dResizeFactor
                            Next
                            sColWidths = Join(vColWidths, ";")
                            ctrl.ColumnWidths = sColWidths
                        End If
                End Select
            End With
        Next
    End With
End Sub

Public Sub PopulateComboBoxFromJSON(ByRef URL As String, ByRef DisplayKey As Variant, ByRef ValueKey As Variant, ByVal c As Object)
    On Error GoTo Handler

    System.Cursor = wdCursorWait

    Dim Response As Dictionary
    Set Response = HTTP.GetReq(URL)

    Dim ArrayItem As Variant
    For Each ArrayItem In Response.Item("body")
        c.AddItem
        c.List(c.ListCount - 1, 0) = ArrayItem.Item(DisplayKey)
        c.List(c.ListCount - 1, 1) = ArrayItem.Item(ValueKey)
    Next ArrayItem

    System.Cursor = wdCursorNormal
    Set Response = Nothing
    Exit Sub

Handler:
    Set Response = Nothing
    System.Cursor = wdCursorNormal
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Public Function GetFileFromDialog(ByVal FilterName As String, ByVal FilterExtension As String, ByVal Title As String, ByVal ButtonText As String) As String
    On Error GoTo Handler

    #If Mac Then
        GetFileFromDialog = AppleScriptTask("Verbatim.scpt", "GetFileFromDialog", "")
    #Else
        ' Show the built-in file picker, only allow picking 1 file at a time
        Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
        Application.FileDialog(msoFileDialogOpen).Filters.Clear
        Application.FileDialog(msoFileDialogOpen).Filters.Add FilterName, FilterExtension
        Application.FileDialog(msoFileDialogOpen).Title = Title
        Application.FileDialog(msoFileDialogOpen).ButtonName = ButtonText
        If Application.FileDialog(msoFileDialogOpen).Show = 0 Then ' Error trap cancel button
            UI.ResetFileDialog msoFileDialogOpen
            Exit Function
        End If
        
        ' Return the first selected filename
        GetFileFromDialog = Application.FileDialog(msoFileDialogOpen).SelectedItems.Item(1)
        
        ' Reset the dialog
        UI.ResetFileDialog msoFileDialogOpen
    #End If
    Exit Function

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Function

Public Function GetFolderFromDialog(ByVal Title As String, ByVal ButtonName As String) As String
    On Error GoTo Handler
    
    #If Mac Then
        GetFolderFromDialog = AppleScriptTask("Verbatim.scpt", "GetFolderFromDialog", "")
    #Else
        ' Show the built-in folder picker, only allow picking 1 folder at a time
        Application.FileDialog(msoFileDialogFolderPicker).AllowMultiSelect = False
        Application.FileDialog(msoFileDialogFolderPicker).Title = Title
        Application.FileDialog(msoFileDialogFolderPicker).ButtonName = ButtonName
        If Application.FileDialog(msoFileDialogFolderPicker).Show = 0 Then 'Error trap cancel button
            UI.ResetFileDialog msoFileDialogFolderPicker
            Exit Function
        End If
        
        ' Return the first selected folder
        GetFolderFromDialog = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems.Item(1)
        
        ' Reset the dialog
        UI.ResetFileDialog msoFileDialogFolderPicker
    #End If
    
    Exit Function

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Function

Public Sub ResetFileDialog(ByVal FD As Byte)
    ' Resets a built-in FileDialog - can pass in a Word constant, also works for folder dialog
    Application.FileDialog(FD).AllowMultiSelect = False
    Application.FileDialog(FD).Filters.Clear
    Application.FileDialog(FD).Title = ""
    Application.FileDialog(FD).ButtonName = ""
    Application.FileDialog(FD).InitialFileName = ""
End Sub
