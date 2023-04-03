Attribute VB_Name = "Speech"
Option Explicit

'@Ignore ProcedureNotUsed
Public Sub SendToSpeechCursor()
    Speech.SendToSpeech PasteAtEnd:=False
End Sub

Public Sub SendToSpeechEnd()
    Speech.SendToSpeech PasteAtEnd:=True
End Sub

Public Sub SendToSpeech(Optional ByVal PasteAtEnd As Boolean)
    Dim wd As Object
    Dim d As Object
    Dim SpeechDoc As Object
    Dim c As Range
    Dim SelectedText As String

    On Error GoTo Handler

    Set wd = GetObject(, "Word.Application")
    
    If wd Is Nothing Then
        MsgBox "Word must be open first!", vbOKOnly
        Exit Sub
    End If
    
    For Each d In wd.Documents
        If InStr(LCase$(d.Name), "speech") Then
            Set SpeechDoc = d
        End If
    Next d
    
    If SpeechDoc Is Nothing Then
        MsgBox "You must have a Word document open with ""Speech"" in the name to send to it.", vbOKOnly
        Exit Sub
    End If
    
    ' Assemble a single string to send to Word
    For Each c In Selection.Cells
        SelectedText = SelectedText & c.Value & vbCrLf
    Next c
    
    If PasteAtEnd = True Then
        SpeechDoc.Content.InsertAfter Text:=vbCrLf & SelectedText
    Else
        ' Check for sending to middle of text
        If SpeechDoc.ActiveWindow.Selection.Start <> SpeechDoc.ActiveWindow.Selection.Paragraphs.Item(1).Range.Start Then
            If MsgBox("Sending to the middle of text. Are you sure?", vbOKCancel) = vbCancel Then Exit Sub
        End If
        If SpeechDoc.ActiveWindow.Selection.Paragraphs.OutlineLevel <> 10 Then
            If MsgBox("Sending into a pocket/block/hat/tag. Are you sure?", vbOKCancel) = vbCancel Then Exit Sub
        End If
        SpeechDoc.ActiveWindow.Selection.InsertAfter Text:=SelectedText
    End If
    
    Set wd = Nothing
    Set d = Nothing
    
    Exit Sub

Handler:
    Set wd = Nothing
    Set d = Nothing
    If Err.Number = 429 Then
        MsgBox "Word must be open first!"
    Else
        MsgBox "Error " & Err.Number & ": " & Err.Description
    End If
End Sub
