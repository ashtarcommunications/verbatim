Attribute VB_Name = "Flow"
Option Explicit

Public Sub SendToFlowCell()
    Flow.SendToFlow SplitParagraphs:=False
End Sub

Public Sub SendToFlowColumn()
    Flow.SendToFlow SplitParagraphs:=True
End Sub

Public Sub SendHeadingsToFlowCell()
    Flow.SendToFlow SplitParagraphs:=False, HeadingsOnly:=True
End Sub

Public Sub SendHeadingsToFlowColumn()
    Flow.SendToFlow SplitParagraphs:=True, HeadingsOnly:=True
End Sub

Public Sub SendToFlow(Optional ByVal SplitParagraphs As Boolean, Optional ByVal HeadingsOnly As Boolean)
    Dim ExcelApp As Object
    Dim w As Variant
    Dim Flow As Object
    Dim p As Paragraph
    Dim Cite As String
    Dim r As Range
    Dim i As Long
    Dim Overwrite As Boolean
    Dim SendText As String
    
    On Error GoTo Handler
    
    Set ExcelApp = GetObject(, "Excel.Application")
    
    If ExcelApp Is Nothing Then
        MsgBox "Excel must be open to send to your flow!"
        Exit Sub
    End If
    
    For Each w In ExcelApp.Workbooks
        If InStr(LCase$(w.Name), "flow") Then Set Flow = w
    Next w
    
    If Flow Is Nothing Then
        MsgBox "You must have an Excel document open with ""Flow"" in the name to send to it!"
        Exit Sub
    End If
    
    If Flow.ActiveSheet Is Nothing Then
        MsgBox "You must have an active sheet in your Flow to send to it!"
        Exit Sub
    End If
    
    If Selection.End = Selection.Start Then
        Paperless.SelectHeadingAndContent
    End If
    
    ' Make sure the Flow is the active sheet or cell selection will fail
    Flow.Activate
    Flow.ActiveSheet.Select
    
    If SplitParagraphs Then
        i = 0
        
        ' Check for overwriting existing content
        For Each p In Selection.Paragraphs
            If Flow.ActiveSheet.Cells(Flow.ActiveSheet.Application.ActiveCell.Offset(i, 0).Row, Flow.ActiveSheet.Application.ActiveCell.Column).Value <> "" Then
                Overwrite = True
            End If
            i = i + 1
        Next p
        
        If Overwrite = True Then
            If MsgBox("There's already text where you're sending.  Are you sure you want to overwrite it?", vbOKCancel) = vbCancel Then Exit Sub
        End If
        
        i = 0

        ' Copy each paragraph to a separate cell
        For Each p In Selection.Paragraphs
            ' In HeadingsOnly mode, just send headings and cites
            If HeadingsOnly <> True Or p.OutlineLevel <> wdOutlineLevelBodyText Then
                ' Add extra space for headings
                If p.OutlineLevel <> wdOutlineLevelBodyText And i > 0 Then
                    Flow.ActiveSheet.Application.ActiveCell.Offset(1, 0).Select
                End If
                Flow.ActiveSheet.Cells(Flow.ActiveSheet.Application.ActiveCell.Row, Flow.ActiveSheet.Application.ActiveCell.Column).Value = p.Range.Text
                Flow.ActiveSheet.Application.ActiveCell.Offset(1, 0).Select
            ElseIf Paperless.IdentifyCiteStyle(p) = True Then
                Cite = ""
                Set r = p.Range
                With r.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .Text = ""
                    .Forward = True
                    .Wrap = wdFindStop
                    .Format = True
                    .Style = "Cite"
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                    
                    Do While .Execute(Forward:=True) = True And r.InRange(p.Range)
                        Cite = Cite & r.Text & " "
                    Loop

                    .ClearFormatting
                    .Replacement.ClearFormatting
                End With
                Set r = Nothing

                Flow.ActiveSheet.Cells(Flow.ActiveSheet.Application.ActiveCell.Row, Flow.ActiveSheet.Application.ActiveCell.Column).Value = Trim$(Cite)
                Flow.ActiveSheet.Application.ActiveCell.Offset(1, 0).Select
            End If
            i = i + 1
        Next p
    Else
        ' Copy selected content into the current cell
        If Flow.ActiveSheet.Cells(Flow.ActiveSheet.Application.ActiveCell.Row, Flow.ActiveSheet.Application.ActiveCell.Column).Value <> "" Then
            If MsgBox("There is already text where you're sending.  Are you sure you want to overwrite it?", vbOKCancel) = vbCancel Then Exit Sub
        End If
        
        i = 0

        For Each p In Selection.Paragraphs
            If HeadingsOnly <> True Or p.OutlineLevel <> wdOutlineLevelBodyText Then
                ' Add extra space for headings
                If p.OutlineLevel <> wdOutlineLevelBodyText And i > 1 Then
                    #If Mac Then
                        SendText = SendText & Chr(13)
                    #Else
                        SendText = SendText & Chr$(10)
                    #End If
                End If

                ' Use correct linebreak for each OS
                #If Mac Then
                    SendText = SendText & p.Range.Text & Chr(13)
                #Else
                    SendText = SendText & p.Range.Text & Chr$(10)
                #End If
            ElseIf Paperless.IdentifyCiteStyle(p) = True Then
                Cite = ""
                Set r = p.Range
                With r.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .Text = ""
                    .Forward = True
                    .Wrap = wdFindStop
                    .Format = True
                    .Style = "Cite"
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                    
                    Do While .Execute(Forward:=True) = True And r.InRange(p.Range)
                        Cite = Cite & r.Text & " "
                    Loop

                    .ClearFormatting
                    .Replacement.ClearFormatting
                End With
                Set r = Nothing
                
                #If Mac Then
                    SendText = SendText & Trim(Cite) & Chr(13)
                #Else
                    SendText = SendText & Trim$(Cite) & Chr$(10)
                #End If
            End If
            i = i + 1
        Next p

        Flow.ActiveSheet.Cells(Flow.ActiveSheet.Application.ActiveCell.Row, Flow.ActiveSheet.Application.ActiveCell.Column).Value = SendText
        Flow.ActiveSheet.Application.ActiveCell.Offset(1, 0).Select
    End If
    
    Set ExcelApp = Nothing
    Set w = Nothing
    Set Flow = Nothing
    
    Exit Sub
    
Handler:
    Set ExcelApp = Nothing
    Set w = Nothing
    Set Flow = Nothing
    If Err.Number = 429 Then
        #If Mac Then
            MsgBox "Excel must be open to send to your flow! If you're on a Mac, sending to Excel is very unreliable, and may not work on your computer."
        #Else
            MsgBox "Excel must be open to send to your flow!"
        #End If
    Else
        MsgBox "Error " & Err.Number & ": " & Err.Description
    End If
End Sub

Public Sub CreateFlow()
    Dim AutoSaveDir As String
    Dim Filename As String
    Dim xl As Object
    Dim w As Object
    
    On Error GoTo Handler
    
    If Filesystem.FileExists(Application.NormalTemplate.Path & Application.PathSeparator & "Debate.xltm") = False Then
        MsgBox "Verbatim Flow is not installed - ensure Debate.xltm is in your Templates folder."
        Exit Sub
    End If
    
    Filename = InputBox("Name for your new flow?", "New Flow", Split(ActiveDocument.Name, ".")(0) & " Flow")
    If Filename = "" Then Exit Sub
    
    AutoSaveDir = GetSetting("Verbatim", "Paperless", "AutoSaveDir", CurDir$())
    If AutoSaveDir = "" Then AutoSaveDir = CurDir$()
    If Right$(AutoSaveDir, 1) <> Application.PathSeparator Then AutoSaveDir = AutoSaveDir & Application.PathSeparator
    Filename = AutoSaveDir & Filename
    If Filename = "" Then Filename = "Flow"
    
    #If Mac Then
        Dim xlwb As Object
        Set xlwb = CreateObject("Excel.Application")
        Set xl = xlwb.Application
    #Else
        Set xl = CreateObject("Excel.Application")
    #End If
        
    xl.Visible = True
    Set w = xl.Workbooks.Add(Template:=Application.NormalTemplate.Path & Application.PathSeparator & "Debate.xltm")
    w.SaveAs Filename:=Filename, FileFormat:=52  ' 52 = Macro-enabled workbook
    
    Set xl = Nothing
    Set w = Nothing
    #If Mac Then
        Set xlwb = Nothing
    #End If
    Exit Sub

Handler:
    Set xl = Nothing
    Set w = Nothing
    #If Mac Then
        Set xlwb = Nothing
    #End If
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub
