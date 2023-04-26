Attribute VB_Name = "Flow"
Option Explicit

'@Ignore ProcedureNotUsed, ParameterNotUsed
Public Sub ToggleInsertMode(ByVal c As IRibbonControl, ByVal pressed As Boolean)
' Ribbon callback for onAction of InsertMode togglebutton
    ' If ToggleButton is turned on
    If pressed Then
        Globals.InsertModeToggle = True
        Flow.EnableInsertMode
        MsgBox "Insert Mode is turned ON. The Enter key will now insert line breaks in the active cell. Use the ESC key to exit editing mode on a cell, " _
            & "and Alt + Enter to move down a cell. Click the button again to turn Insert Mode off."
        Application.StatusBar = "Insert Mode on - press button on ribbon to cancel."
    Else
        Globals.InsertModeToggle = False
        Flow.DisableInsertMode
        MsgBox "Insert Mode is turned OFF. The Enter key will now behave normally, and Alt + Enter will insert linebreaks while editing a cell."
        Application.StatusBar = "Insert Mode off."
    End If
End Sub

Public Sub EnableInsertMode()
    Application.OnKey "~", "Flow.InsertLineBreak" ' The tilde "~" means the Enter key
    Application.OnKey "%~", "Flow.MoveCellDown"
End Sub

Public Sub DisableInsertMode()
    ' Passing nothing as the second parameter clears the key mappings
    Application.OnKey "~"
    Application.OnKey "%~"
End Sub

'@Ignore ProcedureNotUsed
Public Sub InsertLineBreak()
    ActiveCell.Value = ActiveCell.Value & vbCrLf
    Application.SendKeys "{F2}" ' Keeps the cursor at the end of the cell
End Sub

Public Sub EnterCell()
    Application.SendKeys "{F2}"
End Sub

'@Ignore ProcedureNotUsed
Public Sub MoveCellDown()
    ActiveCell.Offset(1).Select
End Sub

Public Sub InsertRowAbove()
    Dim r As Range
    Set r = ActiveCell.EntireRow
    r.Insert Shift:=xlDown
    r.Offset(-1).ClearContents
    Set r = Nothing
End Sub

Public Sub InsertRowBelow()
    Dim r As Range
    Set r = ActiveCell.Offset(1, 0).EntireRow
    r.Insert Shift:=xlDown
    r.Offset(-1).ClearContents
    Set r = Nothing
End Sub

Public Sub InsertCellAbove()
    Dim r As Range
    Set r = ActiveCell
    r.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    r.Offset(-1, 0).ClearContents
    Set r = Nothing
End Sub

Public Sub InsertCellBelow()
    Dim r As Range
    Set r = ActiveCell.Offset(1, 0)
    r.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    r.Offset(-1, 0).ClearContents
    Set r = Nothing
End Sub

Public Sub DeleteRow()
    ActiveCell.EntireRow.Delete
End Sub

Public Sub MergeCells()
    Application.DisplayAlerts = False
    
    Dim c As Range
    Dim Value As String

    If Selection.Rows.Count = ActiveSheet.Rows.Count Then
        MsgBox "Can't merge an entire column, please limit your selection."
        Exit Sub
    ElseIf Selection.Columns.Count = ActiveSheet.Columns.Count Then
        MsgBox "Can't merge an entire row, please limit your selection."
        Exit Sub
    End If
    
    For Each c In Selection
        Value = Value & vbCrLf & c.Value
    Next c
    
    Selection.ClearContents
    ActiveCell.Value = Trim$(Value)
    If Left$(Value, 1) = vbLf Or Left$(Value, 1) = vbCrLf Then Value = Right$(Value, Len(Value) - 1)
    ActiveCell.WrapText = True
    ActiveCell.EntireRow.AutoFit
    ActiveCell.Select
    
    Application.DisplayAlerts = True
End Sub

Public Sub MoveUp()
    If Selection.Row < 3 Then Exit Sub
    Selection.Rows(Selection.Rows.Count + 1).Insert Shift:=xlDown
    Selection.Rows(1).Offset(-1).Cut Selection.Rows(Selection.Rows.Count + 1)
    Selection.Rows(1).Offset(-1).Delete Shift:=xlUp
    Selection.Offset(-1).Select
End Sub

Public Sub MoveDown()
    Selection.Rows(1).Insert Shift:=xlDown
    Selection.Rows(Selection.Rows.Count).Offset(2).Cut Selection.Rows(1)
    Selection.Rows(Selection.Rows.Count).Offset(2).Delete Shift:=xlUp
    Selection.Offset(1).Select
End Sub

Public Sub GoToBottom()
    ActiveSheet.Cells(65536, Selection.Column).End(xlUp).Select
End Sub

Public Sub ToggleHighlighting()
    Dim c As Range
    Dim Highlighted As Boolean
    
    ' Make sure any highlighted cell will toggle the entire range
    For Each c In Selection
        If c.Interior.ColorIndex <> xlNone Then Highlighted = True
    Next c

    If Highlighted Then
        Selection.Interior.ColorIndex = xlNone
        Selection.Borders.ColorIndex = xlColorIndexNone
    Else
        Selection.Interior.ColorIndex = 27
        Selection.Borders.ColorIndex = 15
    End If
End Sub

Public Sub ToggleEvidence()
    If ActiveCell.Borders.[_Default](xlEdgeBottom).Weight = xlThick Then
        ActiveCell.Borders.[_Default](xlEdgeBottom).LineStyle = xlLineStyleNone ' Reset to default borders
    Else
        ActiveCell.Borders.[_Default](xlEdgeBottom).LineStyle = xlContinuous
        ActiveCell.Borders.[_Default](xlEdgeBottom).Weight = xlThick
        ActiveCell.Borders.[_Default](xlEdgeBottom).Color = Globals.GREEN
    End If
End Sub

Public Sub ToggleGroup()
    Dim c As Range
    Dim Grouped As Boolean
    
    ' Make sure any grouped cell will toggle the entire range
    For Each c In Selection
        If c.Borders.[_Default](xlEdgeRight).Weight = xlThick Then Grouped = True
    Next c

    If Grouped Then
        Selection.Borders(xlEdgeRight).LineStyle = xlLineStyleNone
    Else
        Selection.Borders(xlEdgeRight).Weight = xlThick
    End If
End Sub

'@Ignore ProcedureNotUsed
Public Sub SwitchSpeech(ByVal c As IRibbonControl)
    Dim s As Variant
    Dim col As Long
    Dim i As Long
    
    For Each s In ActiveWorkbook.Sheets
        ' Find the matching speech column
        col = 0
        For i = 1 To s.UsedRange.Columns.Count
            If StrComp(LCase$(s.Cells(1, i).Value), LCase$(c.Tag), vbTextCompare) = 0 Then
                col = i
                Exit For
            End If
        Next i
        
        If col > 0 Then
            s.Select
            s.Cells(2, col).Select
        End If
    Next s
    
    ' Switch back to first non-CX tab
    If ActiveWorkbook.Sheets.Count > 1 Then
        If ActiveWorkbook.Sheets.[_Default](2).Name = "CX" And ActiveWorkbook.Sheets.Count > 2 Then
            ActiveWorkbook.Sheets.[_Default](3).Activate
        Else
            ActiveWorkbook.Sheets.[_Default](2).Activate
        End If
    End If
End Sub

Public Sub PasteAsText()
    On Error Resume Next
    ActiveSheet.PasteSpecial Format:="Text", Link:=False, DisplayAsIcon:=False
    On Error GoTo 0
End Sub

Public Sub ExtendArgument()
    Dim c As Range
    Dim Overwrite As Boolean
    
    For Each c In Selection
        If ActiveSheet.Cells(c.Row, c.Column + 2).Value <> "" Then
            Overwrite = True
            Exit For
        End If
    Next c
    
    If Overwrite = True Then
        If MsgBox("There's already content in the destination, overwrite?", vbYesNo) = vbNo Then Exit Sub
    End If
    
    For Each c In Selection
        If GetSetting("Verbatim", "Flow", "ExtendWithArrow", False) = True Then
            ActiveSheet.Cells(c.Row, c.Column).Value = ActiveSheet.Cells(c.Row, c.Column).Value & vbCrLf & ChrW$(8594) & ChrW$(8594) & ChrW$(8594)
            ActiveSheet.Cells(c.Row, c.Column + 2).Value = ChrW$(8594) & ChrW$(8594) & ChrW$(8594)
        Else
            ActiveSheet.Cells(c.Row, c.Column + 2).Value = c.Value
        End If
    Next c
End Sub


