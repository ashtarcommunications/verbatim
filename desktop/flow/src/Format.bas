Attribute VB_Name = "Format"
Option Explicit

Public Sub FormatFlow(ByVal Side As String)
    Dim SpeechNames() As String
    Dim AffCount As Long
    Dim NegCount As Long
    Dim OddColor As Long
    Dim EvenColor As Long
    Dim s As Variant
    Dim i As Long

    Globals.InitializeGlobals

    With ActiveSheet
        .Cells.Font.size = GetSetting("Verbatim", "Flow", "FontSize", 8)
        .Cells.RowHeight = GetSetting("Verbatim", "Flow", "RowHeight", 12)
        .Cells.ColumnWidth = GetSetting("Verbatim", "Flow", "ColumnWidth", 36)
        .Cells.WrapText = True
        .Cells.EntireRow.AutoFit
    End With
    
    AffCount = 0
    NegCount = 0
    For Each s In ActiveWorkbook.Sheets
        If InStr(LCase$(s.Name), "oncase") Or InStr(LCase$(s.Name), "on case") Or InStr(LCase$(s.Name), "aff") Or InStr(LCase$(s.Name), "pro") = True Then
            AffCount = AffCount + 1
        End If
        
        If InStr(LCase$(s.Name), "offcase") Or InStr(LCase$(s.Name), "off case") Or InStr(LCase$(s.Name), "neg") Or InStr(LCase$(s.Name), "con") = True Then
            NegCount = NegCount + 1
        End If
    Next s
    
    If Side = "Neg" Then
        ActiveSheet.Cells(2, 1).Value = "OffCase " & NegCount + 1
    Else
        ActiveSheet.Cells(2, 1).Value = "OnCase " & AffCount + 1
    End If
        
    SpeechNames = Split(GetSetting("Verbatim", "Flow", "SpeechNames", "1AC,1NC,2AC,Block,1AR,2NR,2AR"), ",")
    
    If Side = "Neg" Then
        For i = 1 To UBound(SpeechNames)
            SpeechNames(i - 1) = SpeechNames(i)
        Next i
        ReDim Preserve SpeechNames(UBound(SpeechNames) - 1)
    End If
        
    If Side = "Neg" Then
        OddColor = Globals.RED
        EvenColor = Globals.BLUE
        ActiveSheet.Tab.Color = Globals.RED
    Else
        OddColor = Globals.BLUE
        EvenColor = Globals.RED
        ActiveSheet.Tab.Color = Globals.BLUE
    End If
        
    i = 1
    For Each s In SpeechNames
        ActiveSheet.Cells(1, i).Value = Trim$(s)
        ActiveSheet.Cells(1, i).Font.size = 12
        ActiveSheet.Cells(1, i).Font.Bold = True
        ActiveSheet.Cells(1, i).HorizontalAlignment = xlCenter
        ActiveSheet.Cells(1, i).Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
        ActiveSheet.Cells(1, i).RowHeight = 28
        If i Mod 2 = 0 Then
            ActiveSheet.Columns(i).Font.Color = EvenColor
            ActiveSheet.Cells(1, i).Borders(xlEdgeBottom).Color = EvenColor
        Else
            ActiveSheet.Columns(i).Font.Color = OddColor
            ActiveSheet.Cells(1, i).Borders(xlEdgeBottom).Color = OddColor
        End If
        i = i + 1
    Next s
    
    If GetSetting("Verbatim", "Flow", "FreezeSpeechNames", True) = True Then
        ActiveWindow.SplitRow = 1
        ActiveWindow.FreezePanes = True
    End If

    ' Put cursor in A2 for sheet naming
    ActiveSheet.Cells(2, 1).Select
End Sub

Public Sub AddFlowAff()
    ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Worksheets.[_Default](ActiveWorkbook.Worksheets.Count), Count:=1
    Format.FormatFlow "Aff"
End Sub

Public Sub AddFlowNeg()
    ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Worksheets.[_Default](ActiveWorkbook.Worksheets.Count), Count:=1
    Format.FormatFlow "Neg"
End Sub

Public Sub AddFlowCX()
    Dim CXCount As Long
    Dim s As Worksheet
    
    Application.DisplayAlerts = False
       
    Globals.InitializeGlobals
       
    ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Worksheets.[_Default](1), Count:=1
    
    ' Label the sheet
    CXCount = 0
    For Each s In ActiveWorkbook.Sheets
        If InStr(LCase$(s.Name), "cx") > 0 Then
            CXCount = CXCount + 1
        End If
    Next s
    
    If CXCount > 0 Then
        ActiveSheet.Name = "CX" & CXCount
    Else
        ActiveSheet.Name = "CX"
    End If
    
    ActiveSheet.Tab.Color = Globals.GREEN

    ' Setup the sheet
    With ActiveSheet
        .Cells.Font.size = GetSetting("Verbatim", "Flow", "FontSize", 8)
        .Cells.RowHeight = GetSetting("Verbatim", "Flow", "RowHeight", 12)
        .Cells.ColumnWidth = GetSetting("Verbatim", "Flow", "ColumnWidth", 36)
        .Cells.WrapText = True
        .Cells.EntireRow.AutoFit
    End With
            
    ' Alternate speaker colors
    ActiveSheet.Columns(1).Font.Color = Globals.RED
    ActiveSheet.Columns(2).Font.Color = Globals.BLUE
    ActiveSheet.Columns(3).Font.Color = Globals.BLUE
    ActiveSheet.Columns(4).Font.Color = Globals.RED
    ActiveSheet.Columns(5).Font.Color = Globals.RED
    ActiveSheet.Columns(6).Font.Color = Globals.BLUE
    ActiveSheet.Columns(7).Font.Color = Globals.BLUE
    ActiveSheet.Columns(8).Font.Color = Globals.RED
    ActiveSheet.Cells(1, 1).Borders(xlEdgeBottom).Color = Globals.RED
    ActiveSheet.Cells(1, 2).Borders(xlEdgeBottom).Color = Globals.BLUE
    ActiveSheet.Cells(1, 3).Borders(xlEdgeBottom).Color = Globals.BLUE
    ActiveSheet.Cells(1, 4).Borders(xlEdgeBottom).Color = Globals.RED
    ActiveSheet.Cells(1, 5).Borders(xlEdgeBottom).Color = Globals.RED
    ActiveSheet.Cells(1, 6).Borders(xlEdgeBottom).Color = Globals.BLUE
    ActiveSheet.Cells(1, 7).Borders(xlEdgeBottom).Color = Globals.BLUE
    ActiveSheet.Cells(1, 8).Borders(xlEdgeBottom).Color = Globals.RED
        
    ' Set CX names
    If Left$(GetSetting("Verbatim", "Flow", "SpeechNames", "1AC,1NC,2AC,Block,1AR,2NC,2AR"), "3") = "1AC" Then
        ActiveSheet.Range("A1:B1").Value = "1AC CX"
        ActiveSheet.Range("C1:D1").Value = "1NC CX"
        ActiveSheet.Range("E1:F1").Value = "2AC CX"
        ActiveSheet.Range("G1:H1").Value = "2NC CX"
    Else
        ActiveSheet.Range("A1:B1").Value = "CX #1"
        ActiveSheet.Range("C1:D1").Value = "CX #2"
        ActiveSheet.Range("E1:F1").Value = "CX #3"
        ActiveSheet.Range("G1:H1").Value = "CX #4"
    End If
    
    ' Style headers
    ActiveSheet.Range("A1:H1").HorizontalAlignment = xlCenter
    ActiveSheet.Range("A1:H1").Font.size = 12
    ActiveSheet.Range("A1:H1").Font.Bold = True
    ActiveSheet.Range("A1:H1").Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
    ActiveSheet.Range("A1:H1").Borders(xlEdgeBottom).Weight = xlMedium
    ActiveSheet.Range("A1:H1").RowHeight = 28
    
    ' Merge the header cells
    ActiveSheet.Range("A1:B1").Merge Across:=True
    ActiveSheet.Range("C1:D1").Merge Across:=True
    ActiveSheet.Range("E1:F1").Merge Across:=True
    ActiveSheet.Range("G1:H1").Merge Across:=True
    
    ' Label each speaker
    ActiveSheet.Range("A2").Value = "Question"
    ActiveSheet.Range("B2").Value = "Response"
    ActiveSheet.Range("C2").Value = "Question"
    ActiveSheet.Range("D2").Value = "Response"
    ActiveSheet.Range("E2").Value = "Question"
    ActiveSheet.Range("F2").Value = "Response"
    ActiveSheet.Range("G2").Value = "Question"
    ActiveSheet.Range("H2").Value = "Response"
    ActiveSheet.Range("A2:H2").HorizontalAlignment = xlCenter
    
    ' Add a border between each CX
    ActiveSheet.Columns(2).Borders(xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
    ActiveSheet.Columns(2).Borders(xlEdgeRight).Weight = xlMedium
    ActiveSheet.Columns(4).Borders(xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
    ActiveSheet.Columns(4).Borders(xlEdgeRight).Weight = xlMedium
    ActiveSheet.Columns(6).Borders(xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
    ActiveSheet.Columns(6).Borders(xlEdgeRight).Weight = xlMedium
    ActiveSheet.Columns(8).Borders(xlEdgeRight).LineStyle = XlLineStyle.xlContinuous
    ActiveSheet.Columns(8).Borders(xlEdgeRight).Weight = xlMedium
        
    ' Freeze panes
    If GetSetting("Verbatim", "Flow", "FreezeSpeechNames", True) = True Then
        ActiveWindow.SplitRow = 1
        ActiveWindow.FreezePanes = True
    End If

    ' Select the first content cell
    ActiveSheet.Cells(3, 1).Select
    
    Application.DisplayAlerts = True
End Sub

Public Sub DeleteEmptyFlows()
    Dim ws As Worksheet
    
    On Error Resume Next
    
    If MsgBox("This will delete all empty sheets in the document, and is irreversible. Are you sure?", vbOKCancel) = vbCancel Then Exit Sub
    
    Application.DisplayAlerts = False
    For Each ws In ActiveWorkbook.Sheets
        If ws.Index <> 1 And Application.WorksheetFunction.CountA(ws.Range("A2:XFD1048576")) < 2 Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
    
    Set ws = Nothing
    
    On Error GoTo 0
End Sub

Public Sub DeleteFlow()
    On Error Resume Next
    Application.DisplayAlerts = False
    If ActiveWorkbook.Sheets.Count = 1 Then
        MsgBox "This is the only sheet, you can't delete it!"
        Exit Sub
    End If
    
    If Application.WorksheetFunction.CountA(ActiveSheet.Range("A2:XFD1048576")) > 1 Then
        If MsgBox("This sheet has content, are you sure you want to delete it?", vbOKCancel) = vbCancel Then Exit Sub
    End If
    
    ActiveSheet.Delete
    Application.DisplayAlerts = True
    
    On Error GoTo 0
End Sub

Public Sub AutoScoutingInfo()
    Dim Info As Boolean
    Dim i As Long
    
    Dim Aff As String
    Dim Neg As String

    For i = 1 To ActiveWorkbook.Worksheets.Count
        If ActiveWorkbook.Worksheets.[_Default](i).Name = "Info" Then
            Info = True
        End If
    Next i

    If Info = False Then
        MsgBox "It looks like the Info sheet has been deleted - aborting!"
        Exit Sub
    End If

    ' Find Aff or Neg sheets by color - unlikely to have been changed manually, and already set by side on sheet creation
    For i = 1 To ActiveWorkbook.Worksheets.Count
        If ActiveWorkbook.Worksheets.[_Default](i).Tab.Color = Globals.BLUE Then
            Aff = Aff & ActiveWorkbook.Worksheets.[_Default](i).Name & vbCrLf
        ElseIf ActiveWorkbook.Worksheets.[_Default](i).Tab.Color = Globals.RED Then
            Neg = Neg & ActiveWorkbook.Worksheets.[_Default](i).Name & vbCrLf
        End If
    Next i
    
    If ActiveWorkbook.Worksheets.[_Default]("Info").Range("B8").Value <> "" Or ActiveWorkbook.Worksheets.[_Default]("Info").Range("B9").Value <> "" Then
        If MsgBox("There is already content in the scouting sheet, overwrite?", vbOKCancel) = vbCancel Then Exit Sub
    End If
    
    ActiveWorkbook.Worksheets.[_Default]("Info").Range("B8").Value = Aff
    ActiveWorkbook.Worksheets.[_Default]("Info").Range("B9").Value = Neg
End Sub
