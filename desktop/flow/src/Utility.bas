Attribute VB_Name = "Utility"
Option Explicit

Public Function FirstEmptyColumn(ByVal ws As Worksheet) As Long
    Dim c As Long
    
    For c = 1 To ws.Columns.Count
        If ws.Cells.Item(1, c).Value = "" Then
            FirstEmptyColumn = c
            Exit Function
        End If
    Next c
    
    ' If no empty column is found, return 0
    FirstEmptyColumn = 0
End Function

