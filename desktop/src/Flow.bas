Attribute VB_Name = "Flow"
Sub SendToExcel()
    Dim ExcelApp As Excel.Application
    Dim Flow As Excel.Workbook
    Dim w
    
    Set ExcelApp = GetObject(, "Excel.Application")
    For Each w In ExcelApp.Workbooks
        If w.Name = "Debate.xltm" Then Set Flow = w
    Next w
    
    Paperless.SelectHeadingAndContent
    
    Flow.ActiveSheet.Range("A1").Value = Selection.Text
    
End Sub
