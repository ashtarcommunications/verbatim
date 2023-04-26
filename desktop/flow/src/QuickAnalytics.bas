Attribute VB_Name = "QuickAnalytics"
Option Explicit

Public Sub CreateDebateAnalyticsWorkbook()
    If Filesystem.FileExists(Application.TemplatesPath & "DebateAnalytics.xlsx") = False Then
        Dim xl As Object
        Dim wb As Excel.Workbook
        
        #If Mac Then
            Dim xlwb As Object
            Set xlwb = CreateObject("Excel.Application")
            Set xl = xlwb.Application
        #Else
            Set xl = New Excel.Application
        #End If
    
        Application.StatusBar = "DebateAnalytics.xlsx spreadsheet not found in your templates folder, creating a blank one."

        Application.ScreenUpdating = False
        Set wb = xl.Workbooks.Add
        wb.Sheets.Add Count:=9 ' Ensure 10 sheets for different profiles
        wb.SaveAs Application.TemplatesPath & "DebateAnalytics.xlsx"
        wb.Close SaveChanges:=True
        #If Mac Then
            xlwb.Close SaveChanges:=False
        #Else
            xl.Quit
        #End If
        
        Set xl = Nothing
        Set wb = Nothing
        #If Mac Then
            Set xlwb = Nothing
        #End If

        Application.ScreenUpdating = True
    End If
End Sub

Public Sub AddQuickAnalytic()
    Dim Profile As Long
    Dim wb As Workbook
    Dim Name As String
    Dim OpenCol As Long
    
    On Error GoTo Handler
    
    If Not Selection.Areas.Count = 1 Then
        MsgBox "Selection is not contiguous - please select only adjacent cells and try again.", vbOKOnly
        Exit Sub
    End If
    
    If Selection.Columns.Count > 1 Then
        MsgBox "You can only select cells from a single column for a Quick Analytic.", vbOKOnly
        Exit Sub
    End If
    
    If Application.WorksheetFunction.CountA(Selection) = 0 Then
        MsgBox "Selected cells must contain some text to save a Quick Analytic", vbOKOnly
        Exit Sub
    End If
    
    ' Ensure the workbook exists before accessing it
    QuickAnalytics.CreateDebateAnalyticsWorkbook
    Set wb = GetObject(Application.TemplatesPath & "DebateAnalytics.xlsx")
    
    Name = InputBox("What shortcut word/phrase do you want to use for your Quick Analytic? This should be something short and memorable.", "Add Quick Analytic", "")
    If Name = "" Then Exit Sub

    Profile = CLng(Replace(GetSetting("Verbatim", "Flow", "QuickAnalyticsProfile", "Profile 1"), "Profile ", ""))
    If Profile < 1 Then Profile = 1

    If Not Application.WorksheetFunction.CountIf(wb.Sheets.[_Default](Profile).Rows(1), Name) = 0 Then
        MsgBox "There's already a Quick Analytic with that name, try again with a different name!", vbOKOnly, "Failed To Add Quick Analytic"
        Exit Sub
    End If

    ' Find the first unused column in the Analytics sheet
    OpenCol = Utility.FirstEmptyColumn(wb.Sheets.[_Default](Profile))
    
    ' This should never happen in normal usage, but just in case
    If OpenCol = 0 Then
        MsgBox "There's no space left for Quick Analytics! Try deleting some.", vbOKOnly, "Failed To Add Quick Analytic"
        Exit Sub
    End If
    
    ' Copy selected range to first open column in new sheet
    wb.Sheets.[_Default](Profile).Cells(1, OpenCol).Value = Name
    Selection.Copy destination:=wb.Sheets.[_Default](Profile).Cells(2, OpenCol)
    
    wb.Close SaveChanges:=True
    Set wb = Nothing
    
    MsgBox "Successfully created Quick Analytic with the shortcut """ & Name & """"

    Exit Sub

Handler:
    Set wb = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

'@Ignore ProcedureNotUsed
Public Sub InsertCurrentQuickAnalytic()
    QuickAnalytics.InsertQuickAnalytic ActiveCell.Value
End Sub

Public Sub InsertQuickAnalytic(ByRef QuickCardAnalyticName As String)
    Dim Profile As Long
    Dim wb As Workbook
    Dim c As Long
    Dim i As Long
    Dim LastRow As Long
    
    On Error GoTo Handler
    
    QuickAnalytics.CreateDebateAnalyticsWorkbook
    Set wb = GetObject(Application.TemplatesPath & "DebateAnalytics.xlsx")
    
    Profile = CLng(Replace(GetSetting("Verbatim", "Flow", "QuickAnalyticsProfile", "Profile 1"), "Profile ", ""))
    If Profile < 1 Then Profile = 1
    
    ' Find the Quick Analytic with the specified name in the first row
    For i = 1 To wb.Sheets.[_Default](Profile).UsedRange.Columns.Count
        If StrComp(LCase$(wb.Sheets.[_Default](Profile).Cells(1, i).Value), LCase$(QuickCardAnalyticName), vbTextCompare) = 0 Then
            c = i
            Exit For
        End If
    Next i
    
    ' If the Quick Analytic was found, copy the rest of the column to the active cell
    If c > 0 Then
        LastRow = wb.Sheets.[_Default](Profile).Cells(wb.Sheets.[_Default](Profile).Rows.Count, c).End(xlUp).Row
        wb.Sheets.[_Default](Profile).Range(wb.Sheets.[_Default](Profile).Cells(2, c), wb.Sheets.[_Default](Profile).Cells(LastRow, c)).Copy destination:=ActiveCell
    Else
        Application.StatusBar = "No Quick Analytic with that name found!"
    End If
    
    wb.Close SaveChanges:=False
    Set wb = Nothing
    
    Exit Sub

Handler:
    Set wb = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Public Sub DeleteAllQuickAnalytics()
    Dim Profile As Long
    Dim wb As Workbook
    
    On Error GoTo Handler
    
    If MsgBox("Are you sure you want to delete all saved Quick Analytics in this profile? This cannot be reversed.", vbYesNo, "Are you sure?") = vbNo Then Exit Sub
    
    QuickAnalytics.CreateDebateAnalyticsWorkbook
    Set wb = GetObject(Application.TemplatesPath & "DebateAnalytics.xlsx")
    
    Profile = CLng(Replace(GetSetting("Verbatim", "Flow", "QuickAnalyticsProfile", "Profile 1"), "Profile ", ""))
    If Profile < 1 Then Profile = 1
    
    wb.Sheets.[_Default](Profile).Cells.Clear

    wb.Close SaveChanges:=True
    Set wb = Nothing

    Exit Sub

Handler:
    Set wb = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Public Sub DeleteQuickAnalytic(Optional ByRef QuickAnalyticName As String)
    Dim Profile As Long
    Dim wb As Workbook
    Dim c As Long
    Dim i As Long
    
    On Error GoTo Handler
       
    If QuickAnalyticName <> "" Or IsNull(QuickAnalyticName) Then
        If MsgBox("Are you sure you want to delete the Quick Analytic """ & QuickAnalyticName & """? This cannot be reversed.", vbYesNo, "Are you sure?") = vbNo Then Exit Sub
    End If
       
    QuickAnalytics.CreateDebateAnalyticsWorkbook
    Set wb = GetObject(Application.TemplatesPath & "DebateAnalytics.xlsx")
    
    Profile = CLng(Replace(GetSetting("Verbatim", "Flow", "QuickAnalyticsProfile", "Profile 1"), "Profile ", ""))
    If Profile < 1 Then Profile = 1
    
    ' Find the Quick Analytic with the specified name in the first row
    For i = 1 To wb.Sheets.[_Default](Profile).UsedRange.Columns.Count
        If StrComp(LCase$(wb.Sheets.[_Default](Profile).Cells(1, i).Value), LCase$(QuickAnalyticName), vbTextCompare) = 0 Then
            c = i
            Exit For
        End If
    Next i
    
    ' If the column was found, delete it
    If c > 0 Then
        wb.Sheets.[_Default](Profile).Columns(c).Delete
    End If
        
    wb.Close SaveChanges:=True
    Set wb = Nothing
    
    Exit Sub

Handler:
    Set wb = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

'@Ignore ParameterNotUsed, ProcedureNotUsed
'@Ignore ProcedureCanBeWrittenAsFunction
Public Sub GetQuickAnalyticsContent(ByVal c As IRibbonControl, ByRef returnedVal As Variant)
' Get content for dynamic menu for Quick Analytics
    Dim Profile As Long
    Dim i As Long
    Dim wb As Workbook
    Dim xml As String
    Dim QuickAnalyticName As String
    Dim DisplayName As String
       
    On Error Resume Next

    QuickAnalytics.CreateDebateAnalyticsWorkbook
    Set wb = GetObject(Application.TemplatesPath & "DebateAnalytics.xlsx")

    Profile = CLng(Replace(GetSetting("Verbatim", "Flow", "QuickAnalyticsProfile", "Profile 1"), "Profile ", ""))
    If Profile < 1 Then Profile = 1

    ' Start the menu
    xml = "<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">"
    
    ' Populate the list of current Quick Analytics
    For i = 1 To wb.Sheets.[_Default](Profile).UsedRange.Columns.Count
        QuickAnalyticName = wb.Sheets.[_Default](Profile).Cells(1, i)
        DisplayName = Strings.OnlySafeChars(QuickAnalyticName)
                         
        If DisplayName <> "" Then
            xml = xml & "<button id=""QuickAnalytic" & Replace(DisplayName, " ", "") & """ label=""" & DisplayName & """ tag=""" & QuickAnalyticName & """ onAction=""QuickAnalytics.InsertQuickAnalyticFromRibbon"" imageMso=""AutoSummaryResummarize"" />"
        End If
    Next i
    
    ' Close the menu
    xml = xml & "<button id=""QuickAnalyticsSettings"" label=""Quick Analytics Settings"" onAction=""Ribbon.RibbonMain"" imageMso=""AddInManager""" & " />"
    xml = xml & "</menu>"
    
    wb.Close SaveChanges:=False
    Set wb = Nothing
    
    returnedVal = xml
        
    On Error GoTo 0
        
    Exit Sub
End Sub

'@Ignore ProcedureNotUsed
Public Sub InsertQuickAnalyticFromRibbon(ByVal c As IRibbonControl)
    QuickAnalytics.InsertQuickAnalytic c.Tag
End Sub

