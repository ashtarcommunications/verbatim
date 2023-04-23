Attribute VB_Name = "Startup"
Option Explicit

Public Sub Start()
    On Error Resume Next

    ' Don't run if in Protected View or file isn't visible on Windows
    #If Mac Then
        ' Do Nothing
    #Else
        If Not Application.ActiveProtectedViewWindow Is Nothing Or ActiveWindow.Visible = False Then Exit Sub
    #End If
    
    If InStr(Right$(ActiveWorkbook.Name, 4), "xlt") Then
        MsgBox "You're editing the master Verbatim Flow template. Any changes will affect every future flow." _
            & vbCrLf & vbCrLf _
            & "Instead, you probably want to open a new file based on this template, by double clicking or " _
            & "right-clicking and selecting ""New""."
    End If
            
    Globals.InitializeGlobals
    Settings.ResetKeyboardShortcuts
    
    ' Set Insert Mode - Ribbon toggle state is set in InitializeGlobals
    If GetSetting("Verbatim", "Flow", "InsertMode", False) = True Then
        Flow.EnableInsertMode
    Else
        Flow.DisableInsertMode
    End If
    
    ' Check for Analytics workbook
    QuickAnalytics.CreateDebateAnalyticsWorkbook

    On Error GoTo 0
End Sub
