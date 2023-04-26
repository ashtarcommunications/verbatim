Attribute VB_Name = "UI"
'@IgnoreModule ProcedureNotUsed
Option Explicit

Public Sub ShowForm(ByVal FormName As String)
    Dim Form As Object
    
    On Error Resume Next
    
    Select Case FormName
        Case "CheatSheet"
            Set Form = New frmCheatSheet
        Case "Settings"
            Set Form = New frmSettings
        Case "QuickAnalytics"
            Set Form = New frmQuickAnalytics
        Case Else
            ' Do nothing
            Exit Sub
    End Select

    Form.Show
    
    On Error GoTo 0
End Sub

' Functions for assigning keyboard shortcuts to forms
Public Sub ShowFormSettings()
    UI.ShowForm "Settings"
End Sub

Public Sub ShowFormQuickAnalytics()
    UI.ShowForm "QuickAnalytics"
End Sub

Public Sub ShowFormCheatSheet()
    UI.ShowForm "CheatSheet"
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

Public Sub NextSheet()
    If ActiveSheet.Index = ActiveWorkbook.Worksheets.Count Then
        ActiveWorkbook.Worksheets.[_Default](1).Select
    Else
        ActiveSheet.Next.Select
    End If
End Sub

Public Sub PreviousSheet()
    If ActiveSheet.Index = 1 Then
        ActiveWorkbook.Worksheets.[_Default](ActiveWorkbook.Worksheets.Count).Select
    Else
        ActiveSheet.Previous.Select
    End If
End Sub

'@Ignore ParameterNotUsed, ProcedureNotUsed
'@Ignore ProcedureCanBeWrittenAsFunction
Public Sub GetSwitchSpeechContent(ByVal c As IRibbonControl, ByRef returnedVal As Variant)
' Get content for dynamic menu for switching speeches
    Dim xml As String
    Dim SpeechNames() As String
    Dim s As Variant
    
    On Error Resume Next

    ' Start the menu
    xml = "<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">"
    
    SpeechNames = Split(GetSetting("Verbatim", "Flow", "SpeechNames", "1AC,1NC,2AC,Block,1AR,2NR,2AR"), ",")
    For Each s In SpeechNames
        xml = xml & "<button id=""SwitchSpeech" & Replace(Strings.OnlySafeChars(s), " ", "") & """ label=""" & s & """ tag=""" & s & """ onAction=""Flow.SwitchSpeech"" imageMso=""FillRight"" />"
    Next s
    
    ' Close the menu
    xml = xml & "</menu>"
    
    returnedVal = xml
        
    On Error GoTo 0
        
    Exit Sub
End Sub

Public Sub SplitWithWord()
    Dim wd As Object
    Dim w As Variant
    Dim MaxLeft As Long
    Dim MaxWidth As Long
    Dim MaxTop As Long
    Dim MaxHeight As Long
    
    On Error GoTo Handler
        
    Set wd = GetObject(, "Word.Application")
    
    If wd Is Nothing Then
        MsgBox "Word must be open first!", vbOKOnly
        Exit Sub
    End If
   
    ' Find largest usable window size
    ActiveWindow.WindowState = xlMaximized
    MaxLeft = ActiveWindow.Left
    MaxWidth = ActiveWindow.Width
    MaxTop = ActiveWindow.Top
    MaxHeight = ActiveWindow.Height
        
    ' Set to zero if maximized window returns a negative number
    If MaxLeft < 0 Then MaxLeft = 0
    If MaxTop < 0 Then MaxTop = 0
                 
    ' Put Excel on the left
    ActiveWindow.WindowState = xlNormal
    ActiveWindow.Width = MaxWidth / 2
    ActiveWindow.Left = 0
                 
    ' Set height
    #If Mac Then
        ' On Mac, Application.UsableHeight automatically positions vertically.
        ' Need extra check and no Top setting to avoid maximized window bug.
        If ActiveWindow.WindowState <> xlMaximized Then ActiveWindow.Height = Application.UsableHeight
    #Else
        ActiveWindow.Height = MaxHeight
        ActiveWindow.Top = MaxTop
    #End If
                 
    ' Loop through open Word windows and move to Right
    For Each w In wd.Windows
        w.WindowState = 0
        
        ' Mac Word has a bug that treats full-size windows as maximized even in normal mode
        ' and won't let you change window state directly, so fake it out with a small fixed size
        #If Mac Then
            w.Width = 100
            w.Height = 100
        #End If
                
        ' Put Word docs on the right
        w.Width = MaxWidth / 2
        w.Left = MaxLeft + (MaxWidth / 2)
        
        #If Mac Then
            If w.WindowState <> 1 Then w.Height = Application.UsableHeight
        #Else
            w.Height = MaxHeight
            w.Top = MaxTop
        #End If
    Next w
    
    ActiveWindow.Activate
    
    Set wd = Nothing

    Exit Sub
    
Handler:
    Set wd = Nothing
    If Err.Number = 429 Then
        MsgBox "Word must be open first!"
    Else
        MsgBox "Error " & Err.Number & ": " & Err.Description
    End If
End Sub
