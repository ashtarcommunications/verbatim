Attribute VB_Name = "Settings"
Option Explicit

Public Function GetModifierKey() As String
    #If Mac Then
        GetModifierKey = "*" ' Command Key
    #Else
        GetModifierKey = "^" ' Ctrl Key
    #End If
End Function

Public Function GetTildeCode() As Long
    If GetSetting("Verbatim", "Flow", "AlternateTildeCode", False) = True Then
        GetTildeCode = 192
    Else
        GetTildeCode = 96
    End If
End Function

Public Sub ResetKeyboardShortcuts()
' + = Shift, ^ = Ctrl, % = Alt, * = Command
    Dim Modifier As String
    Modifier = Settings.GetModifierKey
    
    Application.OnKey Chr$(Settings.GetTildeCode), "Speech.SendToSpeechCursor"
    Application.OnKey "%" & Chr$(Settings.GetTildeCode), "Speech.SendToSpeechEnd"
        
    Application.OnKey "{F3}", "Flow.InsertCellAbove"
    Application.OnKey "%{F3}", "Flow.InsertCellBelow"
        
    Application.OnKey "{F4}", "Flow.MergeCells"
        
    Application.OnKey "{F5}", "Flow.InsertRowAbove"
    Application.OnKey "%{F5}", "Flow.InsertRowBelow"
    Application.OnKey Modifier & "%{F5}", "Flow.DeleteRow"
        
    Application.OnKey "{F6}", "Flow.PasteAsText"
    Application.OnKey Modifier & "%6", "Flow.PasteAsText" ' Alternative for Mac
    
    Application.OnKey "{F7}", "Flow.ToggleEvidence"
    
    Application.OnKey "{F8}", "Flow.ToggleGroup"
    
    Application.OnKey "{F9}", "Flow.ExtendArgument"
    
    'Application.OnKey "{F10}"
    
    Application.OnKey "{F11}", "Flow.ToggleHighlighting"
    
    Application.OnKey "{F12}", "UI.ShowFormCheatSheet"
    
    Application.OnKey Modifier & "%a", "Format.AddFlowAff"
    Application.OnKey Modifier & "%n", "Format.AddFlowNeg"
    Application.OnKey Modifier & "%x", "Format.AddFlowCX"
    
    Application.OnKey Modifier & "%{UP}", "Flow.MoveUp"
    Application.OnKey Modifier & "%{DOWN}", "Flow.MoveDown"
    Application.OnKey Modifier & "%+{DOWN}", "Flow.GoToBottom"
    
End Sub

'@Ignore ProcedureNotUsed
Public Sub RemoveKeyBindings()
    Dim Modifier As String
    Modifier = Settings.GetModifierKey
    
    ' Leaving off second parameter to Application.OnKey clears the key binding
    Application.OnKey Chr$(Settings.GetTildeCode)
    Application.OnKey "%" & Chr$(Settings.GetTildeCode)
    
    Application.OnKey "{F2}"
    
    Application.OnKey "{F3}"
    Application.OnKey "%{F3}"
    
    Application.OnKey "{F4}"
    
    Application.OnKey "{F5}"
    Application.OnKey "%{F5}"
    Application.OnKey Modifier & "{F5}"
    
    Application.OnKey "{F6}"
    Application.OnKey Modifier & "%6"
    
    Application.OnKey "{F7}"
    Application.OnKey "{F8}"
    Application.OnKey "{F9}"
    Application.OnKey "{F10}"
    Application.OnKey "{F11}"
    Application.OnKey "{F12}"
    
    Application.OnKey Modifier & "%a"
    Application.OnKey Modifier & "%n"
    Application.OnKey Modifier & "%x"
    
    Application.OnKey Modifier & "%{UP}"
    Application.OnKey Modifier & "%{DOWN}"
    Application.OnKey Modifier & "%+{DOWN}"
End Sub
