Attribute VB_Name = "Ribbon"
'@IgnoreModule ProcedureCanBeWrittenAsFunction, ProcedureNotUsed
Option Explicit

Public Sub OnLoad(ByVal Ribbon As IRibbonUI)
    Dim SavedState As Boolean
    Set Globals.DebateRibbon = Ribbon
    
    ' Save a pointer to the Ribbon in case it gets lost
    SavedState = ActiveWorkbook.Saved
    ActiveWorkbook.Names.Add "DebateRibbonPointer", ObjPtr(Ribbon), False
    ActiveWorkbook.Saved = SavedState
End Sub

Public Function GetRibbon(ByVal lRibbonPointer As LongPtr) As Object
    Dim objRibbon As Object
    #If Mac Then
        CopyMemory_byVar objRibbon, lRibbonPointer, LenB(lRibbonPointer)
    #Else
        '@Ignore ImplicitUnboundDefaultMemberAccess
        CopyMemory objRibbon, lRibbonPointer, LenB(lRibbonPointer)
    #End If
    Set GetRibbon = objRibbon
    Set objRibbon = Nothing
End Function

Public Sub RefreshRibbon()
    If Globals.DebateRibbon Is Nothing Then
        Set Globals.DebateRibbon = GetRibbon(Replace(ActiveWorkbook.Names.[_Default]("DebateRibbonPointer").[_Default], "=", ""))
        Globals.DebateRibbon.Invalidate
    Else
        Globals.DebateRibbon.Invalidate
    End If
End Sub

Public Sub RibbonMain(ByVal c As IRibbonControl)
    Select Case c.ID
        
        ' Speech
        Case Is = "SendToSpeechCursor"
            Speech.SendToSpeechCursor
        
        Case Is = "SendToSpeechEnd"
            Speech.SendToSpeechEnd
            
        Case Is = "QuickAnalyticsSettings"
            UI.ShowFormQuickAnalytics

        ' Cells
        Case Is = "InsertCellAbove"
            Flow.InsertCellAbove
        
        Case Is = "InsertCellBelow"
            Flow.InsertCellBelow
            
        Case Is = "MergeCells"
            Flow.MergeCells
        
        Case Is = "ToggleGroup"
            Flow.ToggleGroup
            
        Case Is = "ToggleHighlighting"
            Flow.ToggleHighlighting
            
        Case Is = "ToggleEvidence"
            Flow.ToggleEvidence
        
        Case Is = "ExtendArgument"
            Flow.ExtendArgument
        
        ' Rows
        Case Is = "InsertRowAbove"
            Flow.InsertRowAbove
        
        Case Is = "InsertRowBelow"
            Flow.InsertRowBelow
        
        Case Is = "DeleteRow"
            Flow.DeleteRow

        Case Is = "MoveUp"
            Flow.MoveUp

        Case Is = "MoveDown"
            Flow.MoveDown

        Case Is = "GoToBottom"
            Flow.GoToBottom
        
        ' Sheets
        Case Is = "AddFlowAff"
            Format.AddFlowAff
            
        Case Is = "AddFlowNeg"
            Format.AddFlowNeg
        
        Case Is = "AddFlowCX"
            Format.AddFlowCX
        
        Case Is = "DeleteFlow"
            Format.DeleteFlow
        
        Case Is = "DeleteEmptyFlows"
            Format.DeleteEmptyFlows
        
        Case Is = "AutoScoutingInfo"
            Format.AutoScoutingInfo

        ' Insert
        Case Is = "EnterCell"
            Flow.EnterCell
        
        Case Is = "PasteAsText"
            Flow.PasteAsText
        
        ' View
        Case Is = "SplitWithWord"
            UI.SplitWithWord
        
        ' Settings
        Case Is = "CheatSheet"
            UI.ShowFormCheatSheet

        Case Is = "FlowSettings", "FlowSettings1", "FlowSettings2"
            UI.ShowFormSettings

        Case Else
            ' Do nothing
    End Select
End Sub

Public Sub GetRibbonToggles(ByVal c As IRibbonControl, ByRef state As Variant)
    Select Case c.ID
        
    Case Is = "InsertMode"
        state = GetSetting("Verbatim", "Flow", "InsertMode", False)
        
    Case Else
        state = False
        
    End Select
End Sub
