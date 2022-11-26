VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStats 
   Caption         =   "Document Stats"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   2760
   OleObjectBlob   =   "frmStats.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ComputingStats As Boolean

Private Sub UserForm_Initialize()
    #If Mac Then
        UI.ResizeUserForm Me
    #End If

    ComputingStats = False

    ' Save parent doc
    Me.txtParentDoc.Value = ActiveDocument.Name
    
    ' Set form caption
    If InStr(ActiveDocument.Name, ".") > 1 Then
        Me.Caption = "Stats - " & Left(ActiveDocument.Name, InStrRev(ActiveDocument.Name, ".") - 1)
    Else
        Me.Caption = "Stats - " & ActiveDocument.Name
    End If

    Select Case GetSetting("Verbatim", "Profile", "CollegeHS", "College")
        Case Is = "College"
            If InStr(ActiveDocument.Name, "1AC") > 0 Or _
            InStr(ActiveDocument.Name, "1NC") > 0 Or _
            InStr(ActiveDocument.Name, "2AC") > 0 Or _
            InStr(ActiveDocument.Name, "2NC") > 0 Then
                Me.spnSpeechLength.Value = 9
            ElseIf InStr(ActiveDocument.Name, "1NR") > 0 Or _
            InStr(ActiveDocument.Name, "1AR") > 0 Or _
            InStr(ActiveDocument.Name, "2NR") > 0 Or _
            InStr(ActiveDocument.Name, "2AR") > 0 Then
                Me.spnSpeechLength.Value = 6
            Else
                Me.spnSpeechLength.Value = 9
            End If
        Case Is = "HS"
            If InStr(ActiveDocument.Name, "1AC") > 0 Or _
            InStr(ActiveDocument.Name, "1NC") > 0 Or _
            InStr(ActiveDocument.Name, "2AC") > 0 Or _
            InStr(ActiveDocument.Name, "2NC") > 0 Then
                Me.spnSpeechLength.Value = 8
            ElseIf InStr(ActiveDocument.Name, "1NR") > 0 Or _
            InStr(ActiveDocument.Name, "1AR") > 0 Or _
            InStr(ActiveDocument.Name, "2NR") > 0 Or _
            InStr(ActiveDocument.Name, "2AR") > 0 Then
                Me.spnSpeechLength.Value = 5
            Else
                Me.spnSpeechLength.Value = 8
            End If
    End Select
End Sub

Private Sub UserForm_Activate()
    On Error GoTo Handler
    
    If Me.lblHighlightCount.Caption = "xxxx" Then Calculate
        
    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub UserForm_Terminate()
    Unload Me
End Sub

Private Sub spnSpeechLength_Change()
    Me.txtSpeechLength.Value = Me.spnSpeechLength.Value
    If Me.lblEstimate <> "mm:ss" Then ComputeTotal
End Sub

Private Sub btnRefreshStats_Click()
    Calculate
End Sub

Private Sub chkAutoRefresh_Click()
    Dim Start
    On Error GoTo Handler
    If Me.chkAutoRefresh.Value = True Then
        Start = Now
        Do While Me.chkAutoRefresh.Value = True And Me.Visible = True
            DoEvents
            If Now > Start + TimeValue("00:00:20") Then
                Calculate
                Start = Now
            End If
        Loop
    End If
    
    Exit Sub
    
Handler:
    ' Error -2147418105 when form is terminated during loop
    Exit Sub
End Sub

Public Sub Calculate()
    If ComputingStats = True Then Exit Sub
    
    If Globals.InvisibilityToggle = True Then
        MsgBox "Cannot compute stats while Invisibility Mode is on. Closing stats form."
        Unload Me
        Exit Sub
    End If
    
    ComputingStats = True
    
    On Error GoTo Handler
    If Me.Visible = False Then Exit Sub
    ComputeHighlightedWords
    ComputeTagWords
    ComputeNumberOfCards
    ComputeTotal
    
    ComputingStats = False
    
    Exit Sub
    
Handler:
    Exit Sub
End Sub

Private Sub ComputeHighlightedWords()
    Dim r As Range
    Dim HighlightCount As Long
    Set r = Documents(Me.txtParentDoc.Value).Content
    
    On Error GoTo Handler
    
    Me.lblHighlightCount.Caption = "...."
    Me.lblHighlightCount.ForeColor = vbRed
    Me.lblHighlightCount1.Caption = "Computing..."
    Me.lblHighlightCount1.ForeColor = vbRed
    
    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .Highlight = True
        Do While .Execute
            DoEvents
            If Globals.InvisibilityToggle = True Then
                MsgBox "Cannot compute stats while Invisibility Mode is on. Closing stats form."
                Unload Me
                Exit Sub
            End If
            If Me.Visible = False Then Exit Sub
            HighlightCount = HighlightCount + r.ComputeStatistics(wdStatisticWords)
            Application.ScreenRefresh
        Loop
    End With
    
    ' Print results
    Me.lblHighlightCount.Caption = HighlightCount
    Me.lblHighlightCount.ForeColor = vbBlack
    Me.lblHighlightCount1.Caption = "Highlighted Words"
    Me.lblHighlightCount1.ForeColor = vbBlack
    Exit Sub
    
Handler:
    Exit Sub
End Sub

Private Sub ComputeTagWords()
    Dim r As Range
    Dim TagCount As Long
    Set r = Documents(Me.txtParentDoc.Value).Content
    
    On Error GoTo Handler
    
    Me.lblTagCount.Caption = "..."
    Me.lblTagCount.ForeColor = vbRed
    Me.lblTagCount1.Caption = "Computing..."
    Me.lblTagCount1.ForeColor = vbRed
    
    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .Format = True
        .ParagraphFormat.outlineLevel = wdOutlineLevel4
        Do While .Execute
            DoEvents
            If Globals.InvisibilityToggle = True Then
                MsgBox "Cannot compute stats while Invisibility Mode is on. Closing stats form."
                Unload Me
                Exit Sub
            End If
            If Me.Visible = False Then Exit Sub
            TagCount = TagCount + r.ComputeStatistics(wdStatisticWords)
            Application.ScreenRefresh
        Loop
        
    End With
    
    ' Print results
    Me.lblTagCount.Caption = TagCount
    Me.lblTagCount.ForeColor = vbBlack
    Me.lblTagCount1.Caption = "Words in Tags"
    Me.lblTagCount1.ForeColor = vbBlack
    
    Exit Sub
    
Handler:
    Exit Sub
End Sub

Private Sub ComputeNumberOfCards()
    Dim NumCards As Long
    Dim p As Paragraph
    
    Me.lblNumberOfCards.Caption = "..."
    Me.lblNumberOfCards.ForeColor = vbRed
    Me.lblNumberOfCards1.Caption = "Computing..."
    Me.lblNumberOfCards1.ForeColor = vbRed
    
    For Each p In Documents(Me.txtParentDoc.Value).Paragraphs
        If p.outlineLevel = wdOutlineLevel4 Then
            If p.Range.End <> ActiveDocument.Range.End Then
                If p.Next.outlineLevel > wdOutlineLevel4 And p.Next.Next.outlineLevel > wdOutlineLevel4 Then
                    NumCards = NumCards + 1
                    Application.ScreenRefresh
                End If
            End If
        End If
    Next p
    
    ' Print results
    Me.lblNumberOfCards.Caption = NumCards
    Me.lblNumberOfCards.ForeColor = vbBlack
    Me.lblNumberOfCards1.Caption = "# of Cards"
    Me.lblNumberOfCards1.ForeColor = vbBlack
End Sub

Private Sub ComputeTotal()
    Dim WPM As Integer
    Dim Remain
    
    WPM = Int(GetSetting("Verbatim", "Profile", "WPM", "350"))
    
    Me.lblTotal.Caption = Int(Me.lblHighlightCount.Caption) + Int(Me.lblTagCount.Caption)
    Remain = Round(((Int(Me.lblTotal.Caption) Mod WPM) / WPM) * 60, 0)
    If Remain < 10 Then Remain = "0" & Remain
    Me.lblEstimate.Caption = Int(Me.lblTotal.Caption) \ WPM & ":" & Remain
    Me.lblEstimate1.Caption = "ESTIMATE" & vbCrLf & "(@" & Str(WPM) & "wpm)"
    
    If Int(Me.lblTotal.Caption) / WPM < (0.75 * Me.spnSpeechLength.Value) Then Me.lblEstimate.BackColor = vbGreen
    If Int(Me.lblTotal.Caption) / WPM > (0.75 * Me.spnSpeechLength.Value) Then Me.lblEstimate.BackColor = vbYellow
    If Int(Me.lblTotal.Caption) / WPM > Me.spnSpeechLength.Value Then Me.lblEstimate.BackColor = vbRed
End Sub
