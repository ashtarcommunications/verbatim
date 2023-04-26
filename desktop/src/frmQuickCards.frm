VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQuickCards 
   Caption         =   "Quick Cards"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8460
   OleObjectBlob   =   "frmQuickCards.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmQuickCards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim Profile As String
    
    On Error GoTo Handler
    
    Globals.InitializeGlobals
    
    #If Mac Then
        UI.ResizeUserForm Me
        Me.btnAdd.ForeColor = Globals.GREEN
        Me.btnDelete.ForeColor = Globals.ORANGE
        Me.btnDeleteAll.ForeColor = Globals.RED
        Me.btnClose.ForeColor = Globals.BLUE
    #End If
      
    Me.cboQuickCardsProfile.AddItem "Profile 1"
    Me.cboQuickCardsProfile.AddItem "Profile 2"
    Me.cboQuickCardsProfile.AddItem "Profile 3"
    Me.cboQuickCardsProfile.AddItem "Profile 4"
    Me.cboQuickCardsProfile.AddItem "Profile 5"
    Me.cboQuickCardsProfile.AddItem "Profile 6"
    Me.cboQuickCardsProfile.AddItem "Profile 7"
    Me.cboQuickCardsProfile.AddItem "Profile 8"
    Me.cboQuickCardsProfile.AddItem "Profile 9"
    Me.cboQuickCardsProfile.AddItem "Profile 10"
    
    Profile = GetSetting("Verbatim", "QuickCards", "QuickCardsProfile", "Verbatim1")
    If Not Profile Like "Verbatim*" Then Profile = "Verbatim1"
    
    Me.cboQuickCardsProfile.Value = Replace(Profile, "Verbatim", "Profile ")
    
    PopulateQuickCards
        
    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub cboQuickCardsProfile_Change()
    Dim Profile As String
    Profile = Replace(Me.cboQuickCardsProfile.Value, "Profile ", "Verbatim")
    SaveSetting "Verbatim", "QuickCards", "QuickCardsProfile", Profile
    PopulateQuickCards
End Sub

Private Sub PopulateQuickCards()
    Dim Profile As String
    Dim t As Template
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    On Error GoTo Handler
    
    Me.lboxQuickCards.Clear
    
    Profile = GetSetting("Verbatim", "QuickCards", "QuickCardsProfile", "Verbatim1")
    If Not Profile Like "Verbatim*" Then Profile = "Verbatim1"
    
    Set t = ActiveDocument.AttachedTemplate

    ' Populate the list of current Quick Cards in the Custom 1 / Verbatim gallery
    For i = 1 To t.BuildingBlockTypes.Count
        If t.BuildingBlockTypes.Item(i).Name = "Custom 1" Then
            For j = 1 To t.BuildingBlockTypes.Item(i).Categories.Count
                If t.BuildingBlockTypes.Item(i).Categories.Item(j).Name = Profile Then
                    For k = 1 To t.BuildingBlockTypes.Item(i).Categories.Item(j).BuildingBlocks.Count
                        Me.lboxQuickCards.AddItem
                        Me.lboxQuickCards.List(Me.lboxQuickCards.ListCount - 1, 0) = t.BuildingBlockTypes.Item(i).Categories.Item(j).BuildingBlocks.Item(k).Name
                        Me.lboxQuickCards.List(Me.lboxQuickCards.ListCount - 1, 1) = Left$(t.BuildingBlockTypes.Item(i).Categories.Item(j).BuildingBlocks.Item(k).Value, 50) & "..."
                    Next k
                End If
            Next j
        End If
    Next i
    
    Set t = Nothing
    
    Exit Sub
    
Handler:
    Set t = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

#If Mac Then
    ' Do Nothing
#Else
Public Sub btnAdd_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.btnAdd.BackColor = Globals.LIGHT_GREEN
End Sub
Public Sub btnDelete_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.btnDelete.BackColor = Globals.LIGHT_ORANGE
End Sub
Public Sub btnDeleteAll_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.btnDeleteAll.BackColor = Globals.LIGHT_RED
End Sub
Public Sub btnClose_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.btnClose.BackColor = Globals.LIGHT_BLUE
End Sub
Public Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Me.btnAdd.BackColor = Globals.GREEN
    Me.btnDelete.BackColor = Globals.ORANGE
    Me.btnDeleteAll.BackColor = Globals.RED
    Me.btnClose.BackColor = Globals.BLUE
End Sub
#End If

Private Sub btnAdd_Click()
    QuickCards.AddQuickCard
    
    ' Refresh the list to get the new Quick Card
    PopulateQuickCards
End Sub

Private Sub btnDelete_Click()
    On Error GoTo Handler
    
    If Me.lboxQuickCards.Value = "" Or IsNull(Me.lboxQuickCards.Value) Then
        MsgBox "Please select a Quick Card to delete first.", vbOKOnly
        Exit Sub
    End If
        
    QuickCards.DeleteQuickCard Me.lboxQuickCards.Value
    Me.lboxQuickCards.RemoveItem (Me.lboxQuickCards.ListIndex)
    
    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub btnDeleteAll_Click()
    On Error GoTo Handler
    
    ' Calling without a name parameter deletes all
    QuickCards.DeleteQuickCard ""
    Me.lboxQuickCards.Clear

    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub btnClose_Click()
    Ribbon.RefreshRibbon
    Unload Me
End Sub
