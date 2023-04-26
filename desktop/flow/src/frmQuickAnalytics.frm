VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQuickAnalytics 
   Caption         =   "Quick Analytics"
   ClientHeight    =   7110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8370
   OleObjectBlob   =   "frmQuickAnalytics.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmQuickAnalytics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    On Error GoTo Handler
    
    Globals.InitializeGlobals
    
    #If Mac Then
        UI.ResizeUserForm Me
        Me.btnAdd.ForeColor = Globals.GREEN
        Me.btnDelete.ForeColor = Globals.ORANGE
        Me.btnDeleteAll.ForeColor = Globals.RED
        Me.btnClose.ForeColor = Globals.BLUE
    #End If
    
    Me.cboQuickAnalyticsProfile.AddItem "Profile 1"
    Me.cboQuickAnalyticsProfile.AddItem "Profile 2"
    Me.cboQuickAnalyticsProfile.AddItem "Profile 3"
    Me.cboQuickAnalyticsProfile.AddItem "Profile 4"
    Me.cboQuickAnalyticsProfile.AddItem "Profile 5"
    Me.cboQuickAnalyticsProfile.AddItem "Profile 6"
    Me.cboQuickAnalyticsProfile.AddItem "Profile 7"
    Me.cboQuickAnalyticsProfile.AddItem "Profile 8"
    Me.cboQuickAnalyticsProfile.AddItem "Profile 9"
    Me.cboQuickAnalyticsProfile.AddItem "Profile 10"
    
    Me.cboQuickAnalyticsProfile.Value = GetSetting("Verbatim", "Flow", "QuickAnalyticsProfile", "Profile 1")
    
    PopulateQuickAnalytics
        
    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub cboQuickAnalyticsProfile_Change()
    SaveSetting "Verbatim", "Flow", "QuickAnalyticsProfile", Me.cboQuickAnalyticsProfile.Value
    PopulateQuickAnalytics
End Sub

Private Sub PopulateQuickAnalytics()
    Dim Profile As Long
    Dim xl As Object
    Dim wb As Workbook
    Dim i As Long
    Dim QuickAnalyticName As String
    
    On Error GoTo Handler
    
    ' Have to use a new Excel instance to avoid the current workbook being set to invisible
    #If Mac Then
        Dim xlwb As Object
        Set xlwb = CreateObject("Excel.Application")
        Set xl = xlwb.Application
    #Else
        Set xl = New Excel.Application
    #End If
    xl.AutomationSecurity = msoAutomationSecurityForceDisable
    xl.EnableEvents = False
    
    ' Make sure the analytics file exists
    QuickAnalytics.CreateDebateAnalyticsWorkbook
    Set wb = xl.Workbooks.Open(Application.TemplatesPath & "DebateAnalytics.xlsx")
    
    Me.lboxQuickAnalytics.Clear
    
    Profile = CLng(Replace(GetSetting("Verbatim", "Flow", "QuickAnalyticsProfile", "Profile 1"), "Profile ", ""))
    If Profile < 1 Then Profile = 1
    
    For i = 1 To wb.Sheets.[_Default](Profile).UsedRange.Columns.Count
        QuickAnalyticName = wb.Sheets.[_Default](Profile).Cells(1, i).Value
        If QuickAnalyticName <> "" Then
            Me.lboxQuickAnalytics.AddItem
            Me.lboxQuickAnalytics.List(Me.lboxQuickAnalytics.ListCount - 1, 0) = QuickAnalyticName
            Me.lboxQuickAnalytics.List(Me.lboxQuickAnalytics.ListCount - 1, 1) = Left$(wb.Sheets.[_Default](Profile).Cells(2, i).Value, 50) & "..."
        End If
    Next i
    
    wb.Close SaveChanges:=False

    #If Mac Then
        xlwb.Close SaveChanges:=False
    #Else
        xl.Quit
    #End If
    
    Set wb = Nothing
    Set xl = Nothing
    #If Mac Then
        Set xlwb = Nothing
    #End If
    
    Exit Sub
    
Handler:
    Set wb = Nothing
    Set xl = Nothing
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
    QuickAnalytics.AddQuickAnalytic
    
    ' Refresh the list to get the new Quick Analytic
    PopulateQuickAnalytics
End Sub

Private Sub btnDelete_Click()
    On Error GoTo Handler
    
    If Me.lboxQuickAnalytics.Value = "" Or IsNull(Me.lboxQuickAnalytics.Value) Then
        MsgBox "Please select a Quick Analytic to delete first.", vbOKOnly
        Exit Sub
    End If
        
    QuickAnalytics.DeleteQuickAnalytic Me.lboxQuickAnalytics.Value
    Me.lboxQuickAnalytics.RemoveItem (Me.lboxQuickAnalytics.ListIndex)
    
    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub btnDeleteAll_Click()
    On Error GoTo Handler
    
    QuickAnalytics.DeleteAllQuickAnalytics
    Me.lboxQuickAnalytics.Clear

    Exit Sub

Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub btnClose_Click()
    Ribbon.RefreshRibbon
    Unload Me
End Sub

