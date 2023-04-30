Imports System.Management.Automation
Imports Microsoft.Win32

Public Class frmMain
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If Me.CheckVerbatimInstalled() = True Then
            Dim Version = Me.GetVerbatimVersion()
            If Version = "x.x.x" Then
                Me.lblStatusVerbatim.Text = "Verbatim appears to be installed, but version cannot be detected"
            Else
                Me.lblStatusVerbatim.Text = "Verbatim is installed at version " & Version
            End If
            Me.UnblockTemplate()
        Else
            Me.lblStatusVerbatim.Text = "Verbatim is NOT installed"
        End If

        Dim MacroSecurity = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Office\16.0\Word\Security", False)
        If Not MacroSecurity Is Nothing Then
            Select Case MacroSecurity.GetValue("VBAWarnings")
                Case 1
                    Me.cboMacroSecurity.SelectedIndex = 3
                Case 2
                    Me.cboMacroSecurity.SelectedIndex = 1
                Case 3
                    Me.cboMacroSecurity.SelectedIndex = 2
                Case 4
                    Me.cboMacroSecurity.SelectedIndex = 0
                Case Else
                    Me.cboMacroSecurity.SelectedIndex = 0
            End Select
            MacroSecurity.Close()
        End If

        Dim AccessVBOM = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Office\16.0\Word\Security", False)
        If Not AccessVBOM Is Nothing Then
            If AccessVBOM.GetValue("AccessVBOM") = 1 Then
                Me.chkAccessVBOM.Checked = True
            Else
                Me.chkAccessVBOM.Checked = False
            End If
            AccessVBOM.Close()
        Else
            Me.chkAccessVBOM.Checked = False
        End If

        Dim ProtectedView = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Office\16.0\Word\Security\ProtectedView", False)
        If Not ProtectedView Is Nothing Then
            If ProtectedView.GetValue("DisableInternetFilesInPV") = 1 Then
                Me.chkProtectedView.Checked = True
            Else
                Me.chkProtectedView.Checked = False
            End If
            ProtectedView.Close()
        Else
            Me.chkProtectedView.Checked = False
        End If

        Dim PreviewPane = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", False)
        If Not PreviewPane Is Nothing Then
            If PreviewPane.GetValue("ShowPreviewHandlers") = 0 Then
                Me.chkPreviewPane.Checked = True
            Else
                Me.chkPreviewPane.Checked = False
            End If
            PreviewPane.Close()
        Else
            Me.chkPreviewPane.Checked = False
        End If

        Dim HardwareAcceleration = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Office\16.0\Common\Graphics", False)
        If Not HardwareAcceleration Is Nothing Then
            If HardwareAcceleration.GetValue("DisableHardwareAcceleration") = 1 Then
                Me.chkHardwareAcceleration.Checked = True
            Else
                Me.chkHardwareAcceleration.Checked = False
            End If
            HardwareAcceleration.Close()
        Else
            Me.chkHardwareAcceleration.Checked = False
        End If
    End Sub

    Private Function CheckVerbatimInstalled() As Boolean
        Return My.Computer.FileSystem.FileExists(Environ("APPDATA") & "\Microsoft\Templates\Debate.dotm")
    End Function

    Private Function GetVerbatimVersion() As String
        Dim Profile = Registry.CurrentUser.OpenSubKey("Software\VB And VBA Program Settings\Verbatim\Profile", False)
        If Profile Is Nothing Then
            Return "x.x.x"
        Else
            Dim Version = Profile.GetValue("Version")
            If Version Is Nothing Then
                Return "x.x.x"
            Else
                Return Version
            End If
        End If
    End Function

    Private Sub UnblockTemplate()
        Dim ps As PowerShell = PowerShell.Create()

        Dim DebatePath = Environ("APPDATA") & "\Microsoft\Templates\Debate.dotm"
        Dim StartupPath = Environ("APPDATA") & "\Microsoft\Word\STARTUP\DebateStartup.dotm"

        If Not DebatePath Is Nothing Then
            ps.AddScript("Unblock-File -Path " & DebatePath)
            ps.Invoke()
        End If

        If Not StartupPath Is Nothing Then
            ps.AddScript("Unblock-File -Path " & StartupPath)
            ps.Invoke()
        End If
    End Sub

    Private Sub cboMacroSecurity_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboMacroSecurity.SelectedIndexChanged
        Dim MacroSecurity = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Office\16.0\Word\Security", True)
        If MacroSecurity Is Nothing Then Exit Sub

        If Me.cboMacroSecurity.SelectedIndex = 0 Then
            MacroSecurity.SetValue("VBAWarnings", 4)
        ElseIf Me.cboMacroSecurity.SelectedIndex = 1 Then
            MacroSecurity.SetValue("VBAWarnings", 2)
        ElseIf Me.cboMacroSecurity.SelectedIndex = 2 Then
            MacroSecurity.SetValue("VBAWarnings", 3)
        ElseIf Me.cboMacroSecurity.SelectedIndex = 3 Then
            MacroSecurity.SetValue("VBAWarnings", 1)
        End If

        MacroSecurity.Close()
    End Sub

    Private Sub chkAccessVBOM_CheckedChanged(sender As Object, e As EventArgs) Handles chkAccessVBOM.CheckedChanged
        Dim AccessVBOM = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Office\16.0\Word\Security", True)
        If AccessVBOM Is Nothing Then Exit Sub

        If Me.chkAccessVBOM.Checked = True Then
            AccessVBOM.SetValue("AccessVBOM", 1)
        Else
            AccessVBOM.SetValue("AccessVBOM", 0)
        End If

        AccessVBOM.Close()
    End Sub

    Private Sub chkProtectedView_CheckedChanged(sender As Object, e As EventArgs) Handles chkProtectedView.CheckedChanged
        Dim ProtectedView = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Office\16.0\Word\Security\ProtectedView", True)
        If ProtectedView Is Nothing Then
            ProtectedView = Registry.CurrentUser.CreateSubKey("Software\Microsoft\Office\16.0\Word\Security\ProtectedView", True)
        End If
        If ProtectedView Is Nothing Then Exit Sub

        If Me.chkProtectedView.Checked = True Then
            ProtectedView.SetValue("DisableInternetFilesInPV", 1)
        Else
            ProtectedView.SetValue("DisableInternetFilesInPV", 0)
        End If

        ProtectedView.Close()
    End Sub

    Private Sub chkPreviewPane_CheckedChanged(sender As Object, e As EventArgs) Handles chkPreviewPane.CheckedChanged
        Dim PreviewPane = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", True)
        If PreviewPane Is Nothing Then Exit Sub

        If Me.chkPreviewPane.Checked = True Then
            PreviewPane.SetValue("ShowPreviewHandlers", 0)
        Else
            PreviewPane.SetValue("ShowPreviewHandlers", 1)
        End If

        PreviewPane.Close()
    End Sub

    Private Sub chkHardwareAcceleration_CheckedChanged(sender As Object, e As EventArgs) Handles chkHardwareAcceleration.CheckedChanged
        Dim HardwareAcceleration = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Office\16.0\Common\Graphics", True)
        If HardwareAcceleration Is Nothing Then
            HardwareAcceleration = Registry.CurrentUser.CreateSubKey("Software\Microsoft\Office\16.0\Common\Graphics", True)
        End If
        If HardwareAcceleration Is Nothing Then Exit Sub

        If Me.chkHardwareAcceleration.Checked = True Then
            HardwareAcceleration.SetValue("DisableHardwareAcceleration", 1)
        Else
            HardwareAcceleration.SetValue("DisableHardwareAcceleration", 0)
        End If

        HardwareAcceleration.Close()
    End Sub

    Private Sub lnkPaperlessDebate_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lnkPaperlessDebate.LinkClicked
        System.Diagnostics.Process.Start("explorer.exe", "https://paperlessdebate.com")
    End Sub
End Class
