Imports System.Management.Automation
Imports System.Net
Imports Microsoft.Win32
Imports Namotion.Reflection

Public Class frmMain

    Private Function CheckVerbatimInstalled() As Boolean
        Return My.Computer.FileSystem.FileExists(Environ("APPDATA") & "\Microsoft\Templates\Debate.dotm")
    End Function

    Private Function GetVerbatimVersion() As String
        Return Registry.CurrentUser.OpenSubKey("Software\VB And VBA Program Settings\Verbatim\Main").GetValue("Version")
    End Function

    Private Sub UnblockTemplate()
        Dim ps As PowerShell = PowerShell.Create()

        ps.AddScript("Unblock-File -Path " & "C:\Users\hardy\Desktop\Debate.dotm")
        Dim results = ps.Invoke()
        'MessageBox.Show(results.Count)
        For Each result As PSObject In ps.Invoke
            MessageBox.Show(result.ToString())
        Next
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load

        If Me.CheckVerbatimInstalled() = True Then
            Me.lblStatusVerbatim.Text = "Verbatim is installed at version " & Me.GetVerbatimVersion()
        Else
            Me.lblStatusVerbatim.Text = "Verbatim is NOT installed"
        End If

        Dim MacroSecurity = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Office\16.0\Word\Security", False)
        ' These are not correct enum
        Select Case MacroSecurity.GetValue("VBAWarnings")
            Case 1
                Me.cboMacroSecurity.SelectedIndex = 1
            Case 2
                Me.cboMacroSecurity.SelectedIndex = 2
            Case 3
                Me.cboMacroSecurity.SelectedIndex = 3
            Case 4
                Me.cboMacroSecurity.SelectedIndex = 4
            Case Else
                Me.cboMacroSecurity.SelectedIndex = 1
        End Select
        MacroSecurity.Close()

        Dim AccessVBOM = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Office\16.0\Word\Security", False)
        If AccessVBOM.GetValue("AccessVBOM") = 0 Then
            Me.chkAccessVBOM.Checked = False
        Else
            Me.chkAccessVBOM.Checked = True
        End If
        AccessVBOM.Close()

        If Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", False).GetValue("ShowPreviewHandlers") = 0 Then
            Me.chkPreviewPane.Checked = True
        Else
            Me.chkPreviewPane.Checked = False
        End If
    End Sub

    Private Sub chkMacroSecurity_CheckedChanged(sender As Object, e As EventArgs) Handles chkMacroSecurity.CheckedChanged
        Dim MacroSecurity = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Office\16.0\Word\Security", True)

        If Me.chkMacroSecurity.Checked = True Then
            MacroSecurity.SetValue("VBAWarnings", 1)
        Else
            MacroSecurity.SetValue("VBAWarnings", 2)
        End If

        MacroSecurity.Close()
    End Sub

    Private Sub chkAccessVBOM_CheckedChanged(sender As Object, e As EventArgs) Handles chkAccessVBOM.CheckedChanged
        Dim AccessVBOM = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Office\16.0\Word\Security", True)

        If Me.chkAccessVBOM.Checked = True Then
            AccessVBOM.SetValue("AccessVBOM", 1)
        Else
            AccessVBOM.SetValue("AccessVBOM", 0)
        End If

        AccessVBOM.Close()
    End Sub

    Private Sub btnInstallOCR_Click(sender As Object, e As EventArgs)

        Dim URL = "https://www.learningcontainer.com/wp-content/uploads/2020/05/sample-zip-file.zip"

        Dim WC = New WebClient()
        WC.DownloadFile(URL, "C:\Users\hardy\Desktop\temp.zip")

        Dim ps As PowerShell = PowerShell.Create()

        ps.AddScript("Expand-Archive c:\Users\hardy\Desktop\temp.zip -DestinationPath c:\Users\hardy\Desktop\temp")
        Dim results = ps.Invoke()
    End Sub

    Private Sub btnUninstallOCR_Click(sender As Object, e As EventArgs)
        My.Computer.FileSystem.DeleteDirectory("C:\Users\hardy\Desktop\temp", FileIO.DeleteDirectoryOption.DeleteAllContents)

    End Sub

End Class
