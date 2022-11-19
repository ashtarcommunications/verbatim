Imports System.Management.Automation
Imports System.Net
Imports Microsoft.Win32

Public Class frmMain

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load

        Dim ps As PowerShell = PowerShell.Create()

        ps.AddScript("Unblock-File -Path " & "C:\Users\hardy\Desktop\Debate.dotm")
        Dim results = ps.Invoke()
        'MessageBox.Show(results.Count)
        For Each result As PSObject In ps.Invoke
            MessageBox.Show(result.ToString())
        Next

        Dim MacroSecurity = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Office\16.0\Word\Security", False)
        If MacroSecurity.GetValue("VBAWarnings") = 1 Then
            Me.chkMacroSecurity.Checked = True
        Else
            Me.chkMacroSecurity.Checked = False
        End If
        MacroSecurity.Close()

        Dim AccessVBOM = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Office\16.0\Word\Security", False)
        If AccessVBOM.GetValue("AccessVBOM") = 0 Then
            Me.chkAccessVBOM.Checked = False
        Else
            Me.chkAccessVBOM.Checked = True
        End If
        AccessVBOM.Close()
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

    Private Sub btnInstallOCR_Click(sender As Object, e As EventArgs) Handles btnInstallOCR.Click

        Dim URL = "https://www.learningcontainer.com/wp-content/uploads/2020/05/sample-zip-file.zip"

        Dim WC = New WebClient()
        WC.DownloadFile(URL, "C:\Users\hardy\Desktop\temp.zip")

        Dim ps As PowerShell = PowerShell.Create()

        ps.AddScript("Expand-Archive c:\Users\hardy\Desktop\temp.zip -DestinationPath c:\Users\hardy\Desktop\temp")
        Dim results = ps.Invoke()
    End Sub

    Private Sub btnUninstallOCR_Click(sender As Object, e As EventArgs) Handles btnUninstallOCR.Click
        My.Computer.FileSystem.DeleteDirectory("C:\Users\hardy\Desktop\temp", FileIO.DeleteDirectoryOption.DeleteAllContents)

    End Sub
End Class
