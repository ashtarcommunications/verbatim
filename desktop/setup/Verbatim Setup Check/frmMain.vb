Imports System.Management.Automation
Imports Microsoft.Win32

Public Class frmMain

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load

        Dim ps As PowerShell = PowerShell.Create()

        ps.AddScript("Unblock-File -Path " & "C:\Users\hardy\Desktop\Debate.dotm")
        Dim results = ps.Invoke()
        MessageBox.Show(results.Count)
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

End Class
