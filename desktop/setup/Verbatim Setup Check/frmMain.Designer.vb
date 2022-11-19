<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.chkMacroSecurity = New System.Windows.Forms.CheckBox()
        Me.chkAccessVBOM = New System.Windows.Forms.CheckBox()
        Me.btnInstallOCR = New System.Windows.Forms.Button()
        Me.btnUninstallOCR = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'chkMacroSecurity
        '
        Me.chkMacroSecurity.AutoSize = True
        Me.chkMacroSecurity.Location = New System.Drawing.Point(42, 52)
        Me.chkMacroSecurity.Name = "chkMacroSecurity"
        Me.chkMacroSecurity.Size = New System.Drawing.Size(127, 19)
        Me.chkMacroSecurity.TabIndex = 1
        Me.chkMacroSecurity.Text = "All Macros Enabled"
        Me.chkMacroSecurity.UseVisualStyleBackColor = True
        '
        'chkAccessVBOM
        '
        Me.chkAccessVBOM.AutoSize = True
        Me.chkAccessVBOM.Location = New System.Drawing.Point(42, 86)
        Me.chkAccessVBOM.Name = "chkAccessVBOM"
        Me.chkAccessVBOM.Size = New System.Drawing.Size(99, 19)
        Me.chkAccessVBOM.TabIndex = 2
        Me.chkAccessVBOM.Text = "Access VBOM"
        Me.chkAccessVBOM.UseVisualStyleBackColor = True
        '
        'btnInstallOCR
        '
        Me.btnInstallOCR.Location = New System.Drawing.Point(261, 54)
        Me.btnInstallOCR.Name = "btnInstallOCR"
        Me.btnInstallOCR.Size = New System.Drawing.Size(111, 23)
        Me.btnInstallOCR.TabIndex = 3
        Me.btnInstallOCR.Text = "Install OCR Plugin"
        Me.btnInstallOCR.UseVisualStyleBackColor = True
        '
        'btnUninstallOCR
        '
        Me.btnUninstallOCR.Location = New System.Drawing.Point(378, 54)
        Me.btnUninstallOCR.Name = "btnUninstallOCR"
        Me.btnUninstallOCR.Size = New System.Drawing.Size(170, 23)
        Me.btnUninstallOCR.TabIndex = 4
        Me.btnUninstallOCR.Text = "Uninstall OCR Plugin"
        Me.btnUninstallOCR.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.btnUninstallOCR)
        Me.Controls.Add(Me.btnInstallOCR)
        Me.Controls.Add(Me.chkAccessVBOM)
        Me.Controls.Add(Me.chkMacroSecurity)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMain"
        Me.Text = "Verbatim Setup Check"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents chkMacroSecurity As CheckBox
    Friend WithEvents chkAccessVBOM As CheckBox
    Friend WithEvents btnInstallOCR As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents btnUninstallOCR As Button
End Class
