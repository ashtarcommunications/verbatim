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
        Me.lblInstructions = New System.Windows.Forms.Label()
        Me.grpSecurity = New System.Windows.Forms.GroupBox()
        Me.chkProtectedView = New System.Windows.Forms.CheckBox()
        Me.grpInstallation = New System.Windows.Forms.GroupBox()
        Me.lblStatusVerbatim = New System.Windows.Forms.Label()
        Me.chkDDE = New System.Windows.Forms.CheckBox()
        Me.chkPreviewPane = New System.Windows.Forms.CheckBox()
        Me.grpAdditional = New System.Windows.Forms.GroupBox()
        Me.chkHardwareAcceleration = New System.Windows.Forms.CheckBox()
        Me.cboMacroSecurity = New System.Windows.Forms.ComboBox()
        Me.lblMacroSecurity = New System.Windows.Forms.Label()
        Me.lnkPaperlessDebate = New System.Windows.Forms.LinkLabel()
        Me.grpSecurity.SuspendLayout()
        Me.grpInstallation.SuspendLayout()
        Me.grpAdditional.SuspendLayout()
        Me.SuspendLayout()
        '
        'chkMacroSecurity
        '
        Me.chkMacroSecurity.AutoSize = True
        Me.chkMacroSecurity.Location = New System.Drawing.Point(22, 22)
        Me.chkMacroSecurity.Name = "chkMacroSecurity"
        Me.chkMacroSecurity.Size = New System.Drawing.Size(127, 19)
        Me.chkMacroSecurity.TabIndex = 1
        Me.chkMacroSecurity.Text = "All Macros Enabled"
        Me.chkMacroSecurity.UseVisualStyleBackColor = True
        '
        'chkAccessVBOM
        '
        Me.chkAccessVBOM.AutoSize = True
        Me.chkAccessVBOM.Location = New System.Drawing.Point(22, 56)
        Me.chkAccessVBOM.Name = "chkAccessVBOM"
        Me.chkAccessVBOM.Size = New System.Drawing.Size(99, 19)
        Me.chkAccessVBOM.TabIndex = 2
        Me.chkAccessVBOM.Text = "Access VBOM"
        Me.chkAccessVBOM.UseVisualStyleBackColor = True
        '
        'lblInstructions
        '
        Me.lblInstructions.AccessibleDescription = "Instructions"
        Me.lblInstructions.AccessibleName = "Instructions"
        Me.lblInstructions.AutoSize = True
        Me.lblInstructions.Location = New System.Drawing.Point(25, 23)
        Me.lblInstructions.Name = "lblInstructions"
        Me.lblInstructions.Size = New System.Drawing.Size(322, 15)
        Me.lblInstructions.TabIndex = 5
        Me.lblInstructions.Text = "This tool helps configure your system for use with Verbatim."
        '
        'grpSecurity
        '
        Me.grpSecurity.AccessibleDescription = "Security"
        Me.grpSecurity.AccessibleName = "Security"
        Me.grpSecurity.Controls.Add(Me.chkProtectedView)
        Me.grpSecurity.Controls.Add(Me.chkMacroSecurity)
        Me.grpSecurity.Controls.Add(Me.chkAccessVBOM)
        Me.grpSecurity.Location = New System.Drawing.Point(25, 285)
        Me.grpSecurity.Name = "grpSecurity"
        Me.grpSecurity.Size = New System.Drawing.Size(222, 140)
        Me.grpSecurity.TabIndex = 6
        Me.grpSecurity.TabStop = False
        Me.grpSecurity.Text = "Macro security settings"
        '
        'chkProtectedView
        '
        Me.chkProtectedView.AutoSize = True
        Me.chkProtectedView.Location = New System.Drawing.Point(22, 81)
        Me.chkProtectedView.Name = "chkProtectedView"
        Me.chkProtectedView.Size = New System.Drawing.Size(105, 19)
        Me.chkProtectedView.TabIndex = 3
        Me.chkProtectedView.Text = "Protected View"
        Me.chkProtectedView.UseVisualStyleBackColor = True
        '
        'grpInstallation
        '
        Me.grpInstallation.AccessibleDescription = "Installation Status"
        Me.grpInstallation.AccessibleName = "Installation Status"
        Me.grpInstallation.Controls.Add(Me.lblStatusVerbatim)
        Me.grpInstallation.Location = New System.Drawing.Point(25, 54)
        Me.grpInstallation.Name = "grpInstallation"
        Me.grpInstallation.Size = New System.Drawing.Size(322, 145)
        Me.grpInstallation.TabIndex = 7
        Me.grpInstallation.TabStop = False
        Me.grpInstallation.Text = "Installation Status"
        '
        'lblStatusVerbatim
        '
        Me.lblStatusVerbatim.AutoSize = True
        Me.lblStatusVerbatim.Location = New System.Drawing.Point(51, 39)
        Me.lblStatusVerbatim.Name = "lblStatusVerbatim"
        Me.lblStatusVerbatim.Size = New System.Drawing.Size(123, 15)
        Me.lblStatusVerbatim.TabIndex = 0
        Me.lblStatusVerbatim.Text = "Verbatim Install Status"
        '
        'chkDDE
        '
        Me.chkDDE.AutoSize = True
        Me.chkDDE.Location = New System.Drawing.Point(6, 22)
        Me.chkDDE.Name = "chkDDE"
        Me.chkDDE.Size = New System.Drawing.Size(130, 19)
        Me.chkDDE.TabIndex = 8
        Me.chkDDE.Text = "DDE Single Instance"
        Me.chkDDE.UseVisualStyleBackColor = True
        '
        'chkPreviewPane
        '
        Me.chkPreviewPane.AutoSize = True
        Me.chkPreviewPane.Location = New System.Drawing.Point(6, 57)
        Me.chkPreviewPane.Name = "chkPreviewPane"
        Me.chkPreviewPane.Size = New System.Drawing.Size(142, 19)
        Me.chkPreviewPane.TabIndex = 9
        Me.chkPreviewPane.Text = "Explorer Preview Pane"
        Me.chkPreviewPane.UseVisualStyleBackColor = True
        '
        'grpAdditional
        '
        Me.grpAdditional.AccessibleDescription = "Additional Settings"
        Me.grpAdditional.AccessibleName = "AdditionalSettings"
        Me.grpAdditional.Controls.Add(Me.chkHardwareAcceleration)
        Me.grpAdditional.Controls.Add(Me.chkDDE)
        Me.grpAdditional.Controls.Add(Me.chkPreviewPane)
        Me.grpAdditional.Location = New System.Drawing.Point(541, 136)
        Me.grpAdditional.Name = "grpAdditional"
        Me.grpAdditional.Size = New System.Drawing.Size(200, 128)
        Me.grpAdditional.TabIndex = 10
        Me.grpAdditional.TabStop = False
        Me.grpAdditional.Text = "Additional Settings"
        '
        'chkHardwareAcceleration
        '
        Me.chkHardwareAcceleration.AutoSize = True
        Me.chkHardwareAcceleration.Location = New System.Drawing.Point(6, 82)
        Me.chkHardwareAcceleration.Name = "chkHardwareAcceleration"
        Me.chkHardwareAcceleration.Size = New System.Drawing.Size(187, 19)
        Me.chkHardwareAcceleration.TabIndex = 13
        Me.chkHardwareAcceleration.Text = "Disable Hardware Acceleration"
        Me.chkHardwareAcceleration.UseVisualStyleBackColor = True
        '
        'cboMacroSecurity
        '
        Me.cboMacroSecurity.AccessibleDescription = "Macro Security"
        Me.cboMacroSecurity.AccessibleName = "MacroSecurity"
        Me.cboMacroSecurity.FormattingEnabled = True
        Me.cboMacroSecurity.Items.AddRange(New Object() {"Very High (Disables macros, Verbatim can't run)", "High (Prompts to allow macros on every file)", "Medium (Prompts to allow macros on every file)", "Low (Disables all prompts, including for non-Verbatim code)"})
        Me.cboMacroSecurity.Location = New System.Drawing.Point(461, 307)
        Me.cboMacroSecurity.Name = "cboMacroSecurity"
        Me.cboMacroSecurity.Size = New System.Drawing.Size(280, 23)
        Me.cboMacroSecurity.TabIndex = 11
        '
        'lblMacroSecurity
        '
        Me.lblMacroSecurity.AutoSize = True
        Me.lblMacroSecurity.Location = New System.Drawing.Point(462, 286)
        Me.lblMacroSecurity.Name = "lblMacroSecurity"
        Me.lblMacroSecurity.Size = New System.Drawing.Size(116, 15)
        Me.lblMacroSecurity.TabIndex = 12
        Me.lblMacroSecurity.Text = "Macro Security Level"
        '
        'lnkPaperlessDebate
        '
        Me.lnkPaperlessDebate.AutoSize = True
        Me.lnkPaperlessDebate.Location = New System.Drawing.Point(376, 23)
        Me.lnkPaperlessDebate.Name = "lnkPaperlessDebate"
        Me.lnkPaperlessDebate.Size = New System.Drawing.Size(119, 15)
        Me.lnkPaperlessDebate.TabIndex = 13
        Me.lnkPaperlessDebate.TabStop = True
        Me.lnkPaperlessDebate.Text = "paperlessdebate.com"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.lnkPaperlessDebate)
        Me.Controls.Add(Me.lblMacroSecurity)
        Me.Controls.Add(Me.cboMacroSecurity)
        Me.Controls.Add(Me.grpAdditional)
        Me.Controls.Add(Me.grpInstallation)
        Me.Controls.Add(Me.grpSecurity)
        Me.Controls.Add(Me.lblInstructions)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMain"
        Me.Text = "Verbatim Setup Check"
        Me.grpSecurity.ResumeLayout(False)
        Me.grpSecurity.PerformLayout()
        Me.grpInstallation.ResumeLayout(False)
        Me.grpInstallation.PerformLayout()
        Me.grpAdditional.ResumeLayout(False)
        Me.grpAdditional.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents chkMacroSecurity As CheckBox
    Friend WithEvents chkAccessVBOM As CheckBox
    Friend WithEvents Button1 As Button
    Friend WithEvents lblInstructions As Label
    Friend WithEvents grpSecurity As GroupBox
    Friend WithEvents chkProtectedView As CheckBox
    Friend WithEvents grpInstallation As GroupBox
    Friend WithEvents lblStatusVerbatim As Label
    Friend WithEvents chkDDE As CheckBox
    Friend WithEvents chkPreviewPane As CheckBox
    Friend WithEvents grpAdditional As GroupBox
    Friend WithEvents cboMacroSecurity As ComboBox
    Friend WithEvents lblMacroSecurity As Label
    Friend WithEvents chkHardwareAcceleration As CheckBox
    Friend WithEvents lnkPaperlessDebate As LinkLabel
End Class
