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
        Dim resources As ComponentModel.ComponentResourceManager = New ComponentModel.ComponentResourceManager(GetType(frmMain))
        chkAccessVBOM = New CheckBox()
        lblInstructions = New Label()
        grpSecurity = New GroupBox()
        lblSecurityWarning = New Label()
        chkProtectedView = New CheckBox()
        lblMacroSecurity = New Label()
        cboMacroSecurity = New ComboBox()
        grpInstallation = New GroupBox()
        lblStatusVerbatim = New Label()
        chkPreviewPane = New CheckBox()
        grpAdditional = New GroupBox()
        lblAdditionalSettings = New Label()
        chkHardwareAcceleration = New CheckBox()
        lnkPaperlessDebate = New LinkLabel()
        grpSecurity.SuspendLayout()
        grpInstallation.SuspendLayout()
        grpAdditional.SuspendLayout()
        SuspendLayout()
        ' 
        ' chkAccessVBOM
        ' 
        chkAccessVBOM.AutoSize = True
        chkAccessVBOM.Location = New Point(6, 132)
        chkAccessVBOM.Name = "chkAccessVBOM"
        chkAccessVBOM.Size = New Size(308, 19)
        chkAccessVBOM.TabIndex = 2
        chkAccessVBOM.Text = "Allow access to VBOM (for verbatimizing documents)"
        chkAccessVBOM.UseVisualStyleBackColor = True
        ' 
        ' lblInstructions
        ' 
        lblInstructions.AccessibleDescription = "Instructions"
        lblInstructions.AccessibleName = "Instructions"
        lblInstructions.Location = New Point(12, 19)
        lblInstructions.Name = "lblInstructions"
        lblInstructions.Size = New Size(400, 40)
        lblInstructions.TabIndex = 5
        lblInstructions.Text = "This tool helps configure your system for use with Verbatim. For more information on each option, see the online help at paperlessdebate.com."
        ' 
        ' grpSecurity
        ' 
        grpSecurity.AccessibleDescription = "Security"
        grpSecurity.AccessibleName = "Security"
        grpSecurity.Controls.Add(lblSecurityWarning)
        grpSecurity.Controls.Add(chkProtectedView)
        grpSecurity.Controls.Add(lblMacroSecurity)
        grpSecurity.Controls.Add(cboMacroSecurity)
        grpSecurity.Controls.Add(chkAccessVBOM)
        grpSecurity.Location = New Point(12, 168)
        grpSecurity.Name = "grpSecurity"
        grpSecurity.Size = New Size(385, 180)
        grpSecurity.TabIndex = 6
        grpSecurity.TabStop = False
        grpSecurity.Text = "Macro Security Settings"
        ' 
        ' lblSecurityWarning
        ' 
        lblSecurityWarning.Location = New Point(6, 19)
        lblSecurityWarning.Name = "lblSecurityWarning"
        lblSecurityWarning.Size = New Size(370, 66)
        lblSecurityWarning.TabIndex = 14
        lblSecurityWarning.Text = resources.GetString("lblSecurityWarning.Text")
        ' 
        ' chkProtectedView
        ' 
        chkProtectedView.AutoSize = True
        chkProtectedView.Location = New Point(6, 157)
        chkProtectedView.Name = "chkProtectedView"
        chkProtectedView.Size = New Size(341, 19)
        chkProtectedView.TabIndex = 3
        chkProtectedView.Text = "Disable Protected View (suppresses prompts when opening)"
        chkProtectedView.UseVisualStyleBackColor = True
        ' 
        ' lblMacroSecurity
        ' 
        lblMacroSecurity.AutoSize = True
        lblMacroSecurity.Location = New Point(6, 85)
        lblMacroSecurity.Name = "lblMacroSecurity"
        lblMacroSecurity.Size = New Size(116, 15)
        lblMacroSecurity.TabIndex = 12
        lblMacroSecurity.Text = "Macro Security Level"
        ' 
        ' cboMacroSecurity
        ' 
        cboMacroSecurity.AccessibleDescription = "Macro Security"
        cboMacroSecurity.AccessibleName = "MacroSecurity"
        cboMacroSecurity.FormattingEnabled = True
        cboMacroSecurity.Items.AddRange(New Object() {"Very High (Disables macros, Verbatim can't run)", "High (Prompts to allow macros on every file)", "Medium (Prompts to allow macros on every file)", "Low (Disables all prompts, including for non-Verbatim code)"})
        cboMacroSecurity.Location = New Point(6, 103)
        cboMacroSecurity.Name = "cboMacroSecurity"
        cboMacroSecurity.Size = New Size(358, 23)
        cboMacroSecurity.TabIndex = 11
        ' 
        ' grpInstallation
        ' 
        grpInstallation.AccessibleDescription = "Installation Status"
        grpInstallation.AccessibleName = "Installation Status"
        grpInstallation.Controls.Add(lblStatusVerbatim)
        grpInstallation.Location = New Point(12, 62)
        grpInstallation.Name = "grpInstallation"
        grpInstallation.Size = New Size(385, 100)
        grpInstallation.TabIndex = 7
        grpInstallation.TabStop = False
        grpInstallation.Text = "Installation Status"
        ' 
        ' lblStatusVerbatim
        ' 
        lblStatusVerbatim.AutoSize = True
        lblStatusVerbatim.Location = New Point(6, 19)
        lblStatusVerbatim.Name = "lblStatusVerbatim"
        lblStatusVerbatim.Size = New Size(123, 15)
        lblStatusVerbatim.TabIndex = 0
        lblStatusVerbatim.Text = "Verbatim Install Status"
        ' 
        ' chkPreviewPane
        ' 
        chkPreviewPane.AutoSize = True
        chkPreviewPane.Location = New Point(6, 54)
        chkPreviewPane.Name = "chkPreviewPane"
        chkPreviewPane.Size = New Size(358, 19)
        chkPreviewPane.TabIndex = 9
        chkPreviewPane.Text = "Disable Explorer Preview Pane (for issues choosing speech doc)"
        chkPreviewPane.UseVisualStyleBackColor = True
        ' 
        ' grpAdditional
        ' 
        grpAdditional.AccessibleDescription = "Additional Settings"
        grpAdditional.AccessibleName = "AdditionalSettings"
        grpAdditional.Controls.Add(lblAdditionalSettings)
        grpAdditional.Controls.Add(chkHardwareAcceleration)
        grpAdditional.Controls.Add(chkPreviewPane)
        grpAdditional.Location = New Point(12, 354)
        grpAdditional.Name = "grpAdditional"
        grpAdditional.Size = New Size(385, 100)
        grpAdditional.TabIndex = 10
        grpAdditional.TabStop = False
        grpAdditional.Text = "Additional Settings"
        ' 
        ' lblAdditionalSettings
        ' 
        lblAdditionalSettings.Location = New Point(6, 19)
        lblAdditionalSettings.Name = "lblAdditionalSettings"
        lblAdditionalSettings.Size = New Size(370, 32)
        lblAdditionalSettings.TabIndex = 15
        lblAdditionalSettings.Text = "These tweaks can help solve specific issues while using Verbatim. Only check them if you run into the listed problem."
        ' 
        ' chkHardwareAcceleration
        ' 
        chkHardwareAcceleration.AutoSize = True
        chkHardwareAcceleration.Location = New Point(6, 75)
        chkHardwareAcceleration.Name = "chkHardwareAcceleration"
        chkHardwareAcceleration.Size = New Size(336, 19)
        chkHardwareAcceleration.TabIndex = 13
        chkHardwareAcceleration.Text = "Disable Hardware Acceleration (for Word slowdown issues)"
        chkHardwareAcceleration.UseVisualStyleBackColor = True
        ' 
        ' lnkPaperlessDebate
        ' 
        lnkPaperlessDebate.AutoSize = True
        lnkPaperlessDebate.Location = New Point(278, 467)
        lnkPaperlessDebate.Name = "lnkPaperlessDebate"
        lnkPaperlessDebate.Size = New Size(119, 15)
        lnkPaperlessDebate.TabIndex = 13
        lnkPaperlessDebate.TabStop = True
        lnkPaperlessDebate.Text = "paperlessdebate.com"
        ' 
        ' frmMain
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(414, 491)
        Controls.Add(lnkPaperlessDebate)
        Controls.Add(grpAdditional)
        Controls.Add(grpInstallation)
        Controls.Add(grpSecurity)
        Controls.Add(lblInstructions)
        Icon = CType(resources.GetObject("$this.Icon"), Icon)
        Name = "frmMain"
        Text = "Verbatim Setup Check"
        grpSecurity.ResumeLayout(False)
        grpSecurity.PerformLayout()
        grpInstallation.ResumeLayout(False)
        grpInstallation.PerformLayout()
        grpAdditional.ResumeLayout(False)
        grpAdditional.PerformLayout()
        ResumeLayout(False)
        PerformLayout()
    End Sub
    Friend WithEvents chkAccessVBOM As CheckBox
    Friend WithEvents Button1 As Button
    Friend WithEvents lblInstructions As Label
    Friend WithEvents grpSecurity As GroupBox
    Friend WithEvents chkProtectedView As CheckBox
    Friend WithEvents grpInstallation As GroupBox
    Friend WithEvents lblStatusVerbatim As Label
    Friend WithEvents chkPreviewPane As CheckBox
    Friend WithEvents grpAdditional As GroupBox
    Friend WithEvents cboMacroSecurity As ComboBox
    Friend WithEvents lblMacroSecurity As Label
    Friend WithEvents chkHardwareAcceleration As CheckBox
    Friend WithEvents lnkPaperlessDebate As LinkLabel
    Friend WithEvents lblSecurityWarning As Label
    Friend WithEvents lblAdditionalSettings As Label
End Class
