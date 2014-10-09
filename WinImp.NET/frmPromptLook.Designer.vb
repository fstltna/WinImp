<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPromptLook
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents chkLimitMobility As System.Windows.Forms.CheckBox
	Public WithEvents cmbUpdates As System.Windows.Forms.ComboBox
	Public WithEvents chkIncludeUpdateEff As System.Windows.Forms.CheckBox
	Public WithEvents chkIncludeUpdateMob As System.Windows.Forms.CheckBox
	Public WithEvents cmdHelp As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents txtUnit As System.Windows.Forms.TextBox
	Public WithEvents lblUpdates As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPromptLook))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.chkLimitMobility = New System.Windows.Forms.CheckBox
		Me.cmbUpdates = New System.Windows.Forms.ComboBox
		Me.chkIncludeUpdateEff = New System.Windows.Forms.CheckBox
		Me.chkIncludeUpdateMob = New System.Windows.Forms.CheckBox
		Me.cmdHelp = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.txtUnit = New System.Windows.Forms.TextBox
		Me.lblUpdates = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.ControlBox = False
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.ClientSize = New System.Drawing.Size(297, 127)
		Me.Location = New System.Drawing.Point(1, 1)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmPromptLook"
		Me.chkLimitMobility.Text = "Limit the Mob. Gain to Max. Mob."
		Me.chkLimitMobility.Size = New System.Drawing.Size(177, 17)
		Me.chkLimitMobility.Location = New System.Drawing.Point(24, 80)
		Me.chkLimitMobility.TabIndex = 9
		Me.chkLimitMobility.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkLimitMobility.Visible = False
		Me.chkLimitMobility.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkLimitMobility.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkLimitMobility.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkLimitMobility.BackColor = System.Drawing.SystemColors.Control
		Me.chkLimitMobility.CausesValidation = True
		Me.chkLimitMobility.Enabled = True
		Me.chkLimitMobility.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkLimitMobility.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkLimitMobility.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkLimitMobility.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkLimitMobility.TabStop = True
		Me.chkLimitMobility.Name = "chkLimitMobility"
		Me.cmbUpdates.Enabled = False
		Me.cmbUpdates.Size = New System.Drawing.Size(49, 21)
		Me.cmbUpdates.Location = New System.Drawing.Point(112, 32)
		Me.cmbUpdates.Items.AddRange(New Object(){"1", "2", "3"})
		Me.cmbUpdates.TabIndex = 7
		Me.cmbUpdates.Text = "1"
		Me.cmbUpdates.Visible = False
		Me.cmbUpdates.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbUpdates.BackColor = System.Drawing.SystemColors.Window
		Me.cmbUpdates.CausesValidation = True
		Me.cmbUpdates.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbUpdates.IntegralHeight = True
		Me.cmbUpdates.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbUpdates.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbUpdates.Sorted = False
		Me.cmbUpdates.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cmbUpdates.TabStop = True
		Me.cmbUpdates.Name = "cmbUpdates"
		Me.chkIncludeUpdateEff.Text = "Include Efficiency Gain from Update"
		Me.chkIncludeUpdateEff.Size = New System.Drawing.Size(193, 17)
		Me.chkIncludeUpdateEff.Location = New System.Drawing.Point(8, 104)
		Me.chkIncludeUpdateEff.TabIndex = 6
		Me.chkIncludeUpdateEff.Visible = False
		Me.chkIncludeUpdateEff.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkIncludeUpdateEff.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkIncludeUpdateEff.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkIncludeUpdateEff.BackColor = System.Drawing.SystemColors.Control
		Me.chkIncludeUpdateEff.CausesValidation = True
		Me.chkIncludeUpdateEff.Enabled = True
		Me.chkIncludeUpdateEff.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkIncludeUpdateEff.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkIncludeUpdateEff.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkIncludeUpdateEff.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkIncludeUpdateEff.TabStop = True
		Me.chkIncludeUpdateEff.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkIncludeUpdateEff.Name = "chkIncludeUpdateEff"
		Me.chkIncludeUpdateMob.Text = "Include Mobility Gain from Update"
		Me.chkIncludeUpdateMob.Size = New System.Drawing.Size(185, 17)
		Me.chkIncludeUpdateMob.Location = New System.Drawing.Point(8, 56)
		Me.chkIncludeUpdateMob.TabIndex = 5
		Me.chkIncludeUpdateMob.Visible = False
		Me.chkIncludeUpdateMob.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkIncludeUpdateMob.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkIncludeUpdateMob.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkIncludeUpdateMob.BackColor = System.Drawing.SystemColors.Control
		Me.chkIncludeUpdateMob.CausesValidation = True
		Me.chkIncludeUpdateMob.Enabled = True
		Me.chkIncludeUpdateMob.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkIncludeUpdateMob.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkIncludeUpdateMob.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkIncludeUpdateMob.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkIncludeUpdateMob.TabStop = True
		Me.chkIncludeUpdateMob.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkIncludeUpdateMob.Name = "chkIncludeUpdateMob"
		Me.cmdHelp.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdHelp.Text = "?"
		Me.cmdHelp.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdHelp.Size = New System.Drawing.Size(25, 25)
		Me.cmdHelp.Location = New System.Drawing.Point(272, 0)
		Me.cmdHelp.TabIndex = 4
		Me.ToolTip1.SetToolTip(Me.cmdHelp, "Click for Help")
		Me.cmdHelp.BackColor = System.Drawing.SystemColors.Control
		Me.cmdHelp.CausesValidation = True
		Me.cmdHelp.Enabled = True
		Me.cmdHelp.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdHelp.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdHelp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdHelp.TabStop = True
		Me.cmdHelp.Name = "cmdHelp"
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "OK"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Size = New System.Drawing.Size(81, 25)
		Me.cmdOK.Location = New System.Drawing.Point(208, 48)
		Me.cmdOK.TabIndex = 2
		Me.cmdOK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(81, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(208, 88)
		Me.cmdCancel.TabIndex = 1
		Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.txtUnit.AutoSize = False
		Me.txtUnit.Size = New System.Drawing.Size(137, 19)
		Me.txtUnit.Location = New System.Drawing.Point(96, 8)
		Me.txtUnit.TabIndex = 0
		Me.txtUnit.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtUnit.AcceptsReturn = True
		Me.txtUnit.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtUnit.BackColor = System.Drawing.SystemColors.Window
		Me.txtUnit.CausesValidation = True
		Me.txtUnit.Enabled = True
		Me.txtUnit.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtUnit.HideSelection = True
		Me.txtUnit.ReadOnly = False
		Me.txtUnit.Maxlength = 0
		Me.txtUnit.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtUnit.MultiLine = False
		Me.txtUnit.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtUnit.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtUnit.TabStop = True
		Me.txtUnit.Visible = True
		Me.txtUnit.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtUnit.Name = "txtUnit"
		Me.lblUpdates.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblUpdates.Text = "Number of Updates"
		Me.lblUpdates.Size = New System.Drawing.Size(97, 17)
		Me.lblUpdates.Location = New System.Drawing.Point(8, 32)
		Me.lblUpdates.TabIndex = 8
		Me.lblUpdates.Visible = False
		Me.lblUpdates.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblUpdates.BackColor = System.Drawing.SystemColors.Control
		Me.lblUpdates.Enabled = True
		Me.lblUpdates.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblUpdates.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblUpdates.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblUpdates.UseMnemonic = True
		Me.lblUpdates.AutoSize = False
		Me.lblUpdates.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblUpdates.Name = "lblUpdates"
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label2.Text = "Look"
		Me.Label2.Font = New System.Drawing.Font("Arial Black", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(81, 17)
		Me.Label2.Location = New System.Drawing.Point(8, 8)
		Me.Label2.TabIndex = 3
		Me.Label2.BackColor = System.Drawing.SystemColors.Control
		Me.Label2.Enabled = True
		Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me.Controls.Add(chkLimitMobility)
		Me.Controls.Add(cmbUpdates)
		Me.Controls.Add(chkIncludeUpdateEff)
		Me.Controls.Add(chkIncludeUpdateMob)
		Me.Controls.Add(cmdHelp)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(txtUnit)
		Me.Controls.Add(lblUpdates)
		Me.Controls.Add(Label2)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class