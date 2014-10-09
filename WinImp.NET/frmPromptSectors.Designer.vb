<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPromptSectors
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
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents txtParm As System.Windows.Forms.TextBox
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents cbField As System.Windows.Forms.ComboBox
	Public WithEvents optCond As System.Windows.Forms.RadioButton
	Public WithEvents txtTerr As System.Windows.Forms.TextBox
	Public WithEvents optTerr As System.Windows.Forms.RadioButton
	Public WithEvents txtDest As System.Windows.Forms.TextBox
	Public WithEvents txtOrigin As System.Windows.Forms.TextBox
	Public WithEvents txtSector As System.Windows.Forms.TextBox
	Public WithEvents optRange As System.Windows.Forms.RadioButton
	Public WithEvents txtRealm As System.Windows.Forms.ComboBox
	Public WithEvents OptRealm As System.Windows.Forms.RadioButton
	Public WithEvents optSect As System.Windows.Forms.RadioButton
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPromptSectors))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdOK = New System.Windows.Forms.Button
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.Frame2 = New System.Windows.Forms.GroupBox
		Me.txtParm = New System.Windows.Forms.TextBox
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.cbField = New System.Windows.Forms.ComboBox
		Me.optCond = New System.Windows.Forms.RadioButton
		Me.txtTerr = New System.Windows.Forms.TextBox
		Me.optTerr = New System.Windows.Forms.RadioButton
		Me.txtDest = New System.Windows.Forms.TextBox
		Me.txtOrigin = New System.Windows.Forms.TextBox
		Me.txtSector = New System.Windows.Forms.TextBox
		Me.optRange = New System.Windows.Forms.RadioButton
		Me.txtRealm = New System.Windows.Forms.ComboBox
		Me.OptRealm = New System.Windows.Forms.RadioButton
		Me.optSect = New System.Windows.Forms.RadioButton
		Me.Label2 = New System.Windows.Forms.Label
		Me.Frame2.SuspendLayout()
		Me.Frame1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "Select Sector(s)"
		Me.ClientSize = New System.Drawing.Size(262, 322)
		Me.Location = New System.Drawing.Point(4, 23)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmPromptSectors"
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "OK"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Size = New System.Drawing.Size(81, 25)
		Me.cmdOK.Location = New System.Drawing.Point(32, 280)
		Me.cmdOK.TabIndex = 13
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
		Me.cmdCancel.Location = New System.Drawing.Point(128, 280)
		Me.cmdCancel.TabIndex = 14
		Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.Frame2.Text = "Optional Paramters"
		Me.Frame2.Size = New System.Drawing.Size(217, 65)
		Me.Frame2.Location = New System.Drawing.Point(16, 200)
		Me.Frame2.TabIndex = 11
		Me.Frame2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame2.BackColor = System.Drawing.SystemColors.Control
		Me.Frame2.Enabled = True
		Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame2.Visible = True
		Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame2.Name = "Frame2"
		Me.txtParm.AutoSize = False
		Me.txtParm.Size = New System.Drawing.Size(177, 25)
		Me.txtParm.Location = New System.Drawing.Point(16, 24)
		Me.txtParm.TabIndex = 12
		Me.txtParm.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtParm.AcceptsReturn = True
		Me.txtParm.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtParm.BackColor = System.Drawing.SystemColors.Window
		Me.txtParm.CausesValidation = True
		Me.txtParm.Enabled = True
		Me.txtParm.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtParm.HideSelection = True
		Me.txtParm.ReadOnly = False
		Me.txtParm.Maxlength = 0
		Me.txtParm.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtParm.MultiLine = False
		Me.txtParm.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtParm.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtParm.TabStop = True
		Me.txtParm.Visible = True
		Me.txtParm.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtParm.Name = "txtParm"
		Me.Frame1.Text = "Select "
		Me.Frame1.Size = New System.Drawing.Size(217, 185)
		Me.Frame1.Location = New System.Drawing.Point(16, 8)
		Me.Frame1.TabIndex = 1
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.cbField.Size = New System.Drawing.Size(73, 21)
		Me.cbField.Location = New System.Drawing.Point(136, 120)
		Me.cbField.Items.AddRange(New Object(){"Terr(0)", "Terr1", "Terr2", "Terr3"})
		Me.cbField.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cbField.TabIndex = 16
		Me.cbField.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbField.BackColor = System.Drawing.SystemColors.Window
		Me.cbField.CausesValidation = True
		Me.cbField.Enabled = True
		Me.cbField.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cbField.IntegralHeight = True
		Me.cbField.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbField.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbField.Sorted = False
		Me.cbField.TabStop = True
		Me.cbField.Visible = True
		Me.cbField.Name = "cbField"
		Me.optCond.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optCond.Text = "Select All F4 - Conditional Sectors"
		Me.optCond.Enabled = False
		Me.optCond.Size = New System.Drawing.Size(201, 17)
		Me.optCond.Location = New System.Drawing.Point(8, 152)
		Me.optCond.TabIndex = 15
		Me.optCond.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optCond.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optCond.BackColor = System.Drawing.SystemColors.Control
		Me.optCond.CausesValidation = True
		Me.optCond.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optCond.Cursor = System.Windows.Forms.Cursors.Default
		Me.optCond.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optCond.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optCond.TabStop = True
		Me.optCond.Checked = False
		Me.optCond.Visible = True
		Me.optCond.Name = "optCond"
		Me.txtTerr.AutoSize = False
		Me.txtTerr.Size = New System.Drawing.Size(49, 19)
		Me.txtTerr.Location = New System.Drawing.Point(80, 120)
		Me.txtTerr.TabIndex = 10
		Me.txtTerr.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtTerr.AcceptsReturn = True
		Me.txtTerr.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtTerr.BackColor = System.Drawing.SystemColors.Window
		Me.txtTerr.CausesValidation = True
		Me.txtTerr.Enabled = True
		Me.txtTerr.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtTerr.HideSelection = True
		Me.txtTerr.ReadOnly = False
		Me.txtTerr.Maxlength = 0
		Me.txtTerr.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtTerr.MultiLine = False
		Me.txtTerr.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtTerr.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtTerr.TabStop = True
		Me.txtTerr.Visible = True
		Me.txtTerr.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtTerr.Name = "txtTerr"
		Me.optTerr.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optTerr.Text = "Territory"
		Me.optTerr.Size = New System.Drawing.Size(65, 17)
		Me.optTerr.Location = New System.Drawing.Point(8, 120)
		Me.optTerr.TabIndex = 9
		Me.optTerr.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optTerr.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optTerr.BackColor = System.Drawing.SystemColors.Control
		Me.optTerr.CausesValidation = True
		Me.optTerr.Enabled = True
		Me.optTerr.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optTerr.Cursor = System.Windows.Forms.Cursors.Default
		Me.optTerr.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optTerr.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optTerr.TabStop = True
		Me.optTerr.Checked = False
		Me.optTerr.Visible = True
		Me.optTerr.Name = "optTerr"
		Me.txtDest.AutoSize = False
		Me.txtDest.Size = New System.Drawing.Size(41, 19)
		Me.txtDest.Location = New System.Drawing.Point(160, 88)
		Me.txtDest.TabIndex = 7
		Me.txtDest.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtDest.AcceptsReturn = True
		Me.txtDest.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtDest.BackColor = System.Drawing.SystemColors.Window
		Me.txtDest.CausesValidation = True
		Me.txtDest.Enabled = True
		Me.txtDest.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtDest.HideSelection = True
		Me.txtDest.ReadOnly = False
		Me.txtDest.Maxlength = 0
		Me.txtDest.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtDest.MultiLine = False
		Me.txtDest.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtDest.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtDest.TabStop = True
		Me.txtDest.Visible = True
		Me.txtDest.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtDest.Name = "txtDest"
		Me.txtOrigin.AutoSize = False
		Me.txtOrigin.Size = New System.Drawing.Size(49, 19)
		Me.txtOrigin.Location = New System.Drawing.Point(80, 88)
		Me.txtOrigin.TabIndex = 6
		Me.txtOrigin.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtOrigin.AcceptsReturn = True
		Me.txtOrigin.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtOrigin.BackColor = System.Drawing.SystemColors.Window
		Me.txtOrigin.CausesValidation = True
		Me.txtOrigin.Enabled = True
		Me.txtOrigin.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtOrigin.HideSelection = True
		Me.txtOrigin.ReadOnly = False
		Me.txtOrigin.Maxlength = 0
		Me.txtOrigin.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtOrigin.MultiLine = False
		Me.txtOrigin.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtOrigin.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtOrigin.TabStop = True
		Me.txtOrigin.Visible = True
		Me.txtOrigin.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtOrigin.Name = "txtOrigin"
		Me.txtSector.AutoSize = False
		Me.txtSector.Size = New System.Drawing.Size(49, 19)
		Me.txtSector.Location = New System.Drawing.Point(80, 24)
		Me.txtSector.TabIndex = 2
		Me.txtSector.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtSector.AcceptsReturn = True
		Me.txtSector.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtSector.BackColor = System.Drawing.SystemColors.Window
		Me.txtSector.CausesValidation = True
		Me.txtSector.Enabled = True
		Me.txtSector.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtSector.HideSelection = True
		Me.txtSector.ReadOnly = False
		Me.txtSector.Maxlength = 0
		Me.txtSector.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtSector.MultiLine = False
		Me.txtSector.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtSector.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtSector.TabStop = True
		Me.txtSector.Visible = True
		Me.txtSector.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtSector.Name = "txtSector"
		Me.optRange.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optRange.Text = "Range"
		Me.optRange.Size = New System.Drawing.Size(65, 17)
		Me.optRange.Location = New System.Drawing.Point(8, 88)
		Me.optRange.TabIndex = 5
		Me.optRange.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optRange.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optRange.BackColor = System.Drawing.SystemColors.Control
		Me.optRange.CausesValidation = True
		Me.optRange.Enabled = True
		Me.optRange.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optRange.Cursor = System.Windows.Forms.Cursors.Default
		Me.optRange.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optRange.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optRange.TabStop = True
		Me.optRange.Checked = False
		Me.optRange.Visible = True
		Me.optRange.Name = "optRange"
		Me.txtRealm.Size = New System.Drawing.Size(49, 21)
		Me.txtRealm.Location = New System.Drawing.Point(80, 56)
		Me.txtRealm.TabIndex = 4
		Me.txtRealm.Text = "Combo1"
		Me.txtRealm.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtRealm.BackColor = System.Drawing.SystemColors.Window
		Me.txtRealm.CausesValidation = True
		Me.txtRealm.Enabled = True
		Me.txtRealm.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtRealm.IntegralHeight = True
		Me.txtRealm.Cursor = System.Windows.Forms.Cursors.Default
		Me.txtRealm.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtRealm.Sorted = False
		Me.txtRealm.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.txtRealm.TabStop = True
		Me.txtRealm.Visible = True
		Me.txtRealm.Name = "txtRealm"
		Me.OptRealm.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.OptRealm.Text = "Realm #"
		Me.OptRealm.Size = New System.Drawing.Size(65, 17)
		Me.OptRealm.Location = New System.Drawing.Point(8, 56)
		Me.OptRealm.TabIndex = 3
		Me.OptRealm.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.OptRealm.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.OptRealm.BackColor = System.Drawing.SystemColors.Control
		Me.OptRealm.CausesValidation = True
		Me.OptRealm.Enabled = True
		Me.OptRealm.ForeColor = System.Drawing.SystemColors.ControlText
		Me.OptRealm.Cursor = System.Windows.Forms.Cursors.Default
		Me.OptRealm.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.OptRealm.Appearance = System.Windows.Forms.Appearance.Normal
		Me.OptRealm.TabStop = True
		Me.OptRealm.Checked = False
		Me.OptRealm.Visible = True
		Me.OptRealm.Name = "OptRealm"
		Me.optSect.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optSect.Text = "Sector"
		Me.optSect.Size = New System.Drawing.Size(65, 17)
		Me.optSect.Location = New System.Drawing.Point(8, 24)
		Me.optSect.TabIndex = 0
		Me.optSect.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optSect.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optSect.BackColor = System.Drawing.SystemColors.Control
		Me.optSect.CausesValidation = True
		Me.optSect.Enabled = True
		Me.optSect.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optSect.Cursor = System.Windows.Forms.Cursors.Default
		Me.optSect.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optSect.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optSect.TabStop = True
		Me.optSect.Checked = False
		Me.optSect.Visible = True
		Me.optSect.Name = "optSect"
		Me.Label2.Text = "To"
		Me.Label2.Size = New System.Drawing.Size(17, 17)
		Me.Label2.Location = New System.Drawing.Point(136, 88)
		Me.Label2.TabIndex = 8
		Me.Label2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
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
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(Frame2)
		Me.Controls.Add(Frame1)
		Me.Frame2.Controls.Add(txtParm)
		Me.Frame1.Controls.Add(cbField)
		Me.Frame1.Controls.Add(optCond)
		Me.Frame1.Controls.Add(txtTerr)
		Me.Frame1.Controls.Add(optTerr)
		Me.Frame1.Controls.Add(txtDest)
		Me.Frame1.Controls.Add(txtOrigin)
		Me.Frame1.Controls.Add(txtSector)
		Me.Frame1.Controls.Add(optRange)
		Me.Frame1.Controls.Add(txtRealm)
		Me.Frame1.Controls.Add(OptRealm)
		Me.Frame1.Controls.Add(optSect)
		Me.Frame1.Controls.Add(Label2)
		Me.Frame2.ResumeLayout(False)
		Me.Frame1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class