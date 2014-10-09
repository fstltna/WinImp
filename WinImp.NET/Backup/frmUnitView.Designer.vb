<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmUnitView
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
	Public WithEvents cbExpiry As System.Windows.Forms.ComboBox
	Public WithEvents _chkUnits_15 As System.Windows.Forms.CheckBox
	Public WithEvents _chkUnits_14 As System.Windows.Forms.CheckBox
	Public WithEvents _chkUnits_13 As System.Windows.Forms.CheckBox
	Public WithEvents lblExpiry As System.Windows.Forms.Label
	Public WithEvents frmExpiry As System.Windows.Forms.GroupBox
	Public WithEvents cmdApply As System.Windows.Forms.Button
	Public WithEvents optFriendlyNeutral As System.Windows.Forms.RadioButton
	Public WithEvents optFriendlyAlly As System.Windows.Forms.RadioButton
	Public WithEvents frmFriendly As System.Windows.Forms.GroupBox
	Public WithEvents cmdClearAll As System.Windows.Forms.Button
	Public WithEvents cmdSetAll As System.Windows.Forms.Button
	Public WithEvents lstSort As System.Windows.Forms.ListBox
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents cmdDown As System.Windows.Forms.Button
	Public WithEvents cmdUp As System.Windows.Forms.Button
	Public WithEvents _chkUnits_10 As System.Windows.Forms.CheckBox
	Public WithEvents _chkUnits_11 As System.Windows.Forms.CheckBox
	Public WithEvents _chkUnits_12 As System.Windows.Forms.CheckBox
	Public WithEvents frmNeutral As System.Windows.Forms.GroupBox
	Public WithEvents _chkUnits_9 As System.Windows.Forms.CheckBox
	Public WithEvents _chkUnits_8 As System.Windows.Forms.CheckBox
	Public WithEvents _chkUnits_7 As System.Windows.Forms.CheckBox
	Public WithEvents frmAlly As System.Windows.Forms.GroupBox
	Public WithEvents _chkUnits_4 As System.Windows.Forms.CheckBox
	Public WithEvents _chkUnits_5 As System.Windows.Forms.CheckBox
	Public WithEvents _chkUnits_6 As System.Windows.Forms.CheckBox
	Public WithEvents frmEnemy As System.Windows.Forms.GroupBox
	Public WithEvents _chkUnits_3 As System.Windows.Forms.CheckBox
	Public WithEvents _chkUnits_2 As System.Windows.Forms.CheckBox
	Public WithEvents _chkUnits_1 As System.Windows.Forms.CheckBox
	Public WithEvents _chkUnits_0 As System.Windows.Forms.CheckBox
	Public WithEvents frmOur As System.Windows.Forms.GroupBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents chkUnits As Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmUnitView))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.frmExpiry = New System.Windows.Forms.GroupBox
		Me.cbExpiry = New System.Windows.Forms.ComboBox
		Me._chkUnits_15 = New System.Windows.Forms.CheckBox
		Me._chkUnits_14 = New System.Windows.Forms.CheckBox
		Me._chkUnits_13 = New System.Windows.Forms.CheckBox
		Me.lblExpiry = New System.Windows.Forms.Label
		Me.cmdApply = New System.Windows.Forms.Button
		Me.frmFriendly = New System.Windows.Forms.GroupBox
		Me.optFriendlyNeutral = New System.Windows.Forms.RadioButton
		Me.optFriendlyAlly = New System.Windows.Forms.RadioButton
		Me.cmdClearAll = New System.Windows.Forms.Button
		Me.cmdSetAll = New System.Windows.Forms.Button
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.lstSort = New System.Windows.Forms.ListBox
		Me.cmdDown = New System.Windows.Forms.Button
		Me.cmdUp = New System.Windows.Forms.Button
		Me.frmNeutral = New System.Windows.Forms.GroupBox
		Me._chkUnits_10 = New System.Windows.Forms.CheckBox
		Me._chkUnits_11 = New System.Windows.Forms.CheckBox
		Me._chkUnits_12 = New System.Windows.Forms.CheckBox
		Me.frmAlly = New System.Windows.Forms.GroupBox
		Me._chkUnits_9 = New System.Windows.Forms.CheckBox
		Me._chkUnits_8 = New System.Windows.Forms.CheckBox
		Me._chkUnits_7 = New System.Windows.Forms.CheckBox
		Me.frmEnemy = New System.Windows.Forms.GroupBox
		Me._chkUnits_4 = New System.Windows.Forms.CheckBox
		Me._chkUnits_5 = New System.Windows.Forms.CheckBox
		Me._chkUnits_6 = New System.Windows.Forms.CheckBox
		Me.frmOur = New System.Windows.Forms.GroupBox
		Me._chkUnits_3 = New System.Windows.Forms.CheckBox
		Me._chkUnits_2 = New System.Windows.Forms.CheckBox
		Me._chkUnits_1 = New System.Windows.Forms.CheckBox
		Me._chkUnits_0 = New System.Windows.Forms.CheckBox
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me.chkUnits = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(components)
		Me.frmExpiry.SuspendLayout()
		Me.frmFriendly.SuspendLayout()
		Me.Frame1.SuspendLayout()
		Me.frmNeutral.SuspendLayout()
		Me.frmAlly.SuspendLayout()
		Me.frmEnemy.SuspendLayout()
		Me.frmOur.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.chkUnits, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Unit Viewing Options"
		Me.ClientSize = New System.Drawing.Size(593, 263)
		Me.Location = New System.Drawing.Point(184, 250)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmUnitView"
		Me.frmExpiry.Text = "Expired Units"
		Me.frmExpiry.Size = New System.Drawing.Size(129, 137)
		Me.frmExpiry.Location = New System.Drawing.Point(216, 8)
		Me.frmExpiry.TabIndex = 28
		Me.frmExpiry.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmExpiry.BackColor = System.Drawing.SystemColors.Control
		Me.frmExpiry.Enabled = True
		Me.frmExpiry.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmExpiry.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmExpiry.Visible = True
		Me.frmExpiry.Padding = New System.Windows.Forms.Padding(0)
		Me.frmExpiry.Name = "frmExpiry"
		Me.cbExpiry.Size = New System.Drawing.Size(113, 21)
		Me.cbExpiry.Location = New System.Drawing.Point(8, 104)
		Me.cbExpiry.Items.AddRange(New Object(){"No Expiry", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "30"})
		Me.cbExpiry.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cbExpiry.TabIndex = 33
		Me.cbExpiry.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbExpiry.BackColor = System.Drawing.SystemColors.Window
		Me.cbExpiry.CausesValidation = True
		Me.cbExpiry.Enabled = True
		Me.cbExpiry.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cbExpiry.IntegralHeight = True
		Me.cbExpiry.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbExpiry.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbExpiry.Sorted = False
		Me.cbExpiry.TabStop = True
		Me.cbExpiry.Visible = True
		Me.cbExpiry.Name = "cbExpiry"
		Me._chkUnits_15.Text = "Planes"
		Me._chkUnits_15.Size = New System.Drawing.Size(57, 17)
		Me._chkUnits_15.Location = New System.Drawing.Point(8, 64)
		Me._chkUnits_15.TabIndex = 31
		Me._chkUnits_15.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkUnits_15.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkUnits_15.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkUnits_15.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkUnits_15.BackColor = System.Drawing.SystemColors.Control
		Me._chkUnits_15.CausesValidation = True
		Me._chkUnits_15.Enabled = True
		Me._chkUnits_15.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkUnits_15.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkUnits_15.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkUnits_15.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkUnits_15.TabStop = True
		Me._chkUnits_15.Visible = True
		Me._chkUnits_15.Name = "_chkUnits_15"
		Me._chkUnits_14.Text = "Land Units"
		Me._chkUnits_14.Size = New System.Drawing.Size(73, 17)
		Me._chkUnits_14.Location = New System.Drawing.Point(8, 40)
		Me._chkUnits_14.TabIndex = 30
		Me._chkUnits_14.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkUnits_14.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkUnits_14.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkUnits_14.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkUnits_14.BackColor = System.Drawing.SystemColors.Control
		Me._chkUnits_14.CausesValidation = True
		Me._chkUnits_14.Enabled = True
		Me._chkUnits_14.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkUnits_14.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkUnits_14.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkUnits_14.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkUnits_14.TabStop = True
		Me._chkUnits_14.Visible = True
		Me._chkUnits_14.Name = "_chkUnits_14"
		Me._chkUnits_13.Text = "Ships"
		Me._chkUnits_13.Size = New System.Drawing.Size(73, 17)
		Me._chkUnits_13.Location = New System.Drawing.Point(8, 16)
		Me._chkUnits_13.TabIndex = 29
		Me._chkUnits_13.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkUnits_13.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkUnits_13.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkUnits_13.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkUnits_13.BackColor = System.Drawing.SystemColors.Control
		Me._chkUnits_13.CausesValidation = True
		Me._chkUnits_13.Enabled = True
		Me._chkUnits_13.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkUnits_13.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkUnits_13.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkUnits_13.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkUnits_13.TabStop = True
		Me._chkUnits_13.Visible = True
		Me._chkUnits_13.Name = "_chkUnits_13"
		Me.lblExpiry.Text = "Expiry after days"
		Me.lblExpiry.Size = New System.Drawing.Size(81, 17)
		Me.lblExpiry.Location = New System.Drawing.Point(8, 88)
		Me.lblExpiry.TabIndex = 32
		Me.lblExpiry.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblExpiry.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblExpiry.BackColor = System.Drawing.SystemColors.Control
		Me.lblExpiry.Enabled = True
		Me.lblExpiry.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblExpiry.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblExpiry.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblExpiry.UseMnemonic = True
		Me.lblExpiry.Visible = True
		Me.lblExpiry.AutoSize = False
		Me.lblExpiry.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblExpiry.Name = "lblExpiry"
		Me.cmdApply.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdApply.Text = "Test"
		Me.cmdApply.Size = New System.Drawing.Size(57, 25)
		Me.cmdApply.Location = New System.Drawing.Point(216, 224)
		Me.cmdApply.TabIndex = 27
		Me.cmdApply.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdApply.BackColor = System.Drawing.SystemColors.Control
		Me.cmdApply.CausesValidation = True
		Me.cmdApply.Enabled = True
		Me.cmdApply.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdApply.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdApply.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdApply.TabStop = True
		Me.cmdApply.Name = "cmdApply"
		Me.frmFriendly.Text = "Friendly Relations"
		Me.frmFriendly.Size = New System.Drawing.Size(129, 41)
		Me.frmFriendly.Location = New System.Drawing.Point(216, 176)
		Me.frmFriendly.TabIndex = 24
		Me.frmFriendly.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmFriendly.BackColor = System.Drawing.SystemColors.Control
		Me.frmFriendly.Enabled = True
		Me.frmFriendly.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmFriendly.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmFriendly.Visible = True
		Me.frmFriendly.Padding = New System.Windows.Forms.Padding(0)
		Me.frmFriendly.Name = "frmFriendly"
		Me.optFriendlyNeutral.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optFriendlyNeutral.Text = "Neutral"
		Me.optFriendlyNeutral.Size = New System.Drawing.Size(57, 17)
		Me.optFriendlyNeutral.Location = New System.Drawing.Point(64, 16)
		Me.optFriendlyNeutral.TabIndex = 26
		Me.optFriendlyNeutral.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optFriendlyNeutral.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optFriendlyNeutral.BackColor = System.Drawing.SystemColors.Control
		Me.optFriendlyNeutral.CausesValidation = True
		Me.optFriendlyNeutral.Enabled = True
		Me.optFriendlyNeutral.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optFriendlyNeutral.Cursor = System.Windows.Forms.Cursors.Default
		Me.optFriendlyNeutral.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optFriendlyNeutral.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optFriendlyNeutral.TabStop = True
		Me.optFriendlyNeutral.Checked = False
		Me.optFriendlyNeutral.Visible = True
		Me.optFriendlyNeutral.Name = "optFriendlyNeutral"
		Me.optFriendlyAlly.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optFriendlyAlly.Text = "Allied"
		Me.optFriendlyAlly.Size = New System.Drawing.Size(49, 17)
		Me.optFriendlyAlly.Location = New System.Drawing.Point(8, 16)
		Me.optFriendlyAlly.TabIndex = 25
		Me.optFriendlyAlly.Checked = True
		Me.optFriendlyAlly.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optFriendlyAlly.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optFriendlyAlly.BackColor = System.Drawing.SystemColors.Control
		Me.optFriendlyAlly.CausesValidation = True
		Me.optFriendlyAlly.Enabled = True
		Me.optFriendlyAlly.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optFriendlyAlly.Cursor = System.Windows.Forms.Cursors.Default
		Me.optFriendlyAlly.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optFriendlyAlly.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optFriendlyAlly.TabStop = True
		Me.optFriendlyAlly.Visible = True
		Me.optFriendlyAlly.Name = "optFriendlyAlly"
		Me.cmdClearAll.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdClearAll.Text = "Clear All"
		Me.cmdClearAll.Size = New System.Drawing.Size(57, 25)
		Me.cmdClearAll.Location = New System.Drawing.Point(64, 224)
		Me.cmdClearAll.TabIndex = 23
		Me.cmdClearAll.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdClearAll.BackColor = System.Drawing.SystemColors.Control
		Me.cmdClearAll.CausesValidation = True
		Me.cmdClearAll.Enabled = True
		Me.cmdClearAll.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdClearAll.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdClearAll.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdClearAll.TabStop = True
		Me.cmdClearAll.Name = "cmdClearAll"
		Me.cmdSetAll.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdSetAll.Text = "Set All"
		Me.cmdSetAll.Size = New System.Drawing.Size(49, 25)
		Me.cmdSetAll.Location = New System.Drawing.Point(8, 224)
		Me.cmdSetAll.TabIndex = 22
		Me.cmdSetAll.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdSetAll.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSetAll.CausesValidation = True
		Me.cmdSetAll.Enabled = True
		Me.cmdSetAll.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSetAll.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSetAll.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSetAll.TabStop = True
		Me.cmdSetAll.Name = "cmdSetAll"
		Me.Frame1.Text = "Priority List"
		Me.Frame1.Size = New System.Drawing.Size(129, 225)
		Me.Frame1.Location = New System.Drawing.Point(360, 24)
		Me.Frame1.TabIndex = 20
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.lstSort.Size = New System.Drawing.Size(113, 202)
		Me.lstSort.Location = New System.Drawing.Point(8, 16)
		Me.lstSort.TabIndex = 21
		Me.lstSort.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstSort.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstSort.BackColor = System.Drawing.SystemColors.Window
		Me.lstSort.CausesValidation = True
		Me.lstSort.Enabled = True
		Me.lstSort.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstSort.IntegralHeight = True
		Me.lstSort.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstSort.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstSort.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstSort.Sorted = False
		Me.lstSort.TabStop = True
		Me.lstSort.Visible = True
		Me.lstSort.MultiColumn = False
		Me.lstSort.Name = "lstSort"
		Me.cmdDown.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		Me.cmdDown.Size = New System.Drawing.Size(41, 41)
		Me.cmdDown.Location = New System.Drawing.Point(496, 128)
		Me.cmdDown.Image = CType(resources.GetObject("cmdDown.Image"), System.Drawing.Image)
		Me.cmdDown.TabIndex = 19
		Me.ToolTip1.SetToolTip(Me.cmdDown, "Lower the Priority")
		Me.cmdDown.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdDown.BackColor = System.Drawing.SystemColors.Control
		Me.cmdDown.CausesValidation = True
		Me.cmdDown.Enabled = True
		Me.cmdDown.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdDown.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdDown.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdDown.TabStop = True
		Me.cmdDown.Name = "cmdDown"
		Me.cmdUp.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		Me.cmdUp.Size = New System.Drawing.Size(41, 41)
		Me.cmdUp.Location = New System.Drawing.Point(496, 72)
		Me.cmdUp.Image = CType(resources.GetObject("cmdUp.Image"), System.Drawing.Image)
		Me.cmdUp.TabIndex = 18
		Me.ToolTip1.SetToolTip(Me.cmdUp, "Raise the Priority")
		Me.cmdUp.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdUp.BackColor = System.Drawing.SystemColors.Control
		Me.cmdUp.CausesValidation = True
		Me.cmdUp.Enabled = True
		Me.cmdUp.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdUp.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdUp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdUp.TabStop = True
		Me.cmdUp.Name = "cmdUp"
		Me.frmNeutral.Text = "Neutral Units"
		Me.frmNeutral.Size = New System.Drawing.Size(89, 89)
		Me.frmNeutral.Location = New System.Drawing.Point(112, 128)
		Me.frmNeutral.TabIndex = 14
		Me.frmNeutral.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmNeutral.BackColor = System.Drawing.SystemColors.Control
		Me.frmNeutral.Enabled = True
		Me.frmNeutral.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmNeutral.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmNeutral.Visible = True
		Me.frmNeutral.Padding = New System.Windows.Forms.Padding(0)
		Me.frmNeutral.Name = "frmNeutral"
		Me._chkUnits_10.Text = "Ships"
		Me._chkUnits_10.Size = New System.Drawing.Size(73, 17)
		Me._chkUnits_10.Location = New System.Drawing.Point(8, 16)
		Me._chkUnits_10.TabIndex = 17
		Me._chkUnits_10.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkUnits_10.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkUnits_10.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkUnits_10.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkUnits_10.BackColor = System.Drawing.SystemColors.Control
		Me._chkUnits_10.CausesValidation = True
		Me._chkUnits_10.Enabled = True
		Me._chkUnits_10.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkUnits_10.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkUnits_10.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkUnits_10.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkUnits_10.TabStop = True
		Me._chkUnits_10.Visible = True
		Me._chkUnits_10.Name = "_chkUnits_10"
		Me._chkUnits_11.Text = "Land Units"
		Me._chkUnits_11.Size = New System.Drawing.Size(73, 17)
		Me._chkUnits_11.Location = New System.Drawing.Point(8, 40)
		Me._chkUnits_11.TabIndex = 16
		Me._chkUnits_11.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkUnits_11.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkUnits_11.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkUnits_11.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkUnits_11.BackColor = System.Drawing.SystemColors.Control
		Me._chkUnits_11.CausesValidation = True
		Me._chkUnits_11.Enabled = True
		Me._chkUnits_11.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkUnits_11.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkUnits_11.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkUnits_11.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkUnits_11.TabStop = True
		Me._chkUnits_11.Visible = True
		Me._chkUnits_11.Name = "_chkUnits_11"
		Me._chkUnits_12.Text = "Planes"
		Me._chkUnits_12.Size = New System.Drawing.Size(73, 17)
		Me._chkUnits_12.Location = New System.Drawing.Point(8, 64)
		Me._chkUnits_12.TabIndex = 15
		Me._chkUnits_12.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkUnits_12.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkUnits_12.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkUnits_12.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkUnits_12.BackColor = System.Drawing.SystemColors.Control
		Me._chkUnits_12.CausesValidation = True
		Me._chkUnits_12.Enabled = True
		Me._chkUnits_12.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkUnits_12.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkUnits_12.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkUnits_12.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkUnits_12.TabStop = True
		Me._chkUnits_12.Visible = True
		Me._chkUnits_12.Name = "_chkUnits_12"
		Me.frmAlly.Text = "Allied Units"
		Me.frmAlly.Size = New System.Drawing.Size(89, 89)
		Me.frmAlly.Location = New System.Drawing.Point(112, 8)
		Me.frmAlly.TabIndex = 10
		Me.frmAlly.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmAlly.BackColor = System.Drawing.SystemColors.Control
		Me.frmAlly.Enabled = True
		Me.frmAlly.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmAlly.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmAlly.Visible = True
		Me.frmAlly.Padding = New System.Windows.Forms.Padding(0)
		Me.frmAlly.Name = "frmAlly"
		Me._chkUnits_9.Text = "Planes"
		Me._chkUnits_9.Size = New System.Drawing.Size(73, 17)
		Me._chkUnits_9.Location = New System.Drawing.Point(8, 64)
		Me._chkUnits_9.TabIndex = 13
		Me._chkUnits_9.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkUnits_9.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkUnits_9.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkUnits_9.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkUnits_9.BackColor = System.Drawing.SystemColors.Control
		Me._chkUnits_9.CausesValidation = True
		Me._chkUnits_9.Enabled = True
		Me._chkUnits_9.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkUnits_9.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkUnits_9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkUnits_9.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkUnits_9.TabStop = True
		Me._chkUnits_9.Visible = True
		Me._chkUnits_9.Name = "_chkUnits_9"
		Me._chkUnits_8.Text = "Land Units"
		Me._chkUnits_8.Size = New System.Drawing.Size(73, 17)
		Me._chkUnits_8.Location = New System.Drawing.Point(8, 40)
		Me._chkUnits_8.TabIndex = 12
		Me._chkUnits_8.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkUnits_8.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkUnits_8.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkUnits_8.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkUnits_8.BackColor = System.Drawing.SystemColors.Control
		Me._chkUnits_8.CausesValidation = True
		Me._chkUnits_8.Enabled = True
		Me._chkUnits_8.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkUnits_8.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkUnits_8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkUnits_8.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkUnits_8.TabStop = True
		Me._chkUnits_8.Visible = True
		Me._chkUnits_8.Name = "_chkUnits_8"
		Me._chkUnits_7.Text = "Ships"
		Me._chkUnits_7.Size = New System.Drawing.Size(73, 17)
		Me._chkUnits_7.Location = New System.Drawing.Point(8, 16)
		Me._chkUnits_7.TabIndex = 11
		Me._chkUnits_7.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkUnits_7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkUnits_7.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkUnits_7.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkUnits_7.BackColor = System.Drawing.SystemColors.Control
		Me._chkUnits_7.CausesValidation = True
		Me._chkUnits_7.Enabled = True
		Me._chkUnits_7.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkUnits_7.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkUnits_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkUnits_7.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkUnits_7.TabStop = True
		Me._chkUnits_7.Visible = True
		Me._chkUnits_7.Name = "_chkUnits_7"
		Me.frmEnemy.Text = "Enemy Units"
		Me.frmEnemy.Size = New System.Drawing.Size(89, 89)
		Me.frmEnemy.Location = New System.Drawing.Point(8, 128)
		Me.frmEnemy.TabIndex = 6
		Me.frmEnemy.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmEnemy.BackColor = System.Drawing.SystemColors.Control
		Me.frmEnemy.Enabled = True
		Me.frmEnemy.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmEnemy.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmEnemy.Visible = True
		Me.frmEnemy.Padding = New System.Windows.Forms.Padding(0)
		Me.frmEnemy.Name = "frmEnemy"
		Me._chkUnits_4.Text = "Ships"
		Me._chkUnits_4.Size = New System.Drawing.Size(73, 17)
		Me._chkUnits_4.Location = New System.Drawing.Point(8, 16)
		Me._chkUnits_4.TabIndex = 9
		Me._chkUnits_4.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkUnits_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkUnits_4.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkUnits_4.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkUnits_4.BackColor = System.Drawing.SystemColors.Control
		Me._chkUnits_4.CausesValidation = True
		Me._chkUnits_4.Enabled = True
		Me._chkUnits_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkUnits_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkUnits_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkUnits_4.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkUnits_4.TabStop = True
		Me._chkUnits_4.Visible = True
		Me._chkUnits_4.Name = "_chkUnits_4"
		Me._chkUnits_5.Text = "Land Units"
		Me._chkUnits_5.Size = New System.Drawing.Size(73, 17)
		Me._chkUnits_5.Location = New System.Drawing.Point(8, 40)
		Me._chkUnits_5.TabIndex = 8
		Me._chkUnits_5.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkUnits_5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkUnits_5.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkUnits_5.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkUnits_5.BackColor = System.Drawing.SystemColors.Control
		Me._chkUnits_5.CausesValidation = True
		Me._chkUnits_5.Enabled = True
		Me._chkUnits_5.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkUnits_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkUnits_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkUnits_5.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkUnits_5.TabStop = True
		Me._chkUnits_5.Visible = True
		Me._chkUnits_5.Name = "_chkUnits_5"
		Me._chkUnits_6.Text = "Planes"
		Me._chkUnits_6.Size = New System.Drawing.Size(73, 17)
		Me._chkUnits_6.Location = New System.Drawing.Point(8, 64)
		Me._chkUnits_6.TabIndex = 7
		Me._chkUnits_6.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkUnits_6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkUnits_6.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkUnits_6.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkUnits_6.BackColor = System.Drawing.SystemColors.Control
		Me._chkUnits_6.CausesValidation = True
		Me._chkUnits_6.Enabled = True
		Me._chkUnits_6.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkUnits_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkUnits_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkUnits_6.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkUnits_6.TabStop = True
		Me._chkUnits_6.Visible = True
		Me._chkUnits_6.Name = "_chkUnits_6"
		Me.frmOur.Text = "Our Units"
		Me.frmOur.Size = New System.Drawing.Size(89, 113)
		Me.frmOur.Location = New System.Drawing.Point(8, 8)
		Me.frmOur.TabIndex = 2
		Me.frmOur.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmOur.BackColor = System.Drawing.SystemColors.Control
		Me.frmOur.Enabled = True
		Me.frmOur.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmOur.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmOur.Visible = True
		Me.frmOur.Padding = New System.Windows.Forms.Padding(0)
		Me.frmOur.Name = "frmOur"
		Me._chkUnits_3.Text = "Nukes"
		Me._chkUnits_3.Size = New System.Drawing.Size(73, 17)
		Me._chkUnits_3.Location = New System.Drawing.Point(8, 88)
		Me._chkUnits_3.TabIndex = 34
		Me._chkUnits_3.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkUnits_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkUnits_3.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkUnits_3.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkUnits_3.BackColor = System.Drawing.SystemColors.Control
		Me._chkUnits_3.CausesValidation = True
		Me._chkUnits_3.Enabled = True
		Me._chkUnits_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkUnits_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkUnits_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkUnits_3.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkUnits_3.TabStop = True
		Me._chkUnits_3.Visible = True
		Me._chkUnits_3.Name = "_chkUnits_3"
		Me._chkUnits_2.Text = "Planes"
		Me._chkUnits_2.Size = New System.Drawing.Size(73, 17)
		Me._chkUnits_2.Location = New System.Drawing.Point(8, 64)
		Me._chkUnits_2.TabIndex = 5
		Me._chkUnits_2.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkUnits_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkUnits_2.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkUnits_2.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkUnits_2.BackColor = System.Drawing.SystemColors.Control
		Me._chkUnits_2.CausesValidation = True
		Me._chkUnits_2.Enabled = True
		Me._chkUnits_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkUnits_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkUnits_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkUnits_2.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkUnits_2.TabStop = True
		Me._chkUnits_2.Visible = True
		Me._chkUnits_2.Name = "_chkUnits_2"
		Me._chkUnits_1.Text = "Land Units"
		Me._chkUnits_1.Size = New System.Drawing.Size(73, 17)
		Me._chkUnits_1.Location = New System.Drawing.Point(8, 40)
		Me._chkUnits_1.TabIndex = 4
		Me._chkUnits_1.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkUnits_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkUnits_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkUnits_1.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkUnits_1.BackColor = System.Drawing.SystemColors.Control
		Me._chkUnits_1.CausesValidation = True
		Me._chkUnits_1.Enabled = True
		Me._chkUnits_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkUnits_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkUnits_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkUnits_1.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkUnits_1.TabStop = True
		Me._chkUnits_1.Visible = True
		Me._chkUnits_1.Name = "_chkUnits_1"
		Me._chkUnits_0.Text = "Ships"
		Me._chkUnits_0.Size = New System.Drawing.Size(73, 17)
		Me._chkUnits_0.Location = New System.Drawing.Point(8, 16)
		Me._chkUnits_0.TabIndex = 3
		Me._chkUnits_0.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkUnits_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkUnits_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkUnits_0.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkUnits_0.BackColor = System.Drawing.SystemColors.Control
		Me._chkUnits_0.CausesValidation = True
		Me._chkUnits_0.Enabled = True
		Me._chkUnits_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkUnits_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkUnits_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkUnits_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkUnits_0.TabStop = True
		Me._chkUnits_0.Visible = True
		Me._chkUnits_0.Name = "_chkUnits_0"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(65, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(136, 224)
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
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "OK"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Size = New System.Drawing.Size(57, 25)
		Me.cmdOK.Location = New System.Drawing.Point(288, 224)
		Me.cmdOK.TabIndex = 0
		Me.cmdOK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.Controls.Add(frmExpiry)
		Me.Controls.Add(cmdApply)
		Me.Controls.Add(frmFriendly)
		Me.Controls.Add(cmdClearAll)
		Me.Controls.Add(cmdSetAll)
		Me.Controls.Add(Frame1)
		Me.Controls.Add(cmdDown)
		Me.Controls.Add(cmdUp)
		Me.Controls.Add(frmNeutral)
		Me.Controls.Add(frmAlly)
		Me.Controls.Add(frmEnemy)
		Me.Controls.Add(frmOur)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOK)
		Me.frmExpiry.Controls.Add(cbExpiry)
		Me.frmExpiry.Controls.Add(_chkUnits_15)
		Me.frmExpiry.Controls.Add(_chkUnits_14)
		Me.frmExpiry.Controls.Add(_chkUnits_13)
		Me.frmExpiry.Controls.Add(lblExpiry)
		Me.frmFriendly.Controls.Add(optFriendlyNeutral)
		Me.frmFriendly.Controls.Add(optFriendlyAlly)
		Me.Frame1.Controls.Add(lstSort)
		Me.frmNeutral.Controls.Add(_chkUnits_10)
		Me.frmNeutral.Controls.Add(_chkUnits_11)
		Me.frmNeutral.Controls.Add(_chkUnits_12)
		Me.frmAlly.Controls.Add(_chkUnits_9)
		Me.frmAlly.Controls.Add(_chkUnits_8)
		Me.frmAlly.Controls.Add(_chkUnits_7)
		Me.frmEnemy.Controls.Add(_chkUnits_4)
		Me.frmEnemy.Controls.Add(_chkUnits_5)
		Me.frmEnemy.Controls.Add(_chkUnits_6)
		Me.frmOur.Controls.Add(_chkUnits_3)
		Me.frmOur.Controls.Add(_chkUnits_2)
		Me.frmOur.Controls.Add(_chkUnits_1)
		Me.frmOur.Controls.Add(_chkUnits_0)
		Me.chkUnits.SetIndex(_chkUnits_15, CType(15, Short))
		Me.chkUnits.SetIndex(_chkUnits_14, CType(14, Short))
		Me.chkUnits.SetIndex(_chkUnits_13, CType(13, Short))
		Me.chkUnits.SetIndex(_chkUnits_10, CType(10, Short))
		Me.chkUnits.SetIndex(_chkUnits_11, CType(11, Short))
		Me.chkUnits.SetIndex(_chkUnits_12, CType(12, Short))
		Me.chkUnits.SetIndex(_chkUnits_9, CType(9, Short))
		Me.chkUnits.SetIndex(_chkUnits_8, CType(8, Short))
		Me.chkUnits.SetIndex(_chkUnits_7, CType(7, Short))
		Me.chkUnits.SetIndex(_chkUnits_4, CType(4, Short))
		Me.chkUnits.SetIndex(_chkUnits_5, CType(5, Short))
		Me.chkUnits.SetIndex(_chkUnits_6, CType(6, Short))
		Me.chkUnits.SetIndex(_chkUnits_3, CType(3, Short))
		Me.chkUnits.SetIndex(_chkUnits_2, CType(2, Short))
		Me.chkUnits.SetIndex(_chkUnits_1, CType(1, Short))
		Me.chkUnits.SetIndex(_chkUnits_0, CType(0, Short))
		CType(Me.chkUnits, System.ComponentModel.ISupportInitialize).EndInit()
		Me.frmExpiry.ResumeLayout(False)
		Me.frmFriendly.ResumeLayout(False)
		Me.Frame1.ResumeLayout(False)
		Me.frmNeutral.ResumeLayout(False)
		Me.frmAlly.ResumeLayout(False)
		Me.frmEnemy.ResumeLayout(False)
		Me.frmOur.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class