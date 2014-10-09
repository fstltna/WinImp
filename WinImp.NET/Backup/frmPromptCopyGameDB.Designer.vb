<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPromptCopyGameDB
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
	Public WithEvents cbGameName As System.Windows.Forms.ComboBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents _chkItems_10 As System.Windows.Forms.CheckBox
	Public WithEvents _chkItems_0 As System.Windows.Forms.CheckBox
	Public WithEvents _chkItems_1 As System.Windows.Forms.CheckBox
	Public WithEvents _chkItems_2 As System.Windows.Forms.CheckBox
	Public WithEvents _chkItems_3 As System.Windows.Forms.CheckBox
	Public WithEvents _chkItems_4 As System.Windows.Forms.CheckBox
	Public WithEvents _chkItems_5 As System.Windows.Forms.CheckBox
	Public WithEvents _chkItems_6 As System.Windows.Forms.CheckBox
	Public WithEvents _chkItems_7 As System.Windows.Forms.CheckBox
	Public WithEvents _chkItems_8 As System.Windows.Forms.CheckBox
	Public WithEvents _chkItems_9 As System.Windows.Forms.CheckBox
	Public WithEvents frameType As System.Windows.Forms.GroupBox
	Public WithEvents txtOffsetX As System.Windows.Forms.TextBox
	Public WithEvents txtOffsetY As System.Windows.Forms.TextBox
	Public WithEvents lblOffsetX As System.Windows.Forms.Label
	Public WithEvents lblOffsetY As System.Windows.Forms.Label
	Public WithEvents frameOffset As System.Windows.Forms.GroupBox
	Public WithEvents lblTitle As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents chkItems As Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPromptCopyGameDB))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cbGameName = New System.Windows.Forms.ComboBox
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me.frameType = New System.Windows.Forms.GroupBox
		Me._chkItems_10 = New System.Windows.Forms.CheckBox
		Me._chkItems_0 = New System.Windows.Forms.CheckBox
		Me._chkItems_1 = New System.Windows.Forms.CheckBox
		Me._chkItems_2 = New System.Windows.Forms.CheckBox
		Me._chkItems_3 = New System.Windows.Forms.CheckBox
		Me._chkItems_4 = New System.Windows.Forms.CheckBox
		Me._chkItems_5 = New System.Windows.Forms.CheckBox
		Me._chkItems_6 = New System.Windows.Forms.CheckBox
		Me._chkItems_7 = New System.Windows.Forms.CheckBox
		Me._chkItems_8 = New System.Windows.Forms.CheckBox
		Me._chkItems_9 = New System.Windows.Forms.CheckBox
		Me.frameOffset = New System.Windows.Forms.GroupBox
		Me.txtOffsetX = New System.Windows.Forms.TextBox
		Me.txtOffsetY = New System.Windows.Forms.TextBox
		Me.lblOffsetX = New System.Windows.Forms.Label
		Me.lblOffsetY = New System.Windows.Forms.Label
		Me.lblTitle = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.chkItems = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(components)
		Me.frameType.SuspendLayout()
		Me.frameOffset.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.chkItems, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.ControlBox = False
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.ClientSize = New System.Drawing.Size(536, 144)
		Me.Location = New System.Drawing.Point(4, 4)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmPromptCopyGameDB"
		Me.cbGameName.Size = New System.Drawing.Size(185, 21)
		Me.cbGameName.Location = New System.Drawing.Point(8, 64)
		Me.cbGameName.TabIndex = 1
		Me.cbGameName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbGameName.BackColor = System.Drawing.SystemColors.Window
		Me.cbGameName.CausesValidation = True
		Me.cbGameName.Enabled = True
		Me.cbGameName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cbGameName.IntegralHeight = True
		Me.cbGameName.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbGameName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbGameName.Sorted = False
		Me.cbGameName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cbGameName.TabStop = True
		Me.cbGameName.Visible = True
		Me.cbGameName.Name = "cbGameName"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "Cancel"
		Me.AcceptButton = Me.cmdCancel
		Me.cmdCancel.Size = New System.Drawing.Size(57, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(464, 80)
		Me.cmdCancel.TabIndex = 18
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
		Me.cmdOK.Text = "Copy"
		Me.cmdOK.Enabled = False
		Me.cmdOK.Size = New System.Drawing.Size(57, 25)
		Me.cmdOK.Location = New System.Drawing.Point(464, 40)
		Me.cmdOK.TabIndex = 17
		Me.cmdOK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.frameType.Text = "Items"
		Me.frameType.Size = New System.Drawing.Size(241, 129)
		Me.frameType.Location = New System.Drawing.Point(208, 8)
		Me.frameType.TabIndex = 5
		Me.frameType.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frameType.BackColor = System.Drawing.SystemColors.Control
		Me.frameType.Enabled = True
		Me.frameType.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frameType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frameType.Visible = True
		Me.frameType.Padding = New System.Windows.Forms.Padding(0)
		Me.frameType.Name = "frameType"
		Me._chkItems_10.Text = "Nation"
		Me._chkItems_10.Size = New System.Drawing.Size(113, 17)
		Me._chkItems_10.Location = New System.Drawing.Point(8, 103)
		Me._chkItems_10.TabIndex = 11
		Me._chkItems_10.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkItems_10.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkItems_10.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkItems_10.BackColor = System.Drawing.SystemColors.Control
		Me._chkItems_10.CausesValidation = True
		Me._chkItems_10.Enabled = True
		Me._chkItems_10.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkItems_10.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkItems_10.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkItems_10.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkItems_10.TabStop = True
		Me._chkItems_10.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._chkItems_10.Visible = True
		Me._chkItems_10.Name = "_chkItems_10"
		Me._chkItems_0.Text = "Enemy Sectors"
		Me._chkItems_0.Size = New System.Drawing.Size(113, 17)
		Me._chkItems_0.Location = New System.Drawing.Point(120, 16)
		Me._chkItems_0.TabIndex = 12
		Me._chkItems_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkItems_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkItems_0.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkItems_0.BackColor = System.Drawing.SystemColors.Control
		Me._chkItems_0.CausesValidation = True
		Me._chkItems_0.Enabled = True
		Me._chkItems_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkItems_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkItems_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkItems_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkItems_0.TabStop = True
		Me._chkItems_0.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._chkItems_0.Visible = True
		Me._chkItems_0.Name = "_chkItems_0"
		Me._chkItems_1.Text = "Enemy Ships"
		Me._chkItems_1.Size = New System.Drawing.Size(113, 17)
		Me._chkItems_1.Location = New System.Drawing.Point(120, 34)
		Me._chkItems_1.TabIndex = 13
		Me._chkItems_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkItems_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkItems_1.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkItems_1.BackColor = System.Drawing.SystemColors.Control
		Me._chkItems_1.CausesValidation = True
		Me._chkItems_1.Enabled = True
		Me._chkItems_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkItems_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkItems_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkItems_1.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkItems_1.TabStop = True
		Me._chkItems_1.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._chkItems_1.Visible = True
		Me._chkItems_1.Name = "_chkItems_1"
		Me._chkItems_2.Text = "Enemy Planes"
		Me._chkItems_2.Size = New System.Drawing.Size(113, 17)
		Me._chkItems_2.Location = New System.Drawing.Point(120, 51)
		Me._chkItems_2.TabIndex = 14
		Me._chkItems_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkItems_2.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkItems_2.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkItems_2.BackColor = System.Drawing.SystemColors.Control
		Me._chkItems_2.CausesValidation = True
		Me._chkItems_2.Enabled = True
		Me._chkItems_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkItems_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkItems_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkItems_2.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkItems_2.TabStop = True
		Me._chkItems_2.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._chkItems_2.Visible = True
		Me._chkItems_2.Name = "_chkItems_2"
		Me._chkItems_3.Text = "Enemy Land Units"
		Me._chkItems_3.Size = New System.Drawing.Size(113, 17)
		Me._chkItems_3.Location = New System.Drawing.Point(120, 68)
		Me._chkItems_3.TabIndex = 15
		Me._chkItems_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkItems_3.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkItems_3.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkItems_3.BackColor = System.Drawing.SystemColors.Control
		Me._chkItems_3.CausesValidation = True
		Me._chkItems_3.Enabled = True
		Me._chkItems_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkItems_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkItems_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkItems_3.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkItems_3.TabStop = True
		Me._chkItems_3.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._chkItems_3.Visible = True
		Me._chkItems_3.Name = "_chkItems_3"
		Me._chkItems_4.Text = "Enemy Nukes"
		Me._chkItems_4.Enabled = False
		Me._chkItems_4.Size = New System.Drawing.Size(113, 17)
		Me._chkItems_4.Location = New System.Drawing.Point(120, 86)
		Me._chkItems_4.TabIndex = 16
		Me._chkItems_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkItems_4.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkItems_4.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkItems_4.BackColor = System.Drawing.SystemColors.Control
		Me._chkItems_4.CausesValidation = True
		Me._chkItems_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkItems_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkItems_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkItems_4.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkItems_4.TabStop = True
		Me._chkItems_4.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._chkItems_4.Visible = True
		Me._chkItems_4.Name = "_chkItems_4"
		Me._chkItems_5.Text = "Owner Sectors"
		Me._chkItems_5.Size = New System.Drawing.Size(113, 17)
		Me._chkItems_5.Location = New System.Drawing.Point(8, 16)
		Me._chkItems_5.TabIndex = 6
		Me._chkItems_5.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkItems_5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkItems_5.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkItems_5.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkItems_5.BackColor = System.Drawing.SystemColors.Control
		Me._chkItems_5.CausesValidation = True
		Me._chkItems_5.Enabled = True
		Me._chkItems_5.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkItems_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkItems_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkItems_5.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkItems_5.TabStop = True
		Me._chkItems_5.Visible = True
		Me._chkItems_5.Name = "_chkItems_5"
		Me._chkItems_6.Text = "Owner Ships"
		Me._chkItems_6.Size = New System.Drawing.Size(113, 17)
		Me._chkItems_6.Location = New System.Drawing.Point(8, 34)
		Me._chkItems_6.TabIndex = 7
		Me._chkItems_6.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkItems_6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkItems_6.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkItems_6.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkItems_6.BackColor = System.Drawing.SystemColors.Control
		Me._chkItems_6.CausesValidation = True
		Me._chkItems_6.Enabled = True
		Me._chkItems_6.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkItems_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkItems_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkItems_6.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkItems_6.TabStop = True
		Me._chkItems_6.Visible = True
		Me._chkItems_6.Name = "_chkItems_6"
		Me._chkItems_7.Text = "Owner Planes"
		Me._chkItems_7.Size = New System.Drawing.Size(113, 17)
		Me._chkItems_7.Location = New System.Drawing.Point(8, 51)
		Me._chkItems_7.TabIndex = 8
		Me._chkItems_7.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkItems_7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkItems_7.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkItems_7.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkItems_7.BackColor = System.Drawing.SystemColors.Control
		Me._chkItems_7.CausesValidation = True
		Me._chkItems_7.Enabled = True
		Me._chkItems_7.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkItems_7.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkItems_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkItems_7.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkItems_7.TabStop = True
		Me._chkItems_7.Visible = True
		Me._chkItems_7.Name = "_chkItems_7"
		Me._chkItems_8.Text = "Owner Land Units"
		Me._chkItems_8.Size = New System.Drawing.Size(113, 17)
		Me._chkItems_8.Location = New System.Drawing.Point(8, 68)
		Me._chkItems_8.TabIndex = 9
		Me._chkItems_8.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkItems_8.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkItems_8.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkItems_8.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkItems_8.BackColor = System.Drawing.SystemColors.Control
		Me._chkItems_8.CausesValidation = True
		Me._chkItems_8.Enabled = True
		Me._chkItems_8.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkItems_8.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkItems_8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkItems_8.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkItems_8.TabStop = True
		Me._chkItems_8.Visible = True
		Me._chkItems_8.Name = "_chkItems_8"
		Me._chkItems_9.Text = "Owner Nukes"
		Me._chkItems_9.Enabled = False
		Me._chkItems_9.Size = New System.Drawing.Size(113, 17)
		Me._chkItems_9.Location = New System.Drawing.Point(8, 86)
		Me._chkItems_9.TabIndex = 10
		Me._chkItems_9.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkItems_9.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkItems_9.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkItems_9.BackColor = System.Drawing.SystemColors.Control
		Me._chkItems_9.CausesValidation = True
		Me._chkItems_9.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkItems_9.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkItems_9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkItems_9.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkItems_9.TabStop = True
		Me._chkItems_9.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._chkItems_9.Visible = True
		Me._chkItems_9.Name = "_chkItems_9"
		Me.frameOffset.Text = "Offset"
		Me.frameOffset.Size = New System.Drawing.Size(185, 41)
		Me.frameOffset.Location = New System.Drawing.Point(8, 96)
		Me.frameOffset.TabIndex = 2
		Me.frameOffset.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frameOffset.BackColor = System.Drawing.SystemColors.Control
		Me.frameOffset.Enabled = True
		Me.frameOffset.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frameOffset.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frameOffset.Visible = True
		Me.frameOffset.Padding = New System.Windows.Forms.Padding(0)
		Me.frameOffset.Name = "frameOffset"
		Me.txtOffsetX.AutoSize = False
		Me.txtOffsetX.Size = New System.Drawing.Size(33, 19)
		Me.txtOffsetX.Location = New System.Drawing.Point(56, 16)
		Me.txtOffsetX.TabIndex = 3
		Me.txtOffsetX.Text = "0"
		Me.txtOffsetX.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtOffsetX.AcceptsReturn = True
		Me.txtOffsetX.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtOffsetX.BackColor = System.Drawing.SystemColors.Window
		Me.txtOffsetX.CausesValidation = True
		Me.txtOffsetX.Enabled = True
		Me.txtOffsetX.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtOffsetX.HideSelection = True
		Me.txtOffsetX.ReadOnly = False
		Me.txtOffsetX.Maxlength = 0
		Me.txtOffsetX.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtOffsetX.MultiLine = False
		Me.txtOffsetX.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtOffsetX.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtOffsetX.TabStop = True
		Me.txtOffsetX.Visible = True
		Me.txtOffsetX.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtOffsetX.Name = "txtOffsetX"
		Me.txtOffsetY.AutoSize = False
		Me.txtOffsetY.Size = New System.Drawing.Size(33, 19)
		Me.txtOffsetY.Location = New System.Drawing.Point(144, 16)
		Me.txtOffsetY.TabIndex = 4
		Me.txtOffsetY.Text = "0"
		Me.txtOffsetY.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtOffsetY.AcceptsReturn = True
		Me.txtOffsetY.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtOffsetY.BackColor = System.Drawing.SystemColors.Window
		Me.txtOffsetY.CausesValidation = True
		Me.txtOffsetY.Enabled = True
		Me.txtOffsetY.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtOffsetY.HideSelection = True
		Me.txtOffsetY.ReadOnly = False
		Me.txtOffsetY.Maxlength = 0
		Me.txtOffsetY.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtOffsetY.MultiLine = False
		Me.txtOffsetY.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtOffsetY.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtOffsetY.TabStop = True
		Me.txtOffsetY.Visible = True
		Me.txtOffsetY.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtOffsetY.Name = "txtOffsetY"
		Me.lblOffsetX.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblOffsetX.Text = "X Offset"
		Me.lblOffsetX.Size = New System.Drawing.Size(41, 17)
		Me.lblOffsetX.Location = New System.Drawing.Point(8, 16)
		Me.lblOffsetX.TabIndex = 21
		Me.lblOffsetX.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblOffsetX.BackColor = System.Drawing.SystemColors.Control
		Me.lblOffsetX.Enabled = True
		Me.lblOffsetX.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblOffsetX.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblOffsetX.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblOffsetX.UseMnemonic = True
		Me.lblOffsetX.Visible = True
		Me.lblOffsetX.AutoSize = False
		Me.lblOffsetX.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblOffsetX.Name = "lblOffsetX"
		Me.lblOffsetY.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblOffsetY.Text = "Y Offset"
		Me.lblOffsetY.Size = New System.Drawing.Size(41, 17)
		Me.lblOffsetY.Location = New System.Drawing.Point(96, 16)
		Me.lblOffsetY.TabIndex = 20
		Me.lblOffsetY.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblOffsetY.BackColor = System.Drawing.SystemColors.Control
		Me.lblOffsetY.Enabled = True
		Me.lblOffsetY.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblOffsetY.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblOffsetY.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblOffsetY.UseMnemonic = True
		Me.lblOffsetY.Visible = True
		Me.lblOffsetY.AutoSize = False
		Me.lblOffsetY.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblOffsetY.Name = "lblOffsetY"
		Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblTitle.Text = "Copy Game Database"
		Me.lblTitle.Font = New System.Drawing.Font("Arial Black", 9.6!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblTitle.Size = New System.Drawing.Size(185, 17)
		Me.lblTitle.Location = New System.Drawing.Point(8, 8)
		Me.lblTitle.TabIndex = 19
		Me.lblTitle.BackColor = System.Drawing.SystemColors.Control
		Me.lblTitle.Enabled = True
		Me.lblTitle.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblTitle.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblTitle.UseMnemonic = True
		Me.lblTitle.Visible = True
		Me.lblTitle.AutoSize = False
		Me.lblTitle.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblTitle.Name = "lblTitle"
		Me.Label2.Text = "Game Name"
		Me.Label2.Size = New System.Drawing.Size(73, 17)
		Me.Label2.Location = New System.Drawing.Point(8, 40)
		Me.Label2.TabIndex = 0
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
		Me.Controls.Add(cbGameName)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(frameType)
		Me.Controls.Add(frameOffset)
		Me.Controls.Add(lblTitle)
		Me.Controls.Add(Label2)
		Me.frameType.Controls.Add(_chkItems_10)
		Me.frameType.Controls.Add(_chkItems_0)
		Me.frameType.Controls.Add(_chkItems_1)
		Me.frameType.Controls.Add(_chkItems_2)
		Me.frameType.Controls.Add(_chkItems_3)
		Me.frameType.Controls.Add(_chkItems_4)
		Me.frameType.Controls.Add(_chkItems_5)
		Me.frameType.Controls.Add(_chkItems_6)
		Me.frameType.Controls.Add(_chkItems_7)
		Me.frameType.Controls.Add(_chkItems_8)
		Me.frameType.Controls.Add(_chkItems_9)
		Me.frameOffset.Controls.Add(txtOffsetX)
		Me.frameOffset.Controls.Add(txtOffsetY)
		Me.frameOffset.Controls.Add(lblOffsetX)
		Me.frameOffset.Controls.Add(lblOffsetY)
		Me.chkItems.SetIndex(_chkItems_10, CType(10, Short))
		Me.chkItems.SetIndex(_chkItems_0, CType(0, Short))
		Me.chkItems.SetIndex(_chkItems_1, CType(1, Short))
		Me.chkItems.SetIndex(_chkItems_2, CType(2, Short))
		Me.chkItems.SetIndex(_chkItems_3, CType(3, Short))
		Me.chkItems.SetIndex(_chkItems_4, CType(4, Short))
		Me.chkItems.SetIndex(_chkItems_5, CType(5, Short))
		Me.chkItems.SetIndex(_chkItems_6, CType(6, Short))
		Me.chkItems.SetIndex(_chkItems_7, CType(7, Short))
		Me.chkItems.SetIndex(_chkItems_8, CType(8, Short))
		Me.chkItems.SetIndex(_chkItems_9, CType(9, Short))
		CType(Me.chkItems, System.ComponentModel.ISupportInitialize).EndInit()
		Me.frameType.ResumeLayout(False)
		Me.frameOffset.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class