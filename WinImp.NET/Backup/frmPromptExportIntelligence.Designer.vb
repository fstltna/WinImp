<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPromptExportIntelligence
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
	Public WithEvents chkHeader As System.Windows.Forms.CheckBox
	Public WithEvents cmdClearAllData As System.Windows.Forms.Button
	Public WithEvents cmdSetAllData As System.Windows.Forms.Button
	Public WithEvents _chkData_16 As System.Windows.Forms.CheckBox
	Public WithEvents _chkData_14 As System.Windows.Forms.CheckBox
	Public WithEvents _chkData_13 As System.Windows.Forms.CheckBox
	Public WithEvents _chkData_15 As System.Windows.Forms.CheckBox
	Public WithEvents _chkData_0 As System.Windows.Forms.CheckBox
	Public WithEvents _chkData_1 As System.Windows.Forms.CheckBox
	Public WithEvents _chkData_8 As System.Windows.Forms.CheckBox
	Public WithEvents _chkData_7 As System.Windows.Forms.CheckBox
	Public WithEvents _chkData_6 As System.Windows.Forms.CheckBox
	Public WithEvents _chkData_5 As System.Windows.Forms.CheckBox
	Public WithEvents _chkData_4 As System.Windows.Forms.CheckBox
	Public WithEvents _chkData_9 As System.Windows.Forms.CheckBox
	Public WithEvents _chkData_12 As System.Windows.Forms.CheckBox
	Public WithEvents _chkData_11 As System.Windows.Forms.CheckBox
	Public WithEvents _chkData_10 As System.Windows.Forms.CheckBox
	Public WithEvents frameData As System.Windows.Forms.GroupBox
	Public WithEvents txtOffsetY As System.Windows.Forms.TextBox
	Public WithEvents txtOffsetX As System.Windows.Forms.TextBox
	Public WithEvents lblOffsetY As System.Windows.Forms.Label
	Public WithEvents lblOffsetX As System.Windows.Forms.Label
	Public WithEvents frameOffset As System.Windows.Forms.GroupBox
	Public WithEvents _chkItems_9 As System.Windows.Forms.CheckBox
	Public WithEvents _chkItems_8 As System.Windows.Forms.CheckBox
	Public WithEvents _chkItems_7 As System.Windows.Forms.CheckBox
	Public WithEvents _chkItems_6 As System.Windows.Forms.CheckBox
	Public WithEvents _chkItems_5 As System.Windows.Forms.CheckBox
	Public WithEvents _chkItems_4 As System.Windows.Forms.CheckBox
	Public WithEvents _chkItems_3 As System.Windows.Forms.CheckBox
	Public WithEvents _chkItems_2 As System.Windows.Forms.CheckBox
	Public WithEvents _chkItems_1 As System.Windows.Forms.CheckBox
	Public WithEvents _chkItems_0 As System.Windows.Forms.CheckBox
	Public WithEvents frameType As System.Windows.Forms.GroupBox
	Public WithEvents txtMultOrigin As System.Windows.Forms.TextBox
	Public WithEvents txtDesignations As System.Windows.Forms.TextBox
	Public WithEvents cbNations As System.Windows.Forms.ComboBox
	Public WithEvents lblOrigin As System.Windows.Forms.Label
	Public WithEvents lblCountry As System.Windows.Forms.Label
	Public WithEvents lblDesignations As System.Windows.Forms.Label
	Public WithEvents frameFilter As System.Windows.Forms.GroupBox
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cbTelegram As System.Windows.Forms.ComboBox
	Public WithEvents _optDestination_0 As System.Windows.Forms.RadioButton
	Public WithEvents _optDestination_1 As System.Windows.Forms.RadioButton
	Public WithEvents lblTelegram As System.Windows.Forms.Label
	Public WithEvents frameOutput As System.Windows.Forms.GroupBox
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents chkData As Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray
	Public WithEvents chkItems As Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray
	Public WithEvents optDestination As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPromptExportIntelligence))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.chkHeader = New System.Windows.Forms.CheckBox
		Me.frameData = New System.Windows.Forms.GroupBox
		Me.cmdClearAllData = New System.Windows.Forms.Button
		Me.cmdSetAllData = New System.Windows.Forms.Button
		Me._chkData_16 = New System.Windows.Forms.CheckBox
		Me._chkData_14 = New System.Windows.Forms.CheckBox
		Me._chkData_13 = New System.Windows.Forms.CheckBox
		Me._chkData_15 = New System.Windows.Forms.CheckBox
		Me._chkData_0 = New System.Windows.Forms.CheckBox
		Me._chkData_1 = New System.Windows.Forms.CheckBox
		Me._chkData_8 = New System.Windows.Forms.CheckBox
		Me._chkData_7 = New System.Windows.Forms.CheckBox
		Me._chkData_6 = New System.Windows.Forms.CheckBox
		Me._chkData_5 = New System.Windows.Forms.CheckBox
		Me._chkData_4 = New System.Windows.Forms.CheckBox
		Me._chkData_9 = New System.Windows.Forms.CheckBox
		Me._chkData_12 = New System.Windows.Forms.CheckBox
		Me._chkData_11 = New System.Windows.Forms.CheckBox
		Me._chkData_10 = New System.Windows.Forms.CheckBox
		Me.frameOffset = New System.Windows.Forms.GroupBox
		Me.txtOffsetY = New System.Windows.Forms.TextBox
		Me.txtOffsetX = New System.Windows.Forms.TextBox
		Me.lblOffsetY = New System.Windows.Forms.Label
		Me.lblOffsetX = New System.Windows.Forms.Label
		Me.frameType = New System.Windows.Forms.GroupBox
		Me._chkItems_9 = New System.Windows.Forms.CheckBox
		Me._chkItems_8 = New System.Windows.Forms.CheckBox
		Me._chkItems_7 = New System.Windows.Forms.CheckBox
		Me._chkItems_6 = New System.Windows.Forms.CheckBox
		Me._chkItems_5 = New System.Windows.Forms.CheckBox
		Me._chkItems_4 = New System.Windows.Forms.CheckBox
		Me._chkItems_3 = New System.Windows.Forms.CheckBox
		Me._chkItems_2 = New System.Windows.Forms.CheckBox
		Me._chkItems_1 = New System.Windows.Forms.CheckBox
		Me._chkItems_0 = New System.Windows.Forms.CheckBox
		Me.frameFilter = New System.Windows.Forms.GroupBox
		Me.txtMultOrigin = New System.Windows.Forms.TextBox
		Me.txtDesignations = New System.Windows.Forms.TextBox
		Me.cbNations = New System.Windows.Forms.ComboBox
		Me.lblOrigin = New System.Windows.Forms.Label
		Me.lblCountry = New System.Windows.Forms.Label
		Me.lblDesignations = New System.Windows.Forms.Label
		Me.cmdOK = New System.Windows.Forms.Button
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.frameOutput = New System.Windows.Forms.GroupBox
		Me.cbTelegram = New System.Windows.Forms.ComboBox
		Me._optDestination_0 = New System.Windows.Forms.RadioButton
		Me._optDestination_1 = New System.Windows.Forms.RadioButton
		Me.lblTelegram = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.chkData = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(components)
		Me.chkItems = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(components)
		Me.optDestination = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(components)
		Me.frameData.SuspendLayout()
		Me.frameOffset.SuspendLayout()
		Me.frameType.SuspendLayout()
		Me.frameFilter.SuspendLayout()
		Me.frameOutput.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.chkData, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.chkItems, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.optDestination, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.ControlBox = False
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.ClientSize = New System.Drawing.Size(679, 198)
		Me.Location = New System.Drawing.Point(3, 3)
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
		Me.Name = "frmPromptExportIntelligence"
		Me.chkHeader.Text = "Include Headers"
		Me.chkHeader.Size = New System.Drawing.Size(105, 17)
		Me.chkHeader.Location = New System.Drawing.Point(8, 176)
		Me.chkHeader.TabIndex = 47
		Me.chkHeader.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkHeader.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkHeader.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkHeader.BackColor = System.Drawing.SystemColors.Control
		Me.chkHeader.CausesValidation = True
		Me.chkHeader.Enabled = True
		Me.chkHeader.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkHeader.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkHeader.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkHeader.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkHeader.TabStop = True
		Me.chkHeader.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkHeader.Visible = True
		Me.chkHeader.Name = "chkHeader"
		Me.frameData.Text = "Data Fields"
		Me.frameData.Size = New System.Drawing.Size(121, 185)
		Me.frameData.Location = New System.Drawing.Point(552, 8)
		Me.frameData.TabIndex = 31
		Me.frameData.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frameData.BackColor = System.Drawing.SystemColors.Control
		Me.frameData.Enabled = True
		Me.frameData.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frameData.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frameData.Visible = True
		Me.frameData.Padding = New System.Windows.Forms.Padding(0)
		Me.frameData.Name = "frameData"
		Me.cmdClearAllData.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdClearAllData.Text = "Clear All"
		Me.cmdClearAllData.Size = New System.Drawing.Size(49, 25)
		Me.cmdClearAllData.Location = New System.Drawing.Point(64, 152)
		Me.cmdClearAllData.TabIndex = 49
		Me.cmdClearAllData.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdClearAllData.BackColor = System.Drawing.SystemColors.Control
		Me.cmdClearAllData.CausesValidation = True
		Me.cmdClearAllData.Enabled = True
		Me.cmdClearAllData.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdClearAllData.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdClearAllData.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdClearAllData.TabStop = True
		Me.cmdClearAllData.Name = "cmdClearAllData"
		Me.cmdSetAllData.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdSetAllData.Text = "Set All"
		Me.cmdSetAllData.Size = New System.Drawing.Size(49, 25)
		Me.cmdSetAllData.Location = New System.Drawing.Point(8, 152)
		Me.cmdSetAllData.TabIndex = 48
		Me.cmdSetAllData.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdSetAllData.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSetAllData.CausesValidation = True
		Me.cmdSetAllData.Enabled = True
		Me.cmdSetAllData.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSetAllData.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSetAllData.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSetAllData.TabStop = True
		Me.cmdSetAllData.Name = "cmdSetAllData"
		Me._chkData_16.Text = "Bar"
		Me._chkData_16.Size = New System.Drawing.Size(49, 17)
		Me._chkData_16.Location = New System.Drawing.Point(64, 64)
		Me._chkData_16.TabIndex = 43
		Me._chkData_16.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkData_16.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkData_16.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkData_16.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkData_16.BackColor = System.Drawing.SystemColors.Control
		Me._chkData_16.CausesValidation = True
		Me._chkData_16.Enabled = True
		Me._chkData_16.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkData_16.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkData_16.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkData_16.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkData_16.TabStop = True
		Me._chkData_16.Visible = True
		Me._chkData_16.Name = "_chkData_16"
		Me._chkData_14.Text = "Food"
		Me._chkData_14.Size = New System.Drawing.Size(49, 17)
		Me._chkData_14.Location = New System.Drawing.Point(64, 48)
		Me._chkData_14.TabIndex = 32
		Me._chkData_14.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkData_14.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkData_14.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkData_14.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkData_14.BackColor = System.Drawing.SystemColors.Control
		Me._chkData_14.CausesValidation = True
		Me._chkData_14.Enabled = True
		Me._chkData_14.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkData_14.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkData_14.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkData_14.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkData_14.TabStop = True
		Me._chkData_14.Visible = True
		Me._chkData_14.Name = "_chkData_14"
		Me._chkData_13.Text = "Petrol"
		Me._chkData_13.Size = New System.Drawing.Size(49, 17)
		Me._chkData_13.Location = New System.Drawing.Point(64, 32)
		Me._chkData_13.TabIndex = 33
		Me._chkData_13.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkData_13.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkData_13.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkData_13.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkData_13.BackColor = System.Drawing.SystemColors.Control
		Me._chkData_13.CausesValidation = True
		Me._chkData_13.Enabled = True
		Me._chkData_13.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkData_13.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkData_13.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkData_13.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkData_13.TabStop = True
		Me._chkData_13.Visible = True
		Me._chkData_13.Name = "_chkData_13"
		Me._chkData_15.Text = "Iron"
		Me._chkData_15.Size = New System.Drawing.Size(49, 17)
		Me._chkData_15.Location = New System.Drawing.Point(64, 16)
		Me._chkData_15.TabIndex = 42
		Me._chkData_15.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkData_15.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkData_15.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkData_15.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkData_15.BackColor = System.Drawing.SystemColors.Control
		Me._chkData_15.CausesValidation = True
		Me._chkData_15.Enabled = True
		Me._chkData_15.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkData_15.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkData_15.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkData_15.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkData_15.TabStop = True
		Me._chkData_15.Visible = True
		Me._chkData_15.Name = "_chkData_15"
		Me._chkData_0.Text = "Wake"
		Me._chkData_0.Size = New System.Drawing.Size(49, 17)
		Me._chkData_0.Location = New System.Drawing.Point(64, 128)
		Me._chkData_0.TabIndex = 46
		Me._chkData_0.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkData_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkData_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkData_0.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkData_0.BackColor = System.Drawing.SystemColors.Control
		Me._chkData_0.CausesValidation = True
		Me._chkData_0.Enabled = True
		Me._chkData_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkData_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkData_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkData_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkData_0.TabStop = True
		Me._chkData_0.Visible = True
		Me._chkData_0.Name = "_chkData_0"
		Me._chkData_1.Text = "Tech"
		Me._chkData_1.Size = New System.Drawing.Size(57, 17)
		Me._chkData_1.Location = New System.Drawing.Point(8, 128)
		Me._chkData_1.TabIndex = 44
		Me._chkData_1.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkData_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkData_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkData_1.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkData_1.BackColor = System.Drawing.SystemColors.Control
		Me._chkData_1.CausesValidation = True
		Me._chkData_1.Enabled = True
		Me._chkData_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkData_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkData_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkData_1.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkData_1.TabStop = True
		Me._chkData_1.Visible = True
		Me._chkData_1.Name = "_chkData_1"
		Me._chkData_8.Text = "Def."
		Me._chkData_8.Size = New System.Drawing.Size(49, 17)
		Me._chkData_8.Location = New System.Drawing.Point(64, 96)
		Me._chkData_8.TabIndex = 41
		Me.ToolTip1.SetToolTip(Me._chkData_8, "Defense")
		Me._chkData_8.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkData_8.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkData_8.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkData_8.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkData_8.BackColor = System.Drawing.SystemColors.Control
		Me._chkData_8.CausesValidation = True
		Me._chkData_8.Enabled = True
		Me._chkData_8.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkData_8.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkData_8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkData_8.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkData_8.TabStop = True
		Me._chkData_8.Visible = True
		Me._chkData_8.Name = "_chkData_8"
		Me._chkData_7.Text = "Rail"
		Me._chkData_7.Size = New System.Drawing.Size(49, 17)
		Me._chkData_7.Location = New System.Drawing.Point(64, 112)
		Me._chkData_7.TabIndex = 40
		Me._chkData_7.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkData_7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkData_7.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkData_7.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkData_7.BackColor = System.Drawing.SystemColors.Control
		Me._chkData_7.CausesValidation = True
		Me._chkData_7.Enabled = True
		Me._chkData_7.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkData_7.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkData_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkData_7.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkData_7.TabStop = True
		Me._chkData_7.Visible = True
		Me._chkData_7.Name = "_chkData_7"
		Me._chkData_6.Text = "Road"
		Me._chkData_6.Size = New System.Drawing.Size(49, 17)
		Me._chkData_6.Location = New System.Drawing.Point(8, 112)
		Me._chkData_6.TabIndex = 34
		Me._chkData_6.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkData_6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkData_6.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkData_6.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkData_6.BackColor = System.Drawing.SystemColors.Control
		Me._chkData_6.CausesValidation = True
		Me._chkData_6.Enabled = True
		Me._chkData_6.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkData_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkData_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkData_6.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkData_6.TabStop = True
		Me._chkData_6.Visible = True
		Me._chkData_6.Name = "_chkData_6"
		Me._chkData_5.Text = "Eff."
		Me._chkData_5.Size = New System.Drawing.Size(49, 17)
		Me._chkData_5.Location = New System.Drawing.Point(8, 96)
		Me._chkData_5.TabIndex = 35
		Me.ToolTip1.SetToolTip(Me._chkData_5, "Efficiency")
		Me._chkData_5.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkData_5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkData_5.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkData_5.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkData_5.BackColor = System.Drawing.SystemColors.Control
		Me._chkData_5.CausesValidation = True
		Me._chkData_5.Enabled = True
		Me._chkData_5.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkData_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkData_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkData_5.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkData_5.TabStop = True
		Me._chkData_5.Visible = True
		Me._chkData_5.Name = "_chkData_5"
		Me._chkData_4.Text = "Old Owner"
		Me._chkData_4.Size = New System.Drawing.Size(81, 17)
		Me._chkData_4.Location = New System.Drawing.Point(8, 80)
		Me._chkData_4.TabIndex = 45
		Me._chkData_4.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkData_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkData_4.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkData_4.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkData_4.BackColor = System.Drawing.SystemColors.Control
		Me._chkData_4.CausesValidation = True
		Me._chkData_4.Enabled = True
		Me._chkData_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkData_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkData_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkData_4.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkData_4.TabStop = True
		Me._chkData_4.Visible = True
		Me._chkData_4.Name = "_chkData_4"
		Me._chkData_9.Text = "Civ."
		Me._chkData_9.Size = New System.Drawing.Size(49, 17)
		Me._chkData_9.Location = New System.Drawing.Point(8, 64)
		Me._chkData_9.TabIndex = 36
		Me._chkData_9.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkData_9.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkData_9.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkData_9.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkData_9.BackColor = System.Drawing.SystemColors.Control
		Me._chkData_9.CausesValidation = True
		Me._chkData_9.Enabled = True
		Me._chkData_9.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkData_9.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkData_9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkData_9.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkData_9.TabStop = True
		Me._chkData_9.Visible = True
		Me._chkData_9.Name = "_chkData_9"
		Me._chkData_12.Text = "Guns"
		Me._chkData_12.Size = New System.Drawing.Size(49, 17)
		Me._chkData_12.Location = New System.Drawing.Point(8, 48)
		Me._chkData_12.TabIndex = 37
		Me._chkData_12.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkData_12.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkData_12.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkData_12.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkData_12.BackColor = System.Drawing.SystemColors.Control
		Me._chkData_12.CausesValidation = True
		Me._chkData_12.Enabled = True
		Me._chkData_12.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkData_12.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkData_12.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkData_12.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkData_12.TabStop = True
		Me._chkData_12.Visible = True
		Me._chkData_12.Name = "_chkData_12"
		Me._chkData_11.Text = "Shells"
		Me._chkData_11.Size = New System.Drawing.Size(49, 17)
		Me._chkData_11.Location = New System.Drawing.Point(8, 32)
		Me._chkData_11.TabIndex = 38
		Me._chkData_11.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkData_11.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkData_11.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkData_11.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkData_11.BackColor = System.Drawing.SystemColors.Control
		Me._chkData_11.CausesValidation = True
		Me._chkData_11.Enabled = True
		Me._chkData_11.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkData_11.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkData_11.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkData_11.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkData_11.TabStop = True
		Me._chkData_11.Visible = True
		Me._chkData_11.Name = "_chkData_11"
		Me._chkData_10.Text = "Mil."
		Me._chkData_10.Size = New System.Drawing.Size(49, 17)
		Me._chkData_10.Location = New System.Drawing.Point(8, 16)
		Me._chkData_10.TabIndex = 39
		Me._chkData_10.CheckState = System.Windows.Forms.CheckState.Checked
		Me._chkData_10.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkData_10.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkData_10.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkData_10.BackColor = System.Drawing.SystemColors.Control
		Me._chkData_10.CausesValidation = True
		Me._chkData_10.Enabled = True
		Me._chkData_10.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkData_10.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkData_10.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkData_10.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkData_10.TabStop = True
		Me._chkData_10.Visible = True
		Me._chkData_10.Name = "_chkData_10"
		Me.frameOffset.Text = "Offset"
		Me.frameOffset.Size = New System.Drawing.Size(185, 41)
		Me.frameOffset.Location = New System.Drawing.Point(8, 128)
		Me.frameOffset.TabIndex = 25
		Me.frameOffset.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frameOffset.BackColor = System.Drawing.SystemColors.Control
		Me.frameOffset.Enabled = True
		Me.frameOffset.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frameOffset.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frameOffset.Visible = True
		Me.frameOffset.Padding = New System.Windows.Forms.Padding(0)
		Me.frameOffset.Name = "frameOffset"
		Me.txtOffsetY.AutoSize = False
		Me.txtOffsetY.Size = New System.Drawing.Size(33, 19)
		Me.txtOffsetY.Location = New System.Drawing.Point(144, 16)
		Me.txtOffsetY.TabIndex = 28
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
		Me.txtOffsetX.AutoSize = False
		Me.txtOffsetX.Size = New System.Drawing.Size(33, 19)
		Me.txtOffsetX.Location = New System.Drawing.Point(56, 16)
		Me.txtOffsetX.TabIndex = 26
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
		Me.lblOffsetY.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblOffsetY.Text = "Y Offset"
		Me.lblOffsetY.Size = New System.Drawing.Size(41, 17)
		Me.lblOffsetY.Location = New System.Drawing.Point(96, 16)
		Me.lblOffsetY.TabIndex = 29
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
		Me.lblOffsetX.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblOffsetX.Text = "X Offset"
		Me.lblOffsetX.Size = New System.Drawing.Size(41, 17)
		Me.lblOffsetX.Location = New System.Drawing.Point(8, 16)
		Me.lblOffsetX.TabIndex = 27
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
		Me.frameType.Text = "Items"
		Me.frameType.Size = New System.Drawing.Size(129, 185)
		Me.frameType.Location = New System.Drawing.Point(408, 8)
		Me.frameType.TabIndex = 9
		Me.frameType.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frameType.BackColor = System.Drawing.SystemColors.Control
		Me.frameType.Enabled = True
		Me.frameType.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frameType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frameType.Visible = True
		Me.frameType.Padding = New System.Windows.Forms.Padding(0)
		Me.frameType.Name = "frameType"
		Me._chkItems_9.Text = "Our Nukes"
		Me._chkItems_9.Enabled = False
		Me._chkItems_9.Size = New System.Drawing.Size(113, 17)
		Me._chkItems_9.Location = New System.Drawing.Point(8, 160)
		Me._chkItems_9.TabIndex = 20
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
		Me._chkItems_8.Text = "Our Land Units"
		Me._chkItems_8.Size = New System.Drawing.Size(113, 17)
		Me._chkItems_8.Location = New System.Drawing.Point(8, 144)
		Me._chkItems_8.TabIndex = 21
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
		Me._chkItems_8.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._chkItems_8.Visible = True
		Me._chkItems_8.Name = "_chkItems_8"
		Me._chkItems_7.Text = "Our Planes"
		Me._chkItems_7.Size = New System.Drawing.Size(113, 17)
		Me._chkItems_7.Location = New System.Drawing.Point(8, 128)
		Me._chkItems_7.TabIndex = 22
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
		Me._chkItems_7.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._chkItems_7.Visible = True
		Me._chkItems_7.Name = "_chkItems_7"
		Me._chkItems_6.Text = "Our Ships"
		Me._chkItems_6.Size = New System.Drawing.Size(113, 17)
		Me._chkItems_6.Location = New System.Drawing.Point(8, 112)
		Me._chkItems_6.TabIndex = 23
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
		Me._chkItems_6.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._chkItems_6.Visible = True
		Me._chkItems_6.Name = "_chkItems_6"
		Me._chkItems_5.Text = "Our Sectors"
		Me._chkItems_5.Size = New System.Drawing.Size(113, 17)
		Me._chkItems_5.Location = New System.Drawing.Point(8, 96)
		Me._chkItems_5.TabIndex = 24
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
		Me._chkItems_5.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._chkItems_5.Visible = True
		Me._chkItems_5.Name = "_chkItems_5"
		Me._chkItems_4.Text = "Enemy Nukes"
		Me._chkItems_4.Enabled = False
		Me._chkItems_4.Size = New System.Drawing.Size(113, 17)
		Me._chkItems_4.Location = New System.Drawing.Point(8, 80)
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
		Me._chkItems_3.Text = "Enemy Land Units"
		Me._chkItems_3.Size = New System.Drawing.Size(113, 17)
		Me._chkItems_3.Location = New System.Drawing.Point(8, 64)
		Me._chkItems_3.TabIndex = 15
		Me._chkItems_3.CheckState = System.Windows.Forms.CheckState.Checked
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
		Me._chkItems_3.Visible = True
		Me._chkItems_3.Name = "_chkItems_3"
		Me._chkItems_2.Text = "Enemy Planes"
		Me._chkItems_2.Size = New System.Drawing.Size(113, 17)
		Me._chkItems_2.Location = New System.Drawing.Point(8, 48)
		Me._chkItems_2.TabIndex = 14
		Me._chkItems_2.CheckState = System.Windows.Forms.CheckState.Checked
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
		Me._chkItems_2.Visible = True
		Me._chkItems_2.Name = "_chkItems_2"
		Me._chkItems_1.Text = "Enemy Ships"
		Me._chkItems_1.Size = New System.Drawing.Size(113, 17)
		Me._chkItems_1.Location = New System.Drawing.Point(8, 32)
		Me._chkItems_1.TabIndex = 13
		Me._chkItems_1.CheckState = System.Windows.Forms.CheckState.Checked
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
		Me._chkItems_1.Visible = True
		Me._chkItems_1.Name = "_chkItems_1"
		Me._chkItems_0.Text = "Enemy Sectors"
		Me._chkItems_0.Size = New System.Drawing.Size(113, 17)
		Me._chkItems_0.Location = New System.Drawing.Point(8, 16)
		Me._chkItems_0.TabIndex = 12
		Me._chkItems_0.CheckState = System.Windows.Forms.CheckState.Checked
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
		Me._chkItems_0.Visible = True
		Me._chkItems_0.Name = "_chkItems_0"
		Me.frameFilter.Text = "Filters"
		Me.frameFilter.Size = New System.Drawing.Size(185, 113)
		Me.frameFilter.Location = New System.Drawing.Point(208, 32)
		Me.frameFilter.TabIndex = 5
		Me.frameFilter.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frameFilter.BackColor = System.Drawing.SystemColors.Control
		Me.frameFilter.Enabled = True
		Me.frameFilter.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frameFilter.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frameFilter.Visible = True
		Me.frameFilter.Padding = New System.Windows.Forms.Padding(0)
		Me.frameFilter.Name = "frameFilter"
		Me.txtMultOrigin.AutoSize = False
		Me.txtMultOrigin.Size = New System.Drawing.Size(97, 19)
		Me.txtMultOrigin.Location = New System.Drawing.Point(80, 56)
		Me.txtMultOrigin.TabIndex = 17
		Me.txtMultOrigin.Text = "*"
		Me.txtMultOrigin.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtMultOrigin.AcceptsReturn = True
		Me.txtMultOrigin.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtMultOrigin.BackColor = System.Drawing.SystemColors.Window
		Me.txtMultOrigin.CausesValidation = True
		Me.txtMultOrigin.Enabled = True
		Me.txtMultOrigin.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtMultOrigin.HideSelection = True
		Me.txtMultOrigin.ReadOnly = False
		Me.txtMultOrigin.Maxlength = 0
		Me.txtMultOrigin.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtMultOrigin.MultiLine = False
		Me.txtMultOrigin.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtMultOrigin.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtMultOrigin.TabStop = True
		Me.txtMultOrigin.Visible = True
		Me.txtMultOrigin.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtMultOrigin.Name = "txtMultOrigin"
		Me.txtDesignations.AutoSize = False
		Me.txtDesignations.Size = New System.Drawing.Size(97, 19)
		Me.txtDesignations.Location = New System.Drawing.Point(80, 88)
		Me.txtDesignations.TabIndex = 7
		Me.txtDesignations.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtDesignations.AcceptsReturn = True
		Me.txtDesignations.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtDesignations.BackColor = System.Drawing.SystemColors.Window
		Me.txtDesignations.CausesValidation = True
		Me.txtDesignations.Enabled = True
		Me.txtDesignations.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtDesignations.HideSelection = True
		Me.txtDesignations.ReadOnly = False
		Me.txtDesignations.Maxlength = 0
		Me.txtDesignations.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtDesignations.MultiLine = False
		Me.txtDesignations.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtDesignations.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtDesignations.TabStop = True
		Me.txtDesignations.Visible = True
		Me.txtDesignations.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtDesignations.Name = "txtDesignations"
		Me.cbNations.Size = New System.Drawing.Size(169, 21)
		Me.cbNations.Location = New System.Drawing.Point(8, 32)
		Me.cbNations.Sorted = True
		Me.cbNations.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cbNations.TabIndex = 6
		Me.cbNations.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbNations.BackColor = System.Drawing.SystemColors.Window
		Me.cbNations.CausesValidation = True
		Me.cbNations.Enabled = True
		Me.cbNations.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cbNations.IntegralHeight = True
		Me.cbNations.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbNations.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbNations.TabStop = True
		Me.cbNations.Visible = True
		Me.cbNations.Name = "cbNations"
		Me.lblOrigin.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblOrigin.Text = "Sectors (sect,range,*)"
		Me.lblOrigin.Size = New System.Drawing.Size(65, 33)
		Me.lblOrigin.Location = New System.Drawing.Point(8, 56)
		Me.lblOrigin.TabIndex = 18
		Me.lblOrigin.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblOrigin.BackColor = System.Drawing.SystemColors.Control
		Me.lblOrigin.Enabled = True
		Me.lblOrigin.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblOrigin.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblOrigin.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblOrigin.UseMnemonic = True
		Me.lblOrigin.Visible = True
		Me.lblOrigin.AutoSize = False
		Me.lblOrigin.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblOrigin.Name = "lblOrigin"
		Me.lblCountry.Text = "Enemy Nation to include in export"
		Me.lblCountry.Size = New System.Drawing.Size(169, 17)
		Me.lblCountry.Location = New System.Drawing.Point(8, 16)
		Me.lblCountry.TabIndex = 10
		Me.lblCountry.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblCountry.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblCountry.BackColor = System.Drawing.SystemColors.Control
		Me.lblCountry.Enabled = True
		Me.lblCountry.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblCountry.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblCountry.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblCountry.UseMnemonic = True
		Me.lblCountry.Visible = True
		Me.lblCountry.AutoSize = False
		Me.lblCountry.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblCountry.Name = "lblCountry"
		Me.lblDesignations.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblDesignations.Text = "Designations"
		Me.lblDesignations.Size = New System.Drawing.Size(65, 17)
		Me.lblDesignations.Location = New System.Drawing.Point(8, 88)
		Me.lblDesignations.TabIndex = 8
		Me.lblDesignations.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblDesignations.BackColor = System.Drawing.SystemColors.Control
		Me.lblDesignations.Enabled = True
		Me.lblDesignations.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblDesignations.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblDesignations.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblDesignations.UseMnemonic = True
		Me.lblDesignations.Visible = True
		Me.lblDesignations.AutoSize = False
		Me.lblDesignations.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblDesignations.Name = "lblDesignations"
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "OK"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Size = New System.Drawing.Size(81, 25)
		Me.cmdOK.Location = New System.Drawing.Point(208, 160)
		Me.cmdOK.TabIndex = 4
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
		Me.cmdCancel.Location = New System.Drawing.Point(312, 160)
		Me.cmdCancel.TabIndex = 3
		Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.frameOutput.Text = "Output"
		Me.frameOutput.Size = New System.Drawing.Size(185, 105)
		Me.frameOutput.Location = New System.Drawing.Point(8, 8)
		Me.frameOutput.TabIndex = 0
		Me.frameOutput.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frameOutput.BackColor = System.Drawing.SystemColors.Control
		Me.frameOutput.Enabled = True
		Me.frameOutput.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frameOutput.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frameOutput.Visible = True
		Me.frameOutput.Padding = New System.Windows.Forms.Padding(0)
		Me.frameOutput.Name = "frameOutput"
		Me.cbTelegram.Size = New System.Drawing.Size(153, 21)
		Me.cbTelegram.Location = New System.Drawing.Point(24, 72)
		Me.cbTelegram.Sorted = True
		Me.cbTelegram.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cbTelegram.TabIndex = 11
		Me.cbTelegram.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbTelegram.BackColor = System.Drawing.SystemColors.Window
		Me.cbTelegram.CausesValidation = True
		Me.cbTelegram.Enabled = True
		Me.cbTelegram.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cbTelegram.IntegralHeight = True
		Me.cbTelegram.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbTelegram.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbTelegram.TabStop = True
		Me.cbTelegram.Visible = True
		Me.cbTelegram.Name = "cbTelegram"
		Me._optDestination_0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optDestination_0.Text = "File"
		Me._optDestination_0.Size = New System.Drawing.Size(57, 17)
		Me._optDestination_0.Location = New System.Drawing.Point(8, 16)
		Me._optDestination_0.TabIndex = 2
		Me._optDestination_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optDestination_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optDestination_0.BackColor = System.Drawing.SystemColors.Control
		Me._optDestination_0.CausesValidation = True
		Me._optDestination_0.Enabled = True
		Me._optDestination_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optDestination_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._optDestination_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optDestination_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optDestination_0.TabStop = True
		Me._optDestination_0.Checked = False
		Me._optDestination_0.Visible = True
		Me._optDestination_0.Name = "_optDestination_0"
		Me._optDestination_1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optDestination_1.Text = "Telegram"
		Me._optDestination_1.Size = New System.Drawing.Size(73, 17)
		Me._optDestination_1.Location = New System.Drawing.Point(8, 40)
		Me._optDestination_1.TabIndex = 1
		Me._optDestination_1.Checked = True
		Me._optDestination_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optDestination_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optDestination_1.BackColor = System.Drawing.SystemColors.Control
		Me._optDestination_1.CausesValidation = True
		Me._optDestination_1.Enabled = True
		Me._optDestination_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optDestination_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._optDestination_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optDestination_1.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optDestination_1.TabStop = True
		Me._optDestination_1.Visible = True
		Me._optDestination_1.Name = "_optDestination_1"
		Me.lblTelegram.Text = "Nation to telegram"
		Me.lblTelegram.Size = New System.Drawing.Size(97, 17)
		Me.lblTelegram.Location = New System.Drawing.Point(80, 56)
		Me.lblTelegram.TabIndex = 19
		Me.lblTelegram.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblTelegram.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblTelegram.BackColor = System.Drawing.SystemColors.Control
		Me.lblTelegram.Enabled = True
		Me.lblTelegram.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblTelegram.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblTelegram.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblTelegram.UseMnemonic = True
		Me.lblTelegram.Visible = True
		Me.lblTelegram.AutoSize = False
		Me.lblTelegram.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblTelegram.Name = "lblTelegram"
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label1.Text = "Export Intelligence"
		Me.Label1.Font = New System.Drawing.Font("Arial Black", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Size = New System.Drawing.Size(185, 17)
		Me.Label1.Location = New System.Drawing.Point(208, 8)
		Me.Label1.TabIndex = 30
		Me.Label1.BackColor = System.Drawing.SystemColors.Control
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(chkHeader)
		Me.Controls.Add(frameData)
		Me.Controls.Add(frameOffset)
		Me.Controls.Add(frameType)
		Me.Controls.Add(frameFilter)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(frameOutput)
		Me.Controls.Add(Label1)
		Me.frameData.Controls.Add(cmdClearAllData)
		Me.frameData.Controls.Add(cmdSetAllData)
		Me.frameData.Controls.Add(_chkData_16)
		Me.frameData.Controls.Add(_chkData_14)
		Me.frameData.Controls.Add(_chkData_13)
		Me.frameData.Controls.Add(_chkData_15)
		Me.frameData.Controls.Add(_chkData_0)
		Me.frameData.Controls.Add(_chkData_1)
		Me.frameData.Controls.Add(_chkData_8)
		Me.frameData.Controls.Add(_chkData_7)
		Me.frameData.Controls.Add(_chkData_6)
		Me.frameData.Controls.Add(_chkData_5)
		Me.frameData.Controls.Add(_chkData_4)
		Me.frameData.Controls.Add(_chkData_9)
		Me.frameData.Controls.Add(_chkData_12)
		Me.frameData.Controls.Add(_chkData_11)
		Me.frameData.Controls.Add(_chkData_10)
		Me.frameOffset.Controls.Add(txtOffsetY)
		Me.frameOffset.Controls.Add(txtOffsetX)
		Me.frameOffset.Controls.Add(lblOffsetY)
		Me.frameOffset.Controls.Add(lblOffsetX)
		Me.frameType.Controls.Add(_chkItems_9)
		Me.frameType.Controls.Add(_chkItems_8)
		Me.frameType.Controls.Add(_chkItems_7)
		Me.frameType.Controls.Add(_chkItems_6)
		Me.frameType.Controls.Add(_chkItems_5)
		Me.frameType.Controls.Add(_chkItems_4)
		Me.frameType.Controls.Add(_chkItems_3)
		Me.frameType.Controls.Add(_chkItems_2)
		Me.frameType.Controls.Add(_chkItems_1)
		Me.frameType.Controls.Add(_chkItems_0)
		Me.frameFilter.Controls.Add(txtMultOrigin)
		Me.frameFilter.Controls.Add(txtDesignations)
		Me.frameFilter.Controls.Add(cbNations)
		Me.frameFilter.Controls.Add(lblOrigin)
		Me.frameFilter.Controls.Add(lblCountry)
		Me.frameFilter.Controls.Add(lblDesignations)
		Me.frameOutput.Controls.Add(cbTelegram)
		Me.frameOutput.Controls.Add(_optDestination_0)
		Me.frameOutput.Controls.Add(_optDestination_1)
		Me.frameOutput.Controls.Add(lblTelegram)
		Me.chkData.SetIndex(_chkData_16, CType(16, Short))
		Me.chkData.SetIndex(_chkData_14, CType(14, Short))
		Me.chkData.SetIndex(_chkData_13, CType(13, Short))
		Me.chkData.SetIndex(_chkData_15, CType(15, Short))
		Me.chkData.SetIndex(_chkData_0, CType(0, Short))
		Me.chkData.SetIndex(_chkData_1, CType(1, Short))
		Me.chkData.SetIndex(_chkData_8, CType(8, Short))
		Me.chkData.SetIndex(_chkData_7, CType(7, Short))
		Me.chkData.SetIndex(_chkData_6, CType(6, Short))
		Me.chkData.SetIndex(_chkData_5, CType(5, Short))
		Me.chkData.SetIndex(_chkData_4, CType(4, Short))
		Me.chkData.SetIndex(_chkData_9, CType(9, Short))
		Me.chkData.SetIndex(_chkData_12, CType(12, Short))
		Me.chkData.SetIndex(_chkData_11, CType(11, Short))
		Me.chkData.SetIndex(_chkData_10, CType(10, Short))
		Me.chkItems.SetIndex(_chkItems_9, CType(9, Short))
		Me.chkItems.SetIndex(_chkItems_8, CType(8, Short))
		Me.chkItems.SetIndex(_chkItems_7, CType(7, Short))
		Me.chkItems.SetIndex(_chkItems_6, CType(6, Short))
		Me.chkItems.SetIndex(_chkItems_5, CType(5, Short))
		Me.chkItems.SetIndex(_chkItems_4, CType(4, Short))
		Me.chkItems.SetIndex(_chkItems_3, CType(3, Short))
		Me.chkItems.SetIndex(_chkItems_2, CType(2, Short))
		Me.chkItems.SetIndex(_chkItems_1, CType(1, Short))
		Me.chkItems.SetIndex(_chkItems_0, CType(0, Short))
		Me.optDestination.SetIndex(_optDestination_0, CType(0, Short))
		Me.optDestination.SetIndex(_optDestination_1, CType(1, Short))
		CType(Me.optDestination, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.chkItems, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.chkData, System.ComponentModel.ISupportInitialize).EndInit()
		Me.frameData.ResumeLayout(False)
		Me.frameOffset.ResumeLayout(False)
		Me.frameType.ResumeLayout(False)
		Me.frameFilter.ResumeLayout(False)
		Me.frameOutput.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class