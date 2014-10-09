<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPromptRecon
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
	Public WithEvents chkDisplayPath As System.Windows.Forms.CheckBox
	Public WithEvents cmdHelp As System.Windows.Forms.Button
	Public WithEvents cbCom As System.Windows.Forms.ComboBox
	Public WithEvents txtEscort As System.Windows.Forms.TextBox
	Public WithEvents txtPath As System.Windows.Forms.TextBox
	Public WithEvents txtOrigin As System.Windows.Forms.TextBox
	Public WithEvents txtUnit As System.Windows.Forms.TextBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents Check1 As System.Windows.Forms.CheckBox
	Public WithEvents Combo1 As System.Windows.Forms.ComboBox
	Public WithEvents Label8 As System.Windows.Forms.Label
	Public WithEvents Label9 As System.Windows.Forms.Label
	Public WithEvents Label10 As System.Windows.Forms.Label
	Public WithEvents Label7 As System.Windows.Forms.Label
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPromptRecon))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.chkDisplayPath = New System.Windows.Forms.CheckBox
		Me.cmdHelp = New System.Windows.Forms.Button
		Me.cbCom = New System.Windows.Forms.ComboBox
		Me.txtEscort = New System.Windows.Forms.TextBox
		Me.txtPath = New System.Windows.Forms.TextBox
		Me.txtOrigin = New System.Windows.Forms.TextBox
		Me.txtUnit = New System.Windows.Forms.TextBox
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me.Check1 = New System.Windows.Forms.CheckBox
		Me.Combo1 = New System.Windows.Forms.ComboBox
		Me.Label8 = New System.Windows.Forms.Label
		Me.Label9 = New System.Windows.Forms.Label
		Me.Label10 = New System.Windows.Forms.Label
		Me.Label7 = New System.Windows.Forms.Label
		Me.Label6 = New System.Windows.Forms.Label
		Me.Label5 = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.ControlBox = False
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.ClientSize = New System.Drawing.Size(506, 136)
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
		Me.Name = "frmPromptRecon"
		Me.chkDisplayPath.Text = "Display Path"
		Me.chkDisplayPath.Size = New System.Drawing.Size(89, 17)
		Me.chkDisplayPath.Location = New System.Drawing.Point(328, 112)
		Me.chkDisplayPath.TabIndex = 20
		Me.chkDisplayPath.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkDisplayPath.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkDisplayPath.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkDisplayPath.BackColor = System.Drawing.SystemColors.Control
		Me.chkDisplayPath.CausesValidation = True
		Me.chkDisplayPath.Enabled = True
		Me.chkDisplayPath.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkDisplayPath.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkDisplayPath.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkDisplayPath.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkDisplayPath.TabStop = True
		Me.chkDisplayPath.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkDisplayPath.Visible = True
		Me.chkDisplayPath.Name = "chkDisplayPath"
		Me.cmdHelp.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdHelp.Text = "?"
		Me.cmdHelp.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdHelp.Size = New System.Drawing.Size(25, 25)
		Me.cmdHelp.Location = New System.Drawing.Point(480, 0)
		Me.cmdHelp.TabIndex = 19
		Me.ToolTip1.SetToolTip(Me.cmdHelp, "Click for Help")
		Me.cmdHelp.BackColor = System.Drawing.SystemColors.Control
		Me.cmdHelp.CausesValidation = True
		Me.cmdHelp.Enabled = True
		Me.cmdHelp.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdHelp.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdHelp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdHelp.TabStop = True
		Me.cmdHelp.Name = "cmdHelp"
		Me.cbCom.Size = New System.Drawing.Size(137, 21)
		Me.cbCom.Location = New System.Drawing.Point(344, 56)
		Me.cbCom.Items.AddRange(New Object(){"1", "10", "2", "20", "25", "5"})
		Me.cbCom.Sorted = True
		Me.cbCom.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cbCom.TabIndex = 4
		Me.cbCom.Visible = False
		Me.cbCom.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbCom.BackColor = System.Drawing.SystemColors.Window
		Me.cbCom.CausesValidation = True
		Me.cbCom.Enabled = True
		Me.cbCom.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cbCom.IntegralHeight = True
		Me.cbCom.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbCom.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbCom.TabStop = True
		Me.cbCom.Name = "cbCom"
		Me.txtEscort.AutoSize = False
		Me.txtEscort.Size = New System.Drawing.Size(129, 19)
		Me.txtEscort.Location = New System.Drawing.Point(112, 32)
		Me.txtEscort.TabIndex = 1
		Me.txtEscort.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtEscort.AcceptsReturn = True
		Me.txtEscort.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtEscort.BackColor = System.Drawing.SystemColors.Window
		Me.txtEscort.CausesValidation = True
		Me.txtEscort.Enabled = True
		Me.txtEscort.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtEscort.HideSelection = True
		Me.txtEscort.ReadOnly = False
		Me.txtEscort.Maxlength = 0
		Me.txtEscort.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtEscort.MultiLine = False
		Me.txtEscort.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtEscort.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtEscort.TabStop = True
		Me.txtEscort.Visible = True
		Me.txtEscort.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtEscort.Name = "txtEscort"
		Me.txtPath.AutoSize = False
		Me.txtPath.Size = New System.Drawing.Size(129, 19)
		Me.txtPath.Location = New System.Drawing.Point(344, 8)
		Me.txtPath.TabIndex = 2
		Me.txtPath.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtPath.AcceptsReturn = True
		Me.txtPath.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtPath.BackColor = System.Drawing.SystemColors.Window
		Me.txtPath.CausesValidation = True
		Me.txtPath.Enabled = True
		Me.txtPath.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtPath.HideSelection = True
		Me.txtPath.ReadOnly = False
		Me.txtPath.Maxlength = 0
		Me.txtPath.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtPath.MultiLine = False
		Me.txtPath.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtPath.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtPath.TabStop = True
		Me.txtPath.Visible = True
		Me.txtPath.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtPath.Name = "txtPath"
		Me.txtOrigin.AutoSize = False
		Me.txtOrigin.Size = New System.Drawing.Size(129, 19)
		Me.txtOrigin.Location = New System.Drawing.Point(344, 32)
		Me.txtOrigin.TabIndex = 3
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
		Me.txtUnit.AutoSize = False
		Me.txtUnit.Size = New System.Drawing.Size(129, 19)
		Me.txtUnit.Location = New System.Drawing.Point(112, 8)
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
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(73, 33)
		Me.cmdCancel.Location = New System.Drawing.Point(16, 96)
		Me.cmdCancel.TabIndex = 8
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
		Me.cmdOK.Text = "Takeoff !"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Enabled = False
		Me.cmdOK.Size = New System.Drawing.Size(73, 33)
		Me.cmdOK.Location = New System.Drawing.Point(16, 56)
		Me.cmdOK.TabIndex = 7
		Me.cmdOK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.Check1.Text = "Repeat "
		Me.Check1.Size = New System.Drawing.Size(65, 17)
		Me.Check1.Location = New System.Drawing.Point(328, 88)
		Me.Check1.TabIndex = 5
		Me.Check1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Check1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Check1.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.Check1.BackColor = System.Drawing.SystemColors.Control
		Me.Check1.CausesValidation = True
		Me.Check1.Enabled = True
		Me.Check1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Check1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Check1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Check1.Appearance = System.Windows.Forms.Appearance.Normal
		Me.Check1.TabStop = True
		Me.Check1.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.Check1.Visible = True
		Me.Check1.Name = "Check1"
		Me.Combo1.Size = New System.Drawing.Size(49, 21)
		Me.Combo1.Location = New System.Drawing.Point(408, 88)
		Me.Combo1.Items.AddRange(New Object(){"1", "3", "5", "7", "10"})
		Me.Combo1.TabIndex = 6
		Me.Combo1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Combo1.BackColor = System.Drawing.SystemColors.Window
		Me.Combo1.CausesValidation = True
		Me.Combo1.Enabled = True
		Me.Combo1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Combo1.IntegralHeight = True
		Me.Combo1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Combo1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Combo1.Sorted = False
		Me.Combo1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.Combo1.TabStop = True
		Me.Combo1.Visible = True
		Me.Combo1.Name = "Combo1"
		Me.Label8.Text = "Mobility Cost:"
		Me.Label8.Size = New System.Drawing.Size(145, 17)
		Me.Label8.Location = New System.Drawing.Point(112, 64)
		Me.Label8.TabIndex = 18
		Me.Label8.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label8.BackColor = System.Drawing.SystemColors.Control
		Me.Label8.Enabled = True
		Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label8.UseMnemonic = True
		Me.Label8.Visible = True
		Me.Label8.AutoSize = False
		Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label8.Name = "Label8"
		Me.Label9.Text = "Plane Range:"
		Me.Label9.Size = New System.Drawing.Size(145, 17)
		Me.Label9.Location = New System.Drawing.Point(112, 80)
		Me.Label9.TabIndex = 17
		Me.Label9.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label9.BackColor = System.Drawing.SystemColors.Control
		Me.Label9.Enabled = True
		Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label9.UseMnemonic = True
		Me.Label9.Visible = True
		Me.Label9.AutoSize = False
		Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label9.Name = "Label9"
		Me.Label10.Text = "Round Trip Distance:"
		Me.Label10.Size = New System.Drawing.Size(153, 17)
		Me.Label10.Location = New System.Drawing.Point(112, 104)
		Me.Label10.TabIndex = 16
		Me.Label10.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label10.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label10.BackColor = System.Drawing.SystemColors.Control
		Me.Label10.Enabled = True
		Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label10.UseMnemonic = True
		Me.Label10.Visible = True
		Me.Label10.AutoSize = False
		Me.Label10.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label10.Name = "Label10"
		Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label7.Text = "commodity"
		Me.Label7.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label7.Size = New System.Drawing.Size(65, 17)
		Me.Label7.Location = New System.Drawing.Point(272, 56)
		Me.Label7.TabIndex = 15
		Me.Label7.Visible = False
		Me.Label7.BackColor = System.Drawing.SystemColors.Control
		Me.Label7.Enabled = True
		Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label7.UseMnemonic = True
		Me.Label7.AutoSize = False
		Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label7.Name = "Label7"
		Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label6.Text = "Assembly"
		Me.Label6.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label6.Size = New System.Drawing.Size(65, 17)
		Me.Label6.Location = New System.Drawing.Point(272, 32)
		Me.Label6.TabIndex = 14
		Me.Label6.BackColor = System.Drawing.SystemColors.Control
		Me.Label6.Enabled = True
		Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label6.UseMnemonic = True
		Me.Label6.Visible = True
		Me.Label6.AutoSize = False
		Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label6.Name = "Label6"
		Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label5.Text = "path"
		Me.Label5.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label5.Size = New System.Drawing.Size(65, 17)
		Me.Label5.Location = New System.Drawing.Point(272, 8)
		Me.Label5.TabIndex = 13
		Me.Label5.BackColor = System.Drawing.SystemColors.Control
		Me.Label5.Enabled = True
		Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label5.UseMnemonic = True
		Me.Label5.Visible = True
		Me.Label5.AutoSize = False
		Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label5.Name = "Label5"
		Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label4.Text = "escorts"
		Me.Label4.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.Size = New System.Drawing.Size(49, 17)
		Me.Label4.Location = New System.Drawing.Point(56, 32)
		Me.Label4.TabIndex = 12
		Me.Label4.BackColor = System.Drawing.SystemColors.Control
		Me.Label4.Enabled = True
		Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label4.UseMnemonic = True
		Me.Label4.Visible = True
		Me.Label4.AutoSize = False
		Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label4.Name = "Label4"
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label3.Text = "with"
		Me.Label3.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.Size = New System.Drawing.Size(25, 17)
		Me.Label3.Location = New System.Drawing.Point(80, 8)
		Me.Label3.TabIndex = 11
		Me.Label3.BackColor = System.Drawing.SystemColors.Control
		Me.Label3.Enabled = True
		Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label3.UseMnemonic = True
		Me.Label3.Visible = True
		Me.Label3.AutoSize = False
		Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label3.Name = "Label3"
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label2.Text = "Must Fill"
		Me.Label2.Font = New System.Drawing.Font("Arial Black", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(57, 17)
		Me.Label2.Location = New System.Drawing.Point(8, 8)
		Me.Label2.TabIndex = 10
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
		Me.Label1.Text = "times"
		Me.Label1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Size = New System.Drawing.Size(33, 17)
		Me.Label1.Location = New System.Drawing.Point(464, 88)
		Me.Label1.TabIndex = 9
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
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
		Me.Controls.Add(chkDisplayPath)
		Me.Controls.Add(cmdHelp)
		Me.Controls.Add(cbCom)
		Me.Controls.Add(txtEscort)
		Me.Controls.Add(txtPath)
		Me.Controls.Add(txtOrigin)
		Me.Controls.Add(txtUnit)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(Check1)
		Me.Controls.Add(Combo1)
		Me.Controls.Add(Label8)
		Me.Controls.Add(Label9)
		Me.Controls.Add(Label10)
		Me.Controls.Add(Label7)
		Me.Controls.Add(Label6)
		Me.Controls.Add(Label5)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class