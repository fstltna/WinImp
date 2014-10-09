<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmEmpireLogin
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
	Public WithEvents cmbThemeGame As System.Windows.Forms.ComboBox
	Public WithEvents cmdHTTP As System.Windows.Forms.Button
	Public WithEvents chkHTTP As System.Windows.Forms.CheckBox
	Public WithEvents chkSanc As System.Windows.Forms.CheckBox
	Public WithEvents chkKill As System.Windows.Forms.CheckBox
	Public WithEvents chkUTF8 As System.Windows.Forms.CheckBox
	Public WithEvents lblThemeGame As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents cmdOffline As System.Windows.Forms.Button
	Public WithEvents txtUser As System.Windows.Forms.TextBox
	Public WithEvents chkDeity As System.Windows.Forms.CheckBox
	Public WithEvents chkAutoRead As System.Windows.Forms.CheckBox
	Public WithEvents chkUpdate As System.Windows.Forms.CheckBox
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	Public WithEvents cmdConnect As System.Windows.Forms.Button
	Public WithEvents txtPort As System.Windows.Forms.TextBox
	Public WithEvents txtHostName As System.Windows.Forms.TextBox
	Public WithEvents txtUserName As System.Windows.Forms.TextBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents txtPassword As System.Windows.Forms.TextBox
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents _lblLabels_4 As System.Windows.Forms.Label
	Public WithEvents lblGame As System.Windows.Forms.Label
	Public WithEvents Line1 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents _lblLabels_1 As System.Windows.Forms.Label
	Public WithEvents _lblLabels_0 As System.Windows.Forms.Label
	Public WithEvents _lblLabels_2 As System.Windows.Forms.Label
	Public WithEvents _lblLabels_3 As System.Windows.Forms.Label
	Public WithEvents lblLabels As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents ShapeContainer1 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmEmpireLogin))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.cmbThemeGame = New System.Windows.Forms.ComboBox
		Me.cmdHTTP = New System.Windows.Forms.Button
		Me.chkHTTP = New System.Windows.Forms.CheckBox
		Me.chkSanc = New System.Windows.Forms.CheckBox
		Me.chkKill = New System.Windows.Forms.CheckBox
		Me.chkUTF8 = New System.Windows.Forms.CheckBox
		Me.lblThemeGame = New System.Windows.Forms.Label
		Me.cmdOffline = New System.Windows.Forms.Button
		Me.txtUser = New System.Windows.Forms.TextBox
		Me.Frame2 = New System.Windows.Forms.GroupBox
		Me.chkDeity = New System.Windows.Forms.CheckBox
		Me.chkAutoRead = New System.Windows.Forms.CheckBox
		Me.chkUpdate = New System.Windows.Forms.CheckBox
		Me.Timer1 = New System.Windows.Forms.Timer(components)
		Me.cmdConnect = New System.Windows.Forms.Button
		Me.txtPort = New System.Windows.Forms.TextBox
		Me.txtHostName = New System.Windows.Forms.TextBox
		Me.txtUserName = New System.Windows.Forms.TextBox
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.txtPassword = New System.Windows.Forms.TextBox
		Me.Label1 = New System.Windows.Forms.Label
		Me._lblLabels_4 = New System.Windows.Forms.Label
		Me.lblGame = New System.Windows.Forms.Label
		Me.Line1 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me._lblLabels_1 = New System.Windows.Forms.Label
		Me._lblLabels_0 = New System.Windows.Forms.Label
		Me._lblLabels_2 = New System.Windows.Forms.Label
		Me._lblLabels_3 = New System.Windows.Forms.Label
		Me.lblLabels = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.Frame1.SuspendLayout()
		Me.Frame2.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.lblLabels, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Empire Login"
		Me.ClientSize = New System.Drawing.Size(305, 384)
		Me.Location = New System.Drawing.Point(189, 232)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmEmpireLogin"
		Me.Frame1.Text = "Server Options"
		Me.Frame1.Size = New System.Drawing.Size(289, 89)
		Me.Frame1.Location = New System.Drawing.Point(8, 208)
		Me.Frame1.TabIndex = 17
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.cmbThemeGame.Size = New System.Drawing.Size(201, 21)
		Me.cmbThemeGame.Location = New System.Drawing.Point(80, 60)
		Me.cmbThemeGame.TabIndex = 26
		Me.cmbThemeGame.Text = "Combo1"
		Me.cmbThemeGame.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbThemeGame.BackColor = System.Drawing.SystemColors.Window
		Me.cmbThemeGame.CausesValidation = True
		Me.cmbThemeGame.Enabled = True
		Me.cmbThemeGame.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbThemeGame.IntegralHeight = True
		Me.cmbThemeGame.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbThemeGame.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbThemeGame.Sorted = False
		Me.cmbThemeGame.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cmbThemeGame.TabStop = True
		Me.cmbThemeGame.Visible = True
		Me.cmbThemeGame.Name = "cmbThemeGame"
		Me.cmdHTTP.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdHTTP.Text = "Settings"
		Me.cmdHTTP.Size = New System.Drawing.Size(57, 17)
		Me.cmdHTTP.Location = New System.Drawing.Point(80, 40)
		Me.cmdHTTP.TabIndex = 24
		Me.cmdHTTP.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdHTTP.BackColor = System.Drawing.SystemColors.Control
		Me.cmdHTTP.CausesValidation = True
		Me.cmdHTTP.Enabled = True
		Me.cmdHTTP.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdHTTP.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdHTTP.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdHTTP.TabStop = True
		Me.cmdHTTP.Name = "cmdHTTP"
		Me.chkHTTP.Text = "HTTP"
		Me.chkHTTP.Size = New System.Drawing.Size(49, 17)
		Me.chkHTTP.Location = New System.Drawing.Point(8, 40)
		Me.chkHTTP.TabIndex = 23
		Me.ToolTip1.SetToolTip(Me.chkHTTP, "Allows the use of Unicode (Multiple Languages) on the game server")
		Me.chkHTTP.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkHTTP.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkHTTP.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkHTTP.BackColor = System.Drawing.SystemColors.Control
		Me.chkHTTP.CausesValidation = True
		Me.chkHTTP.Enabled = True
		Me.chkHTTP.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkHTTP.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkHTTP.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkHTTP.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkHTTP.TabStop = True
		Me.chkHTTP.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkHTTP.Visible = True
		Me.chkHTTP.Name = "chkHTTP"
		Me.chkSanc.Text = "List Sanctuaries"
		Me.chkSanc.Size = New System.Drawing.Size(97, 17)
		Me.chkSanc.Location = New System.Drawing.Point(120, 16)
		Me.chkSanc.TabIndex = 7
		Me.chkSanc.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkSanc.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkSanc.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkSanc.BackColor = System.Drawing.SystemColors.Control
		Me.chkSanc.CausesValidation = True
		Me.chkSanc.Enabled = True
		Me.chkSanc.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkSanc.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkSanc.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkSanc.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkSanc.TabStop = True
		Me.chkSanc.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkSanc.Visible = True
		Me.chkSanc.Name = "chkSanc"
		Me.chkKill.Text = "Kill Hung Session"
		Me.chkKill.Size = New System.Drawing.Size(105, 17)
		Me.chkKill.Location = New System.Drawing.Point(8, 16)
		Me.chkKill.TabIndex = 6
		Me.chkKill.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkKill.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkKill.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkKill.BackColor = System.Drawing.SystemColors.Control
		Me.chkKill.CausesValidation = True
		Me.chkKill.Enabled = True
		Me.chkKill.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkKill.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkKill.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkKill.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkKill.TabStop = True
		Me.chkKill.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkKill.Visible = True
		Me.chkKill.Name = "chkKill"
		Me.chkUTF8.Text = "UTF-8"
		Me.chkUTF8.Size = New System.Drawing.Size(57, 17)
		Me.chkUTF8.Location = New System.Drawing.Point(224, 16)
		Me.chkUTF8.TabIndex = 22
		Me.ToolTip1.SetToolTip(Me.chkUTF8, "Allows the use of Unicode (Multiple Languages) on the game server")
		Me.chkUTF8.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkUTF8.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkUTF8.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkUTF8.BackColor = System.Drawing.SystemColors.Control
		Me.chkUTF8.CausesValidation = True
		Me.chkUTF8.Enabled = True
		Me.chkUTF8.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkUTF8.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkUTF8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkUTF8.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkUTF8.TabStop = True
		Me.chkUTF8.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkUTF8.Visible = True
		Me.chkUTF8.Name = "chkUTF8"
		Me.lblThemeGame.Text = "Theme Game"
		Me.lblThemeGame.Size = New System.Drawing.Size(73, 17)
		Me.lblThemeGame.Location = New System.Drawing.Point(8, 64)
		Me.lblThemeGame.TabIndex = 25
		Me.lblThemeGame.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblThemeGame.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblThemeGame.BackColor = System.Drawing.SystemColors.Control
		Me.lblThemeGame.Enabled = True
		Me.lblThemeGame.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblThemeGame.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblThemeGame.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblThemeGame.UseMnemonic = True
		Me.lblThemeGame.Visible = True
		Me.lblThemeGame.AutoSize = False
		Me.lblThemeGame.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblThemeGame.Name = "lblThemeGame"
		Me.cmdOffline.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOffline.Text = "Offline"
		Me.cmdOffline.Size = New System.Drawing.Size(84, 26)
		Me.cmdOffline.Location = New System.Drawing.Point(112, 352)
		Me.cmdOffline.TabIndex = 21
		Me.cmdOffline.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOffline.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOffline.CausesValidation = True
		Me.cmdOffline.Enabled = True
		Me.cmdOffline.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOffline.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOffline.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOffline.TabStop = True
		Me.cmdOffline.Name = "cmdOffline"
		Me.txtUser.AutoSize = False
		Me.txtUser.Size = New System.Drawing.Size(211, 23)
		Me.txtUser.IMEMode = System.Windows.Forms.ImeMode.Disable
		Me.txtUser.Location = New System.Drawing.Point(88, 152)
		Me.txtUser.TabIndex = 5
		Me.txtUser.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtUser.AcceptsReturn = True
		Me.txtUser.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtUser.BackColor = System.Drawing.SystemColors.Window
		Me.txtUser.CausesValidation = True
		Me.txtUser.Enabled = True
		Me.txtUser.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtUser.HideSelection = True
		Me.txtUser.ReadOnly = False
		Me.txtUser.Maxlength = 0
		Me.txtUser.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtUser.MultiLine = False
		Me.txtUser.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtUser.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtUser.TabStop = True
		Me.txtUser.Visible = True
		Me.txtUser.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtUser.Name = "txtUser"
		Me.Frame2.Text = "Client Options"
		Me.Frame2.Size = New System.Drawing.Size(289, 41)
		Me.Frame2.Location = New System.Drawing.Point(8, 304)
		Me.Frame2.TabIndex = 18
		Me.Frame2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame2.BackColor = System.Drawing.SystemColors.Control
		Me.Frame2.Enabled = True
		Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame2.Visible = True
		Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame2.Name = "Frame2"
		Me.chkDeity.Text = "Sign on as Deity"
		Me.chkDeity.Size = New System.Drawing.Size(97, 17)
		Me.chkDeity.Location = New System.Drawing.Point(184, 16)
		Me.chkDeity.TabIndex = 20
		Me.chkDeity.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkDeity.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkDeity.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkDeity.BackColor = System.Drawing.SystemColors.Control
		Me.chkDeity.CausesValidation = True
		Me.chkDeity.Enabled = True
		Me.chkDeity.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkDeity.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkDeity.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkDeity.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkDeity.TabStop = True
		Me.chkDeity.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkDeity.Visible = True
		Me.chkDeity.Name = "chkDeity"
		Me.chkAutoRead.Text = "AutoRead"
		Me.chkAutoRead.Size = New System.Drawing.Size(73, 17)
		Me.chkAutoRead.Location = New System.Drawing.Point(104, 16)
		Me.chkAutoRead.TabIndex = 9
		Me.chkAutoRead.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkAutoRead.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkAutoRead.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkAutoRead.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkAutoRead.BackColor = System.Drawing.SystemColors.Control
		Me.chkAutoRead.CausesValidation = True
		Me.chkAutoRead.Enabled = True
		Me.chkAutoRead.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkAutoRead.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkAutoRead.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkAutoRead.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkAutoRead.TabStop = True
		Me.chkAutoRead.Visible = True
		Me.chkAutoRead.Name = "chkAutoRead"
		Me.chkUpdate.Text = "AutoUpdate"
		Me.chkUpdate.Size = New System.Drawing.Size(81, 17)
		Me.chkUpdate.Location = New System.Drawing.Point(8, 16)
		Me.chkUpdate.TabIndex = 8
		Me.chkUpdate.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkUpdate.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkUpdate.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkUpdate.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkUpdate.BackColor = System.Drawing.SystemColors.Control
		Me.chkUpdate.CausesValidation = True
		Me.chkUpdate.Enabled = True
		Me.chkUpdate.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkUpdate.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkUpdate.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkUpdate.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkUpdate.TabStop = True
		Me.chkUpdate.Visible = True
		Me.chkUpdate.Name = "chkUpdate"
		Me.Timer1.Enabled = False
		Me.Timer1.Interval = 1
		Me.cmdConnect.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdConnect.Text = "Connect"
		Me.AcceptButton = Me.cmdConnect
		Me.cmdConnect.Size = New System.Drawing.Size(84, 26)
		Me.cmdConnect.Location = New System.Drawing.Point(8, 352)
		Me.cmdConnect.TabIndex = 10
		Me.cmdConnect.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdConnect.BackColor = System.Drawing.SystemColors.Control
		Me.cmdConnect.CausesValidation = True
		Me.cmdConnect.Enabled = True
		Me.cmdConnect.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdConnect.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdConnect.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdConnect.TabStop = True
		Me.cmdConnect.Name = "cmdConnect"
		Me.txtPort.AutoSize = False
		Me.txtPort.Size = New System.Drawing.Size(209, 23)
		Me.txtPort.IMEMode = System.Windows.Forms.ImeMode.Disable
		Me.txtPort.Location = New System.Drawing.Point(88, 80)
		Me.txtPort.TabIndex = 2
		Me.txtPort.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtPort.AcceptsReturn = True
		Me.txtPort.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtPort.BackColor = System.Drawing.SystemColors.Window
		Me.txtPort.CausesValidation = True
		Me.txtPort.Enabled = True
		Me.txtPort.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtPort.HideSelection = True
		Me.txtPort.ReadOnly = False
		Me.txtPort.Maxlength = 0
		Me.txtPort.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtPort.MultiLine = False
		Me.txtPort.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtPort.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtPort.TabStop = True
		Me.txtPort.Visible = True
		Me.txtPort.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtPort.Name = "txtPort"
		Me.txtHostName.AutoSize = False
		Me.txtHostName.Size = New System.Drawing.Size(209, 23)
		Me.txtHostName.Location = New System.Drawing.Point(88, 56)
		Me.txtHostName.TabIndex = 1
		Me.txtHostName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtHostName.AcceptsReturn = True
		Me.txtHostName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtHostName.BackColor = System.Drawing.SystemColors.Window
		Me.txtHostName.CausesValidation = True
		Me.txtHostName.Enabled = True
		Me.txtHostName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtHostName.HideSelection = True
		Me.txtHostName.ReadOnly = False
		Me.txtHostName.Maxlength = 0
		Me.txtHostName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtHostName.MultiLine = False
		Me.txtHostName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtHostName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtHostName.TabStop = True
		Me.txtHostName.Visible = True
		Me.txtHostName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtHostName.Name = "txtHostName"
		Me.txtUserName.AutoSize = False
		Me.txtUserName.Size = New System.Drawing.Size(211, 23)
		Me.txtUserName.Location = New System.Drawing.Point(88, 104)
		Me.txtUserName.TabIndex = 3
		Me.txtUserName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtUserName.AcceptsReturn = True
		Me.txtUserName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtUserName.BackColor = System.Drawing.SystemColors.Window
		Me.txtUserName.CausesValidation = True
		Me.txtUserName.Enabled = True
		Me.txtUserName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtUserName.HideSelection = True
		Me.txtUserName.ReadOnly = False
		Me.txtUserName.Maxlength = 0
		Me.txtUserName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtUserName.MultiLine = False
		Me.txtUserName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtUserName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtUserName.TabStop = True
		Me.txtUserName.Visible = True
		Me.txtUserName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtUserName.Name = "txtUserName"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(84, 26)
		Me.cmdCancel.Location = New System.Drawing.Point(216, 352)
		Me.cmdCancel.TabIndex = 11
		Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.txtPassword.AutoSize = False
		Me.txtPassword.Size = New System.Drawing.Size(211, 23)
		Me.txtPassword.IMEMode = System.Windows.Forms.ImeMode.Disable
		Me.txtPassword.Location = New System.Drawing.Point(88, 128)
		Me.txtPassword.TabIndex = 4
		Me.txtPassword.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtPassword.AcceptsReturn = True
		Me.txtPassword.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtPassword.BackColor = System.Drawing.SystemColors.Window
		Me.txtPassword.CausesValidation = True
		Me.txtPassword.Enabled = True
		Me.txtPassword.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtPassword.HideSelection = True
		Me.txtPassword.ReadOnly = False
		Me.txtPassword.Maxlength = 0
		Me.txtPassword.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtPassword.MultiLine = False
		Me.txtPassword.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtPassword.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtPassword.TabStop = True
		Me.txtPassword.Visible = True
		Me.txtPassword.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtPassword.Name = "txtPassword"
		Me.Label1.Size = New System.Drawing.Size(289, 17)
		Me.Label1.Location = New System.Drawing.Point(8, 184)
		Me.Label1.TabIndex = 15
		Me.Label1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me._lblLabels_4.Text = "&User:"
		Me._lblLabels_4.Size = New System.Drawing.Size(72, 18)
		Me._lblLabels_4.Location = New System.Drawing.Point(8, 152)
		Me._lblLabels_4.TabIndex = 19
		Me._lblLabels_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblLabels_4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblLabels_4.BackColor = System.Drawing.SystemColors.Control
		Me._lblLabels_4.Enabled = True
		Me._lblLabels_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblLabels_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblLabels_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblLabels_4.UseMnemonic = True
		Me._lblLabels_4.Visible = True
		Me._lblLabels_4.AutoSize = False
		Me._lblLabels_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblLabels_4.Name = "_lblLabels_4"
		Me.lblGame.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblGame.Size = New System.Drawing.Size(305, 25)
		Me.lblGame.Location = New System.Drawing.Point(8, 8)
		Me.lblGame.TabIndex = 16
		Me.lblGame.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblGame.BackColor = System.Drawing.SystemColors.Control
		Me.lblGame.Enabled = True
		Me.lblGame.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblGame.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblGame.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblGame.UseMnemonic = True
		Me.lblGame.Visible = True
		Me.lblGame.AutoSize = False
		Me.lblGame.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblGame.Name = "lblGame"
		Me.Line1.X1 = 8
		Me.Line1.X2 = 293
		Me.Line1.Y1 = 24
		Me.Line1.Y2 = 24
		Me.Line1.BorderColor = System.Drawing.SystemColors.WindowText
		Me.Line1.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line1.BorderWidth = 1
		Me.Line1.Visible = True
		Me.Line1.Name = "Line1"
		Me._lblLabels_1.Text = "&Port Number"
		Me._lblLabels_1.Size = New System.Drawing.Size(72, 18)
		Me._lblLabels_1.Location = New System.Drawing.Point(8, 80)
		Me._lblLabels_1.TabIndex = 14
		Me._lblLabels_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblLabels_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblLabels_1.BackColor = System.Drawing.SystemColors.Control
		Me._lblLabels_1.Enabled = True
		Me._lblLabels_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblLabels_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblLabels_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblLabels_1.UseMnemonic = True
		Me._lblLabels_1.Visible = True
		Me._lblLabels_1.AutoSize = False
		Me._lblLabels_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblLabels_1.Name = "_lblLabels_1"
		Me._lblLabels_0.Text = "&Host Name:"
		Me._lblLabels_0.Size = New System.Drawing.Size(72, 18)
		Me._lblLabels_0.Location = New System.Drawing.Point(8, 56)
		Me._lblLabels_0.TabIndex = 13
		Me._lblLabels_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblLabels_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblLabels_0.BackColor = System.Drawing.SystemColors.Control
		Me._lblLabels_0.Enabled = True
		Me._lblLabels_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblLabels_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblLabels_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblLabels_0.UseMnemonic = True
		Me._lblLabels_0.Visible = True
		Me._lblLabels_0.AutoSize = False
		Me._lblLabels_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblLabels_0.Name = "_lblLabels_0"
		Me._lblLabels_2.Text = "&Country Name:"
		Me._lblLabels_2.Size = New System.Drawing.Size(72, 18)
		Me._lblLabels_2.Location = New System.Drawing.Point(7, 104)
		Me._lblLabels_2.TabIndex = 0
		Me._lblLabels_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblLabels_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblLabels_2.BackColor = System.Drawing.SystemColors.Control
		Me._lblLabels_2.Enabled = True
		Me._lblLabels_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblLabels_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblLabels_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblLabels_2.UseMnemonic = True
		Me._lblLabels_2.Visible = True
		Me._lblLabels_2.AutoSize = False
		Me._lblLabels_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblLabels_2.Name = "_lblLabels_2"
		Me._lblLabels_3.Text = "&Password:"
		Me._lblLabels_3.Size = New System.Drawing.Size(72, 18)
		Me._lblLabels_3.Location = New System.Drawing.Point(7, 128)
		Me._lblLabels_3.TabIndex = 12
		Me._lblLabels_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblLabels_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblLabels_3.BackColor = System.Drawing.SystemColors.Control
		Me._lblLabels_3.Enabled = True
		Me._lblLabels_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblLabels_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblLabels_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblLabels_3.UseMnemonic = True
		Me._lblLabels_3.Visible = True
		Me._lblLabels_3.AutoSize = False
		Me._lblLabels_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblLabels_3.Name = "_lblLabels_3"
		Me.Controls.Add(Frame1)
		Me.Controls.Add(cmdOffline)
		Me.Controls.Add(txtUser)
		Me.Controls.Add(Frame2)
		Me.Controls.Add(cmdConnect)
		Me.Controls.Add(txtPort)
		Me.Controls.Add(txtHostName)
		Me.Controls.Add(txtUserName)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(txtPassword)
		Me.Controls.Add(Label1)
		Me.Controls.Add(_lblLabels_4)
		Me.Controls.Add(lblGame)
		Me.ShapeContainer1.Shapes.Add(Line1)
		Me.Controls.Add(_lblLabels_1)
		Me.Controls.Add(_lblLabels_0)
		Me.Controls.Add(_lblLabels_2)
		Me.Controls.Add(_lblLabels_3)
		Me.Controls.Add(ShapeContainer1)
		Me.Frame1.Controls.Add(cmbThemeGame)
		Me.Frame1.Controls.Add(cmdHTTP)
		Me.Frame1.Controls.Add(chkHTTP)
		Me.Frame1.Controls.Add(chkSanc)
		Me.Frame1.Controls.Add(chkKill)
		Me.Frame1.Controls.Add(chkUTF8)
		Me.Frame1.Controls.Add(lblThemeGame)
		Me.Frame2.Controls.Add(chkDeity)
		Me.Frame2.Controls.Add(chkAutoRead)
		Me.Frame2.Controls.Add(chkUpdate)
		Me.lblLabels.SetIndex(_lblLabels_4, CType(4, Short))
		Me.lblLabels.SetIndex(_lblLabels_1, CType(1, Short))
		Me.lblLabels.SetIndex(_lblLabels_0, CType(0, Short))
		Me.lblLabels.SetIndex(_lblLabels_2, CType(2, Short))
		Me.lblLabels.SetIndex(_lblLabels_3, CType(3, Short))
		CType(Me.lblLabels, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Frame1.ResumeLayout(False)
		Me.Frame2.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class