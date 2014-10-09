<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmHTTPSettings
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
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents chkProxySettings As System.Windows.Forms.CheckBox
	Public WithEvents txtProxyPassword As System.Windows.Forms.TextBox
	Public WithEvents txtProxyServerName As System.Windows.Forms.TextBox
	Public WithEvents txtProxyUser As System.Windows.Forms.TextBox
	Public WithEvents lblProxyPassword As System.Windows.Forms.Label
	Public WithEvents lblProxyServerName As System.Windows.Forms.Label
	Public WithEvents lblProxyUser As System.Windows.Forms.Label
	Public WithEvents fraProxySettings As System.Windows.Forms.GroupBox
	Public WithEvents txtXMLRPCServerParameters As System.Windows.Forms.TextBox
	Public WithEvents txtXMLRPCServerName As System.Windows.Forms.TextBox
	Public WithEvents lblXMLRPCServerParameters As System.Windows.Forms.Label
	Public WithEvents lblXMLRPCServerName As System.Windows.Forms.Label
	Public WithEvents fraXMLRPC As System.Windows.Forms.GroupBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmHTTPSettings))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me.fraProxySettings = New System.Windows.Forms.GroupBox
		Me.chkProxySettings = New System.Windows.Forms.CheckBox
		Me.txtProxyPassword = New System.Windows.Forms.TextBox
		Me.txtProxyServerName = New System.Windows.Forms.TextBox
		Me.txtProxyUser = New System.Windows.Forms.TextBox
		Me.lblProxyPassword = New System.Windows.Forms.Label
		Me.lblProxyServerName = New System.Windows.Forms.Label
		Me.lblProxyUser = New System.Windows.Forms.Label
		Me.fraXMLRPC = New System.Windows.Forms.GroupBox
		Me.txtXMLRPCServerParameters = New System.Windows.Forms.TextBox
		Me.txtXMLRPCServerName = New System.Windows.Forms.TextBox
		Me.lblXMLRPCServerParameters = New System.Windows.Forms.Label
		Me.lblXMLRPCServerName = New System.Windows.Forms.Label
		Me.fraProxySettings.SuspendLayout()
		Me.fraXMLRPC.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "HTTP Proxy Settings"
		Me.ClientSize = New System.Drawing.Size(274, 241)
		Me.Location = New System.Drawing.Point(4, 30)
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
		Me.Name = "frmHTTPSettings"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(89, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(152, 208)
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
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "OK"
		Me.cmdOK.Size = New System.Drawing.Size(89, 25)
		Me.cmdOK.Location = New System.Drawing.Point(32, 208)
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
		Me.fraProxySettings.Text = "Proxy Settings"
		Me.fraProxySettings.Size = New System.Drawing.Size(257, 113)
		Me.fraProxySettings.Location = New System.Drawing.Point(8, 88)
		Me.fraProxySettings.TabIndex = 5
		Me.fraProxySettings.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraProxySettings.BackColor = System.Drawing.SystemColors.Control
		Me.fraProxySettings.Enabled = True
		Me.fraProxySettings.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraProxySettings.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraProxySettings.Visible = True
		Me.fraProxySettings.Padding = New System.Windows.Forms.Padding(0)
		Me.fraProxySettings.Name = "fraProxySettings"
		Me.chkProxySettings.Text = "Use Proxy Settings"
		Me.chkProxySettings.Size = New System.Drawing.Size(121, 17)
		Me.chkProxySettings.Location = New System.Drawing.Point(16, 16)
		Me.chkProxySettings.TabIndex = 12
		Me.chkProxySettings.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkProxySettings.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkProxySettings.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkProxySettings.BackColor = System.Drawing.SystemColors.Control
		Me.chkProxySettings.CausesValidation = True
		Me.chkProxySettings.Enabled = True
		Me.chkProxySettings.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkProxySettings.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkProxySettings.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkProxySettings.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkProxySettings.TabStop = True
		Me.chkProxySettings.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkProxySettings.Visible = True
		Me.chkProxySettings.Name = "chkProxySettings"
		Me.txtProxyPassword.AutoSize = False
		Me.txtProxyPassword.Size = New System.Drawing.Size(177, 19)
		Me.txtProxyPassword.Location = New System.Drawing.Point(72, 88)
		Me.txtProxyPassword.TabIndex = 11
		Me.ToolTip1.SetToolTip(Me.txtProxyPassword, "Leave Blank if not needed")
		Me.txtProxyPassword.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtProxyPassword.AcceptsReturn = True
		Me.txtProxyPassword.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtProxyPassword.BackColor = System.Drawing.SystemColors.Window
		Me.txtProxyPassword.CausesValidation = True
		Me.txtProxyPassword.Enabled = True
		Me.txtProxyPassword.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtProxyPassword.HideSelection = True
		Me.txtProxyPassword.ReadOnly = False
		Me.txtProxyPassword.Maxlength = 0
		Me.txtProxyPassword.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtProxyPassword.MultiLine = False
		Me.txtProxyPassword.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtProxyPassword.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtProxyPassword.TabStop = True
		Me.txtProxyPassword.Visible = True
		Me.txtProxyPassword.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtProxyPassword.Name = "txtProxyPassword"
		Me.txtProxyServerName.AutoSize = False
		Me.txtProxyServerName.Size = New System.Drawing.Size(177, 19)
		Me.txtProxyServerName.Location = New System.Drawing.Point(72, 40)
		Me.txtProxyServerName.TabIndex = 9
		Me.ToolTip1.SetToolTip(Me.txtProxyServerName, "Proxy Server Name")
		Me.txtProxyServerName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtProxyServerName.AcceptsReturn = True
		Me.txtProxyServerName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtProxyServerName.BackColor = System.Drawing.SystemColors.Window
		Me.txtProxyServerName.CausesValidation = True
		Me.txtProxyServerName.Enabled = True
		Me.txtProxyServerName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtProxyServerName.HideSelection = True
		Me.txtProxyServerName.ReadOnly = False
		Me.txtProxyServerName.Maxlength = 0
		Me.txtProxyServerName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtProxyServerName.MultiLine = False
		Me.txtProxyServerName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtProxyServerName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtProxyServerName.TabStop = True
		Me.txtProxyServerName.Visible = True
		Me.txtProxyServerName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtProxyServerName.Name = "txtProxyServerName"
		Me.txtProxyUser.AutoSize = False
		Me.txtProxyUser.Size = New System.Drawing.Size(177, 19)
		Me.txtProxyUser.Location = New System.Drawing.Point(72, 64)
		Me.txtProxyUser.TabIndex = 7
		Me.ToolTip1.SetToolTip(Me.txtProxyUser, "Leave Blank if not needed")
		Me.txtProxyUser.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtProxyUser.AcceptsReturn = True
		Me.txtProxyUser.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtProxyUser.BackColor = System.Drawing.SystemColors.Window
		Me.txtProxyUser.CausesValidation = True
		Me.txtProxyUser.Enabled = True
		Me.txtProxyUser.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtProxyUser.HideSelection = True
		Me.txtProxyUser.ReadOnly = False
		Me.txtProxyUser.Maxlength = 0
		Me.txtProxyUser.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtProxyUser.MultiLine = False
		Me.txtProxyUser.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtProxyUser.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtProxyUser.TabStop = True
		Me.txtProxyUser.Visible = True
		Me.txtProxyUser.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtProxyUser.Name = "txtProxyUser"
		Me.lblProxyPassword.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblProxyPassword.Text = "Password"
		Me.lblProxyPassword.Size = New System.Drawing.Size(49, 17)
		Me.lblProxyPassword.Location = New System.Drawing.Point(16, 90)
		Me.lblProxyPassword.TabIndex = 10
		Me.lblProxyPassword.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblProxyPassword.BackColor = System.Drawing.SystemColors.Control
		Me.lblProxyPassword.Enabled = True
		Me.lblProxyPassword.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblProxyPassword.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblProxyPassword.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblProxyPassword.UseMnemonic = True
		Me.lblProxyPassword.Visible = True
		Me.lblProxyPassword.AutoSize = False
		Me.lblProxyPassword.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblProxyPassword.Name = "lblProxyPassword"
		Me.lblProxyServerName.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblProxyServerName.Text = "Server"
		Me.lblProxyServerName.Size = New System.Drawing.Size(33, 17)
		Me.lblProxyServerName.Location = New System.Drawing.Point(32, 40)
		Me.lblProxyServerName.TabIndex = 8
		Me.lblProxyServerName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblProxyServerName.BackColor = System.Drawing.SystemColors.Control
		Me.lblProxyServerName.Enabled = True
		Me.lblProxyServerName.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblProxyServerName.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblProxyServerName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblProxyServerName.UseMnemonic = True
		Me.lblProxyServerName.Visible = True
		Me.lblProxyServerName.AutoSize = False
		Me.lblProxyServerName.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblProxyServerName.Name = "lblProxyServerName"
		Me.lblProxyUser.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblProxyUser.Text = "User"
		Me.lblProxyUser.Size = New System.Drawing.Size(41, 17)
		Me.lblProxyUser.Location = New System.Drawing.Point(24, 64)
		Me.lblProxyUser.TabIndex = 6
		Me.lblProxyUser.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblProxyUser.BackColor = System.Drawing.SystemColors.Control
		Me.lblProxyUser.Enabled = True
		Me.lblProxyUser.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblProxyUser.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblProxyUser.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblProxyUser.UseMnemonic = True
		Me.lblProxyUser.Visible = True
		Me.lblProxyUser.AutoSize = False
		Me.lblProxyUser.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblProxyUser.Name = "lblProxyUser"
		Me.fraXMLRPC.Text = "XML RPC"
		Me.fraXMLRPC.Size = New System.Drawing.Size(257, 65)
		Me.fraXMLRPC.Location = New System.Drawing.Point(8, 16)
		Me.fraXMLRPC.TabIndex = 0
		Me.fraXMLRPC.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraXMLRPC.BackColor = System.Drawing.SystemColors.Control
		Me.fraXMLRPC.Enabled = True
		Me.fraXMLRPC.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraXMLRPC.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraXMLRPC.Visible = True
		Me.fraXMLRPC.Padding = New System.Windows.Forms.Padding(0)
		Me.fraXMLRPC.Name = "fraXMLRPC"
		Me.txtXMLRPCServerParameters.AutoSize = False
		Me.txtXMLRPCServerParameters.Size = New System.Drawing.Size(177, 19)
		Me.txtXMLRPCServerParameters.Location = New System.Drawing.Point(72, 40)
		Me.txtXMLRPCServerParameters.TabIndex = 4
		Me.txtXMLRPCServerParameters.Text = "Text1"
		Me.txtXMLRPCServerParameters.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtXMLRPCServerParameters.AcceptsReturn = True
		Me.txtXMLRPCServerParameters.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtXMLRPCServerParameters.BackColor = System.Drawing.SystemColors.Window
		Me.txtXMLRPCServerParameters.CausesValidation = True
		Me.txtXMLRPCServerParameters.Enabled = True
		Me.txtXMLRPCServerParameters.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtXMLRPCServerParameters.HideSelection = True
		Me.txtXMLRPCServerParameters.ReadOnly = False
		Me.txtXMLRPCServerParameters.Maxlength = 0
		Me.txtXMLRPCServerParameters.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtXMLRPCServerParameters.MultiLine = False
		Me.txtXMLRPCServerParameters.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtXMLRPCServerParameters.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtXMLRPCServerParameters.TabStop = True
		Me.txtXMLRPCServerParameters.Visible = True
		Me.txtXMLRPCServerParameters.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtXMLRPCServerParameters.Name = "txtXMLRPCServerParameters"
		Me.txtXMLRPCServerName.AutoSize = False
		Me.txtXMLRPCServerName.Size = New System.Drawing.Size(177, 19)
		Me.txtXMLRPCServerName.Location = New System.Drawing.Point(72, 16)
		Me.txtXMLRPCServerName.TabIndex = 2
		Me.txtXMLRPCServerName.Text = "Text1"
		Me.txtXMLRPCServerName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtXMLRPCServerName.AcceptsReturn = True
		Me.txtXMLRPCServerName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtXMLRPCServerName.BackColor = System.Drawing.SystemColors.Window
		Me.txtXMLRPCServerName.CausesValidation = True
		Me.txtXMLRPCServerName.Enabled = True
		Me.txtXMLRPCServerName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtXMLRPCServerName.HideSelection = True
		Me.txtXMLRPCServerName.ReadOnly = False
		Me.txtXMLRPCServerName.Maxlength = 0
		Me.txtXMLRPCServerName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtXMLRPCServerName.MultiLine = False
		Me.txtXMLRPCServerName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtXMLRPCServerName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtXMLRPCServerName.TabStop = True
		Me.txtXMLRPCServerName.Visible = True
		Me.txtXMLRPCServerName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtXMLRPCServerName.Name = "txtXMLRPCServerName"
		Me.lblXMLRPCServerParameters.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblXMLRPCServerParameters.Text = "Parameters"
		Me.lblXMLRPCServerParameters.Size = New System.Drawing.Size(57, 17)
		Me.lblXMLRPCServerParameters.Location = New System.Drawing.Point(8, 40)
		Me.lblXMLRPCServerParameters.TabIndex = 3
		Me.lblXMLRPCServerParameters.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblXMLRPCServerParameters.BackColor = System.Drawing.SystemColors.Control
		Me.lblXMLRPCServerParameters.Enabled = True
		Me.lblXMLRPCServerParameters.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblXMLRPCServerParameters.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblXMLRPCServerParameters.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblXMLRPCServerParameters.UseMnemonic = True
		Me.lblXMLRPCServerParameters.Visible = True
		Me.lblXMLRPCServerParameters.AutoSize = False
		Me.lblXMLRPCServerParameters.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblXMLRPCServerParameters.Name = "lblXMLRPCServerParameters"
		Me.lblXMLRPCServerName.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblXMLRPCServerName.Text = "Server"
		Me.lblXMLRPCServerName.Size = New System.Drawing.Size(41, 17)
		Me.lblXMLRPCServerName.Location = New System.Drawing.Point(24, 16)
		Me.lblXMLRPCServerName.TabIndex = 1
		Me.lblXMLRPCServerName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblXMLRPCServerName.BackColor = System.Drawing.SystemColors.Control
		Me.lblXMLRPCServerName.Enabled = True
		Me.lblXMLRPCServerName.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblXMLRPCServerName.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblXMLRPCServerName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblXMLRPCServerName.UseMnemonic = True
		Me.lblXMLRPCServerName.Visible = True
		Me.lblXMLRPCServerName.AutoSize = False
		Me.lblXMLRPCServerName.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblXMLRPCServerName.Name = "lblXMLRPCServerName"
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(fraProxySettings)
		Me.Controls.Add(fraXMLRPC)
		Me.fraProxySettings.Controls.Add(chkProxySettings)
		Me.fraProxySettings.Controls.Add(txtProxyPassword)
		Me.fraProxySettings.Controls.Add(txtProxyServerName)
		Me.fraProxySettings.Controls.Add(txtProxyUser)
		Me.fraProxySettings.Controls.Add(lblProxyPassword)
		Me.fraProxySettings.Controls.Add(lblProxyServerName)
		Me.fraProxySettings.Controls.Add(lblProxyUser)
		Me.fraXMLRPC.Controls.Add(txtXMLRPCServerParameters)
		Me.fraXMLRPC.Controls.Add(txtXMLRPCServerName)
		Me.fraXMLRPC.Controls.Add(lblXMLRPCServerParameters)
		Me.fraXMLRPC.Controls.Add(lblXMLRPCServerName)
		Me.fraProxySettings.ResumeLayout(False)
		Me.fraXMLRPC.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class