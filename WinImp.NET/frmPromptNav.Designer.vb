<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPromptNav
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
	Public WithEvents chkView As System.Windows.Forms.CheckBox
	Public WithEvents chkRefreshBMapOnStop As System.Windows.Forms.CheckBox
	Public WithEvents chkDisplayPath As System.Windows.Forms.CheckBox
	Public WithEvents chkLookAlongTheWay As System.Windows.Forms.CheckBox
	Public WithEvents chkStopAfterMove As System.Windows.Forms.CheckBox
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents txtRadius As System.Windows.Forms.TextBox
	Public WithEvents chkMission As System.Windows.Forms.CheckBox
	Public WithEvents Combo1 As System.Windows.Forms.ComboBox
	Public WithEvents txtRetreatPath As System.Windows.Forms.TextBox
	Public WithEvents ChkRetreat As System.Windows.Forms.CheckBox
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents cmdHelp As System.Windows.Forms.Button
	Public WithEvents txtPath As System.Windows.Forms.TextBox
	Public WithEvents txtUnit As System.Windows.Forms.TextBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents lblRouteInfo As System.Windows.Forms.Label
	Public WithEvents lblBestPath As System.Windows.Forms.Label
	Public WithEvents lblPathCost As System.Windows.Forms.Label
	Public WithEvents lblCost As System.Windows.Forms.Label
	Public WithEvents lblLeft As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPromptNav))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Frame2 = New System.Windows.Forms.GroupBox
		Me.chkView = New System.Windows.Forms.CheckBox
		Me.chkRefreshBMapOnStop = New System.Windows.Forms.CheckBox
		Me.chkDisplayPath = New System.Windows.Forms.CheckBox
		Me.chkLookAlongTheWay = New System.Windows.Forms.CheckBox
		Me.chkStopAfterMove = New System.Windows.Forms.CheckBox
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.txtRadius = New System.Windows.Forms.TextBox
		Me.chkMission = New System.Windows.Forms.CheckBox
		Me.Combo1 = New System.Windows.Forms.ComboBox
		Me.txtRetreatPath = New System.Windows.Forms.TextBox
		Me.ChkRetreat = New System.Windows.Forms.CheckBox
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.cmdHelp = New System.Windows.Forms.Button
		Me.txtPath = New System.Windows.Forms.TextBox
		Me.txtUnit = New System.Windows.Forms.TextBox
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me.lblRouteInfo = New System.Windows.Forms.Label
		Me.lblBestPath = New System.Windows.Forms.Label
		Me.lblPathCost = New System.Windows.Forms.Label
		Me.lblCost = New System.Windows.Forms.Label
		Me.lblLeft = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Frame2.SuspendLayout()
		Me.Frame1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.ControlBox = False
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.ClientSize = New System.Drawing.Size(583, 159)
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
		Me.Name = "frmPromptNav"
		Me.Frame2.Text = "Options"
		Me.Frame2.Size = New System.Drawing.Size(297, 65)
		Me.Frame2.Location = New System.Drawing.Point(256, 88)
		Me.Frame2.TabIndex = 16
		Me.Frame2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame2.BackColor = System.Drawing.SystemColors.Control
		Me.Frame2.Enabled = True
		Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame2.Visible = True
		Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame2.Name = "Frame2"
		Me.chkView.Text = "View along the way"
		Me.chkView.Size = New System.Drawing.Size(121, 17)
		Me.chkView.Location = New System.Drawing.Point(136, 27)
		Me.chkView.TabIndex = 25
		Me.chkView.Visible = False
		Me.chkView.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkView.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkView.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkView.BackColor = System.Drawing.SystemColors.Control
		Me.chkView.CausesValidation = True
		Me.chkView.Enabled = True
		Me.chkView.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkView.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkView.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkView.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkView.TabStop = True
		Me.chkView.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkView.Name = "chkView"
		Me.chkRefreshBMapOnStop.Text = "Refresh Bmap on stop"
		Me.chkRefreshBMapOnStop.Size = New System.Drawing.Size(145, 17)
		Me.chkRefreshBMapOnStop.Location = New System.Drawing.Point(136, 45)
		Me.chkRefreshBMapOnStop.TabIndex = 23
		Me.chkRefreshBMapOnStop.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkRefreshBMapOnStop.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkRefreshBMapOnStop.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkRefreshBMapOnStop.BackColor = System.Drawing.SystemColors.Control
		Me.chkRefreshBMapOnStop.CausesValidation = True
		Me.chkRefreshBMapOnStop.Enabled = True
		Me.chkRefreshBMapOnStop.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkRefreshBMapOnStop.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkRefreshBMapOnStop.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkRefreshBMapOnStop.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkRefreshBMapOnStop.TabStop = True
		Me.chkRefreshBMapOnStop.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkRefreshBMapOnStop.Visible = True
		Me.chkRefreshBMapOnStop.Name = "chkRefreshBMapOnStop"
		Me.chkDisplayPath.Text = "Display Path"
		Me.chkDisplayPath.Size = New System.Drawing.Size(89, 17)
		Me.chkDisplayPath.Location = New System.Drawing.Point(8, 45)
		Me.chkDisplayPath.TabIndex = 19
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
		Me.chkLookAlongTheWay.Text = "Look along the way"
		Me.chkLookAlongTheWay.Size = New System.Drawing.Size(121, 17)
		Me.chkLookAlongTheWay.Location = New System.Drawing.Point(136, 9)
		Me.chkLookAlongTheWay.TabIndex = 18
		Me.chkLookAlongTheWay.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkLookAlongTheWay.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkLookAlongTheWay.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkLookAlongTheWay.BackColor = System.Drawing.SystemColors.Control
		Me.chkLookAlongTheWay.CausesValidation = True
		Me.chkLookAlongTheWay.Enabled = True
		Me.chkLookAlongTheWay.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkLookAlongTheWay.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkLookAlongTheWay.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkLookAlongTheWay.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkLookAlongTheWay.TabStop = True
		Me.chkLookAlongTheWay.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkLookAlongTheWay.Visible = True
		Me.chkLookAlongTheWay.Name = "chkLookAlongTheWay"
		Me.chkStopAfterMove.Text = "Stop after move"
		Me.chkStopAfterMove.Size = New System.Drawing.Size(97, 17)
		Me.chkStopAfterMove.Location = New System.Drawing.Point(8, 27)
		Me.chkStopAfterMove.TabIndex = 17
		Me.chkStopAfterMove.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkStopAfterMove.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkStopAfterMove.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkStopAfterMove.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkStopAfterMove.BackColor = System.Drawing.SystemColors.Control
		Me.chkStopAfterMove.CausesValidation = True
		Me.chkStopAfterMove.Enabled = True
		Me.chkStopAfterMove.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkStopAfterMove.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkStopAfterMove.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkStopAfterMove.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkStopAfterMove.TabStop = True
		Me.chkStopAfterMove.Visible = True
		Me.chkStopAfterMove.Name = "chkStopAfterMove"
		Me.Frame1.Text = "Post Move Settings"
		Me.Frame1.Size = New System.Drawing.Size(297, 73)
		Me.Frame1.Location = New System.Drawing.Point(256, 8)
		Me.Frame1.TabIndex = 11
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.txtRadius.AutoSize = False
		Me.txtRadius.Size = New System.Drawing.Size(33, 19)
		Me.txtRadius.Location = New System.Drawing.Point(216, 40)
		Me.txtRadius.TabIndex = 21
		Me.txtRadius.Text = "99"
		Me.txtRadius.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtRadius.AcceptsReturn = True
		Me.txtRadius.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtRadius.BackColor = System.Drawing.SystemColors.Window
		Me.txtRadius.CausesValidation = True
		Me.txtRadius.Enabled = True
		Me.txtRadius.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtRadius.HideSelection = True
		Me.txtRadius.ReadOnly = False
		Me.txtRadius.Maxlength = 0
		Me.txtRadius.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtRadius.MultiLine = False
		Me.txtRadius.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtRadius.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtRadius.TabStop = True
		Me.txtRadius.Visible = True
		Me.txtRadius.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtRadius.Name = "txtRadius"
		Me.chkMission.Text = "Interdiction/Reserve Mission"
		Me.chkMission.Size = New System.Drawing.Size(161, 25)
		Me.chkMission.Location = New System.Drawing.Point(8, 40)
		Me.chkMission.TabIndex = 20
		Me.chkMission.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkMission.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkMission.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkMission.BackColor = System.Drawing.SystemColors.Control
		Me.chkMission.CausesValidation = True
		Me.chkMission.Enabled = True
		Me.chkMission.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkMission.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkMission.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkMission.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkMission.TabStop = True
		Me.chkMission.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkMission.Visible = True
		Me.chkMission.Name = "chkMission"
		Me.Combo1.Size = New System.Drawing.Size(89, 21)
		Me.Combo1.Location = New System.Drawing.Point(80, 16)
		Me.Combo1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.Combo1.TabIndex = 15
		Me.Combo1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Combo1.BackColor = System.Drawing.SystemColors.Window
		Me.Combo1.CausesValidation = True
		Me.Combo1.Enabled = True
		Me.Combo1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Combo1.IntegralHeight = True
		Me.Combo1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Combo1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Combo1.Sorted = False
		Me.Combo1.TabStop = True
		Me.Combo1.Visible = True
		Me.Combo1.Name = "Combo1"
		Me.txtRetreatPath.AutoSize = False
		Me.txtRetreatPath.Size = New System.Drawing.Size(73, 19)
		Me.txtRetreatPath.Location = New System.Drawing.Point(216, 16)
		Me.txtRetreatPath.TabIndex = 14
		Me.txtRetreatPath.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtRetreatPath.AcceptsReturn = True
		Me.txtRetreatPath.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtRetreatPath.BackColor = System.Drawing.SystemColors.Window
		Me.txtRetreatPath.CausesValidation = True
		Me.txtRetreatPath.Enabled = True
		Me.txtRetreatPath.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtRetreatPath.HideSelection = True
		Me.txtRetreatPath.ReadOnly = False
		Me.txtRetreatPath.Maxlength = 0
		Me.txtRetreatPath.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtRetreatPath.MultiLine = False
		Me.txtRetreatPath.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtRetreatPath.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtRetreatPath.TabStop = True
		Me.txtRetreatPath.Visible = True
		Me.txtRetreatPath.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtRetreatPath.Name = "txtRetreatPath"
		Me.ChkRetreat.Text = "Retreat on"
		Me.ChkRetreat.Size = New System.Drawing.Size(73, 25)
		Me.ChkRetreat.Location = New System.Drawing.Point(8, 16)
		Me.ChkRetreat.TabIndex = 12
		Me.ChkRetreat.CheckState = System.Windows.Forms.CheckState.Checked
		Me.ChkRetreat.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ChkRetreat.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.ChkRetreat.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.ChkRetreat.BackColor = System.Drawing.SystemColors.Control
		Me.ChkRetreat.CausesValidation = True
		Me.ChkRetreat.Enabled = True
		Me.ChkRetreat.ForeColor = System.Drawing.SystemColors.ControlText
		Me.ChkRetreat.Cursor = System.Windows.Forms.Cursors.Default
		Me.ChkRetreat.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ChkRetreat.Appearance = System.Windows.Forms.Appearance.Normal
		Me.ChkRetreat.TabStop = True
		Me.ChkRetreat.Visible = True
		Me.ChkRetreat.Name = "ChkRetreat"
		Me.Label4.Text = "radius"
		Me.Label4.Size = New System.Drawing.Size(33, 17)
		Me.Label4.Location = New System.Drawing.Point(176, 40)
		Me.Label4.TabIndex = 22
		Me.Label4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopLeft
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
		Me.Label3.Text = "path"
		Me.Label3.Size = New System.Drawing.Size(33, 17)
		Me.Label3.Location = New System.Drawing.Point(176, 16)
		Me.Label3.TabIndex = 13
		Me.Label3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopLeft
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
		Me.cmdHelp.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdHelp.Text = "?"
		Me.cmdHelp.Font = New System.Drawing.Font("Arial", 10.8!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdHelp.Size = New System.Drawing.Size(25, 25)
		Me.cmdHelp.Location = New System.Drawing.Point(560, 0)
		Me.cmdHelp.TabIndex = 10
		Me.ToolTip1.SetToolTip(Me.cmdHelp, "Click for Help")
		Me.cmdHelp.BackColor = System.Drawing.SystemColors.Control
		Me.cmdHelp.CausesValidation = True
		Me.cmdHelp.Enabled = True
		Me.cmdHelp.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdHelp.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdHelp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdHelp.TabStop = True
		Me.cmdHelp.Name = "cmdHelp"
		Me.txtPath.AutoSize = False
		Me.txtPath.Size = New System.Drawing.Size(153, 19)
		Me.txtPath.Location = New System.Drawing.Point(96, 48)
		Me.txtPath.TabIndex = 1
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
		Me.txtUnit.AutoSize = False
		Me.txtUnit.Size = New System.Drawing.Size(153, 19)
		Me.txtUnit.Location = New System.Drawing.Point(96, 16)
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
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(81, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(8, 120)
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
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "OK"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Size = New System.Drawing.Size(81, 25)
		Me.cmdOK.Location = New System.Drawing.Point(8, 80)
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
		Me.lblRouteInfo.Text = "RouteInfo"
		Me.lblRouteInfo.Size = New System.Drawing.Size(153, 17)
		Me.lblRouteInfo.Location = New System.Drawing.Point(96, 136)
		Me.lblRouteInfo.TabIndex = 24
		Me.lblRouteInfo.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblRouteInfo.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblRouteInfo.BackColor = System.Drawing.SystemColors.Control
		Me.lblRouteInfo.Enabled = True
		Me.lblRouteInfo.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblRouteInfo.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblRouteInfo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblRouteInfo.UseMnemonic = True
		Me.lblRouteInfo.Visible = True
		Me.lblRouteInfo.AutoSize = False
		Me.lblRouteInfo.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblRouteInfo.Name = "lblRouteInfo"
		Me.lblBestPath.Text = "best path:"
		Me.lblBestPath.Size = New System.Drawing.Size(145, 17)
		Me.lblBestPath.Location = New System.Drawing.Point(96, 72)
		Me.lblBestPath.TabIndex = 9
		Me.lblBestPath.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblBestPath.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblBestPath.BackColor = System.Drawing.SystemColors.Control
		Me.lblBestPath.Enabled = True
		Me.lblBestPath.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblBestPath.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblBestPath.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblBestPath.UseMnemonic = True
		Me.lblBestPath.Visible = True
		Me.lblBestPath.AutoSize = False
		Me.lblBestPath.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblBestPath.Name = "lblBestPath"
		Me.lblPathCost.Text = "path cost:"
		Me.lblPathCost.Size = New System.Drawing.Size(153, 17)
		Me.lblPathCost.Location = New System.Drawing.Point(96, 88)
		Me.lblPathCost.TabIndex = 8
		Me.lblPathCost.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblPathCost.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblPathCost.BackColor = System.Drawing.SystemColors.Control
		Me.lblPathCost.Enabled = True
		Me.lblPathCost.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblPathCost.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblPathCost.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblPathCost.UseMnemonic = True
		Me.lblPathCost.Visible = True
		Me.lblPathCost.AutoSize = False
		Me.lblPathCost.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblPathCost.Name = "lblPathCost"
		Me.lblCost.Text = "mob cost:"
		Me.lblCost.Size = New System.Drawing.Size(153, 17)
		Me.lblCost.Location = New System.Drawing.Point(96, 104)
		Me.lblCost.TabIndex = 7
		Me.lblCost.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblCost.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblCost.BackColor = System.Drawing.SystemColors.Control
		Me.lblCost.Enabled = True
		Me.lblCost.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblCost.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblCost.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblCost.UseMnemonic = True
		Me.lblCost.Visible = True
		Me.lblCost.AutoSize = False
		Me.lblCost.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblCost.Name = "lblCost"
		Me.lblLeft.Text = "mob left:"
		Me.lblLeft.Size = New System.Drawing.Size(153, 17)
		Me.lblLeft.Location = New System.Drawing.Point(96, 120)
		Me.lblLeft.TabIndex = 6
		Me.lblLeft.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblLeft.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblLeft.BackColor = System.Drawing.SystemColors.Control
		Me.lblLeft.Enabled = True
		Me.lblLeft.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblLeft.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblLeft.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblLeft.UseMnemonic = True
		Me.lblLeft.Visible = True
		Me.lblLeft.AutoSize = False
		Me.lblLeft.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblLeft.Name = "lblLeft"
		Me.Label1.Text = "path/destination"
		Me.Label1.Size = New System.Drawing.Size(89, 17)
		Me.Label1.Location = New System.Drawing.Point(8, 48)
		Me.Label1.TabIndex = 5
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
		Me.Label2.Text = "Navigate"
		Me.Label2.Font = New System.Drawing.Font("Arial Black", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(73, 17)
		Me.Label2.Location = New System.Drawing.Point(16, 16)
		Me.Label2.TabIndex = 4
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
		Me.Controls.Add(Frame2)
		Me.Controls.Add(Frame1)
		Me.Controls.Add(cmdHelp)
		Me.Controls.Add(txtPath)
		Me.Controls.Add(txtUnit)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(lblRouteInfo)
		Me.Controls.Add(lblBestPath)
		Me.Controls.Add(lblPathCost)
		Me.Controls.Add(lblCost)
		Me.Controls.Add(lblLeft)
		Me.Controls.Add(Label1)
		Me.Controls.Add(Label2)
		Me.Frame2.Controls.Add(chkView)
		Me.Frame2.Controls.Add(chkRefreshBMapOnStop)
		Me.Frame2.Controls.Add(chkDisplayPath)
		Me.Frame2.Controls.Add(chkLookAlongTheWay)
		Me.Frame2.Controls.Add(chkStopAfterMove)
		Me.Frame1.Controls.Add(txtRadius)
		Me.Frame1.Controls.Add(chkMission)
		Me.Frame1.Controls.Add(Combo1)
		Me.Frame1.Controls.Add(txtRetreatPath)
		Me.Frame1.Controls.Add(ChkRetreat)
		Me.Frame1.Controls.Add(Label4)
		Me.Frame1.Controls.Add(Label3)
		Me.Frame2.ResumeLayout(False)
		Me.Frame1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class