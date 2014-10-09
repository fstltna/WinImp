<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPromptMove
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
	Public WithEvents cmdWayPoints As System.Windows.Forms.Button
	Public WithEvents cmdHelp As System.Windows.Forms.Button
	Public WithEvents chkShowAll As System.Windows.Forms.CheckBox
	Public WithEvents chkDisplayPath As System.Windows.Forms.CheckBox
	Public WithEvents cbCom As System.Windows.Forms.ListBox
	Public WithEvents txtNum As System.Windows.Forms.TextBox
	Public WithEvents cmdTest As System.Windows.Forms.Button
	Public WithEvents txtMultOrigin As System.Windows.Forms.TextBox
	Public WithEvents txtMultPath As System.Windows.Forms.TextBox
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents lblPathCost As System.Windows.Forms.Label
	Public WithEvents lblBestPath As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents lblLeft As System.Windows.Forms.Label
	Public WithEvents lblCost As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents lblDest As System.Windows.Forms.Label
	Public WithEvents lblOrigin As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPromptMove))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdWayPoints = New System.Windows.Forms.Button
		Me.cmdHelp = New System.Windows.Forms.Button
		Me.chkShowAll = New System.Windows.Forms.CheckBox
		Me.chkDisplayPath = New System.Windows.Forms.CheckBox
		Me.cbCom = New System.Windows.Forms.ListBox
		Me.txtNum = New System.Windows.Forms.TextBox
		Me.cmdTest = New System.Windows.Forms.Button
		Me.txtMultOrigin = New System.Windows.Forms.TextBox
		Me.txtMultPath = New System.Windows.Forms.TextBox
		Me.cmdOK = New System.Windows.Forms.Button
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.lblPathCost = New System.Windows.Forms.Label
		Me.lblBestPath = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.lblLeft = New System.Windows.Forms.Label
		Me.lblCost = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.lblDest = New System.Windows.Forms.Label
		Me.lblOrigin = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.ControlBox = False
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.ClientSize = New System.Drawing.Size(530, 162)
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
		Me.Name = "frmPromptMove"
		Me.cmdWayPoints.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdWayPoints.Text = "Waypoints"
		Me.cmdWayPoints.Size = New System.Drawing.Size(57, 21)
		Me.cmdWayPoints.Location = New System.Drawing.Point(0, 56)
		Me.cmdWayPoints.TabIndex = 20
		Me.cmdWayPoints.BackColor = System.Drawing.SystemColors.Control
		Me.cmdWayPoints.CausesValidation = True
		Me.cmdWayPoints.Enabled = True
		Me.cmdWayPoints.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdWayPoints.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdWayPoints.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdWayPoints.TabStop = True
		Me.cmdWayPoints.Name = "cmdWayPoints"
		Me.cmdHelp.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdHelp.Text = "?"
		Me.cmdHelp.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdHelp.Size = New System.Drawing.Size(25, 25)
		Me.cmdHelp.Location = New System.Drawing.Point(504, 0)
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
		Me.chkShowAll.Text = "Show All"
		Me.chkShowAll.Size = New System.Drawing.Size(65, 17)
		Me.chkShowAll.Location = New System.Drawing.Point(408, 136)
		Me.chkShowAll.TabIndex = 18
		Me.chkShowAll.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkShowAll.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkShowAll.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkShowAll.BackColor = System.Drawing.SystemColors.Control
		Me.chkShowAll.CausesValidation = True
		Me.chkShowAll.Enabled = True
		Me.chkShowAll.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkShowAll.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkShowAll.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkShowAll.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkShowAll.TabStop = True
		Me.chkShowAll.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkShowAll.Visible = True
		Me.chkShowAll.Name = "chkShowAll"
		Me.chkDisplayPath.Text = "Display Path"
		Me.chkDisplayPath.Size = New System.Drawing.Size(81, 17)
		Me.chkDisplayPath.Location = New System.Drawing.Point(312, 136)
		Me.chkDisplayPath.TabIndex = 17
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
		Me.cbCom.Font = New System.Drawing.Font("Arial", 9.6!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbCom.Size = New System.Drawing.Size(192, 119)
		Me.cbCom.IntegralHeight = False
		Me.cbCom.Location = New System.Drawing.Point(312, 8)
		Me.cbCom.TabIndex = 11
		Me.cbCom.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.cbCom.BackColor = System.Drawing.SystemColors.Window
		Me.cbCom.CausesValidation = True
		Me.cbCom.Enabled = True
		Me.cbCom.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cbCom.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbCom.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.cbCom.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbCom.Sorted = False
		Me.cbCom.TabStop = True
		Me.cbCom.Visible = True
		Me.cbCom.MultiColumn = True
		Me.cbCom.ColumnWidth = 96
		Me.cbCom.Name = "cbCom"
		Me.txtNum.AutoSize = False
		Me.txtNum.Size = New System.Drawing.Size(49, 19)
		Me.txtNum.Location = New System.Drawing.Point(88, 8)
		Me.txtNum.TabIndex = 1
		Me.txtNum.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtNum.AcceptsReturn = True
		Me.txtNum.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtNum.BackColor = System.Drawing.SystemColors.Window
		Me.txtNum.CausesValidation = True
		Me.txtNum.Enabled = True
		Me.txtNum.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtNum.HideSelection = True
		Me.txtNum.ReadOnly = False
		Me.txtNum.Maxlength = 0
		Me.txtNum.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtNum.MultiLine = False
		Me.txtNum.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtNum.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtNum.TabStop = True
		Me.txtNum.Visible = True
		Me.txtNum.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtNum.Name = "txtNum"
		Me.cmdTest.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdTest.Text = "Test"
		Me.cmdTest.Size = New System.Drawing.Size(81, 25)
		Me.cmdTest.Location = New System.Drawing.Point(112, 128)
		Me.cmdTest.TabIndex = 4
		Me.cmdTest.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdTest.BackColor = System.Drawing.SystemColors.Control
		Me.cmdTest.CausesValidation = True
		Me.cmdTest.Enabled = True
		Me.cmdTest.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdTest.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdTest.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdTest.TabStop = True
		Me.cmdTest.Name = "cmdTest"
		Me.txtMultOrigin.AutoSize = False
		Me.txtMultOrigin.Size = New System.Drawing.Size(49, 19)
		Me.txtMultOrigin.Location = New System.Drawing.Point(88, 32)
		Me.txtMultOrigin.TabIndex = 2
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
		Me.txtMultPath.AutoSize = False
		Me.txtMultPath.Size = New System.Drawing.Size(49, 19)
		Me.txtMultPath.Location = New System.Drawing.Point(88, 56)
		Me.txtMultPath.TabIndex = 0
		Me.txtMultPath.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtMultPath.AcceptsReturn = True
		Me.txtMultPath.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtMultPath.BackColor = System.Drawing.SystemColors.Window
		Me.txtMultPath.CausesValidation = True
		Me.txtMultPath.Enabled = True
		Me.txtMultPath.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtMultPath.HideSelection = True
		Me.txtMultPath.ReadOnly = False
		Me.txtMultPath.Maxlength = 0
		Me.txtMultPath.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtMultPath.MultiLine = False
		Me.txtMultPath.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtMultPath.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtMultPath.TabStop = True
		Me.txtMultPath.Visible = True
		Me.txtMultPath.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtMultPath.Name = "txtMultPath"
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "Move"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Size = New System.Drawing.Size(81, 25)
		Me.cmdOK.Location = New System.Drawing.Point(16, 128)
		Me.cmdOK.TabIndex = 3
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
		Me.cmdCancel.Location = New System.Drawing.Point(208, 128)
		Me.cmdCancel.TabIndex = 5
		Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.lblPathCost.Text = "path cost:"
		Me.lblPathCost.Size = New System.Drawing.Size(89, 17)
		Me.lblPathCost.Location = New System.Drawing.Point(8, 80)
		Me.lblPathCost.TabIndex = 14
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
		Me.lblBestPath.Text = "best path:"
		Me.lblBestPath.Size = New System.Drawing.Size(297, 17)
		Me.lblBestPath.Location = New System.Drawing.Point(8, 104)
		Me.lblBestPath.TabIndex = 13
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
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label1.Text = "to"
		Me.Label1.Font = New System.Drawing.Font("Arial", 9.6!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Size = New System.Drawing.Size(65, 33)
		Me.Label1.Location = New System.Drawing.Point(8, 56)
		Me.Label1.TabIndex = 7
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
		Me.lblLeft.Text = "mob left:"
		Me.lblLeft.Size = New System.Drawing.Size(81, 17)
		Me.lblLeft.Location = New System.Drawing.Point(216, 80)
		Me.lblLeft.TabIndex = 16
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
		Me.lblCost.Text = "mob cost:"
		Me.lblCost.Size = New System.Drawing.Size(97, 17)
		Me.lblCost.Location = New System.Drawing.Point(112, 80)
		Me.lblCost.TabIndex = 15
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
		Me.Label4.Font = New System.Drawing.Font("Arial", 9.6!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.Size = New System.Drawing.Size(161, 17)
		Me.Label4.Location = New System.Drawing.Point(152, 8)
		Me.Label4.TabIndex = 12
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
		Me.lblDest.Font = New System.Drawing.Font("Arial", 9.6!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblDest.Size = New System.Drawing.Size(145, 17)
		Me.lblDest.Location = New System.Drawing.Point(152, 56)
		Me.lblDest.TabIndex = 10
		Me.lblDest.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblDest.BackColor = System.Drawing.SystemColors.Control
		Me.lblDest.Enabled = True
		Me.lblDest.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblDest.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblDest.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblDest.UseMnemonic = True
		Me.lblDest.Visible = True
		Me.lblDest.AutoSize = False
		Me.lblDest.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblDest.Name = "lblDest"
		Me.lblOrigin.Font = New System.Drawing.Font("Arial", 9.6!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblOrigin.Size = New System.Drawing.Size(145, 17)
		Me.lblOrigin.Location = New System.Drawing.Point(152, 32)
		Me.lblOrigin.TabIndex = 9
		Me.lblOrigin.TextAlign = System.Drawing.ContentAlignment.TopLeft
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
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label2.Text = "Move"
		Me.Label2.Font = New System.Drawing.Font("Arial Black", 9.6!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(65, 17)
		Me.Label2.Location = New System.Drawing.Point(8, 8)
		Me.Label2.TabIndex = 8
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
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label3.Text = "from"
		Me.Label3.Font = New System.Drawing.Font("Arial", 9.6!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.Size = New System.Drawing.Size(33, 17)
		Me.Label3.Location = New System.Drawing.Point(40, 32)
		Me.Label3.TabIndex = 6
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
		Me.Controls.Add(cmdWayPoints)
		Me.Controls.Add(cmdHelp)
		Me.Controls.Add(chkShowAll)
		Me.Controls.Add(chkDisplayPath)
		Me.Controls.Add(cbCom)
		Me.Controls.Add(txtNum)
		Me.Controls.Add(cmdTest)
		Me.Controls.Add(txtMultOrigin)
		Me.Controls.Add(txtMultPath)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(lblPathCost)
		Me.Controls.Add(lblBestPath)
		Me.Controls.Add(Label1)
		Me.Controls.Add(lblLeft)
		Me.Controls.Add(lblCost)
		Me.Controls.Add(Label4)
		Me.Controls.Add(lblDest)
		Me.Controls.Add(lblOrigin)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label3)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class