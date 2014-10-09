<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmToolRange
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
	Public WithEvents cmbShip As System.Windows.Forms.ComboBox
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label16 As System.Windows.Forms.Label
	Public WithEvents Label15 As System.Windows.Forms.Label
	Public WithEvents lblShipMax As System.Windows.Forms.Label
	Public WithEvents Label9 As System.Windows.Forms.Label
	Public WithEvents lblShipTurn As System.Windows.Forms.Label
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents lblShipMoveCost As System.Windows.Forms.Label
	Public WithEvents lblShipRange As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents frmShip As System.Windows.Forms.Panel
	Public WithEvents chkRoundTrip As System.Windows.Forms.CheckBox
	Public WithEvents optGroundBurst As System.Windows.Forms.RadioButton
	Public WithEvents optAirBurst As System.Windows.Forms.RadioButton
	Public WithEvents txtOrigin As System.Windows.Forms.TextBox
	Public WithEvents cmbNuke As System.Windows.Forms.ComboBox
	Public WithEvents lblNukeType As System.Windows.Forms.Label
	Public WithEvents lblTarget As System.Windows.Forms.Label
	Public WithEvents frmNuke As System.Windows.Forms.Panel
	Public WithEvents cmbLand As System.Windows.Forms.ComboBox
	Public WithEvents Label7 As System.Windows.Forms.Label
	Public WithEvents lblLandRange As System.Windows.Forms.Label
	Public WithEvents frmLand As System.Windows.Forms.Panel
	Public WithEvents Line2 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents Line1 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents Label14 As System.Windows.Forms.Label
	Public WithEvents lblCoastwatch As System.Windows.Forms.Label
	Public WithEvents Label13 As System.Windows.Forms.Label
	Public WithEvents lblRadar As System.Windows.Forms.Label
	Public WithEvents Label11 As System.Windows.Forms.Label
	Public WithEvents Label10 As System.Windows.Forms.Label
	Public WithEvents lblFort100 As System.Windows.Forms.Label
	Public WithEvents Label8 As System.Windows.Forms.Label
	Public WithEvents lblFort59 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents frmSectors As System.Windows.Forms.Panel
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents vsTech As System.Windows.Forms.VScrollBar
	Public WithEvents txtTech As System.Windows.Forms.TextBox
	Public WithEvents TabStrip1 As AxMSComctlLib.AxTabStrip
	Public WithEvents lblTech As System.Windows.Forms.Label
	Public WithEvents ShapeContainer1 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmToolRange))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer
		Me.frmShip = New System.Windows.Forms.Panel
		Me.cmbShip = New System.Windows.Forms.ComboBox
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label16 = New System.Windows.Forms.Label
		Me.Label15 = New System.Windows.Forms.Label
		Me.lblShipMax = New System.Windows.Forms.Label
		Me.Label9 = New System.Windows.Forms.Label
		Me.lblShipTurn = New System.Windows.Forms.Label
		Me.Label5 = New System.Windows.Forms.Label
		Me.lblShipMoveCost = New System.Windows.Forms.Label
		Me.lblShipRange = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.frmNuke = New System.Windows.Forms.Panel
		Me.chkRoundTrip = New System.Windows.Forms.CheckBox
		Me.optGroundBurst = New System.Windows.Forms.RadioButton
		Me.optAirBurst = New System.Windows.Forms.RadioButton
		Me.txtOrigin = New System.Windows.Forms.TextBox
		Me.cmbNuke = New System.Windows.Forms.ComboBox
		Me.lblNukeType = New System.Windows.Forms.Label
		Me.lblTarget = New System.Windows.Forms.Label
		Me.frmLand = New System.Windows.Forms.Panel
		Me.cmbLand = New System.Windows.Forms.ComboBox
		Me.Label7 = New System.Windows.Forms.Label
		Me.lblLandRange = New System.Windows.Forms.Label
		Me.frmSectors = New System.Windows.Forms.Panel
		Me.Line2 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.Line1 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.Label14 = New System.Windows.Forms.Label
		Me.lblCoastwatch = New System.Windows.Forms.Label
		Me.Label13 = New System.Windows.Forms.Label
		Me.lblRadar = New System.Windows.Forms.Label
		Me.Label11 = New System.Windows.Forms.Label
		Me.Label10 = New System.Windows.Forms.Label
		Me.lblFort100 = New System.Windows.Forms.Label
		Me.Label8 = New System.Windows.Forms.Label
		Me.lblFort59 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.cmdOK = New System.Windows.Forms.Button
		Me.vsTech = New System.Windows.Forms.VScrollBar
		Me.txtTech = New System.Windows.Forms.TextBox
		Me.TabStrip1 = New AxMSComctlLib.AxTabStrip
		Me.lblTech = New System.Windows.Forms.Label
		Me.frmShip.SuspendLayout()
		Me.frmNuke.SuspendLayout()
		Me.frmLand.SuspendLayout()
		Me.frmSectors.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.TabStrip1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Range Table"
		Me.ClientSize = New System.Drawing.Size(282, 200)
		Me.Location = New System.Drawing.Point(3, 22)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
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
		Me.Name = "frmToolRange"
		Me.frmShip.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.frmShip.Size = New System.Drawing.Size(233, 97)
		Me.frmShip.Location = New System.Drawing.Point(24, 40)
		Me.frmShip.TabIndex = 3
		Me.frmShip.Visible = False
		Me.frmShip.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmShip.BackColor = System.Drawing.SystemColors.Control
		Me.frmShip.Enabled = True
		Me.frmShip.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmShip.Cursor = System.Windows.Forms.Cursors.Default
		Me.frmShip.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmShip.Name = "frmShip"
		Me.cmbShip.Size = New System.Drawing.Size(121, 21)
		Me.cmbShip.Location = New System.Drawing.Point(8, 8)
		Me.cmbShip.Sorted = True
		Me.cmbShip.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cmbShip.TabIndex = 4
		Me.cmbShip.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbShip.BackColor = System.Drawing.SystemColors.Window
		Me.cmbShip.CausesValidation = True
		Me.cmbShip.Enabled = True
		Me.cmbShip.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbShip.IntegralHeight = True
		Me.cmbShip.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbShip.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbShip.TabStop = True
		Me.cmbShip.Visible = True
		Me.cmbShip.Name = "cmbShip"
		Me.Label3.Text = "Movement"
		Me.Label3.Size = New System.Drawing.Size(65, 17)
		Me.Label3.Location = New System.Drawing.Point(152, 16)
		Me.Label3.TabIndex = 31
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
		Me.Label16.Text = "Range in Sectors"
		Me.Label16.Size = New System.Drawing.Size(89, 17)
		Me.Label16.Location = New System.Drawing.Point(136, 32)
		Me.Label16.TabIndex = 30
		Me.Label16.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label16.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label16.BackColor = System.Drawing.SystemColors.Control
		Me.Label16.Enabled = True
		Me.Label16.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label16.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label16.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label16.UseMnemonic = True
		Me.Label16.Visible = True
		Me.Label16.AutoSize = False
		Me.Label16.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label16.Name = "Label16"
		Me.Label15.Text = "Max Mob "
		Me.Label15.Size = New System.Drawing.Size(49, 17)
		Me.Label15.Location = New System.Drawing.Point(184, 56)
		Me.Label15.TabIndex = 29
		Me.Label15.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label15.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label15.BackColor = System.Drawing.SystemColors.Control
		Me.Label15.Enabled = True
		Me.Label15.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label15.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label15.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label15.UseMnemonic = True
		Me.Label15.Visible = True
		Me.Label15.AutoSize = False
		Me.Label15.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label15.Name = "Label15"
		Me.lblShipMax.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblShipMax.Text = "00.0"
		Me.lblShipMax.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblShipMax.Size = New System.Drawing.Size(33, 17)
		Me.lblShipMax.Location = New System.Drawing.Point(184, 80)
		Me.lblShipMax.TabIndex = 28
		Me.lblShipMax.BackColor = System.Drawing.SystemColors.Control
		Me.lblShipMax.Enabled = True
		Me.lblShipMax.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblShipMax.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblShipMax.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblShipMax.UseMnemonic = True
		Me.lblShipMax.Visible = True
		Me.lblShipMax.AutoSize = False
		Me.lblShipMax.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblShipMax.Name = "lblShipMax"
		Me.Label9.Text = "Per Turn "
		Me.Label9.Size = New System.Drawing.Size(41, 17)
		Me.Label9.Location = New System.Drawing.Point(128, 56)
		Me.Label9.TabIndex = 27
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
		Me.lblShipTurn.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblShipTurn.Text = "00.0"
		Me.lblShipTurn.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblShipTurn.Size = New System.Drawing.Size(33, 17)
		Me.lblShipTurn.Location = New System.Drawing.Point(128, 80)
		Me.lblShipTurn.TabIndex = 26
		Me.lblShipTurn.BackColor = System.Drawing.SystemColors.Control
		Me.lblShipTurn.Enabled = True
		Me.lblShipTurn.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblShipTurn.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblShipTurn.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblShipTurn.UseMnemonic = True
		Me.lblShipTurn.Visible = True
		Me.lblShipTurn.AutoSize = False
		Me.lblShipTurn.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblShipTurn.Name = "lblShipTurn"
		Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label5.Text = "Sector Move Cost"
		Me.Label5.Size = New System.Drawing.Size(57, 33)
		Me.Label5.Location = New System.Drawing.Point(64, 48)
		Me.Label5.TabIndex = 25
		Me.Label5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.lblShipMoveCost.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblShipMoveCost.Text = "00.0"
		Me.lblShipMoveCost.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblShipMoveCost.Size = New System.Drawing.Size(33, 17)
		Me.lblShipMoveCost.Location = New System.Drawing.Point(72, 80)
		Me.lblShipMoveCost.TabIndex = 24
		Me.lblShipMoveCost.BackColor = System.Drawing.SystemColors.Control
		Me.lblShipMoveCost.Enabled = True
		Me.lblShipMoveCost.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblShipMoveCost.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblShipMoveCost.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblShipMoveCost.UseMnemonic = True
		Me.lblShipMoveCost.Visible = True
		Me.lblShipMoveCost.AutoSize = False
		Me.lblShipMoveCost.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblShipMoveCost.Name = "lblShipMoveCost"
		Me.lblShipRange.Text = "0.00"
		Me.lblShipRange.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblShipRange.Size = New System.Drawing.Size(41, 17)
		Me.lblShipRange.Location = New System.Drawing.Point(16, 80)
		Me.lblShipRange.TabIndex = 6
		Me.lblShipRange.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblShipRange.BackColor = System.Drawing.SystemColors.Control
		Me.lblShipRange.Enabled = True
		Me.lblShipRange.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblShipRange.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblShipRange.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblShipRange.UseMnemonic = True
		Me.lblShipRange.Visible = True
		Me.lblShipRange.AutoSize = False
		Me.lblShipRange.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblShipRange.Name = "lblShipRange"
		Me.Label4.Text = "Firing Range"
		Me.Label4.Size = New System.Drawing.Size(49, 33)
		Me.Label4.Location = New System.Drawing.Point(16, 48)
		Me.Label4.TabIndex = 5
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
		Me.frmNuke.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.frmNuke.Size = New System.Drawing.Size(233, 97)
		Me.frmNuke.Location = New System.Drawing.Point(24, 40)
		Me.frmNuke.TabIndex = 32
		Me.frmNuke.Visible = False
		Me.frmNuke.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmNuke.BackColor = System.Drawing.SystemColors.Control
		Me.frmNuke.Enabled = True
		Me.frmNuke.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmNuke.Cursor = System.Windows.Forms.Cursors.Default
		Me.frmNuke.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmNuke.Name = "frmNuke"
		Me.chkRoundTrip.Text = "Round Trip"
		Me.chkRoundTrip.Size = New System.Drawing.Size(81, 25)
		Me.chkRoundTrip.Location = New System.Drawing.Point(56, 64)
		Me.chkRoundTrip.TabIndex = 39
		Me.chkRoundTrip.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkRoundTrip.Visible = False
		Me.chkRoundTrip.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkRoundTrip.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkRoundTrip.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkRoundTrip.BackColor = System.Drawing.SystemColors.Control
		Me.chkRoundTrip.CausesValidation = True
		Me.chkRoundTrip.Enabled = True
		Me.chkRoundTrip.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkRoundTrip.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkRoundTrip.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkRoundTrip.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkRoundTrip.TabStop = True
		Me.chkRoundTrip.Name = "chkRoundTrip"
		Me.optGroundBurst.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optGroundBurst.Text = "Ground Burst"
		Me.optGroundBurst.Size = New System.Drawing.Size(97, 13)
		Me.optGroundBurst.Location = New System.Drawing.Point(120, 72)
		Me.optGroundBurst.TabIndex = 38
		Me.optGroundBurst.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optGroundBurst.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optGroundBurst.BackColor = System.Drawing.SystemColors.Control
		Me.optGroundBurst.CausesValidation = True
		Me.optGroundBurst.Enabled = True
		Me.optGroundBurst.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optGroundBurst.Cursor = System.Windows.Forms.Cursors.Default
		Me.optGroundBurst.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optGroundBurst.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optGroundBurst.TabStop = True
		Me.optGroundBurst.Checked = False
		Me.optGroundBurst.Visible = True
		Me.optGroundBurst.Name = "optGroundBurst"
		Me.optAirBurst.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optAirBurst.Text = "Air Burst"
		Me.optAirBurst.Size = New System.Drawing.Size(81, 13)
		Me.optAirBurst.Location = New System.Drawing.Point(24, 72)
		Me.optAirBurst.TabIndex = 37
		Me.optAirBurst.Checked = True
		Me.optAirBurst.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optAirBurst.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optAirBurst.BackColor = System.Drawing.SystemColors.Control
		Me.optAirBurst.CausesValidation = True
		Me.optAirBurst.Enabled = True
		Me.optAirBurst.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optAirBurst.Cursor = System.Windows.Forms.Cursors.Default
		Me.optAirBurst.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optAirBurst.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optAirBurst.TabStop = True
		Me.optAirBurst.Visible = True
		Me.optAirBurst.Name = "optAirBurst"
		Me.txtOrigin.AutoSize = False
		Me.txtOrigin.Size = New System.Drawing.Size(145, 19)
		Me.txtOrigin.Location = New System.Drawing.Point(72, 40)
		Me.txtOrigin.TabIndex = 36
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
		Me.cmbNuke.Size = New System.Drawing.Size(145, 21)
		Me.cmbNuke.Location = New System.Drawing.Point(72, 8)
		Me.cmbNuke.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cmbNuke.TabIndex = 33
		Me.cmbNuke.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbNuke.BackColor = System.Drawing.SystemColors.Window
		Me.cmbNuke.CausesValidation = True
		Me.cmbNuke.Enabled = True
		Me.cmbNuke.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbNuke.IntegralHeight = True
		Me.cmbNuke.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbNuke.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbNuke.Sorted = False
		Me.cmbNuke.TabStop = True
		Me.cmbNuke.Visible = True
		Me.cmbNuke.Name = "cmbNuke"
		Me.lblNukeType.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblNukeType.Text = "Nuke Type"
		Me.lblNukeType.Size = New System.Drawing.Size(57, 17)
		Me.lblNukeType.Location = New System.Drawing.Point(8, 8)
		Me.lblNukeType.TabIndex = 35
		Me.lblNukeType.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblNukeType.BackColor = System.Drawing.SystemColors.Control
		Me.lblNukeType.Enabled = True
		Me.lblNukeType.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblNukeType.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblNukeType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblNukeType.UseMnemonic = True
		Me.lblNukeType.Visible = True
		Me.lblNukeType.AutoSize = False
		Me.lblNukeType.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblNukeType.Name = "lblNukeType"
		Me.lblTarget.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblTarget.Text = "Target"
		Me.lblTarget.Size = New System.Drawing.Size(41, 17)
		Me.lblTarget.Location = New System.Drawing.Point(24, 40)
		Me.lblTarget.TabIndex = 34
		Me.lblTarget.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblTarget.BackColor = System.Drawing.SystemColors.Control
		Me.lblTarget.Enabled = True
		Me.lblTarget.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblTarget.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblTarget.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblTarget.UseMnemonic = True
		Me.lblTarget.Visible = True
		Me.lblTarget.AutoSize = False
		Me.lblTarget.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblTarget.Name = "lblTarget"
		Me.frmLand.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.frmLand.Size = New System.Drawing.Size(233, 97)
		Me.frmLand.Location = New System.Drawing.Point(24, 40)
		Me.frmLand.TabIndex = 7
		Me.frmLand.Visible = False
		Me.frmLand.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmLand.BackColor = System.Drawing.SystemColors.Control
		Me.frmLand.Enabled = True
		Me.frmLand.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmLand.Cursor = System.Windows.Forms.Cursors.Default
		Me.frmLand.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmLand.Name = "frmLand"
		Me.cmbLand.Size = New System.Drawing.Size(193, 21)
		Me.cmbLand.Location = New System.Drawing.Point(16, 24)
		Me.cmbLand.Sorted = True
		Me.cmbLand.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cmbLand.TabIndex = 8
		Me.cmbLand.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbLand.BackColor = System.Drawing.SystemColors.Window
		Me.cmbLand.CausesValidation = True
		Me.cmbLand.Enabled = True
		Me.cmbLand.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbLand.IntegralHeight = True
		Me.cmbLand.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbLand.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbLand.TabStop = True
		Me.cmbLand.Visible = True
		Me.cmbLand.Name = "cmbLand"
		Me.Label7.Text = "Firing Range"
		Me.Label7.Size = New System.Drawing.Size(73, 17)
		Me.Label7.Location = New System.Drawing.Point(16, 56)
		Me.Label7.TabIndex = 10
		Me.Label7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label7.BackColor = System.Drawing.SystemColors.Control
		Me.Label7.Enabled = True
		Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label7.UseMnemonic = True
		Me.Label7.Visible = True
		Me.Label7.AutoSize = False
		Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label7.Name = "Label7"
		Me.lblLandRange.Text = "0.00"
		Me.lblLandRange.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblLandRange.Size = New System.Drawing.Size(41, 17)
		Me.lblLandRange.Location = New System.Drawing.Point(32, 72)
		Me.lblLandRange.TabIndex = 9
		Me.lblLandRange.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblLandRange.BackColor = System.Drawing.SystemColors.Control
		Me.lblLandRange.Enabled = True
		Me.lblLandRange.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblLandRange.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblLandRange.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblLandRange.UseMnemonic = True
		Me.lblLandRange.Visible = True
		Me.lblLandRange.AutoSize = False
		Me.lblLandRange.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblLandRange.Name = "lblLandRange"
		Me.frmSectors.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.frmSectors.Size = New System.Drawing.Size(233, 97)
		Me.frmSectors.Location = New System.Drawing.Point(24, 40)
		Me.frmSectors.TabIndex = 11
		Me.frmSectors.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmSectors.BackColor = System.Drawing.SystemColors.Control
		Me.frmSectors.Enabled = True
		Me.frmSectors.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmSectors.Cursor = System.Windows.Forms.Cursors.Default
		Me.frmSectors.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmSectors.Visible = True
		Me.frmSectors.Name = "frmSectors"
		Me.Line2.X1 = 112
		Me.Line2.X2 = 208
		Me.Line2.Y1 = 32
		Me.Line2.Y2 = 32
		Me.Line2.BorderColor = System.Drawing.SystemColors.WindowText
		Me.Line2.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line2.BorderWidth = 1
		Me.Line2.Visible = True
		Me.Line2.Name = "Line2"
		Me.Line1.X1 = 8
		Me.Line1.X2 = 88
		Me.Line1.Y1 = 32
		Me.Line1.Y2 = 32
		Me.Line1.BorderColor = System.Drawing.SystemColors.WindowText
		Me.Line1.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line1.BorderWidth = 1
		Me.Line1.Visible = True
		Me.Line1.Name = "Line1"
		Me.Label14.Text = "Coastwatch"
		Me.Label14.Size = New System.Drawing.Size(65, 17)
		Me.Label14.Location = New System.Drawing.Point(112, 56)
		Me.Label14.TabIndex = 21
		Me.Label14.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label14.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label14.BackColor = System.Drawing.SystemColors.Control
		Me.Label14.Enabled = True
		Me.Label14.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label14.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label14.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label14.UseMnemonic = True
		Me.Label14.Visible = True
		Me.Label14.AutoSize = False
		Me.Label14.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label14.Name = "Label14"
		Me.lblCoastwatch.Text = "0"
		Me.lblCoastwatch.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblCoastwatch.Size = New System.Drawing.Size(33, 17)
		Me.lblCoastwatch.Location = New System.Drawing.Point(184, 56)
		Me.lblCoastwatch.TabIndex = 20
		Me.lblCoastwatch.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblCoastwatch.BackColor = System.Drawing.SystemColors.Control
		Me.lblCoastwatch.Enabled = True
		Me.lblCoastwatch.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblCoastwatch.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblCoastwatch.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblCoastwatch.UseMnemonic = True
		Me.lblCoastwatch.Visible = True
		Me.lblCoastwatch.AutoSize = False
		Me.lblCoastwatch.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblCoastwatch.Name = "lblCoastwatch"
		Me.Label13.Text = "Radar Range"
		Me.Label13.Size = New System.Drawing.Size(73, 17)
		Me.Label13.Location = New System.Drawing.Point(112, 40)
		Me.Label13.TabIndex = 19
		Me.Label13.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label13.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label13.BackColor = System.Drawing.SystemColors.Control
		Me.Label13.Enabled = True
		Me.Label13.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label13.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label13.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label13.UseMnemonic = True
		Me.Label13.Visible = True
		Me.Label13.AutoSize = False
		Me.Label13.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label13.Name = "Label13"
		Me.lblRadar.Text = "0"
		Me.lblRadar.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblRadar.Size = New System.Drawing.Size(33, 17)
		Me.lblRadar.Location = New System.Drawing.Point(184, 40)
		Me.lblRadar.TabIndex = 18
		Me.lblRadar.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblRadar.BackColor = System.Drawing.SystemColors.Control
		Me.lblRadar.Enabled = True
		Me.lblRadar.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblRadar.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblRadar.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblRadar.UseMnemonic = True
		Me.lblRadar.Visible = True
		Me.lblRadar.AutoSize = False
		Me.lblRadar.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblRadar.Name = "lblRadar"
		Me.Label11.Text = "100% Radar Station"
		Me.Label11.Size = New System.Drawing.Size(97, 17)
		Me.Label11.Location = New System.Drawing.Point(112, 16)
		Me.Label11.TabIndex = 17
		Me.Label11.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label11.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label11.BackColor = System.Drawing.SystemColors.Control
		Me.Label11.Enabled = True
		Me.Label11.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label11.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label11.UseMnemonic = True
		Me.Label11.Visible = True
		Me.Label11.AutoSize = False
		Me.Label11.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label11.Name = "Label11"
		Me.Label10.Text = "60-100%"
		Me.Label10.Size = New System.Drawing.Size(41, 17)
		Me.Label10.Location = New System.Drawing.Point(8, 56)
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
		Me.lblFort100.Text = "0.00"
		Me.lblFort100.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblFort100.Size = New System.Drawing.Size(33, 17)
		Me.lblFort100.Location = New System.Drawing.Point(56, 56)
		Me.lblFort100.TabIndex = 15
		Me.lblFort100.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblFort100.BackColor = System.Drawing.SystemColors.Control
		Me.lblFort100.Enabled = True
		Me.lblFort100.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblFort100.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblFort100.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblFort100.UseMnemonic = True
		Me.lblFort100.Visible = True
		Me.lblFort100.AutoSize = False
		Me.lblFort100.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblFort100.Name = "lblFort100"
		Me.Label8.Text = "01-59%"
		Me.Label8.Size = New System.Drawing.Size(41, 17)
		Me.Label8.Location = New System.Drawing.Point(8, 40)
		Me.Label8.TabIndex = 14
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
		Me.lblFort59.Text = "0.00"
		Me.lblFort59.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblFort59.Size = New System.Drawing.Size(33, 17)
		Me.lblFort59.Location = New System.Drawing.Point(56, 40)
		Me.lblFort59.TabIndex = 13
		Me.lblFort59.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblFort59.BackColor = System.Drawing.SystemColors.Control
		Me.lblFort59.Enabled = True
		Me.lblFort59.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblFort59.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblFort59.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblFort59.UseMnemonic = True
		Me.lblFort59.Visible = True
		Me.lblFort59.AutoSize = False
		Me.lblFort59.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblFort59.Name = "lblFort59"
		Me.Label2.Text = "Fort Firing Range"
		Me.Label2.Size = New System.Drawing.Size(97, 17)
		Me.Label2.Location = New System.Drawing.Point(8, 16)
		Me.Label2.TabIndex = 12
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
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdOK
		Me.cmdOK.Text = "OK"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Size = New System.Drawing.Size(81, 33)
		Me.cmdOK.Location = New System.Drawing.Point(168, 160)
		Me.cmdOK.TabIndex = 23
		Me.cmdOK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.vsTech.Size = New System.Drawing.Size(17, 33)
		Me.vsTech.Location = New System.Drawing.Point(128, 160)
		Me.vsTech.Maximum = 0
		Me.vsTech.Minimum = 450
		Me.vsTech.TabIndex = 2
		Me.vsTech.CausesValidation = True
		Me.vsTech.Enabled = True
		Me.vsTech.LargeChange = 1
		Me.vsTech.Cursor = System.Windows.Forms.Cursors.Default
		Me.vsTech.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.vsTech.SmallChange = 1
		Me.vsTech.TabStop = True
		Me.vsTech.Value = 0
		Me.vsTech.Visible = True
		Me.vsTech.Name = "vsTech"
		Me.txtTech.AutoSize = False
		Me.txtTech.Size = New System.Drawing.Size(41, 19)
		Me.txtTech.Location = New System.Drawing.Point(88, 168)
		Me.txtTech.TabIndex = 1
		Me.txtTech.Text = "Text1"
		Me.txtTech.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtTech.AcceptsReturn = True
		Me.txtTech.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtTech.BackColor = System.Drawing.SystemColors.Window
		Me.txtTech.CausesValidation = True
		Me.txtTech.Enabled = True
		Me.txtTech.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtTech.HideSelection = True
		Me.txtTech.ReadOnly = False
		Me.txtTech.Maxlength = 0
		Me.txtTech.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtTech.MultiLine = False
		Me.txtTech.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtTech.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtTech.TabStop = True
		Me.txtTech.Visible = True
		Me.txtTech.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtTech.Name = "txtTech"
		TabStrip1.OcxState = CType(resources.GetObject("TabStrip1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.TabStrip1.Size = New System.Drawing.Size(265, 137)
		Me.TabStrip1.Location = New System.Drawing.Point(8, 8)
		Me.TabStrip1.TabIndex = 22
		Me.TabStrip1.Name = "TabStrip1"
		Me.lblTech.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblTech.Text = "Tech (Estimate)"
		Me.lblTech.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblTech.Size = New System.Drawing.Size(57, 33)
		Me.lblTech.Location = New System.Drawing.Point(16, 160)
		Me.lblTech.TabIndex = 0
		Me.lblTech.BackColor = System.Drawing.SystemColors.Control
		Me.lblTech.Enabled = True
		Me.lblTech.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblTech.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblTech.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblTech.UseMnemonic = True
		Me.lblTech.Visible = True
		Me.lblTech.AutoSize = False
		Me.lblTech.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblTech.Name = "lblTech"
		Me.Controls.Add(frmShip)
		Me.Controls.Add(frmNuke)
		Me.Controls.Add(frmLand)
		Me.Controls.Add(frmSectors)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(vsTech)
		Me.Controls.Add(txtTech)
		Me.Controls.Add(TabStrip1)
		Me.Controls.Add(lblTech)
		Me.frmShip.Controls.Add(cmbShip)
		Me.frmShip.Controls.Add(Label3)
		Me.frmShip.Controls.Add(Label16)
		Me.frmShip.Controls.Add(Label15)
		Me.frmShip.Controls.Add(lblShipMax)
		Me.frmShip.Controls.Add(Label9)
		Me.frmShip.Controls.Add(lblShipTurn)
		Me.frmShip.Controls.Add(Label5)
		Me.frmShip.Controls.Add(lblShipMoveCost)
		Me.frmShip.Controls.Add(lblShipRange)
		Me.frmShip.Controls.Add(Label4)
		Me.frmNuke.Controls.Add(chkRoundTrip)
		Me.frmNuke.Controls.Add(optGroundBurst)
		Me.frmNuke.Controls.Add(optAirBurst)
		Me.frmNuke.Controls.Add(txtOrigin)
		Me.frmNuke.Controls.Add(cmbNuke)
		Me.frmNuke.Controls.Add(lblNukeType)
		Me.frmNuke.Controls.Add(lblTarget)
		Me.frmLand.Controls.Add(cmbLand)
		Me.frmLand.Controls.Add(Label7)
		Me.frmLand.Controls.Add(lblLandRange)
		Me.ShapeContainer1.Shapes.Add(Line2)
		Me.ShapeContainer1.Shapes.Add(Line1)
		Me.frmSectors.Controls.Add(Label14)
		Me.frmSectors.Controls.Add(lblCoastwatch)
		Me.frmSectors.Controls.Add(Label13)
		Me.frmSectors.Controls.Add(lblRadar)
		Me.frmSectors.Controls.Add(Label11)
		Me.frmSectors.Controls.Add(Label10)
		Me.frmSectors.Controls.Add(lblFort100)
		Me.frmSectors.Controls.Add(Label8)
		Me.frmSectors.Controls.Add(lblFort59)
		Me.frmSectors.Controls.Add(Label2)
		Me.frmSectors.Controls.Add(ShapeContainer1)
		CType(Me.TabStrip1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.frmShip.ResumeLayout(False)
		Me.frmNuke.ResumeLayout(False)
		Me.frmLand.ResumeLayout(False)
		Me.frmSectors.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class