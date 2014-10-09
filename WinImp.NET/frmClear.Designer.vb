<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmClear
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
	Public WithEvents cbDays As System.Windows.Forms.ComboBox
	Public WithEvents chkCoastWatch As System.Windows.Forms.CheckBox
	Public WithEvents chkSpy As System.Windows.Forms.CheckBox
	Public WithEvents chkLook As System.Windows.Forms.CheckBox
	Public WithEvents chkLlook As System.Windows.Forms.CheckBox
	Public WithEvents fraPostOptions As System.Windows.Forms.GroupBox
	Public WithEvents chkEnemy As System.Windows.Forms.CheckBox
	Public WithEvents chkNeutral As System.Windows.Forms.CheckBox
	Public WithEvents chkAllied As System.Windows.Forms.CheckBox
	Public WithEvents fraOptions As System.Windows.Forms.GroupBox
	Public WithEvents chkTelegrams As System.Windows.Forms.CheckBox
	Public WithEvents chkAirCombat As System.Windows.Forms.CheckBox
	Public WithEvents chkPlanes As System.Windows.Forms.CheckBox
	Public WithEvents chkShips As System.Windows.Forms.CheckBox
	Public WithEvents chkLands As System.Windows.Forms.CheckBox
	Public WithEvents chkSectors As System.Windows.Forms.CheckBox
	Public WithEvents frmEnemy As System.Windows.Forms.GroupBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents lblDays As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmClear))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cbDays = New System.Windows.Forms.ComboBox
		Me.fraPostOptions = New System.Windows.Forms.GroupBox
		Me.chkCoastWatch = New System.Windows.Forms.CheckBox
		Me.chkSpy = New System.Windows.Forms.CheckBox
		Me.chkLook = New System.Windows.Forms.CheckBox
		Me.chkLlook = New System.Windows.Forms.CheckBox
		Me.fraOptions = New System.Windows.Forms.GroupBox
		Me.chkEnemy = New System.Windows.Forms.CheckBox
		Me.chkNeutral = New System.Windows.Forms.CheckBox
		Me.chkAllied = New System.Windows.Forms.CheckBox
		Me.chkTelegrams = New System.Windows.Forms.CheckBox
		Me.frmEnemy = New System.Windows.Forms.GroupBox
		Me.chkAirCombat = New System.Windows.Forms.CheckBox
		Me.chkPlanes = New System.Windows.Forms.CheckBox
		Me.chkShips = New System.Windows.Forms.CheckBox
		Me.chkLands = New System.Windows.Forms.CheckBox
		Me.chkSectors = New System.Windows.Forms.CheckBox
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me.lblDays = New System.Windows.Forms.Label
		Me.fraPostOptions.SuspendLayout()
		Me.fraOptions.SuspendLayout()
		Me.frmEnemy.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Clear"
		Me.ClientSize = New System.Drawing.Size(594, 103)
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
		Me.Name = "frmClear"
		Me.cbDays.Size = New System.Drawing.Size(81, 21)
		Me.cbDays.Location = New System.Drawing.Point(424, 72)
		Me.cbDays.Items.AddRange(New Object(){"All Records", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "30"})
		Me.cbDays.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cbDays.TabIndex = 19
		Me.cbDays.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbDays.BackColor = System.Drawing.SystemColors.Window
		Me.cbDays.CausesValidation = True
		Me.cbDays.Enabled = True
		Me.cbDays.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cbDays.IntegralHeight = True
		Me.cbDays.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbDays.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbDays.Sorted = False
		Me.cbDays.TabStop = True
		Me.cbDays.Visible = True
		Me.cbDays.Name = "cbDays"
		Me.fraPostOptions.Text = "Post Clear Commands"
		Me.fraPostOptions.Size = New System.Drawing.Size(121, 89)
		Me.fraPostOptions.Location = New System.Drawing.Point(288, 8)
		Me.fraPostOptions.TabIndex = 13
		Me.fraPostOptions.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraPostOptions.BackColor = System.Drawing.SystemColors.Control
		Me.fraPostOptions.Enabled = True
		Me.fraPostOptions.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraPostOptions.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraPostOptions.Visible = True
		Me.fraPostOptions.Padding = New System.Windows.Forms.Padding(0)
		Me.fraPostOptions.Name = "fraPostOptions"
		Me.chkCoastWatch.Text = "coastwatch"
		Me.chkCoastWatch.Size = New System.Drawing.Size(81, 17)
		Me.chkCoastWatch.Location = New System.Drawing.Point(8, 64)
		Me.chkCoastWatch.TabIndex = 17
		Me.chkCoastWatch.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkCoastWatch.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkCoastWatch.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkCoastWatch.BackColor = System.Drawing.SystemColors.Control
		Me.chkCoastWatch.CausesValidation = True
		Me.chkCoastWatch.Enabled = True
		Me.chkCoastWatch.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkCoastWatch.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkCoastWatch.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkCoastWatch.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkCoastWatch.TabStop = True
		Me.chkCoastWatch.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkCoastWatch.Visible = True
		Me.chkCoastWatch.Name = "chkCoastWatch"
		Me.chkSpy.Text = "spy *"
		Me.chkSpy.Size = New System.Drawing.Size(49, 17)
		Me.chkSpy.Location = New System.Drawing.Point(64, 16)
		Me.chkSpy.TabIndex = 16
		Me.chkSpy.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkSpy.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkSpy.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkSpy.BackColor = System.Drawing.SystemColors.Control
		Me.chkSpy.CausesValidation = True
		Me.chkSpy.Enabled = True
		Me.chkSpy.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkSpy.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkSpy.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkSpy.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkSpy.TabStop = True
		Me.chkSpy.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkSpy.Visible = True
		Me.chkSpy.Name = "chkSpy"
		Me.chkLook.Text = "look *"
		Me.chkLook.Size = New System.Drawing.Size(49, 17)
		Me.chkLook.Location = New System.Drawing.Point(8, 16)
		Me.chkLook.TabIndex = 15
		Me.chkLook.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkLook.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkLook.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkLook.BackColor = System.Drawing.SystemColors.Control
		Me.chkLook.CausesValidation = True
		Me.chkLook.Enabled = True
		Me.chkLook.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkLook.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkLook.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkLook.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkLook.TabStop = True
		Me.chkLook.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkLook.Visible = True
		Me.chkLook.Name = "chkLook"
		Me.chkLlook.Text = "llook *"
		Me.chkLlook.Size = New System.Drawing.Size(49, 17)
		Me.chkLlook.Location = New System.Drawing.Point(8, 40)
		Me.chkLlook.TabIndex = 14
		Me.chkLlook.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkLlook.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkLlook.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkLlook.BackColor = System.Drawing.SystemColors.Control
		Me.chkLlook.CausesValidation = True
		Me.chkLlook.Enabled = True
		Me.chkLlook.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkLlook.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkLlook.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkLlook.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkLlook.TabStop = True
		Me.chkLlook.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkLlook.Visible = True
		Me.chkLlook.Name = "chkLlook"
		Me.fraOptions.Text = "Unit Options"
		Me.fraOptions.Size = New System.Drawing.Size(81, 89)
		Me.fraOptions.Location = New System.Drawing.Point(192, 8)
		Me.fraOptions.TabIndex = 9
		Me.fraOptions.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraOptions.BackColor = System.Drawing.SystemColors.Control
		Me.fraOptions.Enabled = True
		Me.fraOptions.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraOptions.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraOptions.Visible = True
		Me.fraOptions.Padding = New System.Windows.Forms.Padding(0)
		Me.fraOptions.Name = "fraOptions"
		Me.chkEnemy.Text = "Enemy"
		Me.chkEnemy.Size = New System.Drawing.Size(57, 17)
		Me.chkEnemy.Location = New System.Drawing.Point(8, 64)
		Me.chkEnemy.TabIndex = 12
		Me.chkEnemy.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkEnemy.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkEnemy.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkEnemy.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkEnemy.BackColor = System.Drawing.SystemColors.Control
		Me.chkEnemy.CausesValidation = True
		Me.chkEnemy.Enabled = True
		Me.chkEnemy.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkEnemy.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkEnemy.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkEnemy.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkEnemy.TabStop = True
		Me.chkEnemy.Visible = True
		Me.chkEnemy.Name = "chkEnemy"
		Me.chkNeutral.Text = "Neutral"
		Me.chkNeutral.Size = New System.Drawing.Size(57, 17)
		Me.chkNeutral.Location = New System.Drawing.Point(8, 40)
		Me.chkNeutral.TabIndex = 11
		Me.chkNeutral.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkNeutral.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkNeutral.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkNeutral.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkNeutral.BackColor = System.Drawing.SystemColors.Control
		Me.chkNeutral.CausesValidation = True
		Me.chkNeutral.Enabled = True
		Me.chkNeutral.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkNeutral.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkNeutral.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkNeutral.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkNeutral.TabStop = True
		Me.chkNeutral.Visible = True
		Me.chkNeutral.Name = "chkNeutral"
		Me.chkAllied.Text = "Allied"
		Me.chkAllied.Size = New System.Drawing.Size(57, 17)
		Me.chkAllied.Location = New System.Drawing.Point(8, 16)
		Me.chkAllied.TabIndex = 10
		Me.chkAllied.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkAllied.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkAllied.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkAllied.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkAllied.BackColor = System.Drawing.SystemColors.Control
		Me.chkAllied.CausesValidation = True
		Me.chkAllied.Enabled = True
		Me.chkAllied.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkAllied.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkAllied.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkAllied.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkAllied.TabStop = True
		Me.chkAllied.Visible = True
		Me.chkAllied.Name = "chkAllied"
		Me.chkTelegrams.Text = "Telegrams"
		Me.chkTelegrams.Size = New System.Drawing.Size(73, 17)
		Me.chkTelegrams.Location = New System.Drawing.Point(424, 16)
		Me.chkTelegrams.TabIndex = 8
		Me.chkTelegrams.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkTelegrams.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkTelegrams.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkTelegrams.BackColor = System.Drawing.SystemColors.Control
		Me.chkTelegrams.CausesValidation = True
		Me.chkTelegrams.Enabled = True
		Me.chkTelegrams.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkTelegrams.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkTelegrams.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkTelegrams.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkTelegrams.TabStop = True
		Me.chkTelegrams.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkTelegrams.Visible = True
		Me.chkTelegrams.Name = "chkTelegrams"
		Me.frmEnemy.Text = "Enemy Intelligence"
		Me.frmEnemy.Size = New System.Drawing.Size(169, 89)
		Me.frmEnemy.Location = New System.Drawing.Point(8, 8)
		Me.frmEnemy.TabIndex = 2
		Me.frmEnemy.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmEnemy.BackColor = System.Drawing.SystemColors.Control
		Me.frmEnemy.Enabled = True
		Me.frmEnemy.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmEnemy.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmEnemy.Visible = True
		Me.frmEnemy.Padding = New System.Windows.Forms.Padding(0)
		Me.frmEnemy.Name = "frmEnemy"
		Me.chkAirCombat.Text = "Air Combat"
		Me.chkAirCombat.Size = New System.Drawing.Size(73, 17)
		Me.chkAirCombat.Location = New System.Drawing.Point(8, 56)
		Me.chkAirCombat.TabIndex = 7
		Me.chkAirCombat.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkAirCombat.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkAirCombat.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkAirCombat.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkAirCombat.BackColor = System.Drawing.SystemColors.Control
		Me.chkAirCombat.CausesValidation = True
		Me.chkAirCombat.Enabled = True
		Me.chkAirCombat.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkAirCombat.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkAirCombat.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkAirCombat.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkAirCombat.TabStop = True
		Me.chkAirCombat.Visible = True
		Me.chkAirCombat.Name = "chkAirCombat"
		Me.chkPlanes.Text = "Planes"
		Me.chkPlanes.Size = New System.Drawing.Size(73, 17)
		Me.chkPlanes.Location = New System.Drawing.Point(88, 64)
		Me.chkPlanes.TabIndex = 6
		Me.chkPlanes.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkPlanes.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkPlanes.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkPlanes.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkPlanes.BackColor = System.Drawing.SystemColors.Control
		Me.chkPlanes.CausesValidation = True
		Me.chkPlanes.Enabled = True
		Me.chkPlanes.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkPlanes.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkPlanes.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkPlanes.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkPlanes.TabStop = True
		Me.chkPlanes.Visible = True
		Me.chkPlanes.Name = "chkPlanes"
		Me.chkShips.Text = "Ships"
		Me.chkShips.Size = New System.Drawing.Size(73, 17)
		Me.chkShips.Location = New System.Drawing.Point(88, 16)
		Me.chkShips.TabIndex = 5
		Me.chkShips.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkShips.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkShips.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkShips.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkShips.BackColor = System.Drawing.SystemColors.Control
		Me.chkShips.CausesValidation = True
		Me.chkShips.Enabled = True
		Me.chkShips.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkShips.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkShips.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkShips.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkShips.TabStop = True
		Me.chkShips.Visible = True
		Me.chkShips.Name = "chkShips"
		Me.chkLands.Text = "Land Units"
		Me.chkLands.Size = New System.Drawing.Size(73, 17)
		Me.chkLands.Location = New System.Drawing.Point(88, 40)
		Me.chkLands.TabIndex = 4
		Me.chkLands.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkLands.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkLands.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkLands.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkLands.BackColor = System.Drawing.SystemColors.Control
		Me.chkLands.CausesValidation = True
		Me.chkLands.Enabled = True
		Me.chkLands.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkLands.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkLands.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkLands.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkLands.TabStop = True
		Me.chkLands.Visible = True
		Me.chkLands.Name = "chkLands"
		Me.chkSectors.Text = "Sectors"
		Me.chkSectors.Size = New System.Drawing.Size(73, 17)
		Me.chkSectors.Location = New System.Drawing.Point(8, 24)
		Me.chkSectors.TabIndex = 3
		Me.chkSectors.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkSectors.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkSectors.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkSectors.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkSectors.BackColor = System.Drawing.SystemColors.Control
		Me.chkSectors.CausesValidation = True
		Me.chkSectors.Enabled = True
		Me.chkSectors.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkSectors.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkSectors.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkSectors.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkSectors.TabStop = True
		Me.chkSectors.Visible = True
		Me.chkSectors.Name = "chkSectors"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "Abort"
		Me.AcceptButton = Me.cmdCancel
		Me.cmdCancel.Size = New System.Drawing.Size(73, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(512, 64)
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
		Me.cmdOK.Text = "Clear"
		Me.cmdOK.Size = New System.Drawing.Size(73, 25)
		Me.cmdOK.Location = New System.Drawing.Point(512, 16)
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
		Me.lblDays.Text = "Clear records older then days"
		Me.lblDays.Size = New System.Drawing.Size(81, 25)
		Me.lblDays.Location = New System.Drawing.Point(424, 40)
		Me.lblDays.TabIndex = 18
		Me.lblDays.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblDays.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblDays.BackColor = System.Drawing.SystemColors.Control
		Me.lblDays.Enabled = True
		Me.lblDays.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblDays.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblDays.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblDays.UseMnemonic = True
		Me.lblDays.Visible = True
		Me.lblDays.AutoSize = False
		Me.lblDays.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblDays.Name = "lblDays"
		Me.Controls.Add(cbDays)
		Me.Controls.Add(fraPostOptions)
		Me.Controls.Add(fraOptions)
		Me.Controls.Add(chkTelegrams)
		Me.Controls.Add(frmEnemy)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(lblDays)
		Me.fraPostOptions.Controls.Add(chkCoastWatch)
		Me.fraPostOptions.Controls.Add(chkSpy)
		Me.fraPostOptions.Controls.Add(chkLook)
		Me.fraPostOptions.Controls.Add(chkLlook)
		Me.fraOptions.Controls.Add(chkEnemy)
		Me.fraOptions.Controls.Add(chkNeutral)
		Me.fraOptions.Controls.Add(chkAllied)
		Me.frmEnemy.Controls.Add(chkAirCombat)
		Me.frmEnemy.Controls.Add(chkPlanes)
		Me.frmEnemy.Controls.Add(chkShips)
		Me.frmEnemy.Controls.Add(chkLands)
		Me.frmEnemy.Controls.Add(chkSectors)
		Me.fraPostOptions.ResumeLayout(False)
		Me.fraOptions.ResumeLayout(False)
		Me.frmEnemy.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class