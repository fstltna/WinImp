<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPromptBuild
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
	Public WithEvents cmdHelp As System.Windows.Forms.Button
	Public WithEvents txtType As System.Windows.Forms.ComboBox
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	Public WithEvents cmdShow As System.Windows.Forms.Button
	Public WithEvents _cmdDir_6 As System.Windows.Forms.Button
	Public WithEvents _cmdDir_5 As System.Windows.Forms.Button
	Public WithEvents _cmdDir_4 As System.Windows.Forms.Button
	Public WithEvents _cmdDir_3 As System.Windows.Forms.Button
	Public WithEvents _cmdDir_2 As System.Windows.Forms.Button
	Public WithEvents _cmdDir_1 As System.Windows.Forms.Button
	Public WithEvents _cmdDir_0 As System.Windows.Forms.Button
	Public WithEvents frmDir As System.Windows.Forms.Panel
	Public WithEvents txtNum As System.Windows.Forms.TextBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents txtOrigin As System.Windows.Forms.TextBox
	Public WithEvents txtTechLevel As System.Windows.Forms.TextBox
	Public WithEvents Option8 As System.Windows.Forms.RadioButton
	Public WithEvents Option7 As System.Windows.Forms.RadioButton
	Public WithEvents Option4 As System.Windows.Forms.RadioButton
	Public WithEvents Option2 As System.Windows.Forms.RadioButton
	Public WithEvents Option1 As System.Windows.Forms.RadioButton
	Public WithEvents Option3 As System.Windows.Forms.RadioButton
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents cmdDir As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPromptBuild))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdHelp = New System.Windows.Forms.Button
		Me.txtType = New System.Windows.Forms.ComboBox
		Me.Timer1 = New System.Windows.Forms.Timer(components)
		Me.cmdShow = New System.Windows.Forms.Button
		Me.frmDir = New System.Windows.Forms.Panel
		Me._cmdDir_6 = New System.Windows.Forms.Button
		Me._cmdDir_5 = New System.Windows.Forms.Button
		Me._cmdDir_4 = New System.Windows.Forms.Button
		Me._cmdDir_3 = New System.Windows.Forms.Button
		Me._cmdDir_2 = New System.Windows.Forms.Button
		Me._cmdDir_1 = New System.Windows.Forms.Button
		Me._cmdDir_0 = New System.Windows.Forms.Button
		Me.txtNum = New System.Windows.Forms.TextBox
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me.txtOrigin = New System.Windows.Forms.TextBox
		Me.txtTechLevel = New System.Windows.Forms.TextBox
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.Option8 = New System.Windows.Forms.RadioButton
		Me.Option7 = New System.Windows.Forms.RadioButton
		Me.Option4 = New System.Windows.Forms.RadioButton
		Me.Option2 = New System.Windows.Forms.RadioButton
		Me.Option1 = New System.Windows.Forms.RadioButton
		Me.Option3 = New System.Windows.Forms.RadioButton
		Me.Label5 = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.cmdDir = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(components)
		Me.frmDir.SuspendLayout()
		Me.Frame1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.cmdDir, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.ControlBox = False
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.ClientSize = New System.Drawing.Size(491, 156)
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
		Me.Name = "frmPromptBuild"
		Me.cmdHelp.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdHelp.Text = "?"
		Me.cmdHelp.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdHelp.Size = New System.Drawing.Size(25, 25)
		Me.cmdHelp.Location = New System.Drawing.Point(464, 0)
		Me.cmdHelp.TabIndex = 27
		Me.ToolTip1.SetToolTip(Me.cmdHelp, "Click for Help")
		Me.cmdHelp.BackColor = System.Drawing.SystemColors.Control
		Me.cmdHelp.CausesValidation = True
		Me.cmdHelp.Enabled = True
		Me.cmdHelp.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdHelp.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdHelp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdHelp.TabStop = True
		Me.cmdHelp.Name = "cmdHelp"
		Me.txtType.Font = New System.Drawing.Font("Courier New", 8.4!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtType.Size = New System.Drawing.Size(145, 22)
		Me.txtType.Location = New System.Drawing.Point(232, 56)
		Me.txtType.Sorted = True
		Me.txtType.TabIndex = 8
		Me.txtType.BackColor = System.Drawing.SystemColors.Window
		Me.txtType.CausesValidation = True
		Me.txtType.Enabled = True
		Me.txtType.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtType.IntegralHeight = True
		Me.txtType.Cursor = System.Windows.Forms.Cursors.Default
		Me.txtType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.txtType.TabStop = True
		Me.txtType.Visible = True
		Me.txtType.Name = "txtType"
		Me.Timer1.Enabled = False
		Me.Timer1.Interval = 1
		Me.cmdShow.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdShow.Text = "Show Builds"
		Me.cmdShow.Size = New System.Drawing.Size(81, 25)
		Me.cmdShow.Location = New System.Drawing.Point(392, 88)
		Me.cmdShow.TabIndex = 24
		Me.cmdShow.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdShow.BackColor = System.Drawing.SystemColors.Control
		Me.cmdShow.CausesValidation = True
		Me.cmdShow.Enabled = True
		Me.cmdShow.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdShow.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdShow.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdShow.TabStop = True
		Me.cmdShow.Name = "cmdShow"
		Me.frmDir.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.frmDir.Size = New System.Drawing.Size(89, 73)
		Me.frmDir.Location = New System.Drawing.Point(24, 80)
		Me.frmDir.TabIndex = 15
		Me.frmDir.Visible = False
		Me.frmDir.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmDir.BackColor = System.Drawing.SystemColors.Control
		Me.frmDir.Enabled = True
		Me.frmDir.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmDir.Cursor = System.Windows.Forms.Cursors.Default
		Me.frmDir.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmDir.Name = "frmDir"
		Me._cmdDir_6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdDir_6.Text = "N"
		Me._cmdDir_6.Size = New System.Drawing.Size(17, 17)
		Me._cmdDir_6.Location = New System.Drawing.Point(48, 48)
		Me._cmdDir_6.TabIndex = 22
		Me._cmdDir_6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdDir_6.BackColor = System.Drawing.SystemColors.Control
		Me._cmdDir_6.CausesValidation = True
		Me._cmdDir_6.Enabled = True
		Me._cmdDir_6.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdDir_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdDir_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdDir_6.TabStop = True
		Me._cmdDir_6.Name = "_cmdDir_6"
		Me._cmdDir_5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdDir_5.Text = "B"
		Me._cmdDir_5.Size = New System.Drawing.Size(17, 17)
		Me._cmdDir_5.Location = New System.Drawing.Point(16, 48)
		Me._cmdDir_5.TabIndex = 21
		Me._cmdDir_5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdDir_5.BackColor = System.Drawing.SystemColors.Control
		Me._cmdDir_5.CausesValidation = True
		Me._cmdDir_5.Enabled = True
		Me._cmdDir_5.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdDir_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdDir_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdDir_5.TabStop = True
		Me._cmdDir_5.Name = "_cmdDir_5"
		Me._cmdDir_4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdDir_4.Text = "J"
		Me._cmdDir_4.Size = New System.Drawing.Size(17, 17)
		Me._cmdDir_4.Location = New System.Drawing.Point(56, 24)
		Me._cmdDir_4.TabIndex = 20
		Me._cmdDir_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdDir_4.BackColor = System.Drawing.SystemColors.Control
		Me._cmdDir_4.CausesValidation = True
		Me._cmdDir_4.Enabled = True
		Me._cmdDir_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdDir_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdDir_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdDir_4.TabStop = True
		Me._cmdDir_4.Name = "_cmdDir_4"
		Me._cmdDir_3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdDir_3.Text = "H"
		Me._cmdDir_3.Size = New System.Drawing.Size(17, 17)
		Me._cmdDir_3.Location = New System.Drawing.Point(32, 24)
		Me._cmdDir_3.TabIndex = 19
		Me._cmdDir_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdDir_3.BackColor = System.Drawing.SystemColors.Control
		Me._cmdDir_3.CausesValidation = True
		Me._cmdDir_3.Enabled = True
		Me._cmdDir_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdDir_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdDir_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdDir_3.TabStop = True
		Me._cmdDir_3.Name = "_cmdDir_3"
		Me._cmdDir_2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdDir_2.Text = "G"
		Me._cmdDir_2.Size = New System.Drawing.Size(17, 17)
		Me._cmdDir_2.Location = New System.Drawing.Point(8, 24)
		Me._cmdDir_2.TabIndex = 18
		Me._cmdDir_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdDir_2.BackColor = System.Drawing.SystemColors.Control
		Me._cmdDir_2.CausesValidation = True
		Me._cmdDir_2.Enabled = True
		Me._cmdDir_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdDir_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdDir_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdDir_2.TabStop = True
		Me._cmdDir_2.Name = "_cmdDir_2"
		Me._cmdDir_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdDir_1.Text = "U"
		Me._cmdDir_1.Size = New System.Drawing.Size(17, 17)
		Me._cmdDir_1.Location = New System.Drawing.Point(48, 0)
		Me._cmdDir_1.TabIndex = 17
		Me._cmdDir_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdDir_1.BackColor = System.Drawing.SystemColors.Control
		Me._cmdDir_1.CausesValidation = True
		Me._cmdDir_1.Enabled = True
		Me._cmdDir_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdDir_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdDir_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdDir_1.TabStop = True
		Me._cmdDir_1.Name = "_cmdDir_1"
		Me._cmdDir_0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdDir_0.Text = "Y"
		Me._cmdDir_0.Size = New System.Drawing.Size(17, 17)
		Me._cmdDir_0.Location = New System.Drawing.Point(16, 0)
		Me._cmdDir_0.TabIndex = 16
		Me._cmdDir_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdDir_0.BackColor = System.Drawing.SystemColors.Control
		Me._cmdDir_0.CausesValidation = True
		Me._cmdDir_0.Enabled = True
		Me._cmdDir_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdDir_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdDir_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdDir_0.TabStop = True
		Me._cmdDir_0.Name = "_cmdDir_0"
		Me.txtNum.AutoSize = False
		Me.txtNum.Size = New System.Drawing.Size(49, 19)
		Me.txtNum.Location = New System.Drawing.Point(312, 88)
		Me.txtNum.TabIndex = 9
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
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(81, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(392, 120)
		Me.cmdCancel.TabIndex = 25
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
		Me.cmdOK.Location = New System.Drawing.Point(392, 56)
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
		Me.txtOrigin.AutoSize = False
		Me.txtOrigin.Size = New System.Drawing.Size(73, 19)
		Me.txtOrigin.Location = New System.Drawing.Point(232, 24)
		Me.txtOrigin.TabIndex = 7
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
		Me.txtTechLevel.AutoSize = False
		Me.txtTechLevel.Size = New System.Drawing.Size(49, 19)
		Me.txtTechLevel.Location = New System.Drawing.Point(312, 112)
		Me.txtTechLevel.TabIndex = 10
		Me.txtTechLevel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtTechLevel.AcceptsReturn = True
		Me.txtTechLevel.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtTechLevel.BackColor = System.Drawing.SystemColors.Window
		Me.txtTechLevel.CausesValidation = True
		Me.txtTechLevel.Enabled = True
		Me.txtTechLevel.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtTechLevel.HideSelection = True
		Me.txtTechLevel.ReadOnly = False
		Me.txtTechLevel.Maxlength = 0
		Me.txtTechLevel.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtTechLevel.MultiLine = False
		Me.txtTechLevel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtTechLevel.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtTechLevel.TabStop = True
		Me.txtTechLevel.Visible = True
		Me.txtTechLevel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtTechLevel.Name = "txtTechLevel"
		Me.Frame1.Text = "Select "
		Me.Frame1.Size = New System.Drawing.Size(105, 129)
		Me.Frame1.Location = New System.Drawing.Point(64, 8)
		Me.Frame1.TabIndex = 0
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.Option8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option8.Text = "bridge tower"
		Me.Option8.Size = New System.Drawing.Size(81, 17)
		Me.Option8.Location = New System.Drawing.Point(16, 101)
		Me.Option8.TabIndex = 5
		Me.Option8.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Option8.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option8.BackColor = System.Drawing.SystemColors.Control
		Me.Option8.CausesValidation = True
		Me.Option8.Enabled = True
		Me.Option8.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Option8.Cursor = System.Windows.Forms.Cursors.Default
		Me.Option8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Option8.Appearance = System.Windows.Forms.Appearance.Normal
		Me.Option8.TabStop = True
		Me.Option8.Checked = False
		Me.Option8.Visible = True
		Me.Option8.Name = "Option8"
		Me.Option7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option7.Text = "bridge span"
		Me.Option7.Size = New System.Drawing.Size(81, 17)
		Me.Option7.Location = New System.Drawing.Point(16, 84)
		Me.Option7.TabIndex = 4
		Me.Option7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Option7.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option7.BackColor = System.Drawing.SystemColors.Control
		Me.Option7.CausesValidation = True
		Me.Option7.Enabled = True
		Me.Option7.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Option7.Cursor = System.Windows.Forms.Cursors.Default
		Me.Option7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Option7.Appearance = System.Windows.Forms.Appearance.Normal
		Me.Option7.TabStop = True
		Me.Option7.Checked = False
		Me.Option7.Visible = True
		Me.Option7.Name = "Option7"
		Me.Option4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option4.Text = "nuke"
		Me.Option4.Size = New System.Drawing.Size(57, 17)
		Me.Option4.Location = New System.Drawing.Point(16, 67)
		Me.Option4.TabIndex = 26
		Me.Option4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Option4.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option4.BackColor = System.Drawing.SystemColors.Control
		Me.Option4.CausesValidation = True
		Me.Option4.Enabled = True
		Me.Option4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Option4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Option4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Option4.Appearance = System.Windows.Forms.Appearance.Normal
		Me.Option4.TabStop = True
		Me.Option4.Checked = False
		Me.Option4.Visible = True
		Me.Option4.Name = "Option4"
		Me.Option2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option2.Text = "plane"
		Me.Option2.Size = New System.Drawing.Size(57, 17)
		Me.Option2.Location = New System.Drawing.Point(16, 50)
		Me.Option2.TabIndex = 3
		Me.Option2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Option2.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option2.BackColor = System.Drawing.SystemColors.Control
		Me.Option2.CausesValidation = True
		Me.Option2.Enabled = True
		Me.Option2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Option2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Option2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Option2.Appearance = System.Windows.Forms.Appearance.Normal
		Me.Option2.TabStop = True
		Me.Option2.Checked = False
		Me.Option2.Visible = True
		Me.Option2.Name = "Option2"
		Me.Option1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option1.Text = "land"
		Me.Option1.Size = New System.Drawing.Size(49, 17)
		Me.Option1.Location = New System.Drawing.Point(16, 33)
		Me.Option1.TabIndex = 2
		Me.Option1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Option1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option1.BackColor = System.Drawing.SystemColors.Control
		Me.Option1.CausesValidation = True
		Me.Option1.Enabled = True
		Me.Option1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Option1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Option1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Option1.Appearance = System.Windows.Forms.Appearance.Normal
		Me.Option1.TabStop = True
		Me.Option1.Checked = False
		Me.Option1.Visible = True
		Me.Option1.Name = "Option1"
		Me.Option3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option3.Text = "ship"
		Me.Option3.Size = New System.Drawing.Size(49, 17)
		Me.Option3.Location = New System.Drawing.Point(16, 16)
		Me.Option3.TabIndex = 1
		Me.Option3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Option3.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option3.BackColor = System.Drawing.SystemColors.Control
		Me.Option3.CausesValidation = True
		Me.Option3.Enabled = True
		Me.Option3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Option3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Option3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Option3.Appearance = System.Windows.Forms.Appearance.Normal
		Me.Option3.TabStop = True
		Me.Option3.Checked = False
		Me.Option3.Visible = True
		Me.Option3.Name = "Option3"
		Me.Label5.Text = "number"
		Me.Label5.Font = New System.Drawing.Font("Arial", 9.6!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label5.Size = New System.Drawing.Size(49, 17)
		Me.Label5.Location = New System.Drawing.Point(184, 88)
		Me.Label5.TabIndex = 14
		Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopLeft
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
		Me.Label4.Text = "type"
		Me.Label4.Font = New System.Drawing.Font("Arial", 9.6!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.Size = New System.Drawing.Size(33, 17)
		Me.Label4.Location = New System.Drawing.Point(184, 56)
		Me.Label4.TabIndex = 13
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
		Me.Label3.Text = "in"
		Me.Label3.Font = New System.Drawing.Font("Arial", 9.6!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.Size = New System.Drawing.Size(33, 17)
		Me.Label3.Location = New System.Drawing.Point(184, 24)
		Me.Label3.TabIndex = 12
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
		Me.Label2.Text = "Build"
		Me.Label2.Font = New System.Drawing.Font("Arial Black", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(49, 17)
		Me.Label2.Location = New System.Drawing.Point(16, 64)
		Me.Label2.TabIndex = 11
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
		Me.Label1.Text = "Tech Level (Optional)"
		Me.Label1.Size = New System.Drawing.Size(113, 17)
		Me.Label1.Location = New System.Drawing.Point(184, 112)
		Me.Label1.TabIndex = 6
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
		Me.Controls.Add(cmdHelp)
		Me.Controls.Add(txtType)
		Me.Controls.Add(cmdShow)
		Me.Controls.Add(frmDir)
		Me.Controls.Add(txtNum)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(txtOrigin)
		Me.Controls.Add(txtTechLevel)
		Me.Controls.Add(Frame1)
		Me.Controls.Add(Label5)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		Me.frmDir.Controls.Add(_cmdDir_6)
		Me.frmDir.Controls.Add(_cmdDir_5)
		Me.frmDir.Controls.Add(_cmdDir_4)
		Me.frmDir.Controls.Add(_cmdDir_3)
		Me.frmDir.Controls.Add(_cmdDir_2)
		Me.frmDir.Controls.Add(_cmdDir_1)
		Me.frmDir.Controls.Add(_cmdDir_0)
		Me.Frame1.Controls.Add(Option8)
		Me.Frame1.Controls.Add(Option7)
		Me.Frame1.Controls.Add(Option4)
		Me.Frame1.Controls.Add(Option2)
		Me.Frame1.Controls.Add(Option1)
		Me.Frame1.Controls.Add(Option3)
		Me.cmdDir.SetIndex(_cmdDir_6, CType(6, Short))
		Me.cmdDir.SetIndex(_cmdDir_5, CType(5, Short))
		Me.cmdDir.SetIndex(_cmdDir_4, CType(4, Short))
		Me.cmdDir.SetIndex(_cmdDir_3, CType(3, Short))
		Me.cmdDir.SetIndex(_cmdDir_2, CType(2, Short))
		Me.cmdDir.SetIndex(_cmdDir_1, CType(1, Short))
		Me.cmdDir.SetIndex(_cmdDir_0, CType(0, Short))
		CType(Me.cmdDir, System.ComponentModel.ISupportInitialize).EndInit()
		Me.frmDir.ResumeLayout(False)
		Me.Frame1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class