<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmNavCtrl
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
	Public WithEvents chkRetreat As System.Windows.Forms.CheckBox
	Public WithEvents txtRetreatPath As System.Windows.Forms.TextBox
	Public WithEvents Combo1 As System.Windows.Forms.ComboBox
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents lstReport As System.Windows.Forms.ListBox
	Public WithEvents _cmdOption_9 As System.Windows.Forms.Button
	Public WithEvents _cmdOption_6 As System.Windows.Forms.Button
	Public WithEvents _cmdOption_7 As System.Windows.Forms.Button
	Public WithEvents _cmdOption_8 As System.Windows.Forms.Button
	Public WithEvents _cmdOption_3 As System.Windows.Forms.Button
	Public WithEvents _cmdOption_1 As System.Windows.Forms.Button
	Public WithEvents _cmdOption_4 As System.Windows.Forms.Button
	Public WithEvents _cmdOption_5 As System.Windows.Forms.Button
	Public WithEvents _cmdOption_2 As System.Windows.Forms.Button
	Public WithEvents _cmdOption_0 As System.Windows.Forms.Button
	Public WithEvents FrameOptions As System.Windows.Forms.GroupBox
	Public WithEvents _cmdDir_0 As System.Windows.Forms.Button
	Public WithEvents _cmdDir_1 As System.Windows.Forms.Button
	Public WithEvents _cmdDir_2 As System.Windows.Forms.Button
	Public WithEvents _cmdDir_3 As System.Windows.Forms.Button
	Public WithEvents _cmdDir_4 As System.Windows.Forms.Button
	Public WithEvents _cmdDir_5 As System.Windows.Forms.Button
	Public WithEvents _cmdDir_6 As System.Windows.Forms.Button
	Public WithEvents frameDir As System.Windows.Forms.GroupBox
	Public WithEvents cmdDir As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
	Public WithEvents cmdOption As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmNavCtrl))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.chkRetreat = New System.Windows.Forms.CheckBox
		Me.txtRetreatPath = New System.Windows.Forms.TextBox
		Me.Combo1 = New System.Windows.Forms.ComboBox
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.lstReport = New System.Windows.Forms.ListBox
		Me.FrameOptions = New System.Windows.Forms.GroupBox
		Me._cmdOption_9 = New System.Windows.Forms.Button
		Me._cmdOption_6 = New System.Windows.Forms.Button
		Me._cmdOption_7 = New System.Windows.Forms.Button
		Me._cmdOption_8 = New System.Windows.Forms.Button
		Me._cmdOption_3 = New System.Windows.Forms.Button
		Me._cmdOption_1 = New System.Windows.Forms.Button
		Me._cmdOption_4 = New System.Windows.Forms.Button
		Me._cmdOption_5 = New System.Windows.Forms.Button
		Me._cmdOption_2 = New System.Windows.Forms.Button
		Me._cmdOption_0 = New System.Windows.Forms.Button
		Me.frameDir = New System.Windows.Forms.GroupBox
		Me._cmdDir_0 = New System.Windows.Forms.Button
		Me._cmdDir_1 = New System.Windows.Forms.Button
		Me._cmdDir_2 = New System.Windows.Forms.Button
		Me._cmdDir_3 = New System.Windows.Forms.Button
		Me._cmdDir_4 = New System.Windows.Forms.Button
		Me._cmdDir_5 = New System.Windows.Forms.Button
		Me._cmdDir_6 = New System.Windows.Forms.Button
		Me.cmdDir = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(components)
		Me.cmdOption = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(components)
		Me.Frame1.SuspendLayout()
		Me.FrameOptions.SuspendLayout()
		Me.frameDir.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.cmdDir, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.cmdOption, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Select Path"
		Me.ClientSize = New System.Drawing.Size(650, 392)
		Me.Location = New System.Drawing.Point(3, 22)
		Me.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmNavCtrl"
		Me.Frame1.Text = "Retreat Settings"
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.Size = New System.Drawing.Size(169, 105)
		Me.Frame1.Location = New System.Drawing.Point(376, 256)
		Me.Frame1.TabIndex = 20
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.chkRetreat.Text = "Set Retreat "
		Me.chkRetreat.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkRetreat.Size = New System.Drawing.Size(89, 25)
		Me.chkRetreat.Location = New System.Drawing.Point(8, 16)
		Me.chkRetreat.TabIndex = 23
		Me.chkRetreat.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkRetreat.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkRetreat.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkRetreat.BackColor = System.Drawing.SystemColors.Control
		Me.chkRetreat.CausesValidation = True
		Me.chkRetreat.Enabled = True
		Me.chkRetreat.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkRetreat.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkRetreat.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkRetreat.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkRetreat.TabStop = True
		Me.chkRetreat.Visible = True
		Me.chkRetreat.Name = "chkRetreat"
		Me.txtRetreatPath.AutoSize = False
		Me.txtRetreatPath.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtRetreatPath.Size = New System.Drawing.Size(89, 19)
		Me.txtRetreatPath.Location = New System.Drawing.Point(48, 72)
		Me.txtRetreatPath.TabIndex = 22
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
		Me.Combo1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Combo1.Size = New System.Drawing.Size(89, 21)
		Me.Combo1.Location = New System.Drawing.Point(48, 48)
		Me.Combo1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.Combo1.TabIndex = 21
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
		Me.Label3.Text = "path"
		Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.Size = New System.Drawing.Size(33, 17)
		Me.Label3.Location = New System.Drawing.Point(8, 72)
		Me.Label3.TabIndex = 25
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
		Me.Label4.Text = "cond"
		Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.Size = New System.Drawing.Size(33, 17)
		Me.Label4.Location = New System.Drawing.Point(8, 48)
		Me.Label4.TabIndex = 24
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
		Me.lstReport.Size = New System.Drawing.Size(641, 247)
		Me.lstReport.Location = New System.Drawing.Point(0, 0)
		Me.lstReport.TabIndex = 19
		Me.lstReport.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstReport.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstReport.BackColor = System.Drawing.SystemColors.Window
		Me.lstReport.CausesValidation = True
		Me.lstReport.Enabled = True
		Me.lstReport.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstReport.IntegralHeight = True
		Me.lstReport.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstReport.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstReport.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstReport.Sorted = False
		Me.lstReport.TabStop = True
		Me.lstReport.Visible = True
		Me.lstReport.MultiColumn = False
		Me.lstReport.Name = "lstReport"
		Me.FrameOptions.Text = "Options"
		Me.FrameOptions.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.FrameOptions.Size = New System.Drawing.Size(233, 121)
		Me.FrameOptions.Location = New System.Drawing.Point(120, 256)
		Me.FrameOptions.TabIndex = 8
		Me.FrameOptions.BackColor = System.Drawing.SystemColors.Control
		Me.FrameOptions.Enabled = True
		Me.FrameOptions.ForeColor = System.Drawing.SystemColors.ControlText
		Me.FrameOptions.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.FrameOptions.Visible = True
		Me.FrameOptions.Padding = New System.Windows.Forms.Padding(0)
		Me.FrameOptions.Name = "FrameOptions"
		Me._cmdOption_9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdOption_9.Text = "Halt"
		Me._cmdOption_9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdOption_9.Size = New System.Drawing.Size(41, 25)
		Me._cmdOption_9.Location = New System.Drawing.Point(160, 56)
		Me._cmdOption_9.TabIndex = 18
		Me._cmdOption_9.Tag = "h"
		Me._cmdOption_9.BackColor = System.Drawing.SystemColors.Control
		Me._cmdOption_9.CausesValidation = True
		Me._cmdOption_9.Enabled = True
		Me._cmdOption_9.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdOption_9.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdOption_9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdOption_9.TabStop = True
		Me._cmdOption_9.Name = "_cmdOption_9"
		Me._cmdOption_6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdOption_6.Text = "Sweep"
		Me._cmdOption_6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdOption_6.Size = New System.Drawing.Size(41, 25)
		Me._cmdOption_6.Location = New System.Drawing.Point(160, 24)
		Me._cmdOption_6.TabIndex = 17
		Me._cmdOption_6.Tag = "m"
		Me._cmdOption_6.BackColor = System.Drawing.SystemColors.Control
		Me._cmdOption_6.CausesValidation = True
		Me._cmdOption_6.Enabled = True
		Me._cmdOption_6.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdOption_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdOption_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdOption_6.TabStop = True
		Me._cmdOption_6.Name = "_cmdOption_6"
		Me._cmdOption_7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdOption_7.Text = "List Ships"
		Me._cmdOption_7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdOption_7.Size = New System.Drawing.Size(81, 25)
		Me._cmdOption_7.Location = New System.Drawing.Point(16, 88)
		Me._cmdOption_7.TabIndex = 16
		Me._cmdOption_7.Tag = "i"
		Me._cmdOption_7.BackColor = System.Drawing.SystemColors.Control
		Me._cmdOption_7.CausesValidation = True
		Me._cmdOption_7.Enabled = True
		Me._cmdOption_7.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdOption_7.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdOption_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdOption_7.TabStop = True
		Me._cmdOption_7.Name = "_cmdOption_7"
		Me._cmdOption_8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdOption_8.Text = "Change Flagship"
		Me._cmdOption_8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdOption_8.Size = New System.Drawing.Size(113, 25)
		Me._cmdOption_8.Location = New System.Drawing.Point(104, 88)
		Me._cmdOption_8.TabIndex = 15
		Me._cmdOption_8.Tag = "f"
		Me._cmdOption_8.BackColor = System.Drawing.SystemColors.Control
		Me._cmdOption_8.CausesValidation = True
		Me._cmdOption_8.Enabled = True
		Me._cmdOption_8.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdOption_8.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdOption_8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdOption_8.TabStop = True
		Me._cmdOption_8.Name = "_cmdOption_8"
		Me._cmdOption_3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdOption_3.Text = "Sonar"
		Me._cmdOption_3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdOption_3.Size = New System.Drawing.Size(41, 25)
		Me._cmdOption_3.Location = New System.Drawing.Point(64, 56)
		Me._cmdOption_3.TabIndex = 14
		Me._cmdOption_3.Tag = "s"
		Me._cmdOption_3.BackColor = System.Drawing.SystemColors.Control
		Me._cmdOption_3.CausesValidation = True
		Me._cmdOption_3.Enabled = True
		Me._cmdOption_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdOption_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdOption_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdOption_3.TabStop = True
		Me._cmdOption_3.Name = "_cmdOption_3"
		Me._cmdOption_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdOption_1.Text = "View"
		Me._cmdOption_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdOption_1.Size = New System.Drawing.Size(41, 25)
		Me._cmdOption_1.Location = New System.Drawing.Point(16, 56)
		Me._cmdOption_1.TabIndex = 13
		Me._cmdOption_1.Tag = "v"
		Me._cmdOption_1.BackColor = System.Drawing.SystemColors.Control
		Me._cmdOption_1.CausesValidation = True
		Me._cmdOption_1.Enabled = True
		Me._cmdOption_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdOption_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdOption_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdOption_1.TabStop = True
		Me._cmdOption_1.Name = "_cmdOption_1"
		Me._cmdOption_4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdOption_4.Text = "Bmap"
		Me._cmdOption_4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdOption_4.Size = New System.Drawing.Size(41, 25)
		Me._cmdOption_4.Location = New System.Drawing.Point(112, 24)
		Me._cmdOption_4.TabIndex = 12
		Me._cmdOption_4.Tag = "B"
		Me._cmdOption_4.BackColor = System.Drawing.SystemColors.Control
		Me._cmdOption_4.CausesValidation = True
		Me._cmdOption_4.Enabled = True
		Me._cmdOption_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdOption_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdOption_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdOption_4.TabStop = True
		Me._cmdOption_4.Name = "_cmdOption_4"
		Me._cmdOption_5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdOption_5.Text = "Map"
		Me._cmdOption_5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdOption_5.Size = New System.Drawing.Size(41, 25)
		Me._cmdOption_5.Location = New System.Drawing.Point(112, 56)
		Me._cmdOption_5.TabIndex = 11
		Me._cmdOption_5.Tag = "M"
		Me._cmdOption_5.BackColor = System.Drawing.SystemColors.Control
		Me._cmdOption_5.CausesValidation = True
		Me._cmdOption_5.Enabled = True
		Me._cmdOption_5.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdOption_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdOption_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdOption_5.TabStop = True
		Me._cmdOption_5.Name = "_cmdOption_5"
		Me._cmdOption_2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdOption_2.Text = "Radar"
		Me._cmdOption_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdOption_2.Size = New System.Drawing.Size(41, 25)
		Me._cmdOption_2.Location = New System.Drawing.Point(64, 24)
		Me._cmdOption_2.TabIndex = 10
		Me._cmdOption_2.Tag = "r"
		Me._cmdOption_2.BackColor = System.Drawing.SystemColors.Control
		Me._cmdOption_2.CausesValidation = True
		Me._cmdOption_2.Enabled = True
		Me._cmdOption_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdOption_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdOption_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdOption_2.TabStop = True
		Me._cmdOption_2.Name = "_cmdOption_2"
		Me._cmdOption_0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdOption_0.Text = "Look"
		Me._cmdOption_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdOption_0.Size = New System.Drawing.Size(41, 25)
		Me._cmdOption_0.Location = New System.Drawing.Point(16, 24)
		Me._cmdOption_0.TabIndex = 9
		Me._cmdOption_0.Tag = "l"
		Me._cmdOption_0.BackColor = System.Drawing.SystemColors.Control
		Me._cmdOption_0.CausesValidation = True
		Me._cmdOption_0.Enabled = True
		Me._cmdOption_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdOption_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdOption_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdOption_0.TabStop = True
		Me._cmdOption_0.Name = "_cmdOption_0"
		Me.frameDir.Text = "Movement"
		Me.frameDir.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frameDir.Size = New System.Drawing.Size(97, 121)
		Me.frameDir.Location = New System.Drawing.Point(8, 256)
		Me.frameDir.TabIndex = 0
		Me.frameDir.BackColor = System.Drawing.SystemColors.Control
		Me.frameDir.Enabled = True
		Me.frameDir.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frameDir.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frameDir.Visible = True
		Me.frameDir.Padding = New System.Windows.Forms.Padding(0)
		Me.frameDir.Name = "frameDir"
		Me._cmdDir_0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdDir_0.Text = "Y"
		Me._cmdDir_0.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdDir_0.Size = New System.Drawing.Size(17, 17)
		Me._cmdDir_0.Location = New System.Drawing.Point(24, 32)
		Me._cmdDir_0.TabIndex = 7
		Me._cmdDir_0.BackColor = System.Drawing.SystemColors.Control
		Me._cmdDir_0.CausesValidation = True
		Me._cmdDir_0.Enabled = True
		Me._cmdDir_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdDir_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdDir_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdDir_0.TabStop = True
		Me._cmdDir_0.Name = "_cmdDir_0"
		Me._cmdDir_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdDir_1.Text = "U"
		Me._cmdDir_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdDir_1.Size = New System.Drawing.Size(17, 17)
		Me._cmdDir_1.Location = New System.Drawing.Point(56, 32)
		Me._cmdDir_1.TabIndex = 6
		Me._cmdDir_1.BackColor = System.Drawing.SystemColors.Control
		Me._cmdDir_1.CausesValidation = True
		Me._cmdDir_1.Enabled = True
		Me._cmdDir_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdDir_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdDir_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdDir_1.TabStop = True
		Me._cmdDir_1.Name = "_cmdDir_1"
		Me._cmdDir_2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdDir_2.Text = "G"
		Me._cmdDir_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdDir_2.Size = New System.Drawing.Size(17, 17)
		Me._cmdDir_2.Location = New System.Drawing.Point(16, 56)
		Me._cmdDir_2.TabIndex = 5
		Me._cmdDir_2.BackColor = System.Drawing.SystemColors.Control
		Me._cmdDir_2.CausesValidation = True
		Me._cmdDir_2.Enabled = True
		Me._cmdDir_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdDir_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdDir_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdDir_2.TabStop = True
		Me._cmdDir_2.Name = "_cmdDir_2"
		Me._cmdDir_3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdDir_3.Text = "H"
		Me._cmdDir_3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdDir_3.Size = New System.Drawing.Size(17, 17)
		Me._cmdDir_3.Location = New System.Drawing.Point(40, 56)
		Me._cmdDir_3.TabIndex = 4
		Me._cmdDir_3.BackColor = System.Drawing.SystemColors.Control
		Me._cmdDir_3.CausesValidation = True
		Me._cmdDir_3.Enabled = True
		Me._cmdDir_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdDir_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdDir_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdDir_3.TabStop = True
		Me._cmdDir_3.Name = "_cmdDir_3"
		Me._cmdDir_4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdDir_4.Text = "J"
		Me._cmdDir_4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdDir_4.Size = New System.Drawing.Size(17, 17)
		Me._cmdDir_4.Location = New System.Drawing.Point(64, 56)
		Me._cmdDir_4.TabIndex = 3
		Me._cmdDir_4.BackColor = System.Drawing.SystemColors.Control
		Me._cmdDir_4.CausesValidation = True
		Me._cmdDir_4.Enabled = True
		Me._cmdDir_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdDir_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdDir_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdDir_4.TabStop = True
		Me._cmdDir_4.Name = "_cmdDir_4"
		Me._cmdDir_5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdDir_5.Text = "B"
		Me._cmdDir_5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdDir_5.Size = New System.Drawing.Size(17, 17)
		Me._cmdDir_5.Location = New System.Drawing.Point(24, 80)
		Me._cmdDir_5.TabIndex = 2
		Me._cmdDir_5.BackColor = System.Drawing.SystemColors.Control
		Me._cmdDir_5.CausesValidation = True
		Me._cmdDir_5.Enabled = True
		Me._cmdDir_5.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdDir_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdDir_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdDir_5.TabStop = True
		Me._cmdDir_5.Name = "_cmdDir_5"
		Me._cmdDir_6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdDir_6.Text = "N"
		Me._cmdDir_6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdDir_6.Size = New System.Drawing.Size(17, 17)
		Me._cmdDir_6.Location = New System.Drawing.Point(56, 80)
		Me._cmdDir_6.TabIndex = 1
		Me._cmdDir_6.BackColor = System.Drawing.SystemColors.Control
		Me._cmdDir_6.CausesValidation = True
		Me._cmdDir_6.Enabled = True
		Me._cmdDir_6.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdDir_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdDir_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdDir_6.TabStop = True
		Me._cmdDir_6.Name = "_cmdDir_6"
		Me.Controls.Add(Frame1)
		Me.Controls.Add(lstReport)
		Me.Controls.Add(FrameOptions)
		Me.Controls.Add(frameDir)
		Me.Frame1.Controls.Add(chkRetreat)
		Me.Frame1.Controls.Add(txtRetreatPath)
		Me.Frame1.Controls.Add(Combo1)
		Me.Frame1.Controls.Add(Label3)
		Me.Frame1.Controls.Add(Label4)
		Me.FrameOptions.Controls.Add(_cmdOption_9)
		Me.FrameOptions.Controls.Add(_cmdOption_6)
		Me.FrameOptions.Controls.Add(_cmdOption_7)
		Me.FrameOptions.Controls.Add(_cmdOption_8)
		Me.FrameOptions.Controls.Add(_cmdOption_3)
		Me.FrameOptions.Controls.Add(_cmdOption_1)
		Me.FrameOptions.Controls.Add(_cmdOption_4)
		Me.FrameOptions.Controls.Add(_cmdOption_5)
		Me.FrameOptions.Controls.Add(_cmdOption_2)
		Me.FrameOptions.Controls.Add(_cmdOption_0)
		Me.frameDir.Controls.Add(_cmdDir_0)
		Me.frameDir.Controls.Add(_cmdDir_1)
		Me.frameDir.Controls.Add(_cmdDir_2)
		Me.frameDir.Controls.Add(_cmdDir_3)
		Me.frameDir.Controls.Add(_cmdDir_4)
		Me.frameDir.Controls.Add(_cmdDir_5)
		Me.frameDir.Controls.Add(_cmdDir_6)
		Me.cmdDir.SetIndex(_cmdDir_0, CType(0, Short))
		Me.cmdDir.SetIndex(_cmdDir_1, CType(1, Short))
		Me.cmdDir.SetIndex(_cmdDir_2, CType(2, Short))
		Me.cmdDir.SetIndex(_cmdDir_3, CType(3, Short))
		Me.cmdDir.SetIndex(_cmdDir_4, CType(4, Short))
		Me.cmdDir.SetIndex(_cmdDir_5, CType(5, Short))
		Me.cmdDir.SetIndex(_cmdDir_6, CType(6, Short))
		Me.cmdOption.SetIndex(_cmdOption_9, CType(9, Short))
		Me.cmdOption.SetIndex(_cmdOption_6, CType(6, Short))
		Me.cmdOption.SetIndex(_cmdOption_7, CType(7, Short))
		Me.cmdOption.SetIndex(_cmdOption_8, CType(8, Short))
		Me.cmdOption.SetIndex(_cmdOption_3, CType(3, Short))
		Me.cmdOption.SetIndex(_cmdOption_1, CType(1, Short))
		Me.cmdOption.SetIndex(_cmdOption_4, CType(4, Short))
		Me.cmdOption.SetIndex(_cmdOption_5, CType(5, Short))
		Me.cmdOption.SetIndex(_cmdOption_2, CType(2, Short))
		Me.cmdOption.SetIndex(_cmdOption_0, CType(0, Short))
		CType(Me.cmdOption, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cmdDir, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Frame1.ResumeLayout(False)
		Me.FrameOptions.ResumeLayout(False)
		Me.frameDir.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class