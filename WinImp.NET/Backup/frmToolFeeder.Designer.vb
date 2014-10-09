<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmToolFeeder
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
	Public WithEvents Text2 As System.Windows.Forms.TextBox
	Public WithEvents CheckSetThresh As System.Windows.Forms.CheckBox
	Public WithEvents Command1 As System.Windows.Forms.Button
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents txtOrigin As System.Windows.Forms.TextBox
	Public WithEvents _Option1_1 As System.Windows.Forms.RadioButton
	Public WithEvents _Option1_0 As System.Windows.Forms.RadioButton
	Public WithEvents Frame3 As System.Windows.Forms.GroupBox
	Public WithEvents _Label2_1 As System.Windows.Forms.Label
	Public WithEvents _Label2_0 As System.Windows.Forms.Label
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents _Text1_2 As System.Windows.Forms.TextBox
	Public WithEvents _Text1_1 As System.Windows.Forms.TextBox
	Public WithEvents _Text1_0 As System.Windows.Forms.TextBox
	Public WithEvents _Label1_2 As System.Windows.Forms.Label
	Public WithEvents _Label1_1 As System.Windows.Forms.Label
	Public WithEvents _Label1_0 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents grid1 As AxMSFlexGridLib.AxMSFlexGrid
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents lblError As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents Label2 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents Option1 As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	Public WithEvents Text1 As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmToolFeeder))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Text2 = New System.Windows.Forms.TextBox
		Me.CheckSetThresh = New System.Windows.Forms.CheckBox
		Me.Command1 = New System.Windows.Forms.Button
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me.txtOrigin = New System.Windows.Forms.TextBox
		Me.Frame3 = New System.Windows.Forms.GroupBox
		Me._Option1_1 = New System.Windows.Forms.RadioButton
		Me._Option1_0 = New System.Windows.Forms.RadioButton
		Me.Frame2 = New System.Windows.Forms.GroupBox
		Me._Label2_1 = New System.Windows.Forms.Label
		Me._Label2_0 = New System.Windows.Forms.Label
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me._Text1_2 = New System.Windows.Forms.TextBox
		Me._Text1_1 = New System.Windows.Forms.TextBox
		Me._Text1_0 = New System.Windows.Forms.TextBox
		Me._Label1_2 = New System.Windows.Forms.Label
		Me._Label1_1 = New System.Windows.Forms.Label
		Me._Label1_0 = New System.Windows.Forms.Label
		Me.grid1 = New AxMSFlexGridLib.AxMSFlexGrid
		Me.Label5 = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.lblError = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.Label2 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.Option1 = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(components)
		Me.Text1 = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(components)
		Me.Frame3.SuspendLayout()
		Me.Frame2.SuspendLayout()
		Me.Frame1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.grid1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label2, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Option1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Text1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Famine Relief Tool"
		Me.ClientSize = New System.Drawing.Size(383, 382)
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
		Me.Name = "frmToolFeeder"
		Me.Text2.AutoSize = False
		Me.Text2.Size = New System.Drawing.Size(33, 19)
		Me.Text2.Location = New System.Drawing.Point(216, 328)
		Me.Text2.TabIndex = 22
		Me.Text2.Text = "10"
		Me.Text2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text2.AcceptsReturn = True
		Me.Text2.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text2.BackColor = System.Drawing.SystemColors.Window
		Me.Text2.CausesValidation = True
		Me.Text2.Enabled = True
		Me.Text2.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text2.HideSelection = True
		Me.Text2.ReadOnly = False
		Me.Text2.Maxlength = 0
		Me.Text2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text2.MultiLine = False
		Me.Text2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text2.TabStop = True
		Me.Text2.Visible = True
		Me.Text2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text2.Name = "Text2"
		Me.CheckSetThresh.Text = "Increase Food Thresholds after moving food"
		Me.CheckSetThresh.Size = New System.Drawing.Size(237, 13)
		Me.CheckSetThresh.Location = New System.Drawing.Point(8, 312)
		Me.CheckSetThresh.TabIndex = 20
		Me.CheckSetThresh.CheckState = System.Windows.Forms.CheckState.Checked
		Me.CheckSetThresh.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.CheckSetThresh.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.CheckSetThresh.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.CheckSetThresh.BackColor = System.Drawing.SystemColors.Control
		Me.CheckSetThresh.CausesValidation = True
		Me.CheckSetThresh.Enabled = True
		Me.CheckSetThresh.ForeColor = System.Drawing.SystemColors.ControlText
		Me.CheckSetThresh.Cursor = System.Windows.Forms.Cursors.Default
		Me.CheckSetThresh.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.CheckSetThresh.Appearance = System.Windows.Forms.Appearance.Normal
		Me.CheckSetThresh.TabStop = True
		Me.CheckSetThresh.Visible = True
		Me.CheckSetThresh.Name = "CheckSetThresh"
		Me.Command1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.Command1
		Me.Command1.Text = "Check Starvation"
		Me.Command1.Size = New System.Drawing.Size(97, 25)
		Me.Command1.Location = New System.Drawing.Point(272, 352)
		Me.Command1.TabIndex = 17
		Me.Command1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Command1.BackColor = System.Drawing.SystemColors.Control
		Me.Command1.CausesValidation = True
		Me.Command1.Enabled = True
		Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command1.TabStop = True
		Me.Command1.Name = "Command1"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(97, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(8, 352)
		Me.cmdCancel.TabIndex = 16
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
		Me.cmdOK.Text = "Distribute Food"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Size = New System.Drawing.Size(97, 25)
		Me.cmdOK.Location = New System.Drawing.Point(144, 352)
		Me.cmdOK.TabIndex = 15
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
		Me.txtOrigin.Size = New System.Drawing.Size(65, 19)
		Me.txtOrigin.Location = New System.Drawing.Point(72, 16)
		Me.txtOrigin.TabIndex = 12
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
		Me.Frame3.Text = "Food Level "
		Me.Frame3.Size = New System.Drawing.Size(129, 57)
		Me.Frame3.Location = New System.Drawing.Point(8, 248)
		Me.Frame3.TabIndex = 6
		Me.Frame3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame3.BackColor = System.Drawing.SystemColors.Control
		Me.Frame3.Enabled = True
		Me.Frame3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame3.Visible = True
		Me.Frame3.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame3.Name = "Frame3"
		Me._Option1_1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Option1_1.Text = "Insure Growth"
		Me._Option1_1.Size = New System.Drawing.Size(89, 17)
		Me._Option1_1.Location = New System.Drawing.Point(16, 32)
		Me._Option1_1.TabIndex = 8
		Me._Option1_1.Checked = True
		Me._Option1_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Option1_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Option1_1.BackColor = System.Drawing.SystemColors.Control
		Me._Option1_1.CausesValidation = True
		Me._Option1_1.Enabled = True
		Me._Option1_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Option1_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Option1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Option1_1.Appearance = System.Windows.Forms.Appearance.Normal
		Me._Option1_1.TabStop = True
		Me._Option1_1.Visible = True
		Me._Option1_1.Name = "_Option1_1"
		Me._Option1_0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Option1_0.Text = "Avoid Starvation"
		Me._Option1_0.Size = New System.Drawing.Size(105, 17)
		Me._Option1_0.Location = New System.Drawing.Point(16, 16)
		Me._Option1_0.TabIndex = 7
		Me._Option1_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Option1_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Option1_0.BackColor = System.Drawing.SystemColors.Control
		Me._Option1_0.CausesValidation = True
		Me._Option1_0.Enabled = True
		Me._Option1_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Option1_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Option1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Option1_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._Option1_0.TabStop = True
		Me._Option1_0.Checked = False
		Me._Option1_0.Visible = True
		Me._Option1_0.Name = "_Option1_0"
		Me.Frame2.Text = "Projected Resources Used"
		Me.Frame2.Size = New System.Drawing.Size(185, 57)
		Me.Frame2.Location = New System.Drawing.Point(168, 248)
		Me.Frame2.TabIndex = 1
		Me.Frame2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame2.BackColor = System.Drawing.SystemColors.Control
		Me.Frame2.Enabled = True
		Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame2.Visible = True
		Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame2.Name = "Frame2"
		Me._Label2_1.Text = "Shortage: "
		Me._Label2_1.Size = New System.Drawing.Size(161, 17)
		Me._Label2_1.Location = New System.Drawing.Point(16, 32)
		Me._Label2_1.TabIndex = 10
		Me._Label2_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label2_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label2_1.BackColor = System.Drawing.SystemColors.Control
		Me._Label2_1.Enabled = True
		Me._Label2_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label2_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label2_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label2_1.UseMnemonic = True
		Me._Label2_1.Visible = True
		Me._Label2_1.AutoSize = False
		Me._Label2_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label2_1.Name = "_Label2_1"
		Me._Label2_0.Text = "240 food, 27.4 mobility"
		Me._Label2_0.Size = New System.Drawing.Size(161, 17)
		Me._Label2_0.Location = New System.Drawing.Point(16, 16)
		Me._Label2_0.TabIndex = 9
		Me._Label2_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label2_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label2_0.BackColor = System.Drawing.SystemColors.Control
		Me._Label2_0.Enabled = True
		Me._Label2_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label2_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label2_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label2_0.UseMnemonic = True
		Me._Label2_0.Visible = True
		Me._Label2_0.AutoSize = False
		Me._Label2_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label2_0.Name = "_Label2_0"
		Me.Frame1.Text = "Resources Limits"
		Me.Frame1.Size = New System.Drawing.Size(169, 65)
		Me.Frame1.Location = New System.Drawing.Point(208, 8)
		Me.Frame1.TabIndex = 0
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me._Text1_2.AutoSize = False
		Me._Text1_2.Size = New System.Drawing.Size(33, 19)
		Me._Text1_2.Location = New System.Drawing.Point(128, 16)
		Me._Text1_2.TabIndex = 18
		Me._Text1_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Text1_2.AcceptsReturn = True
		Me._Text1_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._Text1_2.BackColor = System.Drawing.SystemColors.Window
		Me._Text1_2.CausesValidation = True
		Me._Text1_2.Enabled = True
		Me._Text1_2.ForeColor = System.Drawing.SystemColors.WindowText
		Me._Text1_2.HideSelection = True
		Me._Text1_2.ReadOnly = False
		Me._Text1_2.Maxlength = 0
		Me._Text1_2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._Text1_2.MultiLine = False
		Me._Text1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Text1_2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._Text1_2.TabStop = True
		Me._Text1_2.Visible = True
		Me._Text1_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Text1_2.Name = "_Text1_2"
		Me._Text1_1.AutoSize = False
		Me._Text1_1.Size = New System.Drawing.Size(33, 19)
		Me._Text1_1.Location = New System.Drawing.Point(40, 40)
		Me._Text1_1.TabIndex = 5
		Me._Text1_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Text1_1.AcceptsReturn = True
		Me._Text1_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._Text1_1.BackColor = System.Drawing.SystemColors.Window
		Me._Text1_1.CausesValidation = True
		Me._Text1_1.Enabled = True
		Me._Text1_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._Text1_1.HideSelection = True
		Me._Text1_1.ReadOnly = False
		Me._Text1_1.Maxlength = 0
		Me._Text1_1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._Text1_1.MultiLine = False
		Me._Text1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Text1_1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._Text1_1.TabStop = True
		Me._Text1_1.Visible = True
		Me._Text1_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Text1_1.Name = "_Text1_1"
		Me._Text1_0.AutoSize = False
		Me._Text1_0.Size = New System.Drawing.Size(33, 19)
		Me._Text1_0.Location = New System.Drawing.Point(40, 16)
		Me._Text1_0.TabIndex = 3
		Me._Text1_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Text1_0.AcceptsReturn = True
		Me._Text1_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._Text1_0.BackColor = System.Drawing.SystemColors.Window
		Me._Text1_0.CausesValidation = True
		Me._Text1_0.Enabled = True
		Me._Text1_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._Text1_0.HideSelection = True
		Me._Text1_0.ReadOnly = False
		Me._Text1_0.Maxlength = 0
		Me._Text1_0.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._Text1_0.MultiLine = False
		Me._Text1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Text1_0.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._Text1_0.TabStop = True
		Me._Text1_0.Visible = True
		Me._Text1_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._Text1_0.Name = "_Text1_0"
		Me._Label1_2.Text = "range"
		Me._Label1_2.Size = New System.Drawing.Size(33, 17)
		Me._Label1_2.Location = New System.Drawing.Point(88, 16)
		Me._Label1_2.TabIndex = 19
		Me._Label1_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_2.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_2.Enabled = True
		Me._Label1_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_2.UseMnemonic = True
		Me._Label1_2.Visible = True
		Me._Label1_2.AutoSize = False
		Me._Label1_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_2.Name = "_Label1_2"
		Me._Label1_1.Text = "food"
		Me._Label1_1.Size = New System.Drawing.Size(25, 17)
		Me._Label1_1.Location = New System.Drawing.Point(8, 40)
		Me._Label1_1.TabIndex = 4
		Me._Label1_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_1.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_1.Enabled = True
		Me._Label1_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_1.UseMnemonic = True
		Me._Label1_1.Visible = True
		Me._Label1_1.AutoSize = False
		Me._Label1_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_1.Name = "_Label1_1"
		Me._Label1_0.Text = "mob"
		Me._Label1_0.Size = New System.Drawing.Size(33, 17)
		Me._Label1_0.Location = New System.Drawing.Point(8, 16)
		Me._Label1_0.TabIndex = 2
		Me._Label1_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label1_0.BackColor = System.Drawing.SystemColors.Control
		Me._Label1_0.Enabled = True
		Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label1_0.UseMnemonic = True
		Me._Label1_0.Visible = True
		Me._Label1_0.AutoSize = False
		Me._Label1_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label1_0.Name = "_Label1_0"
		grid1.OcxState = CType(resources.GetObject("grid1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.grid1.Size = New System.Drawing.Size(369, 161)
		Me.grid1.Location = New System.Drawing.Point(8, 80)
		Me.grid1.TabIndex = 13
		Me.grid1.Name = "grid1"
		Me.Label5.Text = "Percent"
		Me.Label5.Size = New System.Drawing.Size(49, 17)
		Me.Label5.Location = New System.Drawing.Point(256, 328)
		Me.Label5.TabIndex = 23
		Me.Label5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.Label4.Text = "Set Threshold higher than required by"
		Me.Label4.Size = New System.Drawing.Size(185, 17)
		Me.Label4.Location = New System.Drawing.Point(32, 328)
		Me.Label4.TabIndex = 21
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
		Me.lblError.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblError.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblError.Size = New System.Drawing.Size(129, 17)
		Me.lblError.Location = New System.Drawing.Point(8, 48)
		Me.lblError.TabIndex = 14
		Me.lblError.BackColor = System.Drawing.SystemColors.Control
		Me.lblError.Enabled = True
		Me.lblError.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblError.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblError.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblError.UseMnemonic = True
		Me.lblError.Visible = True
		Me.lblError.AutoSize = False
		Me.lblError.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblError.Name = "lblError"
		Me.Label3.Text = "Sector"
		Me.Label3.Size = New System.Drawing.Size(57, 17)
		Me.Label3.Location = New System.Drawing.Point(8, 16)
		Me.Label3.TabIndex = 11
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
		Me.Controls.Add(Text2)
		Me.Controls.Add(CheckSetThresh)
		Me.Controls.Add(Command1)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(txtOrigin)
		Me.Controls.Add(Frame3)
		Me.Controls.Add(Frame2)
		Me.Controls.Add(Frame1)
		Me.Controls.Add(grid1)
		Me.Controls.Add(Label5)
		Me.Controls.Add(Label4)
		Me.Controls.Add(lblError)
		Me.Controls.Add(Label3)
		Me.Frame3.Controls.Add(_Option1_1)
		Me.Frame3.Controls.Add(_Option1_0)
		Me.Frame2.Controls.Add(_Label2_1)
		Me.Frame2.Controls.Add(_Label2_0)
		Me.Frame1.Controls.Add(_Text1_2)
		Me.Frame1.Controls.Add(_Text1_1)
		Me.Frame1.Controls.Add(_Text1_0)
		Me.Frame1.Controls.Add(_Label1_2)
		Me.Frame1.Controls.Add(_Label1_1)
		Me.Frame1.Controls.Add(_Label1_0)
		Me.Label1.SetIndex(_Label1_2, CType(2, Short))
		Me.Label1.SetIndex(_Label1_1, CType(1, Short))
		Me.Label1.SetIndex(_Label1_0, CType(0, Short))
		Me.Label2.SetIndex(_Label2_1, CType(1, Short))
		Me.Label2.SetIndex(_Label2_0, CType(0, Short))
		Me.Option1.SetIndex(_Option1_1, CType(1, Short))
		Me.Option1.SetIndex(_Option1_0, CType(0, Short))
		Me.Text1.SetIndex(_Text1_2, CType(2, Short))
		Me.Text1.SetIndex(_Text1_1, CType(1, Short))
		Me.Text1.SetIndex(_Text1_0, CType(0, Short))
		CType(Me.Text1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Option1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Label2, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.grid1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Frame3.ResumeLayout(False)
		Me.Frame2.ResumeLayout(False)
		Me.Frame1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class