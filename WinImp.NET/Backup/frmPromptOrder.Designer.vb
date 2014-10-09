<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPromptOrder
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
	Public WithEvents txtAutoFeed As System.Windows.Forms.TextBox
	Public WithEvents chkAutoFeed As System.Windows.Forms.CheckBox
	Public WithEvents lblAutoFeed As System.Windows.Forms.Label
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents cmdShow As System.Windows.Forms.Button
	Public WithEvents cmdClear As System.Windows.Forms.Button
	Public WithEvents cmdSuspend As System.Windows.Forms.Button
	Public WithEvents cmdResume As System.Windows.Forms.Button
	Public WithEvents _txtEndCargo_2 As System.Windows.Forms.TextBox
	Public WithEvents _txtEndCargo_1 As System.Windows.Forms.TextBox
	Public WithEvents _txtEndCargo_3 As System.Windows.Forms.TextBox
	Public WithEvents _txtEndCargo_4 As System.Windows.Forms.TextBox
	Public WithEvents _txtEndCargo_6 As System.Windows.Forms.TextBox
	Public WithEvents _txtEndCargo_5 As System.Windows.Forms.TextBox
	Public WithEvents _txtStartCargo_5 As System.Windows.Forms.TextBox
	Public WithEvents _txtEnd_5 As System.Windows.Forms.TextBox
	Public WithEvents _txtStart_5 As System.Windows.Forms.TextBox
	Public WithEvents _txtStartCargo_6 As System.Windows.Forms.TextBox
	Public WithEvents _txtEnd_6 As System.Windows.Forms.TextBox
	Public WithEvents _txtStart_6 As System.Windows.Forms.TextBox
	Public WithEvents _txtStartCargo_4 As System.Windows.Forms.TextBox
	Public WithEvents _txtEnd_4 As System.Windows.Forms.TextBox
	Public WithEvents _txtStart_4 As System.Windows.Forms.TextBox
	Public WithEvents _txtStartCargo_3 As System.Windows.Forms.TextBox
	Public WithEvents _txtEnd_3 As System.Windows.Forms.TextBox
	Public WithEvents _txtStart_3 As System.Windows.Forms.TextBox
	Public WithEvents _txtStartCargo_1 As System.Windows.Forms.TextBox
	Public WithEvents _txtEnd_1 As System.Windows.Forms.TextBox
	Public WithEvents _txtStart_1 As System.Windows.Forms.TextBox
	Public WithEvents _txtStartCargo_2 As System.Windows.Forms.TextBox
	Public WithEvents _txtEnd_2 As System.Windows.Forms.TextBox
	Public WithEvents _txtStart_2 As System.Windows.Forms.TextBox
	Public WithEvents Label21 As System.Windows.Forms.Label
	Public WithEvents Label19 As System.Windows.Forms.Label
	Public WithEvents Label18 As System.Windows.Forms.Label
	Public WithEvents _lblHold_6 As System.Windows.Forms.Label
	Public WithEvents _lblHold_5 As System.Windows.Forms.Label
	Public WithEvents _lblHold_4 As System.Windows.Forms.Label
	Public WithEvents Label14 As System.Windows.Forms.Label
	Public WithEvents Label12 As System.Windows.Forms.Label
	Public WithEvents _lblHold_3 As System.Windows.Forms.Label
	Public WithEvents _lblHold_2 As System.Windows.Forms.Label
	Public WithEvents _lblHold_1 As System.Windows.Forms.Label
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents txtOrigin As System.Windows.Forms.TextBox
	Public WithEvents txtPath As System.Windows.Forms.TextBox
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents txtUnit As System.Windows.Forms.TextBox
	Public WithEvents cmdHelp As System.Windows.Forms.Button
	Public WithEvents chkAutoScuttle As System.Windows.Forms.CheckBox
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents lblHold As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents txtEnd As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	Public WithEvents txtEndCargo As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	Public WithEvents txtStart As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	Public WithEvents txtStartCargo As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPromptOrder))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Frame2 = New System.Windows.Forms.GroupBox
		Me.txtAutoFeed = New System.Windows.Forms.TextBox
		Me.chkAutoFeed = New System.Windows.Forms.CheckBox
		Me.lblAutoFeed = New System.Windows.Forms.Label
		Me.cmdShow = New System.Windows.Forms.Button
		Me.cmdClear = New System.Windows.Forms.Button
		Me.cmdSuspend = New System.Windows.Forms.Button
		Me.cmdResume = New System.Windows.Forms.Button
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me._txtEndCargo_2 = New System.Windows.Forms.TextBox
		Me._txtEndCargo_1 = New System.Windows.Forms.TextBox
		Me._txtEndCargo_3 = New System.Windows.Forms.TextBox
		Me._txtEndCargo_4 = New System.Windows.Forms.TextBox
		Me._txtEndCargo_6 = New System.Windows.Forms.TextBox
		Me._txtEndCargo_5 = New System.Windows.Forms.TextBox
		Me._txtStartCargo_5 = New System.Windows.Forms.TextBox
		Me._txtEnd_5 = New System.Windows.Forms.TextBox
		Me._txtStart_5 = New System.Windows.Forms.TextBox
		Me._txtStartCargo_6 = New System.Windows.Forms.TextBox
		Me._txtEnd_6 = New System.Windows.Forms.TextBox
		Me._txtStart_6 = New System.Windows.Forms.TextBox
		Me._txtStartCargo_4 = New System.Windows.Forms.TextBox
		Me._txtEnd_4 = New System.Windows.Forms.TextBox
		Me._txtStart_4 = New System.Windows.Forms.TextBox
		Me._txtStartCargo_3 = New System.Windows.Forms.TextBox
		Me._txtEnd_3 = New System.Windows.Forms.TextBox
		Me._txtStart_3 = New System.Windows.Forms.TextBox
		Me._txtStartCargo_1 = New System.Windows.Forms.TextBox
		Me._txtEnd_1 = New System.Windows.Forms.TextBox
		Me._txtStart_1 = New System.Windows.Forms.TextBox
		Me._txtStartCargo_2 = New System.Windows.Forms.TextBox
		Me._txtEnd_2 = New System.Windows.Forms.TextBox
		Me._txtStart_2 = New System.Windows.Forms.TextBox
		Me.Label21 = New System.Windows.Forms.Label
		Me.Label19 = New System.Windows.Forms.Label
		Me.Label18 = New System.Windows.Forms.Label
		Me._lblHold_6 = New System.Windows.Forms.Label
		Me._lblHold_5 = New System.Windows.Forms.Label
		Me._lblHold_4 = New System.Windows.Forms.Label
		Me.Label14 = New System.Windows.Forms.Label
		Me.Label12 = New System.Windows.Forms.Label
		Me._lblHold_3 = New System.Windows.Forms.Label
		Me._lblHold_2 = New System.Windows.Forms.Label
		Me._lblHold_1 = New System.Windows.Forms.Label
		Me.Label6 = New System.Windows.Forms.Label
		Me.txtOrigin = New System.Windows.Forms.TextBox
		Me.txtPath = New System.Windows.Forms.TextBox
		Me.cmdOK = New System.Windows.Forms.Button
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.txtUnit = New System.Windows.Forms.TextBox
		Me.cmdHelp = New System.Windows.Forms.Button
		Me.chkAutoScuttle = New System.Windows.Forms.CheckBox
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.lblHold = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.txtEnd = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(components)
		Me.txtEndCargo = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(components)
		Me.txtStart = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(components)
		Me.txtStartCargo = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(components)
		Me.Frame2.SuspendLayout()
		Me.Frame1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.lblHold, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.txtEnd, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.txtEndCargo, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.txtStart, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.txtStartCargo, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.ControlBox = False
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.ClientSize = New System.Drawing.Size(529, 176)
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
		Me.Name = "frmPromptOrder"
		Me.Frame2.Text = "Autofeed (fb, ft only)"
		Me.Frame2.Size = New System.Drawing.Size(121, 57)
		Me.Frame2.Location = New System.Drawing.Point(8, 88)
		Me.Frame2.TabIndex = 52
		Me.Frame2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame2.BackColor = System.Drawing.SystemColors.Control
		Me.Frame2.Enabled = True
		Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame2.Visible = True
		Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame2.Name = "Frame2"
		Me.txtAutoFeed.AutoSize = False
		Me.txtAutoFeed.Size = New System.Drawing.Size(41, 19)
		Me.txtAutoFeed.Location = New System.Drawing.Point(72, 32)
		Me.txtAutoFeed.TabIndex = 4
		Me.txtAutoFeed.Text = "30"
		Me.txtAutoFeed.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtAutoFeed.AcceptsReturn = True
		Me.txtAutoFeed.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtAutoFeed.BackColor = System.Drawing.SystemColors.Window
		Me.txtAutoFeed.CausesValidation = True
		Me.txtAutoFeed.Enabled = True
		Me.txtAutoFeed.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtAutoFeed.HideSelection = True
		Me.txtAutoFeed.ReadOnly = False
		Me.txtAutoFeed.Maxlength = 0
		Me.txtAutoFeed.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtAutoFeed.MultiLine = False
		Me.txtAutoFeed.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtAutoFeed.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtAutoFeed.TabStop = True
		Me.txtAutoFeed.Visible = True
		Me.txtAutoFeed.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtAutoFeed.Name = "txtAutoFeed"
		Me.chkAutoFeed.Text = "Autofeed"
		Me.chkAutoFeed.Size = New System.Drawing.Size(65, 17)
		Me.chkAutoFeed.Location = New System.Drawing.Point(8, 16)
		Me.chkAutoFeed.TabIndex = 3
		Me.chkAutoFeed.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkAutoFeed.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkAutoFeed.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkAutoFeed.BackColor = System.Drawing.SystemColors.Control
		Me.chkAutoFeed.CausesValidation = True
		Me.chkAutoFeed.Enabled = True
		Me.chkAutoFeed.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkAutoFeed.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkAutoFeed.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkAutoFeed.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkAutoFeed.TabStop = True
		Me.chkAutoFeed.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkAutoFeed.Visible = True
		Me.chkAutoFeed.Name = "chkAutoFeed"
		Me.lblAutoFeed.Text = "Threshold"
		Me.lblAutoFeed.Size = New System.Drawing.Size(49, 17)
		Me.lblAutoFeed.Location = New System.Drawing.Point(16, 32)
		Me.lblAutoFeed.TabIndex = 53
		Me.lblAutoFeed.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblAutoFeed.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblAutoFeed.BackColor = System.Drawing.SystemColors.Control
		Me.lblAutoFeed.Enabled = True
		Me.lblAutoFeed.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblAutoFeed.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblAutoFeed.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblAutoFeed.UseMnemonic = True
		Me.lblAutoFeed.Visible = True
		Me.lblAutoFeed.AutoSize = False
		Me.lblAutoFeed.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblAutoFeed.Name = "lblAutoFeed"
		Me.cmdShow.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdShow.Text = "Show"
		Me.cmdShow.Size = New System.Drawing.Size(57, 25)
		Me.cmdShow.Location = New System.Drawing.Point(197, 144)
		Me.cmdShow.TabIndex = 31
		Me.cmdShow.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdShow.BackColor = System.Drawing.SystemColors.Control
		Me.cmdShow.CausesValidation = True
		Me.cmdShow.Enabled = True
		Me.cmdShow.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdShow.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdShow.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdShow.TabStop = True
		Me.cmdShow.Name = "cmdShow"
		Me.cmdClear.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdClear.Text = "Clear"
		Me.cmdClear.Size = New System.Drawing.Size(57, 25)
		Me.cmdClear.Location = New System.Drawing.Point(380, 144)
		Me.cmdClear.TabIndex = 34
		Me.cmdClear.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdClear.BackColor = System.Drawing.SystemColors.Control
		Me.cmdClear.CausesValidation = True
		Me.cmdClear.Enabled = True
		Me.cmdClear.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdClear.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdClear.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdClear.TabStop = True
		Me.cmdClear.Name = "cmdClear"
		Me.cmdSuspend.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdSuspend.Text = "Suspend"
		Me.cmdSuspend.Size = New System.Drawing.Size(57, 25)
		Me.cmdSuspend.Location = New System.Drawing.Point(319, 144)
		Me.cmdSuspend.TabIndex = 33
		Me.cmdSuspend.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdSuspend.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSuspend.CausesValidation = True
		Me.cmdSuspend.Enabled = True
		Me.cmdSuspend.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSuspend.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSuspend.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSuspend.TabStop = True
		Me.cmdSuspend.Name = "cmdSuspend"
		Me.cmdResume.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdResume.Text = "Resume"
		Me.cmdResume.Size = New System.Drawing.Size(57, 25)
		Me.cmdResume.Location = New System.Drawing.Point(258, 144)
		Me.cmdResume.TabIndex = 32
		Me.cmdResume.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdResume.BackColor = System.Drawing.SystemColors.Control
		Me.cmdResume.CausesValidation = True
		Me.cmdResume.Enabled = True
		Me.cmdResume.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdResume.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdResume.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdResume.TabStop = True
		Me.cmdResume.Name = "cmdResume"
		Me.Frame1.Text = "Cargo Holds"
		Me.Frame1.Size = New System.Drawing.Size(361, 129)
		Me.Frame1.Location = New System.Drawing.Point(136, 8)
		Me.Frame1.TabIndex = 36
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me._txtEndCargo_2.AutoSize = False
		Me._txtEndCargo_2.Size = New System.Drawing.Size(33, 19)
		Me._txtEndCargo_2.Location = New System.Drawing.Point(152, 80)
		Me._txtEndCargo_2.TabIndex = 12
		Me._txtEndCargo_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtEndCargo_2.AcceptsReturn = True
		Me._txtEndCargo_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtEndCargo_2.BackColor = System.Drawing.SystemColors.Window
		Me._txtEndCargo_2.CausesValidation = True
		Me._txtEndCargo_2.Enabled = True
		Me._txtEndCargo_2.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtEndCargo_2.HideSelection = True
		Me._txtEndCargo_2.ReadOnly = False
		Me._txtEndCargo_2.Maxlength = 0
		Me._txtEndCargo_2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtEndCargo_2.MultiLine = False
		Me._txtEndCargo_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtEndCargo_2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtEndCargo_2.TabStop = True
		Me._txtEndCargo_2.Visible = True
		Me._txtEndCargo_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtEndCargo_2.Name = "_txtEndCargo_2"
		Me._txtEndCargo_1.AutoSize = False
		Me._txtEndCargo_1.Size = New System.Drawing.Size(33, 19)
		Me._txtEndCargo_1.Location = New System.Drawing.Point(112, 80)
		Me._txtEndCargo_1.TabIndex = 8
		Me._txtEndCargo_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtEndCargo_1.AcceptsReturn = True
		Me._txtEndCargo_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtEndCargo_1.BackColor = System.Drawing.SystemColors.Window
		Me._txtEndCargo_1.CausesValidation = True
		Me._txtEndCargo_1.Enabled = True
		Me._txtEndCargo_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtEndCargo_1.HideSelection = True
		Me._txtEndCargo_1.ReadOnly = False
		Me._txtEndCargo_1.Maxlength = 0
		Me._txtEndCargo_1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtEndCargo_1.MultiLine = False
		Me._txtEndCargo_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtEndCargo_1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtEndCargo_1.TabStop = True
		Me._txtEndCargo_1.Visible = True
		Me._txtEndCargo_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtEndCargo_1.Name = "_txtEndCargo_1"
		Me._txtEndCargo_3.AutoSize = False
		Me._txtEndCargo_3.Size = New System.Drawing.Size(33, 19)
		Me._txtEndCargo_3.Location = New System.Drawing.Point(192, 80)
		Me._txtEndCargo_3.TabIndex = 16
		Me._txtEndCargo_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtEndCargo_3.AcceptsReturn = True
		Me._txtEndCargo_3.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtEndCargo_3.BackColor = System.Drawing.SystemColors.Window
		Me._txtEndCargo_3.CausesValidation = True
		Me._txtEndCargo_3.Enabled = True
		Me._txtEndCargo_3.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtEndCargo_3.HideSelection = True
		Me._txtEndCargo_3.ReadOnly = False
		Me._txtEndCargo_3.Maxlength = 0
		Me._txtEndCargo_3.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtEndCargo_3.MultiLine = False
		Me._txtEndCargo_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtEndCargo_3.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtEndCargo_3.TabStop = True
		Me._txtEndCargo_3.Visible = True
		Me._txtEndCargo_3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtEndCargo_3.Name = "_txtEndCargo_3"
		Me._txtEndCargo_4.AutoSize = False
		Me._txtEndCargo_4.Size = New System.Drawing.Size(33, 19)
		Me._txtEndCargo_4.Location = New System.Drawing.Point(232, 80)
		Me._txtEndCargo_4.TabIndex = 20
		Me._txtEndCargo_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtEndCargo_4.AcceptsReturn = True
		Me._txtEndCargo_4.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtEndCargo_4.BackColor = System.Drawing.SystemColors.Window
		Me._txtEndCargo_4.CausesValidation = True
		Me._txtEndCargo_4.Enabled = True
		Me._txtEndCargo_4.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtEndCargo_4.HideSelection = True
		Me._txtEndCargo_4.ReadOnly = False
		Me._txtEndCargo_4.Maxlength = 0
		Me._txtEndCargo_4.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtEndCargo_4.MultiLine = False
		Me._txtEndCargo_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtEndCargo_4.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtEndCargo_4.TabStop = True
		Me._txtEndCargo_4.Visible = True
		Me._txtEndCargo_4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtEndCargo_4.Name = "_txtEndCargo_4"
		Me._txtEndCargo_6.AutoSize = False
		Me._txtEndCargo_6.Size = New System.Drawing.Size(33, 19)
		Me._txtEndCargo_6.Location = New System.Drawing.Point(312, 80)
		Me._txtEndCargo_6.TabIndex = 28
		Me._txtEndCargo_6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtEndCargo_6.AcceptsReturn = True
		Me._txtEndCargo_6.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtEndCargo_6.BackColor = System.Drawing.SystemColors.Window
		Me._txtEndCargo_6.CausesValidation = True
		Me._txtEndCargo_6.Enabled = True
		Me._txtEndCargo_6.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtEndCargo_6.HideSelection = True
		Me._txtEndCargo_6.ReadOnly = False
		Me._txtEndCargo_6.Maxlength = 0
		Me._txtEndCargo_6.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtEndCargo_6.MultiLine = False
		Me._txtEndCargo_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtEndCargo_6.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtEndCargo_6.TabStop = True
		Me._txtEndCargo_6.Visible = True
		Me._txtEndCargo_6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtEndCargo_6.Name = "_txtEndCargo_6"
		Me._txtEndCargo_5.AutoSize = False
		Me._txtEndCargo_5.Size = New System.Drawing.Size(33, 19)
		Me._txtEndCargo_5.Location = New System.Drawing.Point(272, 80)
		Me._txtEndCargo_5.TabIndex = 24
		Me._txtEndCargo_5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtEndCargo_5.AcceptsReturn = True
		Me._txtEndCargo_5.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtEndCargo_5.BackColor = System.Drawing.SystemColors.Window
		Me._txtEndCargo_5.CausesValidation = True
		Me._txtEndCargo_5.Enabled = True
		Me._txtEndCargo_5.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtEndCargo_5.HideSelection = True
		Me._txtEndCargo_5.ReadOnly = False
		Me._txtEndCargo_5.Maxlength = 0
		Me._txtEndCargo_5.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtEndCargo_5.MultiLine = False
		Me._txtEndCargo_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtEndCargo_5.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtEndCargo_5.TabStop = True
		Me._txtEndCargo_5.Visible = True
		Me._txtEndCargo_5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtEndCargo_5.Name = "_txtEndCargo_5"
		Me._txtStartCargo_5.AutoSize = False
		Me._txtStartCargo_5.Size = New System.Drawing.Size(33, 19)
		Me._txtStartCargo_5.Location = New System.Drawing.Point(272, 32)
		Me._txtStartCargo_5.TabIndex = 22
		Me._txtStartCargo_5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtStartCargo_5.AcceptsReturn = True
		Me._txtStartCargo_5.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtStartCargo_5.BackColor = System.Drawing.SystemColors.Window
		Me._txtStartCargo_5.CausesValidation = True
		Me._txtStartCargo_5.Enabled = True
		Me._txtStartCargo_5.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtStartCargo_5.HideSelection = True
		Me._txtStartCargo_5.ReadOnly = False
		Me._txtStartCargo_5.Maxlength = 0
		Me._txtStartCargo_5.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtStartCargo_5.MultiLine = False
		Me._txtStartCargo_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtStartCargo_5.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtStartCargo_5.TabStop = True
		Me._txtStartCargo_5.Visible = True
		Me._txtStartCargo_5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtStartCargo_5.Name = "_txtStartCargo_5"
		Me._txtEnd_5.AutoSize = False
		Me._txtEnd_5.Size = New System.Drawing.Size(33, 19)
		Me._txtEnd_5.Location = New System.Drawing.Point(272, 104)
		Me._txtEnd_5.TabIndex = 25
		Me._txtEnd_5.Text = "0"
		Me._txtEnd_5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtEnd_5.AcceptsReturn = True
		Me._txtEnd_5.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtEnd_5.BackColor = System.Drawing.SystemColors.Window
		Me._txtEnd_5.CausesValidation = True
		Me._txtEnd_5.Enabled = True
		Me._txtEnd_5.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtEnd_5.HideSelection = True
		Me._txtEnd_5.ReadOnly = False
		Me._txtEnd_5.Maxlength = 0
		Me._txtEnd_5.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtEnd_5.MultiLine = False
		Me._txtEnd_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtEnd_5.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtEnd_5.TabStop = True
		Me._txtEnd_5.Visible = True
		Me._txtEnd_5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtEnd_5.Name = "_txtEnd_5"
		Me._txtStart_5.AutoSize = False
		Me._txtStart_5.Size = New System.Drawing.Size(33, 19)
		Me._txtStart_5.Location = New System.Drawing.Point(272, 56)
		Me._txtStart_5.TabIndex = 23
		Me._txtStart_5.Text = "0"
		Me._txtStart_5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtStart_5.AcceptsReturn = True
		Me._txtStart_5.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtStart_5.BackColor = System.Drawing.SystemColors.Window
		Me._txtStart_5.CausesValidation = True
		Me._txtStart_5.Enabled = True
		Me._txtStart_5.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtStart_5.HideSelection = True
		Me._txtStart_5.ReadOnly = False
		Me._txtStart_5.Maxlength = 0
		Me._txtStart_5.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtStart_5.MultiLine = False
		Me._txtStart_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtStart_5.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtStart_5.TabStop = True
		Me._txtStart_5.Visible = True
		Me._txtStart_5.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtStart_5.Name = "_txtStart_5"
		Me._txtStartCargo_6.AutoSize = False
		Me._txtStartCargo_6.Size = New System.Drawing.Size(33, 19)
		Me._txtStartCargo_6.Location = New System.Drawing.Point(312, 32)
		Me._txtStartCargo_6.TabIndex = 26
		Me._txtStartCargo_6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtStartCargo_6.AcceptsReturn = True
		Me._txtStartCargo_6.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtStartCargo_6.BackColor = System.Drawing.SystemColors.Window
		Me._txtStartCargo_6.CausesValidation = True
		Me._txtStartCargo_6.Enabled = True
		Me._txtStartCargo_6.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtStartCargo_6.HideSelection = True
		Me._txtStartCargo_6.ReadOnly = False
		Me._txtStartCargo_6.Maxlength = 0
		Me._txtStartCargo_6.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtStartCargo_6.MultiLine = False
		Me._txtStartCargo_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtStartCargo_6.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtStartCargo_6.TabStop = True
		Me._txtStartCargo_6.Visible = True
		Me._txtStartCargo_6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtStartCargo_6.Name = "_txtStartCargo_6"
		Me._txtEnd_6.AutoSize = False
		Me._txtEnd_6.Size = New System.Drawing.Size(33, 19)
		Me._txtEnd_6.Location = New System.Drawing.Point(312, 104)
		Me._txtEnd_6.TabIndex = 29
		Me._txtEnd_6.Text = "0"
		Me._txtEnd_6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtEnd_6.AcceptsReturn = True
		Me._txtEnd_6.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtEnd_6.BackColor = System.Drawing.SystemColors.Window
		Me._txtEnd_6.CausesValidation = True
		Me._txtEnd_6.Enabled = True
		Me._txtEnd_6.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtEnd_6.HideSelection = True
		Me._txtEnd_6.ReadOnly = False
		Me._txtEnd_6.Maxlength = 0
		Me._txtEnd_6.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtEnd_6.MultiLine = False
		Me._txtEnd_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtEnd_6.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtEnd_6.TabStop = True
		Me._txtEnd_6.Visible = True
		Me._txtEnd_6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtEnd_6.Name = "_txtEnd_6"
		Me._txtStart_6.AutoSize = False
		Me._txtStart_6.Size = New System.Drawing.Size(33, 19)
		Me._txtStart_6.Location = New System.Drawing.Point(312, 56)
		Me._txtStart_6.TabIndex = 27
		Me._txtStart_6.Text = "0"
		Me._txtStart_6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtStart_6.AcceptsReturn = True
		Me._txtStart_6.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtStart_6.BackColor = System.Drawing.SystemColors.Window
		Me._txtStart_6.CausesValidation = True
		Me._txtStart_6.Enabled = True
		Me._txtStart_6.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtStart_6.HideSelection = True
		Me._txtStart_6.ReadOnly = False
		Me._txtStart_6.Maxlength = 0
		Me._txtStart_6.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtStart_6.MultiLine = False
		Me._txtStart_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtStart_6.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtStart_6.TabStop = True
		Me._txtStart_6.Visible = True
		Me._txtStart_6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtStart_6.Name = "_txtStart_6"
		Me._txtStartCargo_4.AutoSize = False
		Me._txtStartCargo_4.Size = New System.Drawing.Size(33, 19)
		Me._txtStartCargo_4.Location = New System.Drawing.Point(232, 32)
		Me._txtStartCargo_4.TabIndex = 18
		Me._txtStartCargo_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtStartCargo_4.AcceptsReturn = True
		Me._txtStartCargo_4.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtStartCargo_4.BackColor = System.Drawing.SystemColors.Window
		Me._txtStartCargo_4.CausesValidation = True
		Me._txtStartCargo_4.Enabled = True
		Me._txtStartCargo_4.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtStartCargo_4.HideSelection = True
		Me._txtStartCargo_4.ReadOnly = False
		Me._txtStartCargo_4.Maxlength = 0
		Me._txtStartCargo_4.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtStartCargo_4.MultiLine = False
		Me._txtStartCargo_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtStartCargo_4.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtStartCargo_4.TabStop = True
		Me._txtStartCargo_4.Visible = True
		Me._txtStartCargo_4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtStartCargo_4.Name = "_txtStartCargo_4"
		Me._txtEnd_4.AutoSize = False
		Me._txtEnd_4.Size = New System.Drawing.Size(33, 19)
		Me._txtEnd_4.Location = New System.Drawing.Point(232, 104)
		Me._txtEnd_4.TabIndex = 21
		Me._txtEnd_4.Text = "0"
		Me._txtEnd_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtEnd_4.AcceptsReturn = True
		Me._txtEnd_4.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtEnd_4.BackColor = System.Drawing.SystemColors.Window
		Me._txtEnd_4.CausesValidation = True
		Me._txtEnd_4.Enabled = True
		Me._txtEnd_4.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtEnd_4.HideSelection = True
		Me._txtEnd_4.ReadOnly = False
		Me._txtEnd_4.Maxlength = 0
		Me._txtEnd_4.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtEnd_4.MultiLine = False
		Me._txtEnd_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtEnd_4.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtEnd_4.TabStop = True
		Me._txtEnd_4.Visible = True
		Me._txtEnd_4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtEnd_4.Name = "_txtEnd_4"
		Me._txtStart_4.AutoSize = False
		Me._txtStart_4.Size = New System.Drawing.Size(33, 19)
		Me._txtStart_4.Location = New System.Drawing.Point(232, 56)
		Me._txtStart_4.TabIndex = 19
		Me._txtStart_4.Text = "0"
		Me._txtStart_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtStart_4.AcceptsReturn = True
		Me._txtStart_4.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtStart_4.BackColor = System.Drawing.SystemColors.Window
		Me._txtStart_4.CausesValidation = True
		Me._txtStart_4.Enabled = True
		Me._txtStart_4.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtStart_4.HideSelection = True
		Me._txtStart_4.ReadOnly = False
		Me._txtStart_4.Maxlength = 0
		Me._txtStart_4.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtStart_4.MultiLine = False
		Me._txtStart_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtStart_4.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtStart_4.TabStop = True
		Me._txtStart_4.Visible = True
		Me._txtStart_4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtStart_4.Name = "_txtStart_4"
		Me._txtStartCargo_3.AutoSize = False
		Me._txtStartCargo_3.Size = New System.Drawing.Size(33, 19)
		Me._txtStartCargo_3.Location = New System.Drawing.Point(192, 32)
		Me._txtStartCargo_3.TabIndex = 14
		Me._txtStartCargo_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtStartCargo_3.AcceptsReturn = True
		Me._txtStartCargo_3.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtStartCargo_3.BackColor = System.Drawing.SystemColors.Window
		Me._txtStartCargo_3.CausesValidation = True
		Me._txtStartCargo_3.Enabled = True
		Me._txtStartCargo_3.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtStartCargo_3.HideSelection = True
		Me._txtStartCargo_3.ReadOnly = False
		Me._txtStartCargo_3.Maxlength = 0
		Me._txtStartCargo_3.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtStartCargo_3.MultiLine = False
		Me._txtStartCargo_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtStartCargo_3.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtStartCargo_3.TabStop = True
		Me._txtStartCargo_3.Visible = True
		Me._txtStartCargo_3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtStartCargo_3.Name = "_txtStartCargo_3"
		Me._txtEnd_3.AutoSize = False
		Me._txtEnd_3.Size = New System.Drawing.Size(33, 19)
		Me._txtEnd_3.Location = New System.Drawing.Point(192, 104)
		Me._txtEnd_3.TabIndex = 17
		Me._txtEnd_3.Text = "0"
		Me._txtEnd_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtEnd_3.AcceptsReturn = True
		Me._txtEnd_3.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtEnd_3.BackColor = System.Drawing.SystemColors.Window
		Me._txtEnd_3.CausesValidation = True
		Me._txtEnd_3.Enabled = True
		Me._txtEnd_3.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtEnd_3.HideSelection = True
		Me._txtEnd_3.ReadOnly = False
		Me._txtEnd_3.Maxlength = 0
		Me._txtEnd_3.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtEnd_3.MultiLine = False
		Me._txtEnd_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtEnd_3.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtEnd_3.TabStop = True
		Me._txtEnd_3.Visible = True
		Me._txtEnd_3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtEnd_3.Name = "_txtEnd_3"
		Me._txtStart_3.AutoSize = False
		Me._txtStart_3.Size = New System.Drawing.Size(33, 19)
		Me._txtStart_3.Location = New System.Drawing.Point(192, 56)
		Me._txtStart_3.TabIndex = 15
		Me._txtStart_3.Text = "0"
		Me._txtStart_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtStart_3.AcceptsReturn = True
		Me._txtStart_3.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtStart_3.BackColor = System.Drawing.SystemColors.Window
		Me._txtStart_3.CausesValidation = True
		Me._txtStart_3.Enabled = True
		Me._txtStart_3.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtStart_3.HideSelection = True
		Me._txtStart_3.ReadOnly = False
		Me._txtStart_3.Maxlength = 0
		Me._txtStart_3.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtStart_3.MultiLine = False
		Me._txtStart_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtStart_3.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtStart_3.TabStop = True
		Me._txtStart_3.Visible = True
		Me._txtStart_3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtStart_3.Name = "_txtStart_3"
		Me._txtStartCargo_1.AutoSize = False
		Me._txtStartCargo_1.Size = New System.Drawing.Size(33, 19)
		Me._txtStartCargo_1.Location = New System.Drawing.Point(112, 32)
		Me._txtStartCargo_1.TabIndex = 6
		Me._txtStartCargo_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtStartCargo_1.AcceptsReturn = True
		Me._txtStartCargo_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtStartCargo_1.BackColor = System.Drawing.SystemColors.Window
		Me._txtStartCargo_1.CausesValidation = True
		Me._txtStartCargo_1.Enabled = True
		Me._txtStartCargo_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtStartCargo_1.HideSelection = True
		Me._txtStartCargo_1.ReadOnly = False
		Me._txtStartCargo_1.Maxlength = 0
		Me._txtStartCargo_1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtStartCargo_1.MultiLine = False
		Me._txtStartCargo_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtStartCargo_1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtStartCargo_1.TabStop = True
		Me._txtStartCargo_1.Visible = True
		Me._txtStartCargo_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtStartCargo_1.Name = "_txtStartCargo_1"
		Me._txtEnd_1.AutoSize = False
		Me._txtEnd_1.Size = New System.Drawing.Size(33, 19)
		Me._txtEnd_1.Location = New System.Drawing.Point(112, 104)
		Me._txtEnd_1.TabIndex = 9
		Me._txtEnd_1.Text = "0"
		Me._txtEnd_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtEnd_1.AcceptsReturn = True
		Me._txtEnd_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtEnd_1.BackColor = System.Drawing.SystemColors.Window
		Me._txtEnd_1.CausesValidation = True
		Me._txtEnd_1.Enabled = True
		Me._txtEnd_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtEnd_1.HideSelection = True
		Me._txtEnd_1.ReadOnly = False
		Me._txtEnd_1.Maxlength = 0
		Me._txtEnd_1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtEnd_1.MultiLine = False
		Me._txtEnd_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtEnd_1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtEnd_1.TabStop = True
		Me._txtEnd_1.Visible = True
		Me._txtEnd_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtEnd_1.Name = "_txtEnd_1"
		Me._txtStart_1.AutoSize = False
		Me._txtStart_1.Size = New System.Drawing.Size(33, 19)
		Me._txtStart_1.Location = New System.Drawing.Point(112, 56)
		Me._txtStart_1.TabIndex = 7
		Me._txtStart_1.Text = "0"
		Me._txtStart_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtStart_1.AcceptsReturn = True
		Me._txtStart_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtStart_1.BackColor = System.Drawing.SystemColors.Window
		Me._txtStart_1.CausesValidation = True
		Me._txtStart_1.Enabled = True
		Me._txtStart_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtStart_1.HideSelection = True
		Me._txtStart_1.ReadOnly = False
		Me._txtStart_1.Maxlength = 0
		Me._txtStart_1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtStart_1.MultiLine = False
		Me._txtStart_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtStart_1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtStart_1.TabStop = True
		Me._txtStart_1.Visible = True
		Me._txtStart_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtStart_1.Name = "_txtStart_1"
		Me._txtStartCargo_2.AutoSize = False
		Me._txtStartCargo_2.Size = New System.Drawing.Size(33, 19)
		Me._txtStartCargo_2.Location = New System.Drawing.Point(152, 32)
		Me._txtStartCargo_2.TabIndex = 10
		Me._txtStartCargo_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtStartCargo_2.AcceptsReturn = True
		Me._txtStartCargo_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtStartCargo_2.BackColor = System.Drawing.SystemColors.Window
		Me._txtStartCargo_2.CausesValidation = True
		Me._txtStartCargo_2.Enabled = True
		Me._txtStartCargo_2.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtStartCargo_2.HideSelection = True
		Me._txtStartCargo_2.ReadOnly = False
		Me._txtStartCargo_2.Maxlength = 0
		Me._txtStartCargo_2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtStartCargo_2.MultiLine = False
		Me._txtStartCargo_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtStartCargo_2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtStartCargo_2.TabStop = True
		Me._txtStartCargo_2.Visible = True
		Me._txtStartCargo_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtStartCargo_2.Name = "_txtStartCargo_2"
		Me._txtEnd_2.AutoSize = False
		Me._txtEnd_2.Size = New System.Drawing.Size(33, 19)
		Me._txtEnd_2.Location = New System.Drawing.Point(152, 104)
		Me._txtEnd_2.TabIndex = 13
		Me._txtEnd_2.Text = "0"
		Me._txtEnd_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtEnd_2.AcceptsReturn = True
		Me._txtEnd_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtEnd_2.BackColor = System.Drawing.SystemColors.Window
		Me._txtEnd_2.CausesValidation = True
		Me._txtEnd_2.Enabled = True
		Me._txtEnd_2.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtEnd_2.HideSelection = True
		Me._txtEnd_2.ReadOnly = False
		Me._txtEnd_2.Maxlength = 0
		Me._txtEnd_2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtEnd_2.MultiLine = False
		Me._txtEnd_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtEnd_2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtEnd_2.TabStop = True
		Me._txtEnd_2.Visible = True
		Me._txtEnd_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtEnd_2.Name = "_txtEnd_2"
		Me._txtStart_2.AutoSize = False
		Me._txtStart_2.Size = New System.Drawing.Size(33, 19)
		Me._txtStart_2.Location = New System.Drawing.Point(152, 56)
		Me._txtStart_2.TabIndex = 11
		Me._txtStart_2.Text = "0"
		Me._txtStart_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtStart_2.AcceptsReturn = True
		Me._txtStart_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtStart_2.BackColor = System.Drawing.SystemColors.Window
		Me._txtStart_2.CausesValidation = True
		Me._txtStart_2.Enabled = True
		Me._txtStart_2.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtStart_2.HideSelection = True
		Me._txtStart_2.ReadOnly = False
		Me._txtStart_2.Maxlength = 0
		Me._txtStart_2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtStart_2.MultiLine = False
		Me._txtStart_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtStart_2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtStart_2.TabStop = True
		Me._txtStart_2.Visible = True
		Me._txtStart_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtStart_2.Name = "_txtStart_2"
		Me.Label21.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label21.Text = "type"
		Me.Label21.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label21.Size = New System.Drawing.Size(33, 17)
		Me.Label21.Location = New System.Drawing.Point(72, 80)
		Me.Label21.TabIndex = 54
		Me.Label21.BackColor = System.Drawing.SystemColors.Control
		Me.Label21.Enabled = True
		Me.Label21.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label21.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label21.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label21.UseMnemonic = True
		Me.Label21.Visible = True
		Me.Label21.AutoSize = False
		Me.Label21.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label21.Name = "Label21"
		Me.Label19.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label19.Text = "level"
		Me.Label19.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label19.Size = New System.Drawing.Size(33, 17)
		Me.Label19.Location = New System.Drawing.Point(72, 56)
		Me.Label19.TabIndex = 48
		Me.Label19.BackColor = System.Drawing.SystemColors.Control
		Me.Label19.Enabled = True
		Me.Label19.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label19.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label19.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label19.UseMnemonic = True
		Me.Label19.Visible = True
		Me.Label19.AutoSize = False
		Me.Label19.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label19.Name = "Label19"
		Me.Label18.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label18.Text = "level"
		Me.Label18.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label18.Size = New System.Drawing.Size(33, 17)
		Me.Label18.Location = New System.Drawing.Point(72, 104)
		Me.Label18.TabIndex = 47
		Me.Label18.BackColor = System.Drawing.SystemColors.Control
		Me.Label18.Enabled = True
		Me.Label18.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label18.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label18.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label18.UseMnemonic = True
		Me.Label18.Visible = True
		Me.Label18.AutoSize = False
		Me.Label18.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label18.Name = "Label18"
		Me._lblHold_6.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._lblHold_6.Text = "6"
		Me._lblHold_6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblHold_6.Size = New System.Drawing.Size(17, 17)
		Me._lblHold_6.Location = New System.Drawing.Point(320, 16)
		Me._lblHold_6.TabIndex = 46
		Me._lblHold_6.BackColor = System.Drawing.SystemColors.Control
		Me._lblHold_6.Enabled = True
		Me._lblHold_6.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblHold_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblHold_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblHold_6.UseMnemonic = True
		Me._lblHold_6.Visible = True
		Me._lblHold_6.AutoSize = False
		Me._lblHold_6.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblHold_6.Name = "_lblHold_6"
		Me._lblHold_5.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._lblHold_5.Text = "5"
		Me._lblHold_5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblHold_5.Size = New System.Drawing.Size(17, 17)
		Me._lblHold_5.Location = New System.Drawing.Point(280, 16)
		Me._lblHold_5.TabIndex = 45
		Me._lblHold_5.BackColor = System.Drawing.SystemColors.Control
		Me._lblHold_5.Enabled = True
		Me._lblHold_5.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblHold_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblHold_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblHold_5.UseMnemonic = True
		Me._lblHold_5.Visible = True
		Me._lblHold_5.AutoSize = False
		Me._lblHold_5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblHold_5.Name = "_lblHold_5"
		Me._lblHold_4.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._lblHold_4.Text = "4"
		Me._lblHold_4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblHold_4.Size = New System.Drawing.Size(17, 17)
		Me._lblHold_4.Location = New System.Drawing.Point(240, 16)
		Me._lblHold_4.TabIndex = 44
		Me._lblHold_4.BackColor = System.Drawing.SystemColors.Control
		Me._lblHold_4.Enabled = True
		Me._lblHold_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblHold_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblHold_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblHold_4.UseMnemonic = True
		Me._lblHold_4.Visible = True
		Me._lblHold_4.AutoSize = False
		Me._lblHold_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblHold_4.Name = "_lblHold_4"
		Me.Label14.Text = "end cargo"
		Me.Label14.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label14.Size = New System.Drawing.Size(65, 17)
		Me.Label14.Location = New System.Drawing.Point(8, 88)
		Me.Label14.TabIndex = 43
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
		Me.Label12.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label12.Text = "type"
		Me.Label12.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label12.Size = New System.Drawing.Size(33, 17)
		Me.Label12.Location = New System.Drawing.Point(72, 32)
		Me.Label12.TabIndex = 42
		Me.Label12.BackColor = System.Drawing.SystemColors.Control
		Me.Label12.Enabled = True
		Me.Label12.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label12.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label12.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label12.UseMnemonic = True
		Me.Label12.Visible = True
		Me.Label12.AutoSize = False
		Me.Label12.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label12.Name = "Label12"
		Me._lblHold_3.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._lblHold_3.Text = "3"
		Me._lblHold_3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblHold_3.Size = New System.Drawing.Size(17, 17)
		Me._lblHold_3.Location = New System.Drawing.Point(200, 16)
		Me._lblHold_3.TabIndex = 41
		Me._lblHold_3.BackColor = System.Drawing.SystemColors.Control
		Me._lblHold_3.Enabled = True
		Me._lblHold_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblHold_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblHold_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblHold_3.UseMnemonic = True
		Me._lblHold_3.Visible = True
		Me._lblHold_3.AutoSize = False
		Me._lblHold_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblHold_3.Name = "_lblHold_3"
		Me._lblHold_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._lblHold_2.Text = "2"
		Me._lblHold_2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblHold_2.Size = New System.Drawing.Size(17, 17)
		Me._lblHold_2.Location = New System.Drawing.Point(160, 16)
		Me._lblHold_2.TabIndex = 40
		Me._lblHold_2.BackColor = System.Drawing.SystemColors.Control
		Me._lblHold_2.Enabled = True
		Me._lblHold_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblHold_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblHold_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblHold_2.UseMnemonic = True
		Me._lblHold_2.Visible = True
		Me._lblHold_2.AutoSize = False
		Me._lblHold_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblHold_2.Name = "_lblHold_2"
		Me._lblHold_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._lblHold_1.Text = "1"
		Me._lblHold_1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblHold_1.Size = New System.Drawing.Size(17, 17)
		Me._lblHold_1.Location = New System.Drawing.Point(120, 16)
		Me._lblHold_1.TabIndex = 39
		Me._lblHold_1.BackColor = System.Drawing.SystemColors.Control
		Me._lblHold_1.Enabled = True
		Me._lblHold_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblHold_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblHold_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblHold_1.UseMnemonic = True
		Me._lblHold_1.Visible = True
		Me._lblHold_1.AutoSize = False
		Me._lblHold_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblHold_1.Name = "_lblHold_1"
		Me.Label6.Text = "start cargo"
		Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label6.Size = New System.Drawing.Size(65, 17)
		Me.Label6.Location = New System.Drawing.Point(8, 40)
		Me.Label6.TabIndex = 38
		Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopLeft
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
		Me.txtOrigin.AutoSize = False
		Me.txtOrigin.Size = New System.Drawing.Size(49, 19)
		Me.txtOrigin.Location = New System.Drawing.Point(80, 40)
		Me.txtOrigin.TabIndex = 1
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
		Me.txtPath.AutoSize = False
		Me.txtPath.Size = New System.Drawing.Size(49, 19)
		Me.txtPath.Location = New System.Drawing.Point(80, 64)
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
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "Set"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Size = New System.Drawing.Size(57, 25)
		Me.cmdOK.Location = New System.Drawing.Point(136, 144)
		Me.cmdOK.TabIndex = 30
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
		Me.cmdCancel.Text = "Close"
		Me.cmdCancel.Size = New System.Drawing.Size(57, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(440, 144)
		Me.cmdCancel.TabIndex = 35
		Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.txtUnit.AutoSize = False
		Me.txtUnit.Size = New System.Drawing.Size(65, 19)
		Me.txtUnit.Location = New System.Drawing.Point(64, 8)
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
		Me.cmdHelp.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdHelp.Text = "?"
		Me.cmdHelp.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdHelp.Size = New System.Drawing.Size(25, 25)
		Me.cmdHelp.Location = New System.Drawing.Point(504, 0)
		Me.cmdHelp.TabIndex = 37
		Me.ToolTip1.SetToolTip(Me.cmdHelp, "Click for Help")
		Me.cmdHelp.BackColor = System.Drawing.SystemColors.Control
		Me.cmdHelp.CausesValidation = True
		Me.cmdHelp.Enabled = True
		Me.cmdHelp.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdHelp.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdHelp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdHelp.TabStop = True
		Me.cmdHelp.Name = "cmdHelp"
		Me.chkAutoScuttle.Text = "Auto-scuttle (ts only)"
		Me.chkAutoScuttle.Size = New System.Drawing.Size(121, 17)
		Me.chkAutoScuttle.Location = New System.Drawing.Point(16, 152)
		Me.chkAutoScuttle.TabIndex = 5
		Me.chkAutoScuttle.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkAutoScuttle.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkAutoScuttle.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkAutoScuttle.BackColor = System.Drawing.SystemColors.Control
		Me.chkAutoScuttle.CausesValidation = True
		Me.chkAutoScuttle.Enabled = True
		Me.chkAutoScuttle.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkAutoScuttle.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkAutoScuttle.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkAutoScuttle.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkAutoScuttle.TabStop = True
		Me.chkAutoScuttle.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkAutoScuttle.Visible = True
		Me.chkAutoScuttle.Name = "chkAutoScuttle"
		Me.Label4.Text = "2nd dest"
		Me.Label4.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.Size = New System.Drawing.Size(57, 17)
		Me.Label4.Location = New System.Drawing.Point(16, 64)
		Me.Label4.TabIndex = 51
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
		Me.Label3.Text = "1st dest"
		Me.Label3.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.Size = New System.Drawing.Size(49, 17)
		Me.Label3.Location = New System.Drawing.Point(16, 40)
		Me.Label3.TabIndex = 50
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
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label2.Text = "Order"
		Me.Label2.Font = New System.Drawing.Font("Arial Black", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(41, 17)
		Me.Label2.Location = New System.Drawing.Point(8, 8)
		Me.Label2.TabIndex = 49
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
		Me.Controls.Add(cmdShow)
		Me.Controls.Add(cmdClear)
		Me.Controls.Add(cmdSuspend)
		Me.Controls.Add(cmdResume)
		Me.Controls.Add(Frame1)
		Me.Controls.Add(txtOrigin)
		Me.Controls.Add(txtPath)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(txtUnit)
		Me.Controls.Add(cmdHelp)
		Me.Controls.Add(chkAutoScuttle)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label2)
		Me.Frame2.Controls.Add(txtAutoFeed)
		Me.Frame2.Controls.Add(chkAutoFeed)
		Me.Frame2.Controls.Add(lblAutoFeed)
		Me.Frame1.Controls.Add(_txtEndCargo_2)
		Me.Frame1.Controls.Add(_txtEndCargo_1)
		Me.Frame1.Controls.Add(_txtEndCargo_3)
		Me.Frame1.Controls.Add(_txtEndCargo_4)
		Me.Frame1.Controls.Add(_txtEndCargo_6)
		Me.Frame1.Controls.Add(_txtEndCargo_5)
		Me.Frame1.Controls.Add(_txtStartCargo_5)
		Me.Frame1.Controls.Add(_txtEnd_5)
		Me.Frame1.Controls.Add(_txtStart_5)
		Me.Frame1.Controls.Add(_txtStartCargo_6)
		Me.Frame1.Controls.Add(_txtEnd_6)
		Me.Frame1.Controls.Add(_txtStart_6)
		Me.Frame1.Controls.Add(_txtStartCargo_4)
		Me.Frame1.Controls.Add(_txtEnd_4)
		Me.Frame1.Controls.Add(_txtStart_4)
		Me.Frame1.Controls.Add(_txtStartCargo_3)
		Me.Frame1.Controls.Add(_txtEnd_3)
		Me.Frame1.Controls.Add(_txtStart_3)
		Me.Frame1.Controls.Add(_txtStartCargo_1)
		Me.Frame1.Controls.Add(_txtEnd_1)
		Me.Frame1.Controls.Add(_txtStart_1)
		Me.Frame1.Controls.Add(_txtStartCargo_2)
		Me.Frame1.Controls.Add(_txtEnd_2)
		Me.Frame1.Controls.Add(_txtStart_2)
		Me.Frame1.Controls.Add(Label21)
		Me.Frame1.Controls.Add(Label19)
		Me.Frame1.Controls.Add(Label18)
		Me.Frame1.Controls.Add(_lblHold_6)
		Me.Frame1.Controls.Add(_lblHold_5)
		Me.Frame1.Controls.Add(_lblHold_4)
		Me.Frame1.Controls.Add(Label14)
		Me.Frame1.Controls.Add(Label12)
		Me.Frame1.Controls.Add(_lblHold_3)
		Me.Frame1.Controls.Add(_lblHold_2)
		Me.Frame1.Controls.Add(_lblHold_1)
		Me.Frame1.Controls.Add(Label6)
		Me.lblHold.SetIndex(_lblHold_6, CType(6, Short))
		Me.lblHold.SetIndex(_lblHold_5, CType(5, Short))
		Me.lblHold.SetIndex(_lblHold_4, CType(4, Short))
		Me.lblHold.SetIndex(_lblHold_3, CType(3, Short))
		Me.lblHold.SetIndex(_lblHold_2, CType(2, Short))
		Me.lblHold.SetIndex(_lblHold_1, CType(1, Short))
		Me.txtEnd.SetIndex(_txtEnd_5, CType(5, Short))
		Me.txtEnd.SetIndex(_txtEnd_6, CType(6, Short))
		Me.txtEnd.SetIndex(_txtEnd_4, CType(4, Short))
		Me.txtEnd.SetIndex(_txtEnd_3, CType(3, Short))
		Me.txtEnd.SetIndex(_txtEnd_1, CType(1, Short))
		Me.txtEnd.SetIndex(_txtEnd_2, CType(2, Short))
		Me.txtEndCargo.SetIndex(_txtEndCargo_2, CType(2, Short))
		Me.txtEndCargo.SetIndex(_txtEndCargo_1, CType(1, Short))
		Me.txtEndCargo.SetIndex(_txtEndCargo_3, CType(3, Short))
		Me.txtEndCargo.SetIndex(_txtEndCargo_4, CType(4, Short))
		Me.txtEndCargo.SetIndex(_txtEndCargo_6, CType(6, Short))
		Me.txtEndCargo.SetIndex(_txtEndCargo_5, CType(5, Short))
		Me.txtStart.SetIndex(_txtStart_5, CType(5, Short))
		Me.txtStart.SetIndex(_txtStart_6, CType(6, Short))
		Me.txtStart.SetIndex(_txtStart_4, CType(4, Short))
		Me.txtStart.SetIndex(_txtStart_3, CType(3, Short))
		Me.txtStart.SetIndex(_txtStart_1, CType(1, Short))
		Me.txtStart.SetIndex(_txtStart_2, CType(2, Short))
		Me.txtStartCargo.SetIndex(_txtStartCargo_5, CType(5, Short))
		Me.txtStartCargo.SetIndex(_txtStartCargo_6, CType(6, Short))
		Me.txtStartCargo.SetIndex(_txtStartCargo_4, CType(4, Short))
		Me.txtStartCargo.SetIndex(_txtStartCargo_3, CType(3, Short))
		Me.txtStartCargo.SetIndex(_txtStartCargo_1, CType(1, Short))
		Me.txtStartCargo.SetIndex(_txtStartCargo_2, CType(2, Short))
		CType(Me.txtStartCargo, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.txtStart, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.txtEndCargo, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.txtEnd, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lblHold, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Frame2.ResumeLayout(False)
		Me.Frame1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class