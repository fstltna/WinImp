<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPromptDes
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
	Public WithEvents chkClearOldThresholds As System.Windows.Forms.CheckBox
	Public WithEvents cmdHelp As System.Windows.Forms.Button
	Public WithEvents Check1 As System.Windows.Forms.CheckBox
	Public WithEvents cbType As System.Windows.Forms.ListBox
	Public WithEvents chkSave As System.Windows.Forms.CheckBox
	Public WithEvents _txtThresh_3 As System.Windows.Forms.TextBox
	Public WithEvents _txtThresh_2 As System.Windows.Forms.TextBox
	Public WithEvents _txtThresh_1 As System.Windows.Forms.TextBox
	Public WithEvents _txtThresh_0 As System.Windows.Forms.TextBox
	Public WithEvents _lblThresh_3 As System.Windows.Forms.Label
	Public WithEvents _lblThresh_2 As System.Windows.Forms.Label
	Public WithEvents _lblThresh_1 As System.Windows.Forms.Label
	Public WithEvents _lblThresh_0 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents txtMultOrigin As System.Windows.Forms.TextBox
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents lblDesc As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents lblThresh As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents txtThresh As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPromptDes))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.chkClearOldThresholds = New System.Windows.Forms.CheckBox
		Me.cmdHelp = New System.Windows.Forms.Button
		Me.Check1 = New System.Windows.Forms.CheckBox
		Me.cbType = New System.Windows.Forms.ListBox
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.chkSave = New System.Windows.Forms.CheckBox
		Me._txtThresh_3 = New System.Windows.Forms.TextBox
		Me._txtThresh_2 = New System.Windows.Forms.TextBox
		Me._txtThresh_1 = New System.Windows.Forms.TextBox
		Me._txtThresh_0 = New System.Windows.Forms.TextBox
		Me._lblThresh_3 = New System.Windows.Forms.Label
		Me._lblThresh_2 = New System.Windows.Forms.Label
		Me._lblThresh_1 = New System.Windows.Forms.Label
		Me._lblThresh_0 = New System.Windows.Forms.Label
		Me.txtMultOrigin = New System.Windows.Forms.TextBox
		Me.cmdOK = New System.Windows.Forms.Button
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.lblDesc = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.lblThresh = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.txtThresh = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(components)
		Me.Frame1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.lblThresh, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.txtThresh, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.ControlBox = False
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.ClientSize = New System.Drawing.Size(608, 184)
		Me.Location = New System.Drawing.Point(1, 1)
		Me.KeyPreview = True
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.Enabled = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmPromptDes"
		Me.chkClearOldThresholds.Text = "Clear Old Thresholds when designating"
		Me.chkClearOldThresholds.Size = New System.Drawing.Size(81, 65)
		Me.chkClearOldThresholds.Location = New System.Drawing.Point(520, 112)
		Me.chkClearOldThresholds.TabIndex = 19
		Me.chkClearOldThresholds.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkClearOldThresholds.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkClearOldThresholds.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkClearOldThresholds.BackColor = System.Drawing.SystemColors.Control
		Me.chkClearOldThresholds.CausesValidation = True
		Me.chkClearOldThresholds.Enabled = True
		Me.chkClearOldThresholds.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkClearOldThresholds.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkClearOldThresholds.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkClearOldThresholds.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkClearOldThresholds.TabStop = True
		Me.chkClearOldThresholds.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkClearOldThresholds.Visible = True
		Me.chkClearOldThresholds.Name = "chkClearOldThresholds"
		Me.cmdHelp.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdHelp.Text = "?"
		Me.cmdHelp.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdHelp.Size = New System.Drawing.Size(25, 25)
		Me.cmdHelp.Location = New System.Drawing.Point(584, 0)
		Me.cmdHelp.TabIndex = 18
		Me.ToolTip1.SetToolTip(Me.cmdHelp, "Click for Help")
		Me.cmdHelp.BackColor = System.Drawing.SystemColors.Control
		Me.cmdHelp.CausesValidation = True
		Me.cmdHelp.Enabled = True
		Me.cmdHelp.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdHelp.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdHelp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdHelp.TabStop = True
		Me.cmdHelp.Name = "cmdHelp"
		Me.Check1.Text = "Set Thresholds when you designate"
		Me.Check1.Size = New System.Drawing.Size(201, 17)
		Me.Check1.Location = New System.Drawing.Point(8, 56)
		Me.Check1.TabIndex = 16
		Me.Check1.CheckState = System.Windows.Forms.CheckState.Checked
		Me.Check1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Check1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Check1.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.Check1.BackColor = System.Drawing.SystemColors.Control
		Me.Check1.CausesValidation = True
		Me.Check1.Enabled = True
		Me.Check1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Check1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Check1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Check1.Appearance = System.Windows.Forms.Appearance.Normal
		Me.Check1.TabStop = True
		Me.Check1.Visible = True
		Me.Check1.Name = "Check1"
		Me.cbType.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbType.Size = New System.Drawing.Size(272, 172)
		Me.cbType.IntegralHeight = False
		Me.cbType.Location = New System.Drawing.Point(232, 8)
		Me.cbType.Sorted = True
		Me.cbType.TabIndex = 1
		Me.cbType.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.cbType.BackColor = System.Drawing.SystemColors.Window
		Me.cbType.CausesValidation = True
		Me.cbType.Enabled = True
		Me.cbType.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cbType.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbType.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.cbType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbType.TabStop = True
		Me.cbType.Visible = True
		Me.cbType.MultiColumn = True
		Me.cbType.ColumnWidth = 136
		Me.cbType.Name = "cbType"
		Me.Frame1.Text = "Thresholds"
		Me.Frame1.Size = New System.Drawing.Size(217, 97)
		Me.Frame1.Location = New System.Drawing.Point(8, 80)
		Me.Frame1.TabIndex = 7
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.chkSave.Text = "Save as New Defaults"
		Me.chkSave.Size = New System.Drawing.Size(153, 17)
		Me.chkSave.Location = New System.Drawing.Point(16, 72)
		Me.chkSave.TabIndex = 17
		Me.chkSave.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkSave.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkSave.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkSave.BackColor = System.Drawing.SystemColors.Control
		Me.chkSave.CausesValidation = True
		Me.chkSave.Enabled = True
		Me.chkSave.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkSave.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkSave.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkSave.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkSave.TabStop = True
		Me.chkSave.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkSave.Visible = True
		Me.chkSave.Name = "chkSave"
		Me._txtThresh_3.AutoSize = False
		Me._txtThresh_3.Size = New System.Drawing.Size(33, 19)
		Me._txtThresh_3.Location = New System.Drawing.Point(168, 48)
		Me._txtThresh_3.TabIndex = 11
		Me._txtThresh_3.Text = "Text1"
		Me._txtThresh_3.Visible = False
		Me._txtThresh_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtThresh_3.AcceptsReturn = True
		Me._txtThresh_3.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtThresh_3.BackColor = System.Drawing.SystemColors.Window
		Me._txtThresh_3.CausesValidation = True
		Me._txtThresh_3.Enabled = True
		Me._txtThresh_3.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtThresh_3.HideSelection = True
		Me._txtThresh_3.ReadOnly = False
		Me._txtThresh_3.Maxlength = 0
		Me._txtThresh_3.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtThresh_3.MultiLine = False
		Me._txtThresh_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtThresh_3.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtThresh_3.TabStop = True
		Me._txtThresh_3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtThresh_3.Name = "_txtThresh_3"
		Me._txtThresh_2.AutoSize = False
		Me._txtThresh_2.Size = New System.Drawing.Size(33, 19)
		Me._txtThresh_2.Location = New System.Drawing.Point(168, 24)
		Me._txtThresh_2.TabIndex = 10
		Me._txtThresh_2.Text = "Text1"
		Me._txtThresh_2.Visible = False
		Me._txtThresh_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtThresh_2.AcceptsReturn = True
		Me._txtThresh_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtThresh_2.BackColor = System.Drawing.SystemColors.Window
		Me._txtThresh_2.CausesValidation = True
		Me._txtThresh_2.Enabled = True
		Me._txtThresh_2.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtThresh_2.HideSelection = True
		Me._txtThresh_2.ReadOnly = False
		Me._txtThresh_2.Maxlength = 0
		Me._txtThresh_2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtThresh_2.MultiLine = False
		Me._txtThresh_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtThresh_2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtThresh_2.TabStop = True
		Me._txtThresh_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtThresh_2.Name = "_txtThresh_2"
		Me._txtThresh_1.AutoSize = False
		Me._txtThresh_1.Size = New System.Drawing.Size(33, 19)
		Me._txtThresh_1.Location = New System.Drawing.Point(64, 48)
		Me._txtThresh_1.TabIndex = 9
		Me._txtThresh_1.Text = "Text1"
		Me._txtThresh_1.Visible = False
		Me._txtThresh_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtThresh_1.AcceptsReturn = True
		Me._txtThresh_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtThresh_1.BackColor = System.Drawing.SystemColors.Window
		Me._txtThresh_1.CausesValidation = True
		Me._txtThresh_1.Enabled = True
		Me._txtThresh_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtThresh_1.HideSelection = True
		Me._txtThresh_1.ReadOnly = False
		Me._txtThresh_1.Maxlength = 0
		Me._txtThresh_1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtThresh_1.MultiLine = False
		Me._txtThresh_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtThresh_1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtThresh_1.TabStop = True
		Me._txtThresh_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtThresh_1.Name = "_txtThresh_1"
		Me._txtThresh_0.AutoSize = False
		Me._txtThresh_0.Size = New System.Drawing.Size(33, 19)
		Me._txtThresh_0.Location = New System.Drawing.Point(64, 24)
		Me._txtThresh_0.TabIndex = 8
		Me._txtThresh_0.Text = "Text1"
		Me._txtThresh_0.Visible = False
		Me._txtThresh_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtThresh_0.AcceptsReturn = True
		Me._txtThresh_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtThresh_0.BackColor = System.Drawing.SystemColors.Window
		Me._txtThresh_0.CausesValidation = True
		Me._txtThresh_0.Enabled = True
		Me._txtThresh_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtThresh_0.HideSelection = True
		Me._txtThresh_0.ReadOnly = False
		Me._txtThresh_0.Maxlength = 0
		Me._txtThresh_0.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtThresh_0.MultiLine = False
		Me._txtThresh_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtThresh_0.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtThresh_0.TabStop = True
		Me._txtThresh_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtThresh_0.Name = "_txtThresh_0"
		Me._lblThresh_3.Text = "Label3"
		Me._lblThresh_3.Size = New System.Drawing.Size(41, 17)
		Me._lblThresh_3.Location = New System.Drawing.Point(120, 48)
		Me._lblThresh_3.TabIndex = 15
		Me._lblThresh_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblThresh_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblThresh_3.BackColor = System.Drawing.SystemColors.Control
		Me._lblThresh_3.Enabled = True
		Me._lblThresh_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblThresh_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblThresh_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblThresh_3.UseMnemonic = True
		Me._lblThresh_3.Visible = True
		Me._lblThresh_3.AutoSize = False
		Me._lblThresh_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblThresh_3.Name = "_lblThresh_3"
		Me._lblThresh_2.Text = "Label3"
		Me._lblThresh_2.Size = New System.Drawing.Size(49, 17)
		Me._lblThresh_2.Location = New System.Drawing.Point(120, 24)
		Me._lblThresh_2.TabIndex = 14
		Me._lblThresh_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblThresh_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblThresh_2.BackColor = System.Drawing.SystemColors.Control
		Me._lblThresh_2.Enabled = True
		Me._lblThresh_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblThresh_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblThresh_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblThresh_2.UseMnemonic = True
		Me._lblThresh_2.Visible = True
		Me._lblThresh_2.AutoSize = False
		Me._lblThresh_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblThresh_2.Name = "_lblThresh_2"
		Me._lblThresh_1.Text = "Label3"
		Me._lblThresh_1.Size = New System.Drawing.Size(41, 17)
		Me._lblThresh_1.Location = New System.Drawing.Point(16, 48)
		Me._lblThresh_1.TabIndex = 13
		Me._lblThresh_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblThresh_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblThresh_1.BackColor = System.Drawing.SystemColors.Control
		Me._lblThresh_1.Enabled = True
		Me._lblThresh_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblThresh_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblThresh_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblThresh_1.UseMnemonic = True
		Me._lblThresh_1.Visible = True
		Me._lblThresh_1.AutoSize = False
		Me._lblThresh_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblThresh_1.Name = "_lblThresh_1"
		Me._lblThresh_0.Text = "Label3"
		Me._lblThresh_0.Size = New System.Drawing.Size(41, 17)
		Me._lblThresh_0.Location = New System.Drawing.Point(16, 24)
		Me._lblThresh_0.TabIndex = 12
		Me._lblThresh_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblThresh_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblThresh_0.BackColor = System.Drawing.SystemColors.Control
		Me._lblThresh_0.Enabled = True
		Me._lblThresh_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblThresh_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblThresh_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblThresh_0.UseMnemonic = True
		Me._lblThresh_0.Visible = True
		Me._lblThresh_0.AutoSize = False
		Me._lblThresh_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblThresh_0.Name = "_lblThresh_0"
		Me.txtMultOrigin.AutoSize = False
		Me.txtMultOrigin.Size = New System.Drawing.Size(89, 19)
		Me.txtMultOrigin.Location = New System.Drawing.Point(88, 8)
		Me.txtMultOrigin.TabIndex = 0
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
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "OK"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Size = New System.Drawing.Size(81, 25)
		Me.cmdOK.Location = New System.Drawing.Point(520, 40)
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
		Me.cmdCancel.Location = New System.Drawing.Point(520, 72)
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
		Me.lblDesc.Size = New System.Drawing.Size(3, 13)
		Me.lblDesc.Location = New System.Drawing.Point(8, 32)
		Me.lblDesc.TabIndex = 6
		Me.lblDesc.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblDesc.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblDesc.BackColor = System.Drawing.SystemColors.Control
		Me.lblDesc.Enabled = True
		Me.lblDesc.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblDesc.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblDesc.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblDesc.UseMnemonic = True
		Me.lblDesc.Visible = True
		Me.lblDesc.AutoSize = True
		Me.lblDesc.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblDesc.Name = "lblDesc"
		Me.Label2.Text = "Designate"
		Me.Label2.Font = New System.Drawing.Font("Arial Black", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(73, 17)
		Me.Label2.Location = New System.Drawing.Point(8, 8)
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
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label1.Text = "as"
		Me.Label1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Size = New System.Drawing.Size(33, 17)
		Me.Label1.Location = New System.Drawing.Point(184, 8)
		Me.Label1.TabIndex = 2
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
		Me.Controls.Add(chkClearOldThresholds)
		Me.Controls.Add(cmdHelp)
		Me.Controls.Add(Check1)
		Me.Controls.Add(cbType)
		Me.Controls.Add(Frame1)
		Me.Controls.Add(txtMultOrigin)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(lblDesc)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		Me.Frame1.Controls.Add(chkSave)
		Me.Frame1.Controls.Add(_txtThresh_3)
		Me.Frame1.Controls.Add(_txtThresh_2)
		Me.Frame1.Controls.Add(_txtThresh_1)
		Me.Frame1.Controls.Add(_txtThresh_0)
		Me.Frame1.Controls.Add(_lblThresh_3)
		Me.Frame1.Controls.Add(_lblThresh_2)
		Me.Frame1.Controls.Add(_lblThresh_1)
		Me.Frame1.Controls.Add(_lblThresh_0)
		Me.lblThresh.SetIndex(_lblThresh_3, CType(3, Short))
		Me.lblThresh.SetIndex(_lblThresh_2, CType(2, Short))
		Me.lblThresh.SetIndex(_lblThresh_1, CType(1, Short))
		Me.lblThresh.SetIndex(_lblThresh_0, CType(0, Short))
		Me.txtThresh.SetIndex(_txtThresh_3, CType(3, Short))
		Me.txtThresh.SetIndex(_txtThresh_2, CType(2, Short))
		Me.txtThresh.SetIndex(_txtThresh_1, CType(1, Short))
		Me.txtThresh.SetIndex(_txtThresh_0, CType(0, Short))
		CType(Me.txtThresh, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lblThresh, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Frame1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class