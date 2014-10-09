<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmToolWorldBuild
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
	Public WithEvents cmdCopy As System.Windows.Forms.Button
	Public WithEvents cbCopy As System.Windows.Forms.ComboBox
	Public WithEvents txtPath As System.Windows.Forms.TextBox
	Public WithEvents txtOrigin As System.Windows.Forms.TextBox
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Frame3 As System.Windows.Forms.GroupBox
	Public WithEvents chkDes As System.Windows.Forms.CheckBox
	Public WithEvents txtDest As System.Windows.Forms.TextBox
	Public WithEvents cmdSetResource As System.Windows.Forms.Button
	Public WithEvents _txtThresh_4 As System.Windows.Forms.TextBox
	Public WithEvents _txtThresh_0 As System.Windows.Forms.TextBox
	Public WithEvents _txtThresh_1 As System.Windows.Forms.TextBox
	Public WithEvents _txtThresh_2 As System.Windows.Forms.TextBox
	Public WithEvents _txtThresh_3 As System.Windows.Forms.TextBox
	Public WithEvents Line1 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents _lblThresh_4 As System.Windows.Forms.Label
	Public WithEvents _lblThresh_0 As System.Windows.Forms.Label
	Public WithEvents _lblThresh_1 As System.Windows.Forms.Label
	Public WithEvents _lblThresh_2 As System.Windows.Forms.Label
	Public WithEvents _lblThresh_3 As System.Windows.Forms.Label
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents Check1 As System.Windows.Forms.CheckBox
	Public WithEvents _cbType_1 As System.Windows.Forms.ComboBox
	Public WithEvents _cbType_0 As System.Windows.Forms.ComboBox
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents cbType As Microsoft.VisualBasic.Compatibility.VB6.ComboBoxArray
	Public WithEvents lblThresh As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents txtThresh As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	Public WithEvents ShapeContainer1 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmToolWorldBuild))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer
		Me.Frame3 = New System.Windows.Forms.GroupBox
		Me.cmdCopy = New System.Windows.Forms.Button
		Me.cbCopy = New System.Windows.Forms.ComboBox
		Me.txtPath = New System.Windows.Forms.TextBox
		Me.txtOrigin = New System.Windows.Forms.TextBox
		Me.Label5 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.chkDes = New System.Windows.Forms.CheckBox
		Me.Frame2 = New System.Windows.Forms.GroupBox
		Me.txtDest = New System.Windows.Forms.TextBox
		Me.cmdSetResource = New System.Windows.Forms.Button
		Me._txtThresh_4 = New System.Windows.Forms.TextBox
		Me._txtThresh_0 = New System.Windows.Forms.TextBox
		Me._txtThresh_1 = New System.Windows.Forms.TextBox
		Me._txtThresh_2 = New System.Windows.Forms.TextBox
		Me._txtThresh_3 = New System.Windows.Forms.TextBox
		Me.Line1 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.Label6 = New System.Windows.Forms.Label
		Me._lblThresh_4 = New System.Windows.Forms.Label
		Me._lblThresh_0 = New System.Windows.Forms.Label
		Me._lblThresh_1 = New System.Windows.Forms.Label
		Me._lblThresh_2 = New System.Windows.Forms.Label
		Me._lblThresh_3 = New System.Windows.Forms.Label
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.Check1 = New System.Windows.Forms.CheckBox
		Me._cbType_1 = New System.Windows.Forms.ComboBox
		Me._cbType_0 = New System.Windows.Forms.ComboBox
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.Label2 = New System.Windows.Forms.Label
		Me.cbType = New Microsoft.VisualBasic.Compatibility.VB6.ComboBoxArray(components)
		Me.lblThresh = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.txtThresh = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(components)
		Me.Frame3.SuspendLayout()
		Me.Frame2.SuspendLayout()
		Me.Frame1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.cbType, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lblThresh, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.txtThresh, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.ControlBox = False
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.ClientSize = New System.Drawing.Size(534, 197)
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
		Me.Name = "frmToolWorldBuild"
		Me.Frame3.Text = "Copy/Paste"
		Me.Frame3.Size = New System.Drawing.Size(121, 169)
		Me.Frame3.Location = New System.Drawing.Point(400, 16)
		Me.Frame3.TabIndex = 21
		Me.Frame3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame3.BackColor = System.Drawing.SystemColors.Control
		Me.Frame3.Enabled = True
		Me.Frame3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame3.Visible = True
		Me.Frame3.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame3.Name = "Frame3"
		Me.cmdCopy.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCopy.Text = "Copy"
		Me.cmdCopy.Size = New System.Drawing.Size(65, 25)
		Me.cmdCopy.Location = New System.Drawing.Point(24, 136)
		Me.cmdCopy.TabIndex = 27
		Me.cmdCopy.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCopy.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCopy.CausesValidation = True
		Me.cmdCopy.Enabled = True
		Me.cmdCopy.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCopy.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCopy.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCopy.TabStop = True
		Me.cmdCopy.Name = "cmdCopy"
		Me.cbCopy.Size = New System.Drawing.Size(105, 21)
		Me.cbCopy.Location = New System.Drawing.Point(8, 104)
		Me.cbCopy.Items.AddRange(New Object(){"Copy Sector", "Copy Sector, Resources", "Copy Sector, Resources, Commodities"})
		Me.cbCopy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cbCopy.TabIndex = 26
		Me.cbCopy.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbCopy.BackColor = System.Drawing.SystemColors.Window
		Me.cbCopy.CausesValidation = True
		Me.cbCopy.Enabled = True
		Me.cbCopy.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cbCopy.IntegralHeight = True
		Me.cbCopy.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbCopy.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbCopy.Sorted = False
		Me.cbCopy.TabStop = True
		Me.cbCopy.Visible = True
		Me.cbCopy.Name = "cbCopy"
		Me.txtPath.AutoSize = False
		Me.txtPath.Size = New System.Drawing.Size(105, 19)
		Me.txtPath.Location = New System.Drawing.Point(8, 72)
		Me.txtPath.TabIndex = 25
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
		Me.txtOrigin.AutoSize = False
		Me.txtOrigin.Size = New System.Drawing.Size(105, 19)
		Me.txtOrigin.Location = New System.Drawing.Point(8, 32)
		Me.txtOrigin.TabIndex = 23
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
		Me.Label5.Text = "Destination Sector"
		Me.Label5.Size = New System.Drawing.Size(105, 17)
		Me.Label5.Location = New System.Drawing.Point(8, 56)
		Me.Label5.TabIndex = 24
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
		Me.Label1.Text = "Source Sectors"
		Me.Label1.Size = New System.Drawing.Size(97, 17)
		Me.Label1.Location = New System.Drawing.Point(8, 16)
		Me.Label1.TabIndex = 22
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
		Me.chkDes.Text = "Click to Designate"
		Me.chkDes.Size = New System.Drawing.Size(121, 17)
		Me.chkDes.Location = New System.Drawing.Point(16, 48)
		Me.chkDes.TabIndex = 20
		Me.chkDes.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkDes.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkDes.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkDes.BackColor = System.Drawing.SystemColors.Control
		Me.chkDes.CausesValidation = True
		Me.chkDes.Enabled = True
		Me.chkDes.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkDes.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkDes.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkDes.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkDes.TabStop = True
		Me.chkDes.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkDes.Visible = True
		Me.chkDes.Name = "chkDes"
		Me.Frame2.Text = "Set Resources"
		Me.Frame2.Size = New System.Drawing.Size(153, 169)
		Me.Frame2.Location = New System.Drawing.Point(240, 16)
		Me.Frame2.TabIndex = 8
		Me.Frame2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame2.BackColor = System.Drawing.SystemColors.Control
		Me.Frame2.Enabled = True
		Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame2.Visible = True
		Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame2.Name = "Frame2"
		Me.txtDest.AutoSize = False
		Me.txtDest.Size = New System.Drawing.Size(57, 19)
		Me.txtDest.Location = New System.Drawing.Point(8, 32)
		Me.txtDest.TabIndex = 28
		Me.txtDest.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtDest.AcceptsReturn = True
		Me.txtDest.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtDest.BackColor = System.Drawing.SystemColors.Window
		Me.txtDest.CausesValidation = True
		Me.txtDest.Enabled = True
		Me.txtDest.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtDest.HideSelection = True
		Me.txtDest.ReadOnly = False
		Me.txtDest.Maxlength = 0
		Me.txtDest.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtDest.MultiLine = False
		Me.txtDest.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtDest.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtDest.TabStop = True
		Me.txtDest.Visible = True
		Me.txtDest.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtDest.Name = "txtDest"
		Me.cmdSetResource.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdSetResource.Text = "Set "
		Me.cmdSetResource.Size = New System.Drawing.Size(65, 25)
		Me.cmdSetResource.Location = New System.Drawing.Point(80, 24)
		Me.cmdSetResource.TabIndex = 19
		Me.cmdSetResource.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdSetResource.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSetResource.CausesValidation = True
		Me.cmdSetResource.Enabled = True
		Me.cmdSetResource.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSetResource.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSetResource.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSetResource.TabStop = True
		Me.cmdSetResource.Name = "cmdSetResource"
		Me._txtThresh_4.AutoSize = False
		Me._txtThresh_4.Size = New System.Drawing.Size(25, 19)
		Me._txtThresh_4.Location = New System.Drawing.Point(112, 120)
		Me._txtThresh_4.TabIndex = 13
		Me._txtThresh_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtThresh_4.AcceptsReturn = True
		Me._txtThresh_4.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtThresh_4.BackColor = System.Drawing.SystemColors.Window
		Me._txtThresh_4.CausesValidation = True
		Me._txtThresh_4.Enabled = True
		Me._txtThresh_4.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtThresh_4.HideSelection = True
		Me._txtThresh_4.ReadOnly = False
		Me._txtThresh_4.Maxlength = 0
		Me._txtThresh_4.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtThresh_4.MultiLine = False
		Me._txtThresh_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtThresh_4.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtThresh_4.TabStop = True
		Me._txtThresh_4.Visible = True
		Me._txtThresh_4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtThresh_4.Name = "_txtThresh_4"
		Me._txtThresh_0.AutoSize = False
		Me._txtThresh_0.Size = New System.Drawing.Size(25, 19)
		Me._txtThresh_0.Location = New System.Drawing.Point(40, 72)
		Me._txtThresh_0.TabIndex = 12
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
		Me._txtThresh_0.Visible = True
		Me._txtThresh_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtThresh_0.Name = "_txtThresh_0"
		Me._txtThresh_1.AutoSize = False
		Me._txtThresh_1.Size = New System.Drawing.Size(25, 19)
		Me._txtThresh_1.Location = New System.Drawing.Point(40, 104)
		Me._txtThresh_1.TabIndex = 11
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
		Me._txtThresh_1.Visible = True
		Me._txtThresh_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtThresh_1.Name = "_txtThresh_1"
		Me._txtThresh_2.AutoSize = False
		Me._txtThresh_2.Size = New System.Drawing.Size(25, 19)
		Me._txtThresh_2.Location = New System.Drawing.Point(40, 136)
		Me._txtThresh_2.TabIndex = 10
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
		Me._txtThresh_2.Visible = True
		Me._txtThresh_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtThresh_2.Name = "_txtThresh_2"
		Me._txtThresh_3.AutoSize = False
		Me._txtThresh_3.Size = New System.Drawing.Size(25, 19)
		Me._txtThresh_3.Location = New System.Drawing.Point(112, 88)
		Me._txtThresh_3.TabIndex = 9
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
		Me._txtThresh_3.Visible = True
		Me._txtThresh_3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtThresh_3.Name = "_txtThresh_3"
		Me.Line1.X1 = 8
		Me.Line1.X2 = 144
		Me.Line1.Y1 = 51
		Me.Line1.Y2 = 51
		Me.Line1.BorderColor = System.Drawing.SystemColors.WindowText
		Me.Line1.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line1.BorderWidth = 1
		Me.Line1.Visible = True
		Me.Line1.Name = "Line1"
		Me.Label6.Text = "Sector"
		Me.Label6.Size = New System.Drawing.Size(41, 17)
		Me.Label6.Location = New System.Drawing.Point(8, 16)
		Me.Label6.TabIndex = 29
		Me.Label6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me._lblThresh_4.Text = "Uran"
		Me._lblThresh_4.Size = New System.Drawing.Size(33, 17)
		Me._lblThresh_4.Location = New System.Drawing.Point(80, 120)
		Me._lblThresh_4.TabIndex = 18
		Me._lblThresh_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblThresh_4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblThresh_4.BackColor = System.Drawing.SystemColors.Control
		Me._lblThresh_4.Enabled = True
		Me._lblThresh_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblThresh_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblThresh_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblThresh_4.UseMnemonic = True
		Me._lblThresh_4.Visible = True
		Me._lblThresh_4.AutoSize = False
		Me._lblThresh_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblThresh_4.Name = "_lblThresh_4"
		Me._lblThresh_0.Text = "Iron"
		Me._lblThresh_0.Size = New System.Drawing.Size(33, 17)
		Me._lblThresh_0.Location = New System.Drawing.Point(8, 72)
		Me._lblThresh_0.TabIndex = 17
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
		Me._lblThresh_1.Text = "Gold"
		Me._lblThresh_1.Size = New System.Drawing.Size(33, 17)
		Me._lblThresh_1.Location = New System.Drawing.Point(8, 104)
		Me._lblThresh_1.TabIndex = 16
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
		Me._lblThresh_2.Text = "Fert"
		Me._lblThresh_2.Size = New System.Drawing.Size(33, 17)
		Me._lblThresh_2.Location = New System.Drawing.Point(8, 136)
		Me._lblThresh_2.TabIndex = 15
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
		Me._lblThresh_3.Text = "Oil"
		Me._lblThresh_3.Size = New System.Drawing.Size(25, 17)
		Me._lblThresh_3.Location = New System.Drawing.Point(80, 88)
		Me._lblThresh_3.TabIndex = 14
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
		Me.Frame1.Text = "Select Mouse Buttons"
		Me.Frame1.Size = New System.Drawing.Size(217, 113)
		Me.Frame1.Location = New System.Drawing.Point(16, 72)
		Me.Frame1.TabIndex = 2
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.Check1.Text = "Set Resources when you designate"
		Me.Check1.Size = New System.Drawing.Size(201, 17)
		Me.Check1.Location = New System.Drawing.Point(8, 88)
		Me.Check1.TabIndex = 7
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
		Me._cbType_1.Size = New System.Drawing.Size(137, 21)
		Me._cbType_1.Location = New System.Drawing.Point(64, 56)
		Me._cbType_1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me._cbType_1.TabIndex = 4
		Me._cbType_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cbType_1.BackColor = System.Drawing.SystemColors.Window
		Me._cbType_1.CausesValidation = True
		Me._cbType_1.Enabled = True
		Me._cbType_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._cbType_1.IntegralHeight = True
		Me._cbType_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._cbType_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cbType_1.Sorted = False
		Me._cbType_1.TabStop = True
		Me._cbType_1.Visible = True
		Me._cbType_1.Name = "_cbType_1"
		Me._cbType_0.Size = New System.Drawing.Size(137, 21)
		Me._cbType_0.Location = New System.Drawing.Point(64, 24)
		Me._cbType_0.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me._cbType_0.TabIndex = 3
		Me._cbType_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cbType_0.BackColor = System.Drawing.SystemColors.Window
		Me._cbType_0.CausesValidation = True
		Me._cbType_0.Enabled = True
		Me._cbType_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._cbType_0.IntegralHeight = True
		Me._cbType_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._cbType_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cbType_0.Sorted = False
		Me._cbType_0.TabStop = True
		Me._cbType_0.Visible = True
		Me._cbType_0.Name = "_cbType_0"
		Me.Label4.Text = "Right"
		Me.Label4.Size = New System.Drawing.Size(33, 17)
		Me.Label4.Location = New System.Drawing.Point(16, 56)
		Me.Label4.TabIndex = 6
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
		Me.Label3.Text = "Left "
		Me.Label3.Size = New System.Drawing.Size(25, 17)
		Me.Label3.Location = New System.Drawing.Point(16, 24)
		Me.Label3.TabIndex = 5
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
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "Done"
		Me.cmdCancel.Size = New System.Drawing.Size(81, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(152, 40)
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
		Me.Label2.Text = "World Builder"
		Me.Label2.Font = New System.Drawing.Font("Arial Black", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(97, 17)
		Me.Label2.Location = New System.Drawing.Point(16, 16)
		Me.Label2.TabIndex = 0
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
		Me.Controls.Add(Frame3)
		Me.Controls.Add(chkDes)
		Me.Controls.Add(Frame2)
		Me.Controls.Add(Frame1)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(Label2)
		Me.Frame3.Controls.Add(cmdCopy)
		Me.Frame3.Controls.Add(cbCopy)
		Me.Frame3.Controls.Add(txtPath)
		Me.Frame3.Controls.Add(txtOrigin)
		Me.Frame3.Controls.Add(Label5)
		Me.Frame3.Controls.Add(Label1)
		Me.Frame2.Controls.Add(txtDest)
		Me.Frame2.Controls.Add(cmdSetResource)
		Me.Frame2.Controls.Add(_txtThresh_4)
		Me.Frame2.Controls.Add(_txtThresh_0)
		Me.Frame2.Controls.Add(_txtThresh_1)
		Me.Frame2.Controls.Add(_txtThresh_2)
		Me.Frame2.Controls.Add(_txtThresh_3)
		Me.ShapeContainer1.Shapes.Add(Line1)
		Me.Frame2.Controls.Add(Label6)
		Me.Frame2.Controls.Add(_lblThresh_4)
		Me.Frame2.Controls.Add(_lblThresh_0)
		Me.Frame2.Controls.Add(_lblThresh_1)
		Me.Frame2.Controls.Add(_lblThresh_2)
		Me.Frame2.Controls.Add(_lblThresh_3)
		Me.Frame2.Controls.Add(ShapeContainer1)
		Me.Frame1.Controls.Add(Check1)
		Me.Frame1.Controls.Add(_cbType_1)
		Me.Frame1.Controls.Add(_cbType_0)
		Me.Frame1.Controls.Add(Label4)
		Me.Frame1.Controls.Add(Label3)
		Me.cbType.SetIndex(_cbType_1, CType(1, Short))
		Me.cbType.SetIndex(_cbType_0, CType(0, Short))
		Me.lblThresh.SetIndex(_lblThresh_4, CType(4, Short))
		Me.lblThresh.SetIndex(_lblThresh_0, CType(0, Short))
		Me.lblThresh.SetIndex(_lblThresh_1, CType(1, Short))
		Me.lblThresh.SetIndex(_lblThresh_2, CType(2, Short))
		Me.lblThresh.SetIndex(_lblThresh_3, CType(3, Short))
		Me.txtThresh.SetIndex(_txtThresh_4, CType(4, Short))
		Me.txtThresh.SetIndex(_txtThresh_0, CType(0, Short))
		Me.txtThresh.SetIndex(_txtThresh_1, CType(1, Short))
		Me.txtThresh.SetIndex(_txtThresh_2, CType(2, Short))
		Me.txtThresh.SetIndex(_txtThresh_3, CType(3, Short))
		CType(Me.txtThresh, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lblThresh, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cbType, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Frame3.ResumeLayout(False)
		Me.Frame2.ResumeLayout(False)
		Me.Frame1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class