<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmToolProduction
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
	Public WithEvents rtbReport As System.Windows.Forms.RichTextBox
	Public WithEvents cmdPE As System.Windows.Forms.Button
	Public WithEvents txtProduce As System.Windows.Forms.TextBox
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents txtOrigin As System.Windows.Forms.TextBox
	Public WithEvents _lblMaterial_3 As System.Windows.Forms.Label
	Public WithEvents _lblMaterial_2 As System.Windows.Forms.Label
	Public WithEvents _lblMaterial_1 As System.Windows.Forms.Label
	Public WithEvents _lblMaterial_0 As System.Windows.Forms.Label
	Public WithEvents _lblRequired_3 As System.Windows.Forms.Label
	Public WithEvents _lblRequired_2 As System.Windows.Forms.Label
	Public WithEvents _lblRequired_1 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents _lblRequired_0 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents _lblResource_3 As System.Windows.Forms.Label
	Public WithEvents _lblResource_2 As System.Windows.Forms.Label
	Public WithEvents _lblResource_1 As System.Windows.Forms.Label
	Public WithEvents _lblResource_0 As System.Windows.Forms.Label
	Public WithEvents Line12 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents Line11 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents Line10 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents Line9 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents Line8 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents Line7 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents Line6 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents Line5 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents Line4 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents Line3 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents Line2 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents lblProduct As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents lblType As System.Windows.Forms.Label
	Public WithEvents lblMaterial As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblRequired As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents lblResource As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents ShapeContainer1 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmToolProduction))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer
		Me.rtbReport = New System.Windows.Forms.RichTextBox
		Me.cmdPE = New System.Windows.Forms.Button
		Me.txtProduce = New System.Windows.Forms.TextBox
		Me.cmdOK = New System.Windows.Forms.Button
		Me.txtOrigin = New System.Windows.Forms.TextBox
		Me._lblMaterial_3 = New System.Windows.Forms.Label
		Me._lblMaterial_2 = New System.Windows.Forms.Label
		Me._lblMaterial_1 = New System.Windows.Forms.Label
		Me._lblMaterial_0 = New System.Windows.Forms.Label
		Me._lblRequired_3 = New System.Windows.Forms.Label
		Me._lblRequired_2 = New System.Windows.Forms.Label
		Me._lblRequired_1 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me._lblRequired_0 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me._lblResource_3 = New System.Windows.Forms.Label
		Me._lblResource_2 = New System.Windows.Forms.Label
		Me._lblResource_1 = New System.Windows.Forms.Label
		Me._lblResource_0 = New System.Windows.Forms.Label
		Me.Line12 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.Line11 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.Line10 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.Line9 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.Line8 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.Line7 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.Line6 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.Line5 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.Line4 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.Line3 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.Line2 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.lblProduct = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.lblType = New System.Windows.Forms.Label
		Me.lblMaterial = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.lblRequired = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.lblResource = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.lblMaterial, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lblRequired, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lblResource, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.ControlBox = False
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.ClientSize = New System.Drawing.Size(376, 152)
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
		Me.Name = "frmToolProduction"
		Me.rtbReport.Size = New System.Drawing.Size(169, 65)
		Me.rtbReport.Location = New System.Drawing.Point(200, 80)
		Me.rtbReport.TabIndex = 7
		Me.rtbReport.Enabled = True
		Me.rtbReport.RTF = resources.GetString("rtbReport.TextRTF")
		Me.rtbReport.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.rtbReport.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.rtbReport.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.None
		Me.rtbReport.Name = "rtbReport"
		Me.cmdPE.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdPE.Text = "PE: 97%"
		Me.cmdPE.Size = New System.Drawing.Size(57, 25)
		Me.cmdPE.Location = New System.Drawing.Point(160, 32)
		Me.cmdPE.TabIndex = 6
		Me.cmdPE.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdPE.BackColor = System.Drawing.SystemColors.Control
		Me.cmdPE.CausesValidation = True
		Me.cmdPE.Enabled = True
		Me.cmdPE.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdPE.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdPE.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdPE.TabStop = True
		Me.cmdPE.Name = "cmdPE"
		Me.txtProduce.AutoSize = False
		Me.txtProduce.Size = New System.Drawing.Size(49, 19)
		Me.txtProduce.Location = New System.Drawing.Point(256, 56)
		Me.txtProduce.TabIndex = 3
		Me.txtProduce.Text = "999"
		Me.txtProduce.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtProduce.AcceptsReturn = True
		Me.txtProduce.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtProduce.BackColor = System.Drawing.SystemColors.Window
		Me.txtProduce.CausesValidation = True
		Me.txtProduce.Enabled = True
		Me.txtProduce.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtProduce.HideSelection = True
		Me.txtProduce.ReadOnly = False
		Me.txtProduce.Maxlength = 0
		Me.txtProduce.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtProduce.MultiLine = False
		Me.txtProduce.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtProduce.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtProduce.TabStop = True
		Me.txtProduce.Visible = True
		Me.txtProduce.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtProduce.Name = "txtProduce"
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "Done"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Size = New System.Drawing.Size(49, 25)
		Me.cmdOK.Location = New System.Drawing.Point(320, 8)
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
		Me.txtOrigin.AutoSize = False
		Me.txtOrigin.Size = New System.Drawing.Size(33, 19)
		Me.txtOrigin.Location = New System.Drawing.Point(72, 8)
		Me.txtOrigin.TabIndex = 0
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
		Me._lblMaterial_3.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._lblMaterial_3.Text = "oil"
		Me._lblMaterial_3.ForeColor = System.Drawing.SystemColors.WindowText
		Me._lblMaterial_3.Size = New System.Drawing.Size(33, 17)
		Me._lblMaterial_3.Location = New System.Drawing.Point(56, 128)
		Me._lblMaterial_3.TabIndex = 22
		Me._lblMaterial_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblMaterial_3.BackColor = System.Drawing.SystemColors.Control
		Me._lblMaterial_3.Enabled = True
		Me._lblMaterial_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblMaterial_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblMaterial_3.UseMnemonic = True
		Me._lblMaterial_3.Visible = True
		Me._lblMaterial_3.AutoSize = False
		Me._lblMaterial_3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._lblMaterial_3.Name = "_lblMaterial_3"
		Me._lblMaterial_2.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._lblMaterial_2.Text = "lcm"
		Me._lblMaterial_2.ForeColor = System.Drawing.SystemColors.WindowText
		Me._lblMaterial_2.Size = New System.Drawing.Size(33, 17)
		Me._lblMaterial_2.Location = New System.Drawing.Point(56, 104)
		Me._lblMaterial_2.TabIndex = 21
		Me._lblMaterial_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblMaterial_2.BackColor = System.Drawing.SystemColors.Control
		Me._lblMaterial_2.Enabled = True
		Me._lblMaterial_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblMaterial_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblMaterial_2.UseMnemonic = True
		Me._lblMaterial_2.Visible = True
		Me._lblMaterial_2.AutoSize = False
		Me._lblMaterial_2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._lblMaterial_2.Name = "_lblMaterial_2"
		Me._lblMaterial_1.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._lblMaterial_1.Text = "hcm"
		Me._lblMaterial_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._lblMaterial_1.Size = New System.Drawing.Size(33, 17)
		Me._lblMaterial_1.Location = New System.Drawing.Point(56, 80)
		Me._lblMaterial_1.TabIndex = 20
		Me._lblMaterial_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblMaterial_1.BackColor = System.Drawing.SystemColors.Control
		Me._lblMaterial_1.Enabled = True
		Me._lblMaterial_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblMaterial_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblMaterial_1.UseMnemonic = True
		Me._lblMaterial_1.Visible = True
		Me._lblMaterial_1.AutoSize = False
		Me._lblMaterial_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._lblMaterial_1.Name = "_lblMaterial_1"
		Me._lblMaterial_0.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me._lblMaterial_0.Text = "avail"
		Me._lblMaterial_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._lblMaterial_0.Size = New System.Drawing.Size(33, 17)
		Me._lblMaterial_0.Location = New System.Drawing.Point(56, 56)
		Me._lblMaterial_0.TabIndex = 19
		Me._lblMaterial_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblMaterial_0.BackColor = System.Drawing.SystemColors.Control
		Me._lblMaterial_0.Enabled = True
		Me._lblMaterial_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblMaterial_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblMaterial_0.UseMnemonic = True
		Me._lblMaterial_0.Visible = True
		Me._lblMaterial_0.AutoSize = False
		Me._lblMaterial_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me._lblMaterial_0.Name = "_lblMaterial_0"
		Me._lblRequired_3.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._lblRequired_3.Text = "9999"
		Me._lblRequired_3.Size = New System.Drawing.Size(33, 17)
		Me._lblRequired_3.Location = New System.Drawing.Point(104, 128)
		Me._lblRequired_3.TabIndex = 18
		Me._lblRequired_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblRequired_3.BackColor = System.Drawing.SystemColors.Control
		Me._lblRequired_3.Enabled = True
		Me._lblRequired_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblRequired_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblRequired_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblRequired_3.UseMnemonic = True
		Me._lblRequired_3.Visible = True
		Me._lblRequired_3.AutoSize = False
		Me._lblRequired_3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._lblRequired_3.Name = "_lblRequired_3"
		Me._lblRequired_2.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._lblRequired_2.Text = "9999"
		Me._lblRequired_2.Size = New System.Drawing.Size(33, 17)
		Me._lblRequired_2.Location = New System.Drawing.Point(104, 104)
		Me._lblRequired_2.TabIndex = 17
		Me._lblRequired_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblRequired_2.BackColor = System.Drawing.SystemColors.Control
		Me._lblRequired_2.Enabled = True
		Me._lblRequired_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblRequired_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblRequired_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblRequired_2.UseMnemonic = True
		Me._lblRequired_2.Visible = True
		Me._lblRequired_2.AutoSize = False
		Me._lblRequired_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._lblRequired_2.Name = "_lblRequired_2"
		Me._lblRequired_1.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._lblRequired_1.Text = "9999"
		Me._lblRequired_1.Size = New System.Drawing.Size(33, 17)
		Me._lblRequired_1.Location = New System.Drawing.Point(104, 80)
		Me._lblRequired_1.TabIndex = 16
		Me._lblRequired_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblRequired_1.BackColor = System.Drawing.SystemColors.Control
		Me._lblRequired_1.Enabled = True
		Me._lblRequired_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblRequired_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblRequired_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblRequired_1.UseMnemonic = True
		Me._lblRequired_1.Visible = True
		Me._lblRequired_1.AutoSize = False
		Me._lblRequired_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._lblRequired_1.Name = "_lblRequired_1"
		Me.Label3.Text = "Sector"
		Me.Label3.Size = New System.Drawing.Size(57, 17)
		Me.Label3.Location = New System.Drawing.Point(8, 8)
		Me.Label3.TabIndex = 15
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
		Me._lblRequired_0.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._lblRequired_0.Text = "9999"
		Me._lblRequired_0.Size = New System.Drawing.Size(33, 17)
		Me._lblRequired_0.Location = New System.Drawing.Point(104, 56)
		Me._lblRequired_0.TabIndex = 14
		Me._lblRequired_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblRequired_0.BackColor = System.Drawing.SystemColors.Control
		Me._lblRequired_0.Enabled = True
		Me._lblRequired_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblRequired_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblRequired_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblRequired_0.UseMnemonic = True
		Me._lblRequired_0.Visible = True
		Me._lblRequired_0.AutoSize = False
		Me._lblRequired_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._lblRequired_0.Name = "_lblRequired_0"
		Me.Label2.Text = "Required"
		Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(57, 17)
		Me.Label2.Location = New System.Drawing.Point(96, 32)
		Me.Label2.TabIndex = 13
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
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label1.Text = "Current"
		Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Size = New System.Drawing.Size(41, 17)
		Me.Label1.Location = New System.Drawing.Point(8, 32)
		Me.Label1.TabIndex = 12
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
		Me._lblResource_3.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._lblResource_3.Text = "9999"
		Me._lblResource_3.Size = New System.Drawing.Size(25, 17)
		Me._lblResource_3.Location = New System.Drawing.Point(16, 128)
		Me._lblResource_3.TabIndex = 11
		Me._lblResource_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblResource_3.BackColor = System.Drawing.SystemColors.Control
		Me._lblResource_3.Enabled = True
		Me._lblResource_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblResource_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblResource_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblResource_3.UseMnemonic = True
		Me._lblResource_3.Visible = True
		Me._lblResource_3.AutoSize = False
		Me._lblResource_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblResource_3.Name = "_lblResource_3"
		Me._lblResource_2.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._lblResource_2.Text = "9999"
		Me._lblResource_2.Size = New System.Drawing.Size(25, 17)
		Me._lblResource_2.Location = New System.Drawing.Point(16, 104)
		Me._lblResource_2.TabIndex = 10
		Me._lblResource_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblResource_2.BackColor = System.Drawing.SystemColors.Control
		Me._lblResource_2.Enabled = True
		Me._lblResource_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblResource_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblResource_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblResource_2.UseMnemonic = True
		Me._lblResource_2.Visible = True
		Me._lblResource_2.AutoSize = False
		Me._lblResource_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblResource_2.Name = "_lblResource_2"
		Me._lblResource_1.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._lblResource_1.Text = "9999"
		Me._lblResource_1.Size = New System.Drawing.Size(25, 17)
		Me._lblResource_1.Location = New System.Drawing.Point(16, 80)
		Me._lblResource_1.TabIndex = 9
		Me._lblResource_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblResource_1.BackColor = System.Drawing.SystemColors.Control
		Me._lblResource_1.Enabled = True
		Me._lblResource_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblResource_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblResource_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblResource_1.UseMnemonic = True
		Me._lblResource_1.Visible = True
		Me._lblResource_1.AutoSize = False
		Me._lblResource_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblResource_1.Name = "_lblResource_1"
		Me._lblResource_0.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._lblResource_0.Text = "9999"
		Me._lblResource_0.Size = New System.Drawing.Size(25, 17)
		Me._lblResource_0.Location = New System.Drawing.Point(16, 56)
		Me._lblResource_0.TabIndex = 8
		Me._lblResource_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblResource_0.BackColor = System.Drawing.SystemColors.Control
		Me._lblResource_0.Enabled = True
		Me._lblResource_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblResource_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblResource_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblResource_0.UseMnemonic = True
		Me._lblResource_0.Visible = True
		Me._lblResource_0.AutoSize = False
		Me._lblResource_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblResource_0.Name = "_lblResource_0"
		Me.Line12.BorderWidth = 3
		Me.Line12.X1 = 192
		Me.Line12.X2 = 200
		Me.Line12.Y1 = 64
		Me.Line12.Y2 = 72
		Me.Line12.BorderColor = System.Drawing.SystemColors.WindowText
		Me.Line12.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line12.Visible = True
		Me.Line12.Name = "Line12"
		Me.Line11.BorderWidth = 3
		Me.Line11.X1 = 192
		Me.Line11.X2 = 184
		Me.Line11.Y1 = 64
		Me.Line11.Y2 = 72
		Me.Line11.BorderColor = System.Drawing.SystemColors.WindowText
		Me.Line11.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line11.Visible = True
		Me.Line11.Name = "Line11"
		Me.Line10.BorderWidth = 4
		Me.Line10.X1 = 240
		Me.Line10.X2 = 232
		Me.Line10.Y1 = 64
		Me.Line10.Y2 = 72
		Me.Line10.BorderColor = System.Drawing.SystemColors.WindowText
		Me.Line10.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line10.Visible = True
		Me.Line10.Name = "Line10"
		Me.Line9.BorderWidth = 4
		Me.Line9.X1 = 232
		Me.Line9.X2 = 240
		Me.Line9.Y1 = 56
		Me.Line9.Y2 = 64
		Me.Line9.BorderColor = System.Drawing.SystemColors.WindowText
		Me.Line9.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line9.Visible = True
		Me.Line9.Name = "Line9"
		Me.Line8.BorderWidth = 2
		Me.Line8.X1 = 192
		Me.Line8.X2 = 192
		Me.Line8.Y1 = 112
		Me.Line8.Y2 = 136
		Me.Line8.BorderColor = System.Drawing.SystemColors.WindowText
		Me.Line8.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line8.Visible = True
		Me.Line8.Name = "Line8"
		Me.Line7.BorderWidth = 2
		Me.Line7.X1 = 144
		Me.Line7.X2 = 192
		Me.Line7.Y1 = 136
		Me.Line7.Y2 = 136
		Me.Line7.BorderColor = System.Drawing.SystemColors.WindowText
		Me.Line7.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line7.Visible = True
		Me.Line7.Name = "Line7"
		Me.Line6.BorderWidth = 2
		Me.Line6.X1 = 192
		Me.Line6.X2 = 192
		Me.Line6.Y1 = 88
		Me.Line6.Y2 = 112
		Me.Line6.BorderColor = System.Drawing.SystemColors.WindowText
		Me.Line6.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line6.Visible = True
		Me.Line6.Name = "Line6"
		Me.Line5.BorderWidth = 2
		Me.Line5.X1 = 144
		Me.Line5.X2 = 192
		Me.Line5.Y1 = 112
		Me.Line5.Y2 = 112
		Me.Line5.BorderColor = System.Drawing.SystemColors.WindowText
		Me.Line5.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line5.Visible = True
		Me.Line5.Name = "Line5"
		Me.Line4.BorderWidth = 2
		Me.Line4.X1 = 192
		Me.Line4.X2 = 192
		Me.Line4.Y1 = 64
		Me.Line4.Y2 = 88
		Me.Line4.BorderColor = System.Drawing.SystemColors.WindowText
		Me.Line4.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line4.Visible = True
		Me.Line4.Name = "Line4"
		Me.Line3.BorderWidth = 2
		Me.Line3.X1 = 144
		Me.Line3.X2 = 192
		Me.Line3.Y1 = 88
		Me.Line3.Y2 = 88
		Me.Line3.BorderColor = System.Drawing.SystemColors.WindowText
		Me.Line3.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line3.Visible = True
		Me.Line3.Name = "Line3"
		Me.Line2.BorderWidth = 2
		Me.Line2.X1 = 144
		Me.Line2.X2 = 240
		Me.Line2.Y1 = 64
		Me.Line2.Y2 = 64
		Me.Line2.BorderColor = System.Drawing.SystemColors.WindowText
		Me.Line2.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.Line2.Visible = True
		Me.Line2.Name = "Line2"
		Me.lblProduct.Text = "guns"
		Me.lblProduct.Size = New System.Drawing.Size(41, 17)
		Me.lblProduct.Location = New System.Drawing.Point(320, 56)
		Me.lblProduct.TabIndex = 5
		Me.lblProduct.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblProduct.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblProduct.BackColor = System.Drawing.SystemColors.Control
		Me.lblProduct.Enabled = True
		Me.lblProduct.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblProduct.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblProduct.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblProduct.UseMnemonic = True
		Me.lblProduct.Visible = True
		Me.lblProduct.AutoSize = False
		Me.lblProduct.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblProduct.Name = "lblProduct"
		Me.Label4.Text = "Product"
		Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.Size = New System.Drawing.Size(65, 17)
		Me.Label4.Location = New System.Drawing.Point(256, 32)
		Me.Label4.TabIndex = 4
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
		Me.lblType.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblType.Size = New System.Drawing.Size(193, 17)
		Me.lblType.Location = New System.Drawing.Point(112, 8)
		Me.lblType.TabIndex = 1
		Me.lblType.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblType.BackColor = System.Drawing.SystemColors.Control
		Me.lblType.Enabled = True
		Me.lblType.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblType.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblType.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblType.UseMnemonic = True
		Me.lblType.Visible = True
		Me.lblType.AutoSize = False
		Me.lblType.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblType.Name = "lblType"
		Me.Controls.Add(rtbReport)
		Me.Controls.Add(cmdPE)
		Me.Controls.Add(txtProduce)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(txtOrigin)
		Me.Controls.Add(_lblMaterial_3)
		Me.Controls.Add(_lblMaterial_2)
		Me.Controls.Add(_lblMaterial_1)
		Me.Controls.Add(_lblMaterial_0)
		Me.Controls.Add(_lblRequired_3)
		Me.Controls.Add(_lblRequired_2)
		Me.Controls.Add(_lblRequired_1)
		Me.Controls.Add(Label3)
		Me.Controls.Add(_lblRequired_0)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		Me.Controls.Add(_lblResource_3)
		Me.Controls.Add(_lblResource_2)
		Me.Controls.Add(_lblResource_1)
		Me.Controls.Add(_lblResource_0)
		Me.ShapeContainer1.Shapes.Add(Line12)
		Me.ShapeContainer1.Shapes.Add(Line11)
		Me.ShapeContainer1.Shapes.Add(Line10)
		Me.ShapeContainer1.Shapes.Add(Line9)
		Me.ShapeContainer1.Shapes.Add(Line8)
		Me.ShapeContainer1.Shapes.Add(Line7)
		Me.ShapeContainer1.Shapes.Add(Line6)
		Me.ShapeContainer1.Shapes.Add(Line5)
		Me.ShapeContainer1.Shapes.Add(Line4)
		Me.ShapeContainer1.Shapes.Add(Line3)
		Me.ShapeContainer1.Shapes.Add(Line2)
		Me.Controls.Add(lblProduct)
		Me.Controls.Add(Label4)
		Me.Controls.Add(lblType)
		Me.Controls.Add(ShapeContainer1)
		Me.lblMaterial.SetIndex(_lblMaterial_3, CType(3, Short))
		Me.lblMaterial.SetIndex(_lblMaterial_2, CType(2, Short))
		Me.lblMaterial.SetIndex(_lblMaterial_1, CType(1, Short))
		Me.lblMaterial.SetIndex(_lblMaterial_0, CType(0, Short))
		Me.lblRequired.SetIndex(_lblRequired_3, CType(3, Short))
		Me.lblRequired.SetIndex(_lblRequired_2, CType(2, Short))
		Me.lblRequired.SetIndex(_lblRequired_1, CType(1, Short))
		Me.lblRequired.SetIndex(_lblRequired_0, CType(0, Short))
		Me.lblResource.SetIndex(_lblResource_3, CType(3, Short))
		Me.lblResource.SetIndex(_lblResource_2, CType(2, Short))
		Me.lblResource.SetIndex(_lblResource_1, CType(1, Short))
		Me.lblResource.SetIndex(_lblResource_0, CType(0, Short))
		CType(Me.lblResource, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lblRequired, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lblMaterial, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class