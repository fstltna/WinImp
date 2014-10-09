<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmAbout
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
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	Public WithEvents Picture1 As System.Windows.Forms.PictureBox
	Public WithEvents _lblAck_4 As System.Windows.Forms.Label
	Public WithEvents _lblAck_3 As System.Windows.Forms.Label
	Public WithEvents _lblAck_2 As System.Windows.Forms.Label
	Public WithEvents _lblAck_1 As System.Windows.Forms.Label
	Public WithEvents _lblAck_0 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents cmdSysInfo As System.Windows.Forms.Button
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents _Line1_1 As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents lblDescription As System.Windows.Forms.Label
	Public WithEvents lblTitle As System.Windows.Forms.Label
	Public WithEvents lblVersion As System.Windows.Forms.Label
	Public WithEvents lblDisclaimer As System.Windows.Forms.Label
	Public WithEvents Line1 As LineShapeArray
	Public WithEvents lblAck As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents ShapeContainer1 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAbout))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer
		Me.Timer1 = New System.Windows.Forms.Timer(components)
		Me.Picture1 = New System.Windows.Forms.PictureBox
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me._lblAck_4 = New System.Windows.Forms.Label
		Me._lblAck_3 = New System.Windows.Forms.Label
		Me._lblAck_2 = New System.Windows.Forms.Label
		Me._lblAck_1 = New System.Windows.Forms.Label
		Me._lblAck_0 = New System.Windows.Forms.Label
		Me.cmdOK = New System.Windows.Forms.Button
		Me.cmdSysInfo = New System.Windows.Forms.Button
		Me.Label1 = New System.Windows.Forms.Label
		Me._Line1_1 = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.lblDescription = New System.Windows.Forms.Label
		Me.lblTitle = New System.Windows.Forms.Label
		Me.lblVersion = New System.Windows.Forms.Label
		Me.lblDisclaimer = New System.Windows.Forms.Label
		Me.Line1 = New LineShapeArray(components)
		Me.lblAck = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.Frame1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.Line1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lblAck, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "About MyApp"
		Me.ClientSize = New System.Drawing.Size(553, 362)
		Me.Location = New System.Drawing.Point(156, 129)
		Me.Icon = CType(resources.GetObject("frmAbout.Icon"), System.Drawing.Icon)
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
		Me.Name = "frmAbout"
		Me.Timer1.Enabled = False
		Me.Timer1.Interval = 1
		Me.Picture1.Size = New System.Drawing.Size(193, 249)
		Me.Picture1.Location = New System.Drawing.Point(8, 16)
		Me.Picture1.Image = CType(resources.GetObject("Picture1.Image"), System.Drawing.Image)
		Me.Picture1.TabIndex = 10
		Me.Picture1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Picture1.Dock = System.Windows.Forms.DockStyle.None
		Me.Picture1.BackColor = System.Drawing.SystemColors.Control
		Me.Picture1.CausesValidation = True
		Me.Picture1.Enabled = True
		Me.Picture1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Picture1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Picture1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Picture1.TabStop = True
		Me.Picture1.Visible = True
		Me.Picture1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.Picture1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Picture1.Name = "Picture1"
		Me.Frame1.Text = "Acknowledgements"
		Me.Frame1.Size = New System.Drawing.Size(337, 113)
		Me.Frame1.Location = New System.Drawing.Point(208, 160)
		Me.Frame1.TabIndex = 6
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me._lblAck_4.Text = "Updates and maintence since v2.3.0: Ron Koenderink (RimFire)"
		Me._lblAck_4.Size = New System.Drawing.Size(321, 17)
		Me._lblAck_4.Location = New System.Drawing.Point(8, 16)
		Me._lblAck_4.TabIndex = 13
		Me._lblAck_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblAck_4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblAck_4.BackColor = System.Drawing.SystemColors.Control
		Me._lblAck_4.Enabled = True
		Me._lblAck_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblAck_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblAck_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblAck_4.UseMnemonic = True
		Me._lblAck_4.Visible = True
		Me._lblAck_4.AutoSize = False
		Me._lblAck_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblAck_4.Name = "_lblAck_4"
		Me._lblAck_3.Text = "Updates and maintence since v2.2.0: Daniel Kiracofe (Bwian)"
		Me._lblAck_3.Size = New System.Drawing.Size(321, 17)
		Me._lblAck_3.Location = New System.Drawing.Point(8, 32)
		Me._lblAck_3.TabIndex = 12
		Me._lblAck_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblAck_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblAck_3.BackColor = System.Drawing.SystemColors.Control
		Me._lblAck_3.Enabled = True
		Me._lblAck_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblAck_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblAck_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblAck_3.UseMnemonic = True
		Me._lblAck_3.Visible = True
		Me._lblAck_3.AutoSize = False
		Me._lblAck_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblAck_3.Name = "_lblAck_3"
		Me._lblAck_2.Text = "Much thanks to Jon Butler for the original WS Empire code which provided a good server connection template."
		Me._lblAck_2.Size = New System.Drawing.Size(321, 25)
		Me._lblAck_2.Location = New System.Drawing.Point(8, 80)
		Me._lblAck_2.TabIndex = 9
		Me._lblAck_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblAck_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblAck_2.BackColor = System.Drawing.SystemColors.Control
		Me._lblAck_2.Enabled = True
		Me._lblAck_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblAck_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblAck_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblAck_2.UseMnemonic = True
		Me._lblAck_2.Visible = True
		Me._lblAck_2.AutoSize = False
		Me._lblAck_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblAck_2.Name = "_lblAck_2"
		Me._lblAck_1.Text = "Collaboration and Testing:  Ken McDonald (also Escher)"
		Me._lblAck_1.Size = New System.Drawing.Size(321, 17)
		Me._lblAck_1.Location = New System.Drawing.Point(8, 64)
		Me._lblAck_1.TabIndex = 8
		Me._lblAck_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblAck_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblAck_1.BackColor = System.Drawing.SystemColors.Control
		Me._lblAck_1.Enabled = True
		Me._lblAck_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblAck_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblAck_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblAck_1.UseMnemonic = True
		Me._lblAck_1.Visible = True
		Me._lblAck_1.AutoSize = False
		Me._lblAck_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblAck_1.Name = "_lblAck_1"
		Me._lblAck_0.Text = "Original Design and Programming:  Jim Simons (Escher)"
		Me._lblAck_0.Size = New System.Drawing.Size(321, 17)
		Me._lblAck_0.Location = New System.Drawing.Point(8, 48)
		Me._lblAck_0.TabIndex = 7
		Me._lblAck_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblAck_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblAck_0.BackColor = System.Drawing.SystemColors.Control
		Me._lblAck_0.Enabled = True
		Me._lblAck_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblAck_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblAck_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblAck_0.UseMnemonic = True
		Me._lblAck_0.Visible = True
		Me._lblAck_0.AutoSize = False
		Me._lblAck_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblAck_0.Name = "_lblAck_0"
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdOK
		Me.cmdOK.Text = "OK"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Size = New System.Drawing.Size(84, 23)
		Me.cmdOK.Location = New System.Drawing.Point(464, 296)
		Me.cmdOK.TabIndex = 0
		Me.cmdOK.Visible = False
		Me.cmdOK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.cmdSysInfo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdSysInfo.Text = "&System Info..."
		Me.cmdSysInfo.Size = New System.Drawing.Size(83, 23)
		Me.cmdSysInfo.Location = New System.Drawing.Point(464, 328)
		Me.cmdSysInfo.TabIndex = 1
		Me.cmdSysInfo.Visible = False
		Me.cmdSysInfo.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdSysInfo.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSysInfo.CausesValidation = True
		Me.cmdSysInfo.Enabled = True
		Me.cmdSysInfo.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSysInfo.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSysInfo.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSysInfo.TabStop = True
		Me.cmdSysInfo.Name = "cmdSysInfo"
		Me.Label1.Text = "Questions, comments, and bug reports should be submitted to WinACE group at sourceforge.net/projects/winace"
		Me.Label1.Font = New System.Drawing.Font("Arial", 7.8!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.ForeColor = System.Drawing.Color.Black
		Me.Label1.Size = New System.Drawing.Size(339, 29)
		Me.Label1.Location = New System.Drawing.Point(208, 128)
		Me.Label1.TabIndex = 11
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.BackColor = System.Drawing.SystemColors.Control
		Me.Label1.Enabled = True
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me._Line1_1.BorderColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me._Line1_1.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me._Line1_1.X1 = 8
		Me._Line1_1.X2 = 504
		Me._Line1_1.Y1 = 194
		Me._Line1_1.Y2 = 194
		Me._Line1_1.BorderWidth = 1
		Me._Line1_1.Visible = True
		Me._Line1_1.Name = "_Line1_1"
		Me.lblDescription.Text = "WinACE is a Windows 95/98/2000/NT/XP GUI Empire Client compatible with version 4.3.23 of the Empire Server as published by Wolfpack.  It was developed in Visual Basic 6.0 to provide a GUI client for the Windows environment."
		Me.lblDescription.ForeColor = System.Drawing.Color.Black
		Me.lblDescription.Size = New System.Drawing.Size(339, 53)
		Me.lblDescription.Location = New System.Drawing.Point(208, 72)
		Me.lblDescription.TabIndex = 2
		Me.lblDescription.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblDescription.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblDescription.BackColor = System.Drawing.SystemColors.Control
		Me.lblDescription.Enabled = True
		Me.lblDescription.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblDescription.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblDescription.UseMnemonic = True
		Me.lblDescription.Visible = True
		Me.lblDescription.AutoSize = False
		Me.lblDescription.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblDescription.Name = "lblDescription"
		Me.lblTitle.Text = "Application Title"
		Me.lblTitle.Font = New System.Drawing.Font("Arial", 9.6!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblTitle.ForeColor = System.Drawing.Color.Black
		Me.lblTitle.Size = New System.Drawing.Size(323, 32)
		Me.lblTitle.Location = New System.Drawing.Point(208, 8)
		Me.lblTitle.TabIndex = 4
		Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblTitle.BackColor = System.Drawing.SystemColors.Control
		Me.lblTitle.Enabled = True
		Me.lblTitle.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblTitle.UseMnemonic = True
		Me.lblTitle.Visible = True
		Me.lblTitle.AutoSize = False
		Me.lblTitle.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblTitle.Name = "lblTitle"
		Me.lblVersion.Text = "Version"
		Me.lblVersion.Size = New System.Drawing.Size(315, 15)
		Me.lblVersion.Location = New System.Drawing.Point(208, 48)
		Me.lblVersion.TabIndex = 5
		Me.lblVersion.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblVersion.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblVersion.BackColor = System.Drawing.SystemColors.Control
		Me.lblVersion.Enabled = True
		Me.lblVersion.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblVersion.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblVersion.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblVersion.UseMnemonic = True
		Me.lblVersion.Visible = True
		Me.lblVersion.AutoSize = False
		Me.lblVersion.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblVersion.Name = "lblVersion"
		Me.lblDisclaimer.Text = "WinACE is copyrighted (1997-2001) by James A. Simons, and remains his intellectual property.  However, the author has granted permission for this version of WinACE to be freely used, copied and distributed, as long as no fee is charged for its use or distribution, all files are distributed, and all the original acknowlegements are left intact."
		Me.lblDisclaimer.ForeColor = System.Drawing.Color.Black
		Me.lblDisclaimer.Size = New System.Drawing.Size(449, 55)
		Me.lblDisclaimer.Location = New System.Drawing.Point(8, 296)
		Me.lblDisclaimer.TabIndex = 3
		Me.lblDisclaimer.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblDisclaimer.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblDisclaimer.BackColor = System.Drawing.SystemColors.Control
		Me.lblDisclaimer.Enabled = True
		Me.lblDisclaimer.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblDisclaimer.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblDisclaimer.UseMnemonic = True
		Me.lblDisclaimer.Visible = True
		Me.lblDisclaimer.AutoSize = False
		Me.lblDisclaimer.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblDisclaimer.Name = "lblDisclaimer"
		Me.Controls.Add(Picture1)
		Me.Controls.Add(Frame1)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(cmdSysInfo)
		Me.Controls.Add(Label1)
		Me.ShapeContainer1.Shapes.Add(_Line1_1)
		Me.Controls.Add(lblDescription)
		Me.Controls.Add(lblTitle)
		Me.Controls.Add(lblVersion)
		Me.Controls.Add(lblDisclaimer)
		Me.Controls.Add(ShapeContainer1)
		Me.Frame1.Controls.Add(_lblAck_4)
		Me.Frame1.Controls.Add(_lblAck_3)
		Me.Frame1.Controls.Add(_lblAck_2)
		Me.Frame1.Controls.Add(_lblAck_1)
		Me.Frame1.Controls.Add(_lblAck_0)
		Me.Line1.SetIndex(_Line1_1, CType(1, Short))
		Me.lblAck.SetIndex(_lblAck_4, CType(4, Short))
		Me.lblAck.SetIndex(_lblAck_3, CType(3, Short))
		Me.lblAck.SetIndex(_lblAck_2, CType(2, Short))
		Me.lblAck.SetIndex(_lblAck_1, CType(1, Short))
		Me.lblAck.SetIndex(_lblAck_0, CType(0, Short))
		CType(Me.lblAck, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.Line1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Frame1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class