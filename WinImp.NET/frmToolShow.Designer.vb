<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmToolShow
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
	Public WithEvents _Toolbar1_Button1 As System.Windows.Forms.ToolStripButton
	Public WithEvents _Toolbar1_Button2 As System.Windows.Forms.ToolStripButton
	Public WithEvents _Toolbar1_Button3 As System.Windows.Forms.ToolStripButton
	Public WithEvents _Toolbar1_Button4 As System.Windows.Forms.ToolStripButton
	Public WithEvents _Toolbar1_Button5 As System.Windows.Forms.ToolStripButton
	Public WithEvents _Toolbar1_Button6 As System.Windows.Forms.ToolStripButton
	Public WithEvents _Toolbar1_Button7 As System.Windows.Forms.ToolStripButton
	Public WithEvents _Toolbar1_Button8 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents _Toolbar1_Button9 As System.Windows.Forms.ToolStripButton
	Public WithEvents _Toolbar1_Button10 As System.Windows.Forms.ToolStripButton
	Public WithEvents _Toolbar1_Button11 As System.Windows.Forms.ToolStripButton
	Public WithEvents _Toolbar1_Button12 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents _Toolbar1_Button13 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents _Toolbar1_Button14 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents Toolbar1 As System.Windows.Forms.ToolStrip
	Public WithEvents txtTech As System.Windows.Forms.TextBox
	Public WithEvents cmdClear As System.Windows.Forms.Button
	Public WithEvents cmdRefreshUnits As System.Windows.Forms.Button
	Public WithEvents cmdRefreshAll As System.Windows.Forms.Button
	Public WithEvents TextBox1 As System.Windows.Forms.RichTextBox
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents cmdRefresh As System.Windows.Forms.Button
	Public WithEvents grid1 As AxMSFlexGridLib.AxMSFlexGrid
	Public WithEvents _sb1_Panel1 As System.Windows.Forms.ToolStripStatusLabel
	Public WithEvents sb1 As System.Windows.Forms.StatusStrip
	Public WithEvents ImageList1 As System.Windows.Forms.ImageList
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmToolShow))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Toolbar1 = New System.Windows.Forms.ToolStrip
		Me._Toolbar1_Button1 = New System.Windows.Forms.ToolStripButton
		Me._Toolbar1_Button2 = New System.Windows.Forms.ToolStripButton
		Me._Toolbar1_Button3 = New System.Windows.Forms.ToolStripButton
		Me._Toolbar1_Button4 = New System.Windows.Forms.ToolStripButton
		Me._Toolbar1_Button5 = New System.Windows.Forms.ToolStripButton
		Me._Toolbar1_Button6 = New System.Windows.Forms.ToolStripButton
		Me._Toolbar1_Button7 = New System.Windows.Forms.ToolStripButton
		Me._Toolbar1_Button8 = New System.Windows.Forms.ToolStripSeparator
		Me._Toolbar1_Button9 = New System.Windows.Forms.ToolStripButton
		Me._Toolbar1_Button10 = New System.Windows.Forms.ToolStripButton
		Me._Toolbar1_Button11 = New System.Windows.Forms.ToolStripButton
		Me._Toolbar1_Button12 = New System.Windows.Forms.ToolStripSeparator
		Me._Toolbar1_Button13 = New System.Windows.Forms.ToolStripSeparator
		Me._Toolbar1_Button14 = New System.Windows.Forms.ToolStripSeparator
		Me.txtTech = New System.Windows.Forms.TextBox
		Me.cmdClear = New System.Windows.Forms.Button
		Me.cmdRefreshUnits = New System.Windows.Forms.Button
		Me.cmdRefreshAll = New System.Windows.Forms.Button
		Me.TextBox1 = New System.Windows.Forms.RichTextBox
		Me.cmdOK = New System.Windows.Forms.Button
		Me.cmdRefresh = New System.Windows.Forms.Button
		Me.grid1 = New AxMSFlexGridLib.AxMSFlexGrid
		Me.sb1 = New System.Windows.Forms.StatusStrip
		Me._sb1_Panel1 = New System.Windows.Forms.ToolStripStatusLabel
		Me.ImageList1 = New System.Windows.Forms.ImageList
		Me.Label1 = New System.Windows.Forms.Label
		Me.Toolbar1.SuspendLayout()
		Me.sb1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.grid1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.Text = "Show Tool"
		Me.ClientSize = New System.Drawing.Size(528, 376)
		Me.Location = New System.Drawing.Point(4, 24)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmToolShow"
		Me.Toolbar1.ShowItemToolTips = True
		Me.Toolbar1.Size = New System.Drawing.Size(560, 23)
		Me.Toolbar1.Location = New System.Drawing.Point(0, 0)
		Me.Toolbar1.TabIndex = 1
		Me.Toolbar1.ImageList = ImageList1
		Me.Toolbar1.Name = "Toolbar1"
		Me._Toolbar1_Button1.Size = New System.Drawing.Size(27, 19)
		Me._Toolbar1_Button1.AutoSize = False
		Me._Toolbar1_Button1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._Toolbar1_Button1.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._Toolbar1_Button1.Name = "Ship"
		Me._Toolbar1_Button1.ToolTipText = "Show Ships"
		Me._Toolbar1_Button1.ImageIndex = 2
		Me._Toolbar1_Button1.Checked = True
		Me._Toolbar1_Button1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._Toolbar1_Button2.Size = New System.Drawing.Size(27, 19)
		Me._Toolbar1_Button2.AutoSize = False
		Me._Toolbar1_Button2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._Toolbar1_Button2.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._Toolbar1_Button2.Name = "Land"
		Me._Toolbar1_Button2.ToolTipText = "Show Lands"
		Me._Toolbar1_Button2.ImageIndex = 0
		Me._Toolbar1_Button2.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._Toolbar1_Button3.Size = New System.Drawing.Size(27, 19)
		Me._Toolbar1_Button3.AutoSize = False
		Me._Toolbar1_Button3.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._Toolbar1_Button3.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._Toolbar1_Button3.Name = "Plane"
		Me._Toolbar1_Button3.ToolTipText = "Show Planes"
		Me._Toolbar1_Button3.ImageIndex = 1
		Me._Toolbar1_Button3.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._Toolbar1_Button4.Size = New System.Drawing.Size(27, 19)
		Me._Toolbar1_Button4.AutoSize = False
		Me._Toolbar1_Button4.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._Toolbar1_Button4.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._Toolbar1_Button4.Name = "Nuke"
		Me._Toolbar1_Button4.ToolTipText = "Show Nukes"
		Me._Toolbar1_Button4.ImageIndex = 8
		Me._Toolbar1_Button4.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._Toolbar1_Button5.Size = New System.Drawing.Size(27, 19)
		Me._Toolbar1_Button5.AutoSize = False
		Me._Toolbar1_Button5.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._Toolbar1_Button5.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._Toolbar1_Button5.Name = "Sector"
		Me._Toolbar1_Button5.ToolTipText = "Show Sector"
		Me._Toolbar1_Button5.ImageIndex = 3
		Me._Toolbar1_Button5.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._Toolbar1_Button6.Size = New System.Drawing.Size(27, 19)
		Me._Toolbar1_Button6.AutoSize = False
		Me._Toolbar1_Button6.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._Toolbar1_Button6.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._Toolbar1_Button6.Name = "Bridge"
		Me._Toolbar1_Button6.ToolTipText = "Show Bridges and Towers"
		Me._Toolbar1_Button6.ImageIndex = 6
		Me._Toolbar1_Button6.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._Toolbar1_Button7.Size = New System.Drawing.Size(27, 19)
		Me._Toolbar1_Button7.AutoSize = False
		Me._Toolbar1_Button7.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._Toolbar1_Button7.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._Toolbar1_Button7.Name = "Item"
		Me._Toolbar1_Button7.ToolTipText = "Show Items"
		Me._Toolbar1_Button7.ImageIndex = 9
		Me._Toolbar1_Button7.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._Toolbar1_Button8.Size = New System.Drawing.Size(27, 19)
		Me._Toolbar1_Button8.AutoSize = False
		Me._Toolbar1_Button8.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._Toolbar1_Button8.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._Toolbar1_Button8.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._Toolbar1_Button9.Size = New System.Drawing.Size(27, 19)
		Me._Toolbar1_Button9.AutoSize = False
		Me._Toolbar1_Button9.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._Toolbar1_Button9.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._Toolbar1_Button9.Name = "Build"
		Me._Toolbar1_Button9.ToolTipText = "Show Builds"
		Me._Toolbar1_Button9.ImageIndex = 7
		Me._Toolbar1_Button9.Checked = True
		Me._Toolbar1_Button9.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._Toolbar1_Button10.Size = New System.Drawing.Size(27, 19)
		Me._Toolbar1_Button10.AutoSize = False
		Me._Toolbar1_Button10.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._Toolbar1_Button10.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._Toolbar1_Button10.Name = "Stat"
		Me._Toolbar1_Button10.ToolTipText = "Show Statistics"
		Me._Toolbar1_Button10.ImageIndex = 4
		Me._Toolbar1_Button10.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._Toolbar1_Button11.Size = New System.Drawing.Size(27, 19)
		Me._Toolbar1_Button11.AutoSize = False
		Me._Toolbar1_Button11.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._Toolbar1_Button11.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._Toolbar1_Button11.Name = "Cap"
		Me._Toolbar1_Button11.ToolTipText = "Show Capacities"
		Me._Toolbar1_Button11.ImageIndex = 5
		Me._Toolbar1_Button11.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._Toolbar1_Button12.Size = New System.Drawing.Size(27, 19)
		Me._Toolbar1_Button12.AutoSize = False
		Me._Toolbar1_Button12.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._Toolbar1_Button12.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._Toolbar1_Button12.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._Toolbar1_Button13.Size = New System.Drawing.Size(27, 19)
		Me._Toolbar1_Button13.AutoSize = False
		Me._Toolbar1_Button13.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._Toolbar1_Button13.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._Toolbar1_Button13.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._Toolbar1_Button14.Size = New System.Drawing.Size(27, 19)
		Me._Toolbar1_Button14.AutoSize = False
		Me._Toolbar1_Button14.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._Toolbar1_Button14.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._Toolbar1_Button14.Width = 495
		Me._Toolbar1_Button14.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me.txtTech.AutoSize = False
		Me.txtTech.Size = New System.Drawing.Size(33, 19)
		Me.txtTech.Location = New System.Drawing.Point(408, 152)
		Me.txtTech.TabIndex = 9
		Me.txtTech.Text = "999"
		Me.ToolTip1.SetToolTip(Me.txtTech, "Optional Tech Level for Refresh")
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
		Me.cmdClear.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdClear.Text = "Clear"
		Me.cmdClear.Size = New System.Drawing.Size(81, 33)
		Me.cmdClear.Location = New System.Drawing.Point(360, 208)
		Me.cmdClear.TabIndex = 8
		Me.cmdClear.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdClear.BackColor = System.Drawing.SystemColors.Control
		Me.cmdClear.CausesValidation = True
		Me.cmdClear.Enabled = True
		Me.cmdClear.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdClear.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdClear.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdClear.TabStop = True
		Me.cmdClear.Name = "cmdClear"
		Me.cmdRefreshUnits.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdRefreshUnits.Text = "Refresh Units"
		Me.cmdRefreshUnits.Size = New System.Drawing.Size(81, 33)
		Me.cmdRefreshUnits.Location = New System.Drawing.Point(360, 72)
		Me.cmdRefreshUnits.TabIndex = 7
		Me.cmdRefreshUnits.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdRefreshUnits.BackColor = System.Drawing.SystemColors.Control
		Me.cmdRefreshUnits.CausesValidation = True
		Me.cmdRefreshUnits.Enabled = True
		Me.cmdRefreshUnits.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdRefreshUnits.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdRefreshUnits.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdRefreshUnits.TabStop = True
		Me.cmdRefreshUnits.Name = "cmdRefreshUnits"
		Me.cmdRefreshAll.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdRefreshAll.Text = "Refresh All"
		Me.cmdRefreshAll.Size = New System.Drawing.Size(81, 33)
		Me.cmdRefreshAll.Location = New System.Drawing.Point(360, 112)
		Me.cmdRefreshAll.TabIndex = 6
		Me.cmdRefreshAll.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdRefreshAll.BackColor = System.Drawing.SystemColors.Control
		Me.cmdRefreshAll.CausesValidation = True
		Me.cmdRefreshAll.Enabled = True
		Me.cmdRefreshAll.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdRefreshAll.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdRefreshAll.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdRefreshAll.TabStop = True
		Me.cmdRefreshAll.Name = "cmdRefreshAll"
		Me.TextBox1.Size = New System.Drawing.Size(57, 81)
		Me.TextBox1.Location = New System.Drawing.Point(136, 320)
		Me.TextBox1.TabIndex = 5
		Me.TextBox1.Enabled = True
		Me.TextBox1.RTF = resources.GetString("TextBox1.TextRTF")
		Me.TextBox1.Font = New System.Drawing.Font("Courier New", 9.6!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.TextBox1.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.None
		Me.TextBox1.Name = "TextBox1"
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "Close"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Size = New System.Drawing.Size(81, 33)
		Me.cmdOK.Location = New System.Drawing.Point(360, 248)
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
		Me.cmdRefresh.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdRefresh.Text = "Refresh"
		Me.cmdRefresh.Size = New System.Drawing.Size(81, 33)
		Me.cmdRefresh.Location = New System.Drawing.Point(360, 32)
		Me.cmdRefresh.TabIndex = 2
		Me.cmdRefresh.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdRefresh.BackColor = System.Drawing.SystemColors.Control
		Me.cmdRefresh.CausesValidation = True
		Me.cmdRefresh.Enabled = True
		Me.cmdRefresh.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdRefresh.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdRefresh.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdRefresh.TabStop = True
		Me.cmdRefresh.Name = "cmdRefresh"
		grid1.OcxState = CType(resources.GetObject("grid1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.grid1.Size = New System.Drawing.Size(193, 297)
		Me.grid1.Location = New System.Drawing.Point(0, 24)
		Me.grid1.TabIndex = 0
		Me.grid1.Name = "grid1"
		Me.sb1.Dock = System.Windows.Forms.DockStyle.Bottom
		Me.sb1.Size = New System.Drawing.Size(528, 17)
		Me.sb1.Location = New System.Drawing.Point(0, 359)
		Me.sb1.TabIndex = 4
		Me.sb1.Text = "Printing for tech level"
		Me.sb1.Font = New System.Drawing.Font("Arial", 9.6!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.sb1.Name = "sb1"
		Me._sb1_Panel1.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter
		Me._sb1_Panel1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
		Me._sb1_Panel1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._sb1_Panel1.BorderSides = CType(System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom, System.Windows.Forms.ToolStripStatusLabelBorderSides)
		Me._sb1_Panel1.Margin = New System.Windows.Forms.Padding(0)
		Me._sb1_Panel1.Size = New System.Drawing.Size(96, 17)
		Me._sb1_Panel1.AutoSize = False
		Me.ImageList1.ImageSize = New System.Drawing.Size(20, 13)
		Me.ImageList1.TransparentColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me.ImageList1.Images.SetKeyName(0, "")
		Me.ImageList1.Images.SetKeyName(1, "")
		Me.ImageList1.Images.SetKeyName(2, "")
		Me.ImageList1.Images.SetKeyName(3, "")
		Me.ImageList1.Images.SetKeyName(4, "")
		Me.ImageList1.Images.SetKeyName(5, "")
		Me.ImageList1.Images.SetKeyName(6, "")
		Me.ImageList1.Images.SetKeyName(7, "")
		Me.ImageList1.Images.SetKeyName(8, "")
		Me.ImageList1.Images.SetKeyName(9, "")
		Me.Label1.Text = "Refresh Tech "
		Me.Label1.Size = New System.Drawing.Size(41, 25)
		Me.Label1.Location = New System.Drawing.Point(360, 152)
		Me.Label1.TabIndex = 10
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
		Me.Controls.Add(Toolbar1)
		Me.Controls.Add(txtTech)
		Me.Controls.Add(cmdClear)
		Me.Controls.Add(cmdRefreshUnits)
		Me.Controls.Add(cmdRefreshAll)
		Me.Controls.Add(TextBox1)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(cmdRefresh)
		Me.Controls.Add(grid1)
		Me.Controls.Add(sb1)
		Me.Controls.Add(Label1)
		Me.Toolbar1.Items.Add(_Toolbar1_Button1)
		Me.Toolbar1.Items.Add(_Toolbar1_Button2)
		Me.Toolbar1.Items.Add(_Toolbar1_Button3)
		Me.Toolbar1.Items.Add(_Toolbar1_Button4)
		Me.Toolbar1.Items.Add(_Toolbar1_Button5)
		Me.Toolbar1.Items.Add(_Toolbar1_Button6)
		Me.Toolbar1.Items.Add(_Toolbar1_Button7)
		Me.Toolbar1.Items.Add(_Toolbar1_Button8)
		Me.Toolbar1.Items.Add(_Toolbar1_Button9)
		Me.Toolbar1.Items.Add(_Toolbar1_Button10)
		Me.Toolbar1.Items.Add(_Toolbar1_Button11)
		Me.Toolbar1.Items.Add(_Toolbar1_Button12)
		Me.Toolbar1.Items.Add(_Toolbar1_Button13)
		Me.Toolbar1.Items.Add(_Toolbar1_Button14)
		Me.sb1.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me._sb1_Panel1})
		CType(Me.grid1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Toolbar1.ResumeLayout(False)
		Me.sb1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class