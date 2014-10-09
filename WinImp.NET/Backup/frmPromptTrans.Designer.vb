<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPromptTrans
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
	Public WithEvents txtQuantity As System.Windows.Forms.TextBox
	Public WithEvents cmdHelp As System.Windows.Forms.Button
	Public WithEvents chkDisplayPath As System.Windows.Forms.CheckBox
	Public WithEvents txtPath As System.Windows.Forms.TextBox
	Public WithEvents txtUnit As System.Windows.Forms.TextBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents Option1 As System.Windows.Forms.RadioButton
	Public WithEvents Option2 As System.Windows.Forms.RadioButton
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents lblOrigin As System.Windows.Forms.Label
	Public WithEvents lblQuantity As System.Windows.Forms.Label
	Public WithEvents lblLeft As System.Windows.Forms.Label
	Public WithEvents lblCost As System.Windows.Forms.Label
	Public WithEvents lblPathCost As System.Windows.Forms.Label
	Public WithEvents lblBestPath As System.Windows.Forms.Label
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPromptTrans))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.txtQuantity = New System.Windows.Forms.TextBox
		Me.cmdHelp = New System.Windows.Forms.Button
		Me.chkDisplayPath = New System.Windows.Forms.CheckBox
		Me.txtPath = New System.Windows.Forms.TextBox
		Me.txtUnit = New System.Windows.Forms.TextBox
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.Option1 = New System.Windows.Forms.RadioButton
		Me.Option2 = New System.Windows.Forms.RadioButton
		Me.lblOrigin = New System.Windows.Forms.Label
		Me.lblQuantity = New System.Windows.Forms.Label
		Me.lblLeft = New System.Windows.Forms.Label
		Me.lblCost = New System.Windows.Forms.Label
		Me.lblPathCost = New System.Windows.Forms.Label
		Me.lblBestPath = New System.Windows.Forms.Label
		Me.Label6 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Frame1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.ControlBox = False
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.ClientSize = New System.Drawing.Size(473, 120)
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
		Me.Name = "frmPromptTrans"
		Me.txtQuantity.AutoSize = False
		Me.txtQuantity.Size = New System.Drawing.Size(49, 19)
		Me.txtQuantity.Location = New System.Drawing.Point(368, 8)
		Me.txtQuantity.TabIndex = 1
		Me.txtQuantity.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtQuantity.AcceptsReturn = True
		Me.txtQuantity.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtQuantity.BackColor = System.Drawing.SystemColors.Window
		Me.txtQuantity.CausesValidation = True
		Me.txtQuantity.Enabled = True
		Me.txtQuantity.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtQuantity.HideSelection = True
		Me.txtQuantity.ReadOnly = False
		Me.txtQuantity.Maxlength = 0
		Me.txtQuantity.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtQuantity.MultiLine = False
		Me.txtQuantity.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtQuantity.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtQuantity.TabStop = True
		Me.txtQuantity.Visible = True
		Me.txtQuantity.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtQuantity.Name = "txtQuantity"
		Me.cmdHelp.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdHelp.Text = "?"
		Me.cmdHelp.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdHelp.Size = New System.Drawing.Size(25, 25)
		Me.cmdHelp.Location = New System.Drawing.Point(448, 0)
		Me.cmdHelp.TabIndex = 16
		Me.ToolTip1.SetToolTip(Me.cmdHelp, "Click for Help")
		Me.cmdHelp.BackColor = System.Drawing.SystemColors.Control
		Me.cmdHelp.CausesValidation = True
		Me.cmdHelp.Enabled = True
		Me.cmdHelp.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdHelp.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdHelp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdHelp.TabStop = True
		Me.cmdHelp.Name = "cmdHelp"
		Me.chkDisplayPath.Text = "Display Path"
		Me.chkDisplayPath.Size = New System.Drawing.Size(97, 17)
		Me.chkDisplayPath.Location = New System.Drawing.Point(168, 64)
		Me.chkDisplayPath.TabIndex = 15
		Me.chkDisplayPath.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkDisplayPath.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkDisplayPath.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkDisplayPath.BackColor = System.Drawing.SystemColors.Control
		Me.chkDisplayPath.CausesValidation = True
		Me.chkDisplayPath.Enabled = True
		Me.chkDisplayPath.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkDisplayPath.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkDisplayPath.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkDisplayPath.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkDisplayPath.TabStop = True
		Me.chkDisplayPath.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkDisplayPath.Visible = True
		Me.chkDisplayPath.Name = "chkDisplayPath"
		Me.txtPath.AutoSize = False
		Me.txtPath.Size = New System.Drawing.Size(121, 19)
		Me.txtPath.Location = New System.Drawing.Point(168, 40)
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
		Me.txtUnit.AutoSize = False
		Me.txtUnit.Size = New System.Drawing.Size(121, 19)
		Me.txtUnit.Location = New System.Drawing.Point(168, 8)
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
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(81, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(208, 88)
		Me.cmdCancel.TabIndex = 4
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
		Me.cmdOK.Location = New System.Drawing.Point(104, 88)
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
		Me.Frame1.Text = "Select "
		Me.Frame1.Size = New System.Drawing.Size(81, 65)
		Me.Frame1.Location = New System.Drawing.Point(16, 48)
		Me.Frame1.TabIndex = 7
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.Option1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option1.Text = "plane"
		Me.Option1.Size = New System.Drawing.Size(57, 17)
		Me.Option1.Location = New System.Drawing.Point(16, 16)
		Me.Option1.TabIndex = 5
		Me.Option1.Checked = True
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
		Me.Option1.Visible = True
		Me.Option1.Name = "Option1"
		Me.Option2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option2.Text = "nuke"
		Me.Option2.Size = New System.Drawing.Size(57, 17)
		Me.Option2.Location = New System.Drawing.Point(16, 40)
		Me.Option2.TabIndex = 6
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
		Me.lblOrigin.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblOrigin.Text = "start"
		Me.lblOrigin.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblOrigin.Size = New System.Drawing.Size(49, 17)
		Me.lblOrigin.Location = New System.Drawing.Point(24, 24)
		Me.lblOrigin.TabIndex = 18
		Me.lblOrigin.BackColor = System.Drawing.SystemColors.Control
		Me.lblOrigin.Enabled = True
		Me.lblOrigin.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblOrigin.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblOrigin.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblOrigin.UseMnemonic = True
		Me.lblOrigin.Visible = True
		Me.lblOrigin.AutoSize = False
		Me.lblOrigin.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblOrigin.Name = "lblOrigin"
		Me.lblQuantity.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblQuantity.Text = "quantity"
		Me.lblQuantity.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblQuantity.Size = New System.Drawing.Size(49, 17)
		Me.lblQuantity.Location = New System.Drawing.Point(304, 8)
		Me.lblQuantity.TabIndex = 17
		Me.lblQuantity.BackColor = System.Drawing.SystemColors.Control
		Me.lblQuantity.Enabled = True
		Me.lblQuantity.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblQuantity.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblQuantity.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblQuantity.UseMnemonic = True
		Me.lblQuantity.Visible = True
		Me.lblQuantity.AutoSize = False
		Me.lblQuantity.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblQuantity.Name = "lblQuantity"
		Me.lblLeft.Text = "mob left:"
		Me.lblLeft.Size = New System.Drawing.Size(153, 17)
		Me.lblLeft.Location = New System.Drawing.Point(304, 96)
		Me.lblLeft.TabIndex = 14
		Me.lblLeft.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblLeft.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblLeft.BackColor = System.Drawing.SystemColors.Control
		Me.lblLeft.Enabled = True
		Me.lblLeft.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblLeft.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblLeft.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblLeft.UseMnemonic = True
		Me.lblLeft.Visible = True
		Me.lblLeft.AutoSize = False
		Me.lblLeft.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblLeft.Name = "lblLeft"
		Me.lblCost.Text = "mob cost:"
		Me.lblCost.Size = New System.Drawing.Size(153, 17)
		Me.lblCost.Location = New System.Drawing.Point(304, 80)
		Me.lblCost.TabIndex = 13
		Me.lblCost.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblCost.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblCost.BackColor = System.Drawing.SystemColors.Control
		Me.lblCost.Enabled = True
		Me.lblCost.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblCost.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblCost.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblCost.UseMnemonic = True
		Me.lblCost.Visible = True
		Me.lblCost.AutoSize = False
		Me.lblCost.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblCost.Name = "lblCost"
		Me.lblPathCost.Text = "path cost:"
		Me.lblPathCost.Size = New System.Drawing.Size(153, 17)
		Me.lblPathCost.Location = New System.Drawing.Point(304, 64)
		Me.lblPathCost.TabIndex = 12
		Me.lblPathCost.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblPathCost.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblPathCost.BackColor = System.Drawing.SystemColors.Control
		Me.lblPathCost.Enabled = True
		Me.lblPathCost.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblPathCost.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblPathCost.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblPathCost.UseMnemonic = True
		Me.lblPathCost.Visible = True
		Me.lblPathCost.AutoSize = False
		Me.lblPathCost.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblPathCost.Name = "lblPathCost"
		Me.lblBestPath.Text = "best path:"
		Me.lblBestPath.Size = New System.Drawing.Size(161, 17)
		Me.lblBestPath.Location = New System.Drawing.Point(304, 48)
		Me.lblBestPath.TabIndex = 11
		Me.lblBestPath.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblBestPath.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblBestPath.BackColor = System.Drawing.SystemColors.Control
		Me.lblBestPath.Enabled = True
		Me.lblBestPath.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblBestPath.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblBestPath.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblBestPath.UseMnemonic = True
		Me.lblBestPath.Visible = True
		Me.lblBestPath.AutoSize = False
		Me.lblBestPath.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblBestPath.Name = "lblBestPath"
		Me.Label6.Text = "path/ dest."
		Me.Label6.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label6.Size = New System.Drawing.Size(41, 33)
		Me.Label6.Location = New System.Drawing.Point(112, 32)
		Me.Label6.TabIndex = 10
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
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label3.Text = "unit"
		Me.Label3.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.Size = New System.Drawing.Size(33, 17)
		Me.Label3.Location = New System.Drawing.Point(112, 8)
		Me.Label3.TabIndex = 9
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
		Me.Label2.Text = "Transport"
		Me.Label2.Font = New System.Drawing.Font("Arial Black", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(73, 17)
		Me.Label2.Location = New System.Drawing.Point(16, 8)
		Me.Label2.TabIndex = 8
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
		Me.Controls.Add(txtQuantity)
		Me.Controls.Add(cmdHelp)
		Me.Controls.Add(chkDisplayPath)
		Me.Controls.Add(txtPath)
		Me.Controls.Add(txtUnit)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(Frame1)
		Me.Controls.Add(lblOrigin)
		Me.Controls.Add(lblQuantity)
		Me.Controls.Add(lblLeft)
		Me.Controls.Add(lblCost)
		Me.Controls.Add(lblPathCost)
		Me.Controls.Add(lblBestPath)
		Me.Controls.Add(Label6)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label2)
		Me.Frame1.Controls.Add(Option1)
		Me.Frame1.Controls.Add(Option2)
		Me.Frame1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class