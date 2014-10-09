<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPromptCom
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
	Public WithEvents txtPrice As System.Windows.Forms.TextBox
	Public WithEvents txtLotNumber As System.Windows.Forms.TextBox
	Public WithEvents txtOrigin As System.Windows.Forms.TextBox
	Public WithEvents cmdHelp As System.Windows.Forms.Button
	Public WithEvents cbCom As System.Windows.Forms.ComboBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents lblOrigin As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents lblPrice As System.Windows.Forms.Label
	Public WithEvents lblLotNumber As System.Windows.Forms.Label
	Public WithEvents lblCommodity As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPromptCom))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.txtPrice = New System.Windows.Forms.TextBox
		Me.txtLotNumber = New System.Windows.Forms.TextBox
		Me.txtOrigin = New System.Windows.Forms.TextBox
		Me.cmdHelp = New System.Windows.Forms.Button
		Me.cbCom = New System.Windows.Forms.ComboBox
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me.lblOrigin = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.lblPrice = New System.Windows.Forms.Label
		Me.lblLotNumber = New System.Windows.Forms.Label
		Me.lblCommodity = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.ControlBox = False
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.ClientSize = New System.Drawing.Size(312, 101)
		Me.Location = New System.Drawing.Point(1, 0)
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
		Me.Name = "frmPromptCom"
		Me.txtPrice.AutoSize = False
		Me.txtPrice.Size = New System.Drawing.Size(41, 19)
		Me.txtPrice.Location = New System.Drawing.Point(176, 72)
		Me.txtPrice.TabIndex = 11
		Me.txtPrice.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtPrice.AcceptsReturn = True
		Me.txtPrice.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtPrice.BackColor = System.Drawing.SystemColors.Window
		Me.txtPrice.CausesValidation = True
		Me.txtPrice.Enabled = True
		Me.txtPrice.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtPrice.HideSelection = True
		Me.txtPrice.ReadOnly = False
		Me.txtPrice.Maxlength = 0
		Me.txtPrice.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtPrice.MultiLine = False
		Me.txtPrice.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtPrice.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtPrice.TabStop = True
		Me.txtPrice.Visible = True
		Me.txtPrice.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtPrice.Name = "txtPrice"
		Me.txtLotNumber.AutoSize = False
		Me.txtLotNumber.Size = New System.Drawing.Size(41, 19)
		Me.txtLotNumber.Location = New System.Drawing.Point(72, 72)
		Me.txtLotNumber.TabIndex = 10
		Me.txtLotNumber.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtLotNumber.AcceptsReturn = True
		Me.txtLotNumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtLotNumber.BackColor = System.Drawing.SystemColors.Window
		Me.txtLotNumber.CausesValidation = True
		Me.txtLotNumber.Enabled = True
		Me.txtLotNumber.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtLotNumber.HideSelection = True
		Me.txtLotNumber.ReadOnly = False
		Me.txtLotNumber.Maxlength = 0
		Me.txtLotNumber.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtLotNumber.MultiLine = False
		Me.txtLotNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtLotNumber.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtLotNumber.TabStop = True
		Me.txtLotNumber.Visible = True
		Me.txtLotNumber.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtLotNumber.Name = "txtLotNumber"
		Me.txtOrigin.AutoSize = False
		Me.txtOrigin.Size = New System.Drawing.Size(89, 19)
		Me.txtOrigin.Location = New System.Drawing.Point(128, 8)
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
		Me.cmdHelp.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdHelp.Text = "?"
		Me.cmdHelp.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdHelp.Size = New System.Drawing.Size(25, 25)
		Me.cmdHelp.Location = New System.Drawing.Point(288, 0)
		Me.cmdHelp.TabIndex = 6
		Me.ToolTip1.SetToolTip(Me.cmdHelp, "Click for Help")
		Me.cmdHelp.BackColor = System.Drawing.SystemColors.Control
		Me.cmdHelp.CausesValidation = True
		Me.cmdHelp.Enabled = True
		Me.cmdHelp.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdHelp.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdHelp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdHelp.TabStop = True
		Me.cmdHelp.Name = "cmdHelp"
		Me.cbCom.Size = New System.Drawing.Size(145, 21)
		Me.cbCom.Location = New System.Drawing.Point(72, 40)
		Me.cbCom.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cbCom.TabIndex = 3
		Me.cbCom.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbCom.BackColor = System.Drawing.SystemColors.Window
		Me.cbCom.CausesValidation = True
		Me.cbCom.Enabled = True
		Me.cbCom.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cbCom.IntegralHeight = True
		Me.cbCom.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbCom.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbCom.Sorted = False
		Me.cbCom.TabStop = True
		Me.cbCom.Visible = True
		Me.cbCom.Name = "cbCom"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(81, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(224, 64)
		Me.cmdCancel.TabIndex = 2
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
		Me.cmdOK.Location = New System.Drawing.Point(224, 32)
		Me.cmdOK.TabIndex = 1
		Me.cmdOK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.lblOrigin.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblOrigin.Text = "to"
		Me.lblOrigin.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblOrigin.Size = New System.Drawing.Size(25, 17)
		Me.lblOrigin.Location = New System.Drawing.Point(96, 8)
		Me.lblOrigin.TabIndex = 9
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
		Me.Label2.Text = "Market"
		Me.Label2.Font = New System.Drawing.Font("Arial Black", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(73, 17)
		Me.Label2.Location = New System.Drawing.Point(8, 8)
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
		Me.lblPrice.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblPrice.Text = "Price"
		Me.lblPrice.Size = New System.Drawing.Size(41, 17)
		Me.lblPrice.Location = New System.Drawing.Point(120, 72)
		Me.lblPrice.TabIndex = 5
		Me.lblPrice.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblPrice.BackColor = System.Drawing.SystemColors.Control
		Me.lblPrice.Enabled = True
		Me.lblPrice.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblPrice.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblPrice.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblPrice.UseMnemonic = True
		Me.lblPrice.Visible = True
		Me.lblPrice.AutoSize = False
		Me.lblPrice.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblPrice.Name = "lblPrice"
		Me.lblLotNumber.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblLotNumber.Text = "Lot Number"
		Me.lblLotNumber.Size = New System.Drawing.Size(65, 17)
		Me.lblLotNumber.Location = New System.Drawing.Point(0, 72)
		Me.lblLotNumber.TabIndex = 4
		Me.lblLotNumber.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblLotNumber.BackColor = System.Drawing.SystemColors.Control
		Me.lblLotNumber.Enabled = True
		Me.lblLotNumber.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblLotNumber.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblLotNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblLotNumber.UseMnemonic = True
		Me.lblLotNumber.Visible = True
		Me.lblLotNumber.AutoSize = False
		Me.lblLotNumber.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblLotNumber.Name = "lblLotNumber"
		Me.lblCommodity.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblCommodity.Text = "Commodity"
		Me.lblCommodity.Size = New System.Drawing.Size(57, 17)
		Me.lblCommodity.Location = New System.Drawing.Point(8, 40)
		Me.lblCommodity.TabIndex = 0
		Me.lblCommodity.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblCommodity.BackColor = System.Drawing.SystemColors.Control
		Me.lblCommodity.Enabled = True
		Me.lblCommodity.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblCommodity.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblCommodity.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblCommodity.UseMnemonic = True
		Me.lblCommodity.Visible = True
		Me.lblCommodity.AutoSize = False
		Me.lblCommodity.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblCommodity.Name = "lblCommodity"
		Me.Controls.Add(txtPrice)
		Me.Controls.Add(txtLotNumber)
		Me.Controls.Add(txtOrigin)
		Me.Controls.Add(cmdHelp)
		Me.Controls.Add(cbCom)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(lblOrigin)
		Me.Controls.Add(Label2)
		Me.Controls.Add(lblPrice)
		Me.Controls.Add(lblLotNumber)
		Me.Controls.Add(lblCommodity)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class