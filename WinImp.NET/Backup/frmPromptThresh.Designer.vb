<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPromptThresh
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
	Public WithEvents chkFoodRequired As System.Windows.Forms.CheckBox
	Public WithEvents chkSupply As System.Windows.Forms.CheckBox
	Public WithEvents cmdHelp As System.Windows.Forms.Button
	Public WithEvents cbCom As System.Windows.Forms.ListBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents txtMultOrigin As System.Windows.Forms.TextBox
	Public WithEvents txtNum As System.Windows.Forms.TextBox
	Public WithEvents lblFoodPercent As System.Windows.Forms.Label
	Public WithEvents lblDesc As System.Windows.Forms.Label
	Public WithEvents lblThresh As System.Windows.Forms.Label
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPromptThresh))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.chkFoodRequired = New System.Windows.Forms.CheckBox
		Me.chkSupply = New System.Windows.Forms.CheckBox
		Me.cmdHelp = New System.Windows.Forms.Button
		Me.cbCom = New System.Windows.Forms.ListBox
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me.txtMultOrigin = New System.Windows.Forms.TextBox
		Me.txtNum = New System.Windows.Forms.TextBox
		Me.lblFoodPercent = New System.Windows.Forms.Label
		Me.lblDesc = New System.Windows.Forms.Label
		Me.lblThresh = New System.Windows.Forms.Label
		Me.Label5 = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.ControlBox = False
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.ClientSize = New System.Drawing.Size(511, 147)
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
		Me.Name = "frmPromptThresh"
		Me.chkFoodRequired.Text = "Use Food Required to determine Threshold"
		Me.chkFoodRequired.Size = New System.Drawing.Size(129, 25)
		Me.chkFoodRequired.Location = New System.Drawing.Point(96, 80)
		Me.chkFoodRequired.TabIndex = 14
		Me.chkFoodRequired.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkFoodRequired.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkFoodRequired.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkFoodRequired.BackColor = System.Drawing.SystemColors.Control
		Me.chkFoodRequired.CausesValidation = True
		Me.chkFoodRequired.Enabled = True
		Me.chkFoodRequired.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkFoodRequired.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkFoodRequired.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkFoodRequired.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkFoodRequired.TabStop = True
		Me.chkFoodRequired.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkFoodRequired.Visible = True
		Me.chkFoodRequired.Name = "chkFoodRequired"
		Me.chkSupply.Text = "Auto Push/Pull supply from warehouse to meet new threshold"
		Me.chkSupply.Size = New System.Drawing.Size(185, 25)
		Me.chkSupply.Location = New System.Drawing.Point(96, 112)
		Me.chkSupply.TabIndex = 12
		Me.chkSupply.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkSupply.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkSupply.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkSupply.BackColor = System.Drawing.SystemColors.Control
		Me.chkSupply.CausesValidation = True
		Me.chkSupply.Enabled = True
		Me.chkSupply.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkSupply.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkSupply.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkSupply.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkSupply.TabStop = True
		Me.chkSupply.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkSupply.Visible = True
		Me.chkSupply.Name = "chkSupply"
		Me.cmdHelp.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdHelp.Text = "?"
		Me.cmdHelp.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdHelp.Size = New System.Drawing.Size(25, 25)
		Me.cmdHelp.Location = New System.Drawing.Point(488, 0)
		Me.cmdHelp.TabIndex = 11
		Me.ToolTip1.SetToolTip(Me.cmdHelp, "Click for Help")
		Me.cmdHelp.BackColor = System.Drawing.SystemColors.Control
		Me.cmdHelp.CausesValidation = True
		Me.cmdHelp.Enabled = True
		Me.cmdHelp.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdHelp.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdHelp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdHelp.TabStop = True
		Me.cmdHelp.Name = "cmdHelp"
		Me.cbCom.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbCom.Size = New System.Drawing.Size(192, 119)
		Me.cbCom.Location = New System.Drawing.Point(288, 8)
		Me.cbCom.TabIndex = 9
		Me.cbCom.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.cbCom.BackColor = System.Drawing.SystemColors.Window
		Me.cbCom.CausesValidation = True
		Me.cbCom.Enabled = True
		Me.cbCom.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cbCom.IntegralHeight = True
		Me.cbCom.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbCom.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.cbCom.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbCom.Sorted = False
		Me.cbCom.TabStop = True
		Me.cbCom.Visible = True
		Me.cbCom.MultiColumn = True
		Me.cbCom.ColumnWidth = 96
		Me.cbCom.Name = "cbCom"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(81, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(8, 112)
		Me.cmdCancel.TabIndex = 8
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
		Me.cmdOK.Location = New System.Drawing.Point(8, 80)
		Me.cmdOK.TabIndex = 7
		Me.cmdOK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.txtMultOrigin.AutoSize = False
		Me.txtMultOrigin.Size = New System.Drawing.Size(89, 19)
		Me.txtMultOrigin.Location = New System.Drawing.Point(152, 8)
		Me.txtMultOrigin.TabIndex = 2
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
		Me.txtNum.AutoSize = False
		Me.txtNum.Size = New System.Drawing.Size(49, 19)
		Me.txtNum.Location = New System.Drawing.Point(152, 32)
		Me.txtNum.TabIndex = 0
		Me.txtNum.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtNum.AcceptsReturn = True
		Me.txtNum.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtNum.BackColor = System.Drawing.SystemColors.Window
		Me.txtNum.CausesValidation = True
		Me.txtNum.Enabled = True
		Me.txtNum.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtNum.HideSelection = True
		Me.txtNum.ReadOnly = False
		Me.txtNum.Maxlength = 0
		Me.txtNum.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtNum.MultiLine = False
		Me.txtNum.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtNum.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtNum.TabStop = True
		Me.txtNum.Visible = True
		Me.txtNum.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtNum.Name = "txtNum"
		Me.lblFoodPercent.Text = "Percent Above Food Required"
		Me.lblFoodPercent.Size = New System.Drawing.Size(81, 25)
		Me.lblFoodPercent.Location = New System.Drawing.Point(208, 32)
		Me.lblFoodPercent.TabIndex = 15
		Me.lblFoodPercent.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblFoodPercent.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblFoodPercent.BackColor = System.Drawing.SystemColors.Control
		Me.lblFoodPercent.Enabled = True
		Me.lblFoodPercent.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblFoodPercent.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblFoodPercent.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblFoodPercent.UseMnemonic = True
		Me.lblFoodPercent.Visible = True
		Me.lblFoodPercent.AutoSize = False
		Me.lblFoodPercent.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblFoodPercent.Name = "lblFoodPercent"
		Me.lblDesc.Size = New System.Drawing.Size(107, 13)
		Me.lblDesc.Location = New System.Drawing.Point(8, 32)
		Me.lblDesc.TabIndex = 13
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
		Me.lblThresh.Text = "Current Threshold: 999"
		Me.lblThresh.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblThresh.Size = New System.Drawing.Size(153, 17)
		Me.lblThresh.Location = New System.Drawing.Point(128, 56)
		Me.lblThresh.TabIndex = 10
		Me.lblThresh.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblThresh.BackColor = System.Drawing.SystemColors.Control
		Me.lblThresh.Enabled = True
		Me.lblThresh.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblThresh.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblThresh.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblThresh.UseMnemonic = True
		Me.lblThresh.Visible = True
		Me.lblThresh.AutoSize = False
		Me.lblThresh.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblThresh.Name = "lblThresh"
		Me.Label5.Text = "Set"
		Me.Label5.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label5.Size = New System.Drawing.Size(25, 17)
		Me.Label5.Location = New System.Drawing.Point(8, 8)
		Me.Label5.TabIndex = 6
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
		Me.Label4.Text = "for"
		Me.Label4.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.Size = New System.Drawing.Size(25, 17)
		Me.Label4.Location = New System.Drawing.Point(256, 8)
		Me.Label4.TabIndex = 5
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
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label2.Text = "at"
		Me.Label2.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(25, 17)
		Me.Label2.Location = New System.Drawing.Point(120, 32)
		Me.Label2.TabIndex = 4
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
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label3.Text = "in"
		Me.Label3.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.Size = New System.Drawing.Size(25, 17)
		Me.Label3.Location = New System.Drawing.Point(120, 8)
		Me.Label3.TabIndex = 3
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
		Me.Label1.Text = "Threshold"
		Me.Label1.Font = New System.Drawing.Font("Arial Black", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Size = New System.Drawing.Size(81, 17)
		Me.Label1.Location = New System.Drawing.Point(40, 8)
		Me.Label1.TabIndex = 1
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
		Me.Controls.Add(chkFoodRequired)
		Me.Controls.Add(chkSupply)
		Me.Controls.Add(cmdHelp)
		Me.Controls.Add(cbCom)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(txtMultOrigin)
		Me.Controls.Add(txtNum)
		Me.Controls.Add(lblFoodPercent)
		Me.Controls.Add(lblDesc)
		Me.Controls.Add(lblThresh)
		Me.Controls.Add(Label5)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class