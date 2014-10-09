<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPromptOffer
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
	Public WithEvents Option4 As System.Windows.Forms.RadioButton
	Public WithEvents Option3 As System.Windows.Forms.RadioButton
	Public WithEvents txtNum As System.Windows.Forms.TextBox
	Public WithEvents txtDur As System.Windows.Forms.TextBox
	Public WithEvents txtRate As System.Windows.Forms.TextBox
	Public WithEvents txtAmount As System.Windows.Forms.TextBox
	Public WithEvents Option2 As System.Windows.Forms.RadioButton
	Public WithEvents Option1 As System.Windows.Forms.RadioButton
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents cbNations As System.Windows.Forms.ComboBox
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPromptOffer))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Option4 = New System.Windows.Forms.RadioButton
		Me.Option3 = New System.Windows.Forms.RadioButton
		Me.txtNum = New System.Windows.Forms.TextBox
		Me.txtDur = New System.Windows.Forms.TextBox
		Me.txtRate = New System.Windows.Forms.TextBox
		Me.txtAmount = New System.Windows.Forms.TextBox
		Me.Option2 = New System.Windows.Forms.RadioButton
		Me.Option1 = New System.Windows.Forms.RadioButton
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me.cbNations = New System.Windows.Forms.ComboBox
		Me.Label5 = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.ControlBox = False
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.ClientSize = New System.Drawing.Size(470, 132)
		Me.Location = New System.Drawing.Point(3, 3)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmPromptOffer"
		Me.Option4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option4.Text = "announcements"
		Me.Option4.Size = New System.Drawing.Size(97, 17)
		Me.Option4.Location = New System.Drawing.Point(88, 88)
		Me.Option4.TabIndex = 15
		Me.Option4.Visible = False
		Me.Option4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Option4.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option4.BackColor = System.Drawing.SystemColors.Control
		Me.Option4.CausesValidation = True
		Me.Option4.Enabled = True
		Me.Option4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Option4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Option4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Option4.Appearance = System.Windows.Forms.Appearance.Normal
		Me.Option4.TabStop = True
		Me.Option4.Checked = False
		Me.Option4.Name = "Option4"
		Me.Option3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option3.Text = "mail"
		Me.Option3.Size = New System.Drawing.Size(57, 17)
		Me.Option3.Location = New System.Drawing.Point(88, 64)
		Me.Option3.TabIndex = 14
		Me.Option3.Visible = False
		Me.Option3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Option3.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option3.BackColor = System.Drawing.SystemColors.Control
		Me.Option3.CausesValidation = True
		Me.Option3.Enabled = True
		Me.Option3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Option3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Option3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Option3.Appearance = System.Windows.Forms.Appearance.Normal
		Me.Option3.TabStop = True
		Me.Option3.Checked = False
		Me.Option3.Name = "Option3"
		Me.txtNum.AutoSize = False
		Me.txtNum.Size = New System.Drawing.Size(73, 19)
		Me.txtNum.Location = New System.Drawing.Point(32, 80)
		Me.txtNum.TabIndex = 13
		Me.txtNum.Text = "Text1"
		Me.txtNum.Visible = False
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
		Me.txtNum.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtNum.Name = "txtNum"
		Me.txtDur.AutoSize = False
		Me.txtDur.Size = New System.Drawing.Size(49, 19)
		Me.txtDur.Location = New System.Drawing.Point(304, 96)
		Me.txtDur.TabIndex = 9
		Me.txtDur.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtDur.AcceptsReturn = True
		Me.txtDur.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtDur.BackColor = System.Drawing.SystemColors.Window
		Me.txtDur.CausesValidation = True
		Me.txtDur.Enabled = True
		Me.txtDur.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtDur.HideSelection = True
		Me.txtDur.ReadOnly = False
		Me.txtDur.Maxlength = 0
		Me.txtDur.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtDur.MultiLine = False
		Me.txtDur.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtDur.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtDur.TabStop = True
		Me.txtDur.Visible = True
		Me.txtDur.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtDur.Name = "txtDur"
		Me.txtRate.AutoSize = False
		Me.txtRate.Size = New System.Drawing.Size(49, 19)
		Me.txtRate.Location = New System.Drawing.Point(200, 96)
		Me.txtRate.TabIndex = 8
		Me.txtRate.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtRate.AcceptsReturn = True
		Me.txtRate.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtRate.BackColor = System.Drawing.SystemColors.Window
		Me.txtRate.CausesValidation = True
		Me.txtRate.Enabled = True
		Me.txtRate.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtRate.HideSelection = True
		Me.txtRate.ReadOnly = False
		Me.txtRate.Maxlength = 0
		Me.txtRate.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtRate.MultiLine = False
		Me.txtRate.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtRate.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtRate.TabStop = True
		Me.txtRate.Visible = True
		Me.txtRate.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtRate.Name = "txtRate"
		Me.txtAmount.AutoSize = False
		Me.txtAmount.Size = New System.Drawing.Size(153, 19)
		Me.txtAmount.Location = New System.Drawing.Point(200, 64)
		Me.txtAmount.TabIndex = 7
		Me.txtAmount.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtAmount.AcceptsReturn = True
		Me.txtAmount.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtAmount.BackColor = System.Drawing.SystemColors.Window
		Me.txtAmount.CausesValidation = True
		Me.txtAmount.Enabled = True
		Me.txtAmount.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtAmount.HideSelection = True
		Me.txtAmount.ReadOnly = False
		Me.txtAmount.Maxlength = 0
		Me.txtAmount.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtAmount.MultiLine = False
		Me.txtAmount.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtAmount.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtAmount.TabStop = True
		Me.txtAmount.Visible = True
		Me.txtAmount.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtAmount.Name = "txtAmount"
		Me.Option2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option2.Text = "treaty"
		Me.Option2.Size = New System.Drawing.Size(57, 17)
		Me.Option2.Location = New System.Drawing.Point(88, 40)
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
		Me.Option1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option1.Text = "loan"
		Me.Option1.Size = New System.Drawing.Size(57, 17)
		Me.Option1.Location = New System.Drawing.Point(88, 16)
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
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(81, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(376, 56)
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
		Me.cmdOK.Location = New System.Drawing.Point(376, 16)
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
		Me.cbNations.Size = New System.Drawing.Size(153, 21)
		Me.cbNations.Location = New System.Drawing.Point(200, 24)
		Me.cbNations.Sorted = True
		Me.cbNations.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cbNations.TabIndex = 2
		Me.cbNations.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbNations.BackColor = System.Drawing.SystemColors.Window
		Me.cbNations.CausesValidation = True
		Me.cbNations.Enabled = True
		Me.cbNations.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cbNations.IntegralHeight = True
		Me.cbNations.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbNations.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbNations.TabStop = True
		Me.cbNations.Visible = True
		Me.cbNations.Name = "cbNations"
		Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label5.Text = "Duration"
		Me.Label5.Size = New System.Drawing.Size(49, 17)
		Me.Label5.Location = New System.Drawing.Point(248, 96)
		Me.Label5.TabIndex = 12
		Me.Label5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label4.Text = "Rate"
		Me.Label4.Size = New System.Drawing.Size(41, 17)
		Me.Label4.Location = New System.Drawing.Point(152, 96)
		Me.Label4.TabIndex = 11
		Me.Label4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label3.Text = "Amount"
		Me.Label3.Size = New System.Drawing.Size(49, 17)
		Me.Label3.Location = New System.Drawing.Point(144, 64)
		Me.Label3.TabIndex = 10
		Me.Label3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label2.Text = "Offer"
		Me.Label2.Font = New System.Drawing.Font("Arial Black", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(81, 17)
		Me.Label2.Location = New System.Drawing.Point(0, 24)
		Me.Label2.TabIndex = 1
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
		Me.Label1.Text = "to"
		Me.Label1.Size = New System.Drawing.Size(57, 17)
		Me.Label1.Location = New System.Drawing.Point(184, 24)
		Me.Label1.TabIndex = 0
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
		Me.Controls.Add(Option4)
		Me.Controls.Add(Option3)
		Me.Controls.Add(txtNum)
		Me.Controls.Add(txtDur)
		Me.Controls.Add(txtRate)
		Me.Controls.Add(txtAmount)
		Me.Controls.Add(Option2)
		Me.Controls.Add(Option1)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(cbNations)
		Me.Controls.Add(Label5)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class