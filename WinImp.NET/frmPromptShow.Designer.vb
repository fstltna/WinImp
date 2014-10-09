<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPromptShow
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
	Public WithEvents cmdHelp As System.Windows.Forms.Button
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents txtTechLevel As System.Windows.Forms.TextBox
	Public WithEvents Option6 As System.Windows.Forms.RadioButton
	Public WithEvents Option5 As System.Windows.Forms.RadioButton
	Public WithEvents Option4 As System.Windows.Forms.RadioButton
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents Option10 As System.Windows.Forms.RadioButton
	Public WithEvents Option9 As System.Windows.Forms.RadioButton
	Public WithEvents Option8 As System.Windows.Forms.RadioButton
	Public WithEvents Option7 As System.Windows.Forms.RadioButton
	Public WithEvents Option3 As System.Windows.Forms.RadioButton
	Public WithEvents Option2 As System.Windows.Forms.RadioButton
	Public WithEvents Option1 As System.Windows.Forms.RadioButton
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPromptShow))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdHelp = New System.Windows.Forms.Button
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me.txtTechLevel = New System.Windows.Forms.TextBox
		Me.Frame2 = New System.Windows.Forms.GroupBox
		Me.Option6 = New System.Windows.Forms.RadioButton
		Me.Option5 = New System.Windows.Forms.RadioButton
		Me.Option4 = New System.Windows.Forms.RadioButton
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.Option10 = New System.Windows.Forms.RadioButton
		Me.Option9 = New System.Windows.Forms.RadioButton
		Me.Option8 = New System.Windows.Forms.RadioButton
		Me.Option7 = New System.Windows.Forms.RadioButton
		Me.Option3 = New System.Windows.Forms.RadioButton
		Me.Option2 = New System.Windows.Forms.RadioButton
		Me.Option1 = New System.Windows.Forms.RadioButton
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.Frame2.SuspendLayout()
		Me.Frame1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Show"
		Me.ClientSize = New System.Drawing.Size(400, 170)
		Me.Location = New System.Drawing.Point(3, 22)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
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
		Me.Name = "frmPromptShow"
		Me.cmdHelp.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdHelp.Text = "?"
		Me.cmdHelp.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdHelp.Size = New System.Drawing.Size(25, 25)
		Me.cmdHelp.Location = New System.Drawing.Point(376, 0)
		Me.cmdHelp.TabIndex = 17
		Me.ToolTip1.SetToolTip(Me.cmdHelp, "Click for Help")
		Me.cmdHelp.BackColor = System.Drawing.SystemColors.Control
		Me.cmdHelp.CausesValidation = True
		Me.cmdHelp.Enabled = True
		Me.cmdHelp.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdHelp.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdHelp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdHelp.TabStop = True
		Me.cmdHelp.Name = "cmdHelp"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(81, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(240, 136)
		Me.cmdCancel.TabIndex = 13
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
		Me.cmdOK.Location = New System.Drawing.Point(104, 136)
		Me.cmdOK.TabIndex = 12
		Me.cmdOK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.txtTechLevel.AutoSize = False
		Me.txtTechLevel.Size = New System.Drawing.Size(49, 19)
		Me.txtTechLevel.Location = New System.Drawing.Point(192, 104)
		Me.txtTechLevel.TabIndex = 9
		Me.txtTechLevel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtTechLevel.AcceptsReturn = True
		Me.txtTechLevel.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtTechLevel.BackColor = System.Drawing.SystemColors.Window
		Me.txtTechLevel.CausesValidation = True
		Me.txtTechLevel.Enabled = True
		Me.txtTechLevel.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtTechLevel.HideSelection = True
		Me.txtTechLevel.ReadOnly = False
		Me.txtTechLevel.Maxlength = 0
		Me.txtTechLevel.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtTechLevel.MultiLine = False
		Me.txtTechLevel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtTechLevel.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtTechLevel.TabStop = True
		Me.txtTechLevel.Visible = True
		Me.txtTechLevel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtTechLevel.Name = "txtTechLevel"
		Me.Frame2.Text = "Select"
		Me.Frame2.Size = New System.Drawing.Size(105, 89)
		Me.Frame2.Location = New System.Drawing.Point(256, 8)
		Me.Frame2.TabIndex = 4
		Me.Frame2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame2.BackColor = System.Drawing.SystemColors.Control
		Me.Frame2.Enabled = True
		Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame2.Visible = True
		Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame2.Name = "Frame2"
		Me.Option6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option6.Text = "statistics"
		Me.Option6.Size = New System.Drawing.Size(73, 17)
		Me.Option6.Location = New System.Drawing.Point(16, 64)
		Me.Option6.TabIndex = 7
		Me.Option6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Option6.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option6.BackColor = System.Drawing.SystemColors.Control
		Me.Option6.CausesValidation = True
		Me.Option6.Enabled = True
		Me.Option6.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Option6.Cursor = System.Windows.Forms.Cursors.Default
		Me.Option6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Option6.Appearance = System.Windows.Forms.Appearance.Normal
		Me.Option6.TabStop = True
		Me.Option6.Checked = False
		Me.Option6.Visible = True
		Me.Option6.Name = "Option6"
		Me.Option5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option5.Text = "capacities"
		Me.Option5.Size = New System.Drawing.Size(73, 17)
		Me.Option5.Location = New System.Drawing.Point(16, 40)
		Me.Option5.TabIndex = 6
		Me.Option5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Option5.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option5.BackColor = System.Drawing.SystemColors.Control
		Me.Option5.CausesValidation = True
		Me.Option5.Enabled = True
		Me.Option5.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Option5.Cursor = System.Windows.Forms.Cursors.Default
		Me.Option5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Option5.Appearance = System.Windows.Forms.Appearance.Normal
		Me.Option5.TabStop = True
		Me.Option5.Checked = False
		Me.Option5.Visible = True
		Me.Option5.Name = "Option5"
		Me.Option4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option4.Text = "build"
		Me.Option4.Size = New System.Drawing.Size(73, 17)
		Me.Option4.Location = New System.Drawing.Point(16, 16)
		Me.Option4.TabIndex = 5
		Me.Option4.Checked = True
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
		Me.Option4.Visible = True
		Me.Option4.Name = "Option4"
		Me.Frame1.Text = "Select "
		Me.Frame1.Size = New System.Drawing.Size(177, 89)
		Me.Frame1.Location = New System.Drawing.Point(64, 8)
		Me.Frame1.TabIndex = 0
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.Option10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option10.Text = "nuke"
		Me.Option10.Size = New System.Drawing.Size(65, 17)
		Me.Option10.Location = New System.Drawing.Point(16, 64)
		Me.Option10.TabIndex = 16
		Me.Option10.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Option10.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option10.BackColor = System.Drawing.SystemColors.Control
		Me.Option10.CausesValidation = True
		Me.Option10.Enabled = True
		Me.Option10.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Option10.Cursor = System.Windows.Forms.Cursors.Default
		Me.Option10.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Option10.Appearance = System.Windows.Forms.Appearance.Normal
		Me.Option10.TabStop = True
		Me.Option10.Checked = False
		Me.Option10.Visible = True
		Me.Option10.Name = "Option10"
		Me.Option9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option9.Text = "bridge tower"
		Me.Option9.Size = New System.Drawing.Size(89, 17)
		Me.Option9.Location = New System.Drawing.Point(80, 40)
		Me.Option9.TabIndex = 15
		Me.Option9.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Option9.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option9.BackColor = System.Drawing.SystemColors.Control
		Me.Option9.CausesValidation = True
		Me.Option9.Enabled = True
		Me.Option9.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Option9.Cursor = System.Windows.Forms.Cursors.Default
		Me.Option9.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Option9.Appearance = System.Windows.Forms.Appearance.Normal
		Me.Option9.TabStop = True
		Me.Option9.Checked = False
		Me.Option9.Visible = True
		Me.Option9.Name = "Option9"
		Me.Option8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option8.Text = "sector"
		Me.Option8.Size = New System.Drawing.Size(81, 17)
		Me.Option8.Location = New System.Drawing.Point(80, 64)
		Me.Option8.TabIndex = 11
		Me.Option8.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Option8.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option8.BackColor = System.Drawing.SystemColors.Control
		Me.Option8.CausesValidation = True
		Me.Option8.Enabled = True
		Me.Option8.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Option8.Cursor = System.Windows.Forms.Cursors.Default
		Me.Option8.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Option8.Appearance = System.Windows.Forms.Appearance.Normal
		Me.Option8.TabStop = True
		Me.Option8.Checked = False
		Me.Option8.Visible = True
		Me.Option8.Name = "Option8"
		Me.Option7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option7.Text = "bridge span"
		Me.Option7.Size = New System.Drawing.Size(89, 17)
		Me.Option7.Location = New System.Drawing.Point(80, 16)
		Me.Option7.TabIndex = 10
		Me.Option7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Option7.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option7.BackColor = System.Drawing.SystemColors.Control
		Me.Option7.CausesValidation = True
		Me.Option7.Enabled = True
		Me.Option7.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Option7.Cursor = System.Windows.Forms.Cursors.Default
		Me.Option7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Option7.Appearance = System.Windows.Forms.Appearance.Normal
		Me.Option7.TabStop = True
		Me.Option7.Checked = False
		Me.Option7.Visible = True
		Me.Option7.Name = "Option7"
		Me.Option3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option3.Text = "ship"
		Me.Option3.Size = New System.Drawing.Size(65, 17)
		Me.Option3.Location = New System.Drawing.Point(16, 16)
		Me.Option3.TabIndex = 3
		Me.Option3.Checked = True
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
		Me.Option3.Visible = True
		Me.Option3.Name = "Option3"
		Me.Option2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option2.Text = "plane"
		Me.Option2.Size = New System.Drawing.Size(65, 17)
		Me.Option2.Location = New System.Drawing.Point(16, 48)
		Me.Option2.TabIndex = 2
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
		Me.Option1.Text = "land"
		Me.Option1.Size = New System.Drawing.Size(57, 17)
		Me.Option1.Location = New System.Drawing.Point(16, 32)
		Me.Option1.TabIndex = 1
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
		Me.Option1.Checked = False
		Me.Option1.Visible = True
		Me.Option1.Name = "Option1"
		Me.Label2.Text = "Show"
		Me.Label2.Font = New System.Drawing.Font("Arial Black", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(41, 17)
		Me.Label2.Location = New System.Drawing.Point(8, 40)
		Me.Label2.TabIndex = 14
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
		Me.Label1.Text = "Tech Level (Optional)"
		Me.Label1.Size = New System.Drawing.Size(113, 17)
		Me.Label1.Location = New System.Drawing.Point(64, 104)
		Me.Label1.TabIndex = 8
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
		Me.Controls.Add(cmdHelp)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(txtTechLevel)
		Me.Controls.Add(Frame2)
		Me.Controls.Add(Frame1)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		Me.Frame2.Controls.Add(Option6)
		Me.Frame2.Controls.Add(Option5)
		Me.Frame2.Controls.Add(Option4)
		Me.Frame1.Controls.Add(Option10)
		Me.Frame1.Controls.Add(Option9)
		Me.Frame1.Controls.Add(Option8)
		Me.Frame1.Controls.Add(Option7)
		Me.Frame1.Controls.Add(Option3)
		Me.Frame1.Controls.Add(Option2)
		Me.Frame1.Controls.Add(Option1)
		Me.Frame2.ResumeLayout(False)
		Me.Frame1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class