<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPromptImprove
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
	Public WithEvents optNavigation As System.Windows.Forms.RadioButton
	Public WithEvents optFort As System.Windows.Forms.RadioButton
	Public WithEvents optRadar As System.Windows.Forms.RadioButton
	Public WithEvents optRunway As System.Windows.Forms.RadioButton
	Public WithEvents cmdHelp As System.Windows.Forms.Button
	Public WithEvents txtNum As System.Windows.Forms.TextBox
	Public WithEvents txtOrigin As System.Windows.Forms.TextBox
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents optDefense As System.Windows.Forms.RadioButton
	Public WithEvents optRail As System.Windows.Forms.RadioButton
	Public WithEvents optRoad As System.Windows.Forms.RadioButton
	Public WithEvents lblDesc As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPromptImprove))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.optNavigation = New System.Windows.Forms.RadioButton
		Me.optFort = New System.Windows.Forms.RadioButton
		Me.optRadar = New System.Windows.Forms.RadioButton
		Me.optRunway = New System.Windows.Forms.RadioButton
		Me.cmdHelp = New System.Windows.Forms.Button
		Me.txtNum = New System.Windows.Forms.TextBox
		Me.txtOrigin = New System.Windows.Forms.TextBox
		Me.cmdOK = New System.Windows.Forms.Button
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.optDefense = New System.Windows.Forms.RadioButton
		Me.optRail = New System.Windows.Forms.RadioButton
		Me.optRoad = New System.Windows.Forms.RadioButton
		Me.lblDesc = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.ControlBox = False
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.ClientSize = New System.Drawing.Size(456, 110)
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
		Me.Name = "frmPromptImprove"
		Me.optNavigation.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optNavigation.Text = "navigation"
		Me.optNavigation.Size = New System.Drawing.Size(97, 17)
		Me.optNavigation.Location = New System.Drawing.Point(96, 88)
		Me.optNavigation.TabIndex = 16
		Me.optNavigation.Visible = False
		Me.optNavigation.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optNavigation.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optNavigation.BackColor = System.Drawing.SystemColors.Control
		Me.optNavigation.CausesValidation = True
		Me.optNavigation.Enabled = True
		Me.optNavigation.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optNavigation.Cursor = System.Windows.Forms.Cursors.Default
		Me.optNavigation.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optNavigation.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optNavigation.TabStop = True
		Me.optNavigation.Checked = False
		Me.optNavigation.Name = "optNavigation"
		Me.optFort.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optFort.Text = "fort"
		Me.optFort.Size = New System.Drawing.Size(65, 17)
		Me.optFort.Location = New System.Drawing.Point(16, 88)
		Me.optFort.TabIndex = 15
		Me.optFort.Visible = False
		Me.optFort.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optFort.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optFort.BackColor = System.Drawing.SystemColors.Control
		Me.optFort.CausesValidation = True
		Me.optFort.Enabled = True
		Me.optFort.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optFort.Cursor = System.Windows.Forms.Cursors.Default
		Me.optFort.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optFort.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optFort.TabStop = True
		Me.optFort.Checked = False
		Me.optFort.Name = "optFort"
		Me.optRadar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optRadar.Text = "radar"
		Me.optRadar.Size = New System.Drawing.Size(65, 17)
		Me.optRadar.Location = New System.Drawing.Point(16, 64)
		Me.optRadar.TabIndex = 14
		Me.optRadar.Visible = False
		Me.optRadar.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optRadar.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optRadar.BackColor = System.Drawing.SystemColors.Control
		Me.optRadar.CausesValidation = True
		Me.optRadar.Enabled = True
		Me.optRadar.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optRadar.Cursor = System.Windows.Forms.Cursors.Default
		Me.optRadar.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optRadar.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optRadar.TabStop = True
		Me.optRadar.Checked = False
		Me.optRadar.Name = "optRadar"
		Me.optRunway.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optRunway.Text = "runway"
		Me.optRunway.Size = New System.Drawing.Size(65, 17)
		Me.optRunway.Location = New System.Drawing.Point(16, 40)
		Me.optRunway.TabIndex = 13
		Me.optRunway.Visible = False
		Me.optRunway.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optRunway.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optRunway.BackColor = System.Drawing.SystemColors.Control
		Me.optRunway.CausesValidation = True
		Me.optRunway.Enabled = True
		Me.optRunway.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optRunway.Cursor = System.Windows.Forms.Cursors.Default
		Me.optRunway.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optRunway.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optRunway.TabStop = True
		Me.optRunway.Checked = False
		Me.optRunway.Name = "optRunway"
		Me.cmdHelp.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdHelp.Text = "?"
		Me.cmdHelp.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdHelp.Size = New System.Drawing.Size(25, 25)
		Me.cmdHelp.Location = New System.Drawing.Point(432, 0)
		Me.cmdHelp.TabIndex = 12
		Me.ToolTip1.SetToolTip(Me.cmdHelp, "Click for Help")
		Me.cmdHelp.BackColor = System.Drawing.SystemColors.Control
		Me.cmdHelp.CausesValidation = True
		Me.cmdHelp.Enabled = True
		Me.cmdHelp.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdHelp.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdHelp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdHelp.TabStop = True
		Me.cmdHelp.Name = "cmdHelp"
		Me.txtNum.AutoSize = False
		Me.txtNum.Size = New System.Drawing.Size(49, 19)
		Me.txtNum.Location = New System.Drawing.Point(336, 32)
		Me.txtNum.TabIndex = 1
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
		Me.txtOrigin.AutoSize = False
		Me.txtOrigin.Size = New System.Drawing.Size(89, 19)
		Me.txtOrigin.Location = New System.Drawing.Point(200, 32)
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
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "OK"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Size = New System.Drawing.Size(81, 25)
		Me.cmdOK.Location = New System.Drawing.Point(264, 80)
		Me.cmdOK.TabIndex = 6
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
		Me.cmdCancel.Location = New System.Drawing.Point(360, 80)
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
		Me.optDefense.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optDefense.Text = "defense"
		Me.optDefense.Size = New System.Drawing.Size(65, 17)
		Me.optDefense.Location = New System.Drawing.Point(96, 64)
		Me.optDefense.TabIndex = 4
		Me.optDefense.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optDefense.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optDefense.BackColor = System.Drawing.SystemColors.Control
		Me.optDefense.CausesValidation = True
		Me.optDefense.Enabled = True
		Me.optDefense.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optDefense.Cursor = System.Windows.Forms.Cursors.Default
		Me.optDefense.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optDefense.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optDefense.TabStop = True
		Me.optDefense.Checked = False
		Me.optDefense.Visible = True
		Me.optDefense.Name = "optDefense"
		Me.optRail.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optRail.Text = "rail"
		Me.optRail.Size = New System.Drawing.Size(65, 17)
		Me.optRail.Location = New System.Drawing.Point(96, 40)
		Me.optRail.TabIndex = 3
		Me.optRail.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optRail.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optRail.BackColor = System.Drawing.SystemColors.Control
		Me.optRail.CausesValidation = True
		Me.optRail.Enabled = True
		Me.optRail.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optRail.Cursor = System.Windows.Forms.Cursors.Default
		Me.optRail.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optRail.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optRail.TabStop = True
		Me.optRail.Checked = False
		Me.optRail.Visible = True
		Me.optRail.Name = "optRail"
		Me.optRoad.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optRoad.Text = "road"
		Me.optRoad.Size = New System.Drawing.Size(65, 17)
		Me.optRoad.Location = New System.Drawing.Point(96, 16)
		Me.optRoad.TabIndex = 2
		Me.optRoad.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optRoad.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optRoad.BackColor = System.Drawing.SystemColors.Control
		Me.optRoad.CausesValidation = True
		Me.optRoad.Enabled = True
		Me.optRoad.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optRoad.Cursor = System.Windows.Forms.Cursors.Default
		Me.optRoad.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optRoad.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optRoad.TabStop = True
		Me.optRoad.Checked = False
		Me.optRoad.Visible = True
		Me.optRoad.Name = "optRoad"
		Me.lblDesc.Size = New System.Drawing.Size(257, 17)
		Me.lblDesc.Location = New System.Drawing.Point(200, 56)
		Me.lblDesc.TabIndex = 11
		Me.lblDesc.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblDesc.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblDesc.BackColor = System.Drawing.SystemColors.Control
		Me.lblDesc.Enabled = True
		Me.lblDesc.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblDesc.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblDesc.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblDesc.UseMnemonic = True
		Me.lblDesc.Visible = True
		Me.lblDesc.AutoSize = False
		Me.lblDesc.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblDesc.Name = "lblDesc"
		Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label4.Text = "percent"
		Me.Label4.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.Size = New System.Drawing.Size(57, 17)
		Me.Label4.Location = New System.Drawing.Point(392, 32)
		Me.Label4.TabIndex = 10
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
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label3.Text = "by"
		Me.Label3.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.Size = New System.Drawing.Size(25, 17)
		Me.Label3.Location = New System.Drawing.Point(304, 32)
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
		Me.Label1.Text = "Improve"
		Me.Label1.Font = New System.Drawing.Font("Arial Black", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Size = New System.Drawing.Size(57, 17)
		Me.Label1.Location = New System.Drawing.Point(16, 16)
		Me.Label1.TabIndex = 8
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
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label2.Text = "in"
		Me.Label2.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(25, 17)
		Me.Label2.Location = New System.Drawing.Point(168, 32)
		Me.Label2.TabIndex = 7
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
		Me.Controls.Add(optNavigation)
		Me.Controls.Add(optFort)
		Me.Controls.Add(optRadar)
		Me.Controls.Add(optRunway)
		Me.Controls.Add(cmdHelp)
		Me.Controls.Add(txtNum)
		Me.Controls.Add(txtOrigin)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(optDefense)
		Me.Controls.Add(optRail)
		Me.Controls.Add(optRoad)
		Me.Controls.Add(lblDesc)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label1)
		Me.Controls.Add(Label2)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class