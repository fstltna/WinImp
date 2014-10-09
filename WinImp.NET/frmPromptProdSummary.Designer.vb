<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPromptProdSummary
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
	Public WithEvents cmdAll As System.Windows.Forms.Button
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents txtOrigin As System.Windows.Forms.TextBox
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents chkBuild As System.Windows.Forms.CheckBox
	Public WithEvents chkIdle As System.Windows.Forms.CheckBox
	Public WithEvents ckOvrPop As System.Windows.Forms.CheckBox
	Public WithEvents ckfood As System.Windows.Forms.CheckBox
	Public WithEvents frameShow As System.Windows.Forms.GroupBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents Label2 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPromptProdSummary))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.Frame2 = New System.Windows.Forms.GroupBox
		Me.cmdAll = New System.Windows.Forms.Button
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.cmdOK = New System.Windows.Forms.Button
		Me.txtOrigin = New System.Windows.Forms.TextBox
		Me.Label1 = New System.Windows.Forms.Label
		Me.frameShow = New System.Windows.Forms.GroupBox
		Me.chkBuild = New System.Windows.Forms.CheckBox
		Me.chkIdle = New System.Windows.Forms.CheckBox
		Me.ckOvrPop = New System.Windows.Forms.CheckBox
		Me.ckfood = New System.Windows.Forms.CheckBox
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.Label2 = New System.Windows.Forms.Label
		Me.Frame2.SuspendLayout()
		Me.Frame1.SuspendLayout()
		Me.frameShow.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.ControlBox = False
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.ClientSize = New System.Drawing.Size(328, 198)
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
		Me.Name = "frmPromptProdSummary"
		Me.Frame2.Text = "Show All Distribution Centers"
		Me.Frame2.Size = New System.Drawing.Size(313, 41)
		Me.Frame2.Location = New System.Drawing.Point(8, 32)
		Me.Frame2.TabIndex = 11
		Me.Frame2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame2.BackColor = System.Drawing.SystemColors.Control
		Me.Frame2.Enabled = True
		Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame2.Visible = True
		Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame2.Name = "Frame2"
		Me.cmdAll.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdAll.Text = "Show All"
		Me.cmdAll.Size = New System.Drawing.Size(81, 25)
		Me.cmdAll.Location = New System.Drawing.Point(224, 8)
		Me.cmdAll.TabIndex = 12
		Me.cmdAll.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdAll.BackColor = System.Drawing.SystemColors.Control
		Me.cmdAll.CausesValidation = True
		Me.cmdAll.Enabled = True
		Me.cmdAll.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdAll.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdAll.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdAll.TabStop = True
		Me.cmdAll.Name = "cmdAll"
		Me.Frame1.Text = "Show Single Distribution Center Only"
		Me.Frame1.Size = New System.Drawing.Size(313, 49)
		Me.Frame1.Location = New System.Drawing.Point(8, 80)
		Me.Frame1.TabIndex = 7
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "Show Single"
		Me.cmdOK.Size = New System.Drawing.Size(81, 25)
		Me.cmdOK.Location = New System.Drawing.Point(224, 16)
		Me.cmdOK.TabIndex = 10
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
		Me.txtOrigin.Size = New System.Drawing.Size(65, 19)
		Me.txtOrigin.Location = New System.Drawing.Point(112, 16)
		Me.txtOrigin.TabIndex = 8
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
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label1.Text = "Dist. Center"
		Me.Label1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Size = New System.Drawing.Size(73, 17)
		Me.Label1.Location = New System.Drawing.Point(24, 16)
		Me.Label1.TabIndex = 9
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
		Me.frameShow.Text = "Problem Reports"
		Me.frameShow.Size = New System.Drawing.Size(217, 57)
		Me.frameShow.Location = New System.Drawing.Point(8, 136)
		Me.frameShow.TabIndex = 2
		Me.frameShow.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frameShow.BackColor = System.Drawing.SystemColors.Control
		Me.frameShow.Enabled = True
		Me.frameShow.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frameShow.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frameShow.Visible = True
		Me.frameShow.Padding = New System.Windows.Forms.Padding(0)
		Me.frameShow.Name = "frameShow"
		Me.chkBuild.Text = "Build problems"
		Me.chkBuild.Size = New System.Drawing.Size(89, 17)
		Me.chkBuild.Location = New System.Drawing.Point(120, 32)
		Me.chkBuild.TabIndex = 6
		Me.chkBuild.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkBuild.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkBuild.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkBuild.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkBuild.BackColor = System.Drawing.SystemColors.Control
		Me.chkBuild.CausesValidation = True
		Me.chkBuild.Enabled = True
		Me.chkBuild.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkBuild.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkBuild.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkBuild.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkBuild.TabStop = True
		Me.chkBuild.Visible = True
		Me.chkBuild.Name = "chkBuild"
		Me.chkIdle.Text = "Idle Civs"
		Me.chkIdle.Size = New System.Drawing.Size(81, 17)
		Me.chkIdle.Location = New System.Drawing.Point(120, 16)
		Me.chkIdle.TabIndex = 5
		Me.chkIdle.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkIdle.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkIdle.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkIdle.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkIdle.BackColor = System.Drawing.SystemColors.Control
		Me.chkIdle.CausesValidation = True
		Me.chkIdle.Enabled = True
		Me.chkIdle.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkIdle.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkIdle.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkIdle.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkIdle.TabStop = True
		Me.chkIdle.Visible = True
		Me.chkIdle.Name = "chkIdle"
		Me.ckOvrPop.Text = "Overpopulation"
		Me.ckOvrPop.Size = New System.Drawing.Size(97, 17)
		Me.ckOvrPop.Location = New System.Drawing.Point(16, 32)
		Me.ckOvrPop.TabIndex = 4
		Me.ckOvrPop.CheckState = System.Windows.Forms.CheckState.Checked
		Me.ckOvrPop.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ckOvrPop.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.ckOvrPop.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.ckOvrPop.BackColor = System.Drawing.SystemColors.Control
		Me.ckOvrPop.CausesValidation = True
		Me.ckOvrPop.Enabled = True
		Me.ckOvrPop.ForeColor = System.Drawing.SystemColors.ControlText
		Me.ckOvrPop.Cursor = System.Windows.Forms.Cursors.Default
		Me.ckOvrPop.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ckOvrPop.Appearance = System.Windows.Forms.Appearance.Normal
		Me.ckOvrPop.TabStop = True
		Me.ckOvrPop.Visible = True
		Me.ckOvrPop.Name = "ckOvrPop"
		Me.ckfood.Text = "Food Shortage"
		Me.ckfood.Size = New System.Drawing.Size(97, 17)
		Me.ckfood.Location = New System.Drawing.Point(16, 16)
		Me.ckfood.TabIndex = 3
		Me.ckfood.CheckState = System.Windows.Forms.CheckState.Checked
		Me.ckfood.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ckfood.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.ckfood.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.ckfood.BackColor = System.Drawing.SystemColors.Control
		Me.ckfood.CausesValidation = True
		Me.ckfood.Enabled = True
		Me.ckfood.ForeColor = System.Drawing.SystemColors.ControlText
		Me.ckfood.Cursor = System.Windows.Forms.Cursors.Default
		Me.ckfood.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ckfood.Appearance = System.Windows.Forms.Appearance.Normal
		Me.ckfood.TabStop = True
		Me.ckfood.Visible = True
		Me.ckfood.Name = "ckfood"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(81, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(232, 136)
		Me.cmdCancel.TabIndex = 0
		Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.Label2.Text = "Production Summary"
		Me.Label2.Font = New System.Drawing.Font("Arial Black", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(145, 17)
		Me.Label2.Location = New System.Drawing.Point(16, 8)
		Me.Label2.TabIndex = 1
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
		Me.Controls.Add(Frame2)
		Me.Controls.Add(Frame1)
		Me.Controls.Add(frameShow)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(Label2)
		Me.Frame2.Controls.Add(cmdAll)
		Me.Frame1.Controls.Add(cmdOK)
		Me.Frame1.Controls.Add(txtOrigin)
		Me.Frame1.Controls.Add(Label1)
		Me.frameShow.Controls.Add(chkBuild)
		Me.frameShow.Controls.Add(chkIdle)
		Me.frameShow.Controls.Add(ckOvrPop)
		Me.frameShow.Controls.Add(ckfood)
		Me.Frame2.ResumeLayout(False)
		Me.Frame1.ResumeLayout(False)
		Me.frameShow.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class