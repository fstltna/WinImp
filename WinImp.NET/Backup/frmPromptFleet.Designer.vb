<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPromptFleet
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
	Public WithEvents cmbSub As System.Windows.Forms.ComboBox
	Public WithEvents txtUnit As System.Windows.Forms.TextBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPromptFleet))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdHelp = New System.Windows.Forms.Button
		Me.cmbSub = New System.Windows.Forms.ComboBox
		Me.txtUnit = New System.Windows.Forms.TextBox
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me.Label1 = New System.Windows.Forms.Label
		Me.Label5 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.ControlBox = False
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.ClientSize = New System.Drawing.Size(392, 93)
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
		Me.Name = "frmPromptFleet"
		Me.cmdHelp.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdHelp.Text = "?"
		Me.cmdHelp.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdHelp.Size = New System.Drawing.Size(25, 25)
		Me.cmdHelp.Location = New System.Drawing.Point(368, 0)
		Me.cmdHelp.TabIndex = 7
		Me.ToolTip1.SetToolTip(Me.cmdHelp, "Click for Help")
		Me.cmdHelp.BackColor = System.Drawing.SystemColors.Control
		Me.cmdHelp.CausesValidation = True
		Me.cmdHelp.Enabled = True
		Me.cmdHelp.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdHelp.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdHelp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdHelp.TabStop = True
		Me.cmdHelp.Name = "cmdHelp"
		Me.cmbSub.Size = New System.Drawing.Size(57, 21)
		Me.cmbSub.Location = New System.Drawing.Point(304, 16)
		Me.cmbSub.Sorted = True
		Me.cmbSub.TabIndex = 6
		Me.cmbSub.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbSub.BackColor = System.Drawing.SystemColors.Window
		Me.cmbSub.CausesValidation = True
		Me.cmbSub.Enabled = True
		Me.cmbSub.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbSub.IntegralHeight = True
		Me.cmbSub.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbSub.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbSub.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cmbSub.TabStop = True
		Me.cmbSub.Visible = True
		Me.cmbSub.Name = "cmbSub"
		Me.txtUnit.AutoSize = False
		Me.txtUnit.Size = New System.Drawing.Size(161, 19)
		Me.txtUnit.Location = New System.Drawing.Point(48, 16)
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
		Me.cmdCancel.Location = New System.Drawing.Point(184, 56)
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
		Me.cmdOK.Location = New System.Drawing.Point(80, 56)
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
		Me.Label1.Text = "to"
		Me.Label1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Size = New System.Drawing.Size(17, 17)
		Me.Label1.Location = New System.Drawing.Point(224, 16)
		Me.Label1.TabIndex = 5
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
		Me.Label5.Text = "Add"
		Me.Label5.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label5.Size = New System.Drawing.Size(33, 17)
		Me.Label5.Location = New System.Drawing.Point(8, 16)
		Me.Label5.TabIndex = 4
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
		Me.Label2.Text = "Fleet"
		Me.Label2.Font = New System.Drawing.Font("Arial Black", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(49, 17)
		Me.Label2.Location = New System.Drawing.Point(248, 16)
		Me.Label2.TabIndex = 3
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
		Me.Controls.Add(cmdHelp)
		Me.Controls.Add(cmbSub)
		Me.Controls.Add(txtUnit)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(Label1)
		Me.Controls.Add(Label5)
		Me.Controls.Add(Label2)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class