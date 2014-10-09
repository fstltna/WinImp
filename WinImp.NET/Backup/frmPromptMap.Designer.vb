<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPromptMap
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
	Public WithEvents _Check1_2 As System.Windows.Forms.CheckBox
	Public WithEvents _Check1_1 As System.Windows.Forms.CheckBox
	Public WithEvents _Check1_0 As System.Windows.Forms.CheckBox
	Public WithEvents txtOrigin As System.Windows.Forms.TextBox
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents txtUnit As System.Windows.Forms.TextBox
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Check1 As Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPromptMap))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdHelp = New System.Windows.Forms.Button
		Me._Check1_2 = New System.Windows.Forms.CheckBox
		Me._Check1_1 = New System.Windows.Forms.CheckBox
		Me._Check1_0 = New System.Windows.Forms.CheckBox
		Me.txtOrigin = New System.Windows.Forms.TextBox
		Me.cmdOK = New System.Windows.Forms.Button
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.txtUnit = New System.Windows.Forms.TextBox
		Me.Label2 = New System.Windows.Forms.Label
		Me.Check1 = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.Check1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.ControlBox = False
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.ClientSize = New System.Drawing.Size(376, 106)
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
		Me.Name = "frmPromptMap"
		Me.cmdHelp.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdHelp.Text = "?"
		Me.cmdHelp.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdHelp.Size = New System.Drawing.Size(25, 25)
		Me.cmdHelp.Location = New System.Drawing.Point(352, 0)
		Me.cmdHelp.TabIndex = 8
		Me.ToolTip1.SetToolTip(Me.cmdHelp, "Click for Help")
		Me.cmdHelp.BackColor = System.Drawing.SystemColors.Control
		Me.cmdHelp.CausesValidation = True
		Me.cmdHelp.Enabled = True
		Me.cmdHelp.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdHelp.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdHelp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdHelp.TabStop = True
		Me.cmdHelp.Name = "cmdHelp"
		Me._Check1_2.Text = "Show Planes"
		Me._Check1_2.Size = New System.Drawing.Size(97, 13)
		Me._Check1_2.Location = New System.Drawing.Point(224, 64)
		Me._Check1_2.TabIndex = 7
		Me._Check1_2.Tag = "p"
		Me._Check1_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Check1_2.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Check1_2.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._Check1_2.BackColor = System.Drawing.SystemColors.Control
		Me._Check1_2.CausesValidation = True
		Me._Check1_2.Enabled = True
		Me._Check1_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Check1_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._Check1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Check1_2.Appearance = System.Windows.Forms.Appearance.Normal
		Me._Check1_2.TabStop = True
		Me._Check1_2.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._Check1_2.Visible = True
		Me._Check1_2.Name = "_Check1_2"
		Me._Check1_1.Text = "Show Lands"
		Me._Check1_1.Size = New System.Drawing.Size(97, 17)
		Me._Check1_1.Location = New System.Drawing.Point(224, 40)
		Me._Check1_1.TabIndex = 6
		Me._Check1_1.Tag = "l"
		Me._Check1_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Check1_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Check1_1.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._Check1_1.BackColor = System.Drawing.SystemColors.Control
		Me._Check1_1.CausesValidation = True
		Me._Check1_1.Enabled = True
		Me._Check1_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Check1_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Check1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Check1_1.Appearance = System.Windows.Forms.Appearance.Normal
		Me._Check1_1.TabStop = True
		Me._Check1_1.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._Check1_1.Visible = True
		Me._Check1_1.Name = "_Check1_1"
		Me._Check1_0.Text = "Show Ships"
		Me._Check1_0.Size = New System.Drawing.Size(105, 17)
		Me._Check1_0.Location = New System.Drawing.Point(224, 16)
		Me._Check1_0.TabIndex = 5
		Me._Check1_0.Tag = "s"
		Me._Check1_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Check1_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._Check1_0.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._Check1_0.BackColor = System.Drawing.SystemColors.Control
		Me._Check1_0.CausesValidation = True
		Me._Check1_0.Enabled = True
		Me._Check1_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Check1_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Check1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Check1_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._Check1_0.TabStop = True
		Me._Check1_0.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._Check1_0.Visible = True
		Me._Check1_0.Name = "_Check1_0"
		Me.txtOrigin.AutoSize = False
		Me.txtOrigin.Size = New System.Drawing.Size(89, 19)
		Me.txtOrigin.Location = New System.Drawing.Point(72, 16)
		Me.txtOrigin.TabIndex = 4
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
		Me.cmdOK.Location = New System.Drawing.Point(24, 56)
		Me.cmdOK.TabIndex = 2
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
		Me.cmdCancel.Location = New System.Drawing.Point(112, 56)
		Me.cmdCancel.TabIndex = 1
		Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.txtUnit.AutoSize = False
		Me.txtUnit.Size = New System.Drawing.Size(89, 19)
		Me.txtUnit.Location = New System.Drawing.Point(72, 16)
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
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label2.Text = "Map"
		Me.Label2.Font = New System.Drawing.Font("Arial Black", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(49, 17)
		Me.Label2.Location = New System.Drawing.Point(8, 16)
		Me.Label2.TabIndex = 3
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
		Me.Controls.Add(_Check1_2)
		Me.Controls.Add(_Check1_1)
		Me.Controls.Add(_Check1_0)
		Me.Controls.Add(txtOrigin)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(txtUnit)
		Me.Controls.Add(Label2)
		Me.Check1.SetIndex(_Check1_2, CType(2, Short))
		Me.Check1.SetIndex(_Check1_1, CType(1, Short))
		Me.Check1.SetIndex(_Check1_0, CType(0, Short))
		CType(Me.Check1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class