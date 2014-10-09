<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPromptIsland
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
	Public WithEvents cbNations As System.Windows.Forms.ComboBox
	Public WithEvents txtDest As System.Windows.Forms.TextBox
	Public WithEvents txtOrigin As System.Windows.Forms.TextBox
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPromptIsland))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cbNations = New System.Windows.Forms.ComboBox
		Me.txtDest = New System.Windows.Forms.TextBox
		Me.txtOrigin = New System.Windows.Forms.TextBox
		Me.cmdOK = New System.Windows.Forms.Button
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.ControlBox = False
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.ClientSize = New System.Drawing.Size(320, 132)
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
		Me.Name = "frmPromptIsland"
		Me.cbNations.Size = New System.Drawing.Size(153, 21)
		Me.cbNations.Location = New System.Drawing.Point(128, 8)
		Me.cbNations.Sorted = True
		Me.cbNations.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cbNations.TabIndex = 7
		Me.cbNations.Visible = False
		Me.cbNations.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbNations.BackColor = System.Drawing.SystemColors.Window
		Me.cbNations.CausesValidation = True
		Me.cbNations.Enabled = True
		Me.cbNations.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cbNations.IntegralHeight = True
		Me.cbNations.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbNations.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbNations.TabStop = True
		Me.cbNations.Name = "cbNations"
		Me.txtDest.AutoSize = False
		Me.txtDest.Size = New System.Drawing.Size(65, 19)
		Me.txtDest.Location = New System.Drawing.Point(128, 72)
		Me.txtDest.TabIndex = 1
		Me.txtDest.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtDest.AcceptsReturn = True
		Me.txtDest.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtDest.BackColor = System.Drawing.SystemColors.Window
		Me.txtDest.CausesValidation = True
		Me.txtDest.Enabled = True
		Me.txtDest.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtDest.HideSelection = True
		Me.txtDest.ReadOnly = False
		Me.txtDest.Maxlength = 0
		Me.txtDest.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtDest.MultiLine = False
		Me.txtDest.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtDest.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtDest.TabStop = True
		Me.txtDest.Visible = True
		Me.txtDest.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtDest.Name = "txtDest"
		Me.txtOrigin.AutoSize = False
		Me.txtOrigin.Size = New System.Drawing.Size(65, 19)
		Me.txtOrigin.Location = New System.Drawing.Point(128, 40)
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
		Me.cmdOK.Size = New System.Drawing.Size(81, 25)
		Me.cmdOK.Location = New System.Drawing.Point(232, 40)
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
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(81, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(232, 72)
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
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label3.Text = "to"
		Me.Label3.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.Size = New System.Drawing.Size(33, 17)
		Me.Label3.Location = New System.Drawing.Point(80, 72)
		Me.Label3.TabIndex = 6
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
		Me.Label2.Text = "Island Setup Tool"
		Me.Label2.Font = New System.Drawing.Font("Arial Black", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(145, 17)
		Me.Label2.Location = New System.Drawing.Point(16, 8)
		Me.Label2.TabIndex = 5
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
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label1.Text = "from"
		Me.Label1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Size = New System.Drawing.Size(97, 17)
		Me.Label1.Location = New System.Drawing.Point(16, 40)
		Me.Label1.TabIndex = 4
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
		Me.Controls.Add(cbNations)
		Me.Controls.Add(txtDest)
		Me.Controls.Add(txtOrigin)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class