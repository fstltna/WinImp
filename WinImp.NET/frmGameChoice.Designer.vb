<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmGameChoice
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
	Public WithEvents cmdNew As System.Windows.Forms.Button
	Public WithEvents cmdOpen As System.Windows.Forms.Button
	Public WithEvents _cmdMRU_4 As System.Windows.Forms.Button
	Public WithEvents _cmdMRU_3 As System.Windows.Forms.Button
	Public WithEvents _cmdMRU_2 As System.Windows.Forms.Button
	Public WithEvents _cmdMRU_1 As System.Windows.Forms.Button
	Public WithEvents _cmdMRU_0 As System.Windows.Forms.Button
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents cmdMRU As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmGameChoice))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdNew = New System.Windows.Forms.Button
		Me.cmdOpen = New System.Windows.Forms.Button
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me._cmdMRU_4 = New System.Windows.Forms.Button
		Me._cmdMRU_3 = New System.Windows.Forms.Button
		Me._cmdMRU_2 = New System.Windows.Forms.Button
		Me._cmdMRU_1 = New System.Windows.Forms.Button
		Me._cmdMRU_0 = New System.Windows.Forms.Button
		Me.cmdMRU = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(components)
		Me.Frame1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.cmdMRU, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Choose Empire Game"
		Me.ClientSize = New System.Drawing.Size(422, 257)
		Me.Location = New System.Drawing.Point(3, 22)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmGameChoice"
		Me.cmdNew.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdNew.Text = "New Game"
		Me.cmdNew.Size = New System.Drawing.Size(161, 33)
		Me.cmdNew.Location = New System.Drawing.Point(232, 64)
		Me.cmdNew.TabIndex = 7
		Me.cmdNew.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdNew.BackColor = System.Drawing.SystemColors.Control
		Me.cmdNew.CausesValidation = True
		Me.cmdNew.Enabled = True
		Me.cmdNew.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdNew.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdNew.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdNew.TabStop = True
		Me.cmdNew.Name = "cmdNew"
		Me.cmdOpen.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOpen.Text = "Show Other Games"
		Me.cmdOpen.Size = New System.Drawing.Size(161, 33)
		Me.cmdOpen.Location = New System.Drawing.Point(232, 24)
		Me.cmdOpen.TabIndex = 6
		Me.cmdOpen.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOpen.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOpen.CausesValidation = True
		Me.cmdOpen.Enabled = True
		Me.cmdOpen.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOpen.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOpen.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOpen.TabStop = True
		Me.cmdOpen.Name = "cmdOpen"
		Me.Frame1.Text = "Most Recently Played Games"
		Me.Frame1.Size = New System.Drawing.Size(209, 225)
		Me.Frame1.Location = New System.Drawing.Point(16, 16)
		Me.Frame1.TabIndex = 0
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me._cmdMRU_4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdMRU_4.Text = "cmdMRU"
		Me._cmdMRU_4.Size = New System.Drawing.Size(161, 33)
		Me._cmdMRU_4.Location = New System.Drawing.Point(24, 184)
		Me._cmdMRU_4.TabIndex = 5
		Me._cmdMRU_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdMRU_4.BackColor = System.Drawing.SystemColors.Control
		Me._cmdMRU_4.CausesValidation = True
		Me._cmdMRU_4.Enabled = True
		Me._cmdMRU_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdMRU_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdMRU_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdMRU_4.TabStop = True
		Me._cmdMRU_4.Name = "_cmdMRU_4"
		Me._cmdMRU_3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdMRU_3.Text = "cmdMRU"
		Me._cmdMRU_3.Size = New System.Drawing.Size(161, 33)
		Me._cmdMRU_3.Location = New System.Drawing.Point(24, 144)
		Me._cmdMRU_3.TabIndex = 4
		Me._cmdMRU_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdMRU_3.BackColor = System.Drawing.SystemColors.Control
		Me._cmdMRU_3.CausesValidation = True
		Me._cmdMRU_3.Enabled = True
		Me._cmdMRU_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdMRU_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdMRU_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdMRU_3.TabStop = True
		Me._cmdMRU_3.Name = "_cmdMRU_3"
		Me._cmdMRU_2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdMRU_2.Text = "cmdMRU"
		Me._cmdMRU_2.Size = New System.Drawing.Size(161, 33)
		Me._cmdMRU_2.Location = New System.Drawing.Point(24, 104)
		Me._cmdMRU_2.TabIndex = 3
		Me._cmdMRU_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdMRU_2.BackColor = System.Drawing.SystemColors.Control
		Me._cmdMRU_2.CausesValidation = True
		Me._cmdMRU_2.Enabled = True
		Me._cmdMRU_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdMRU_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdMRU_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdMRU_2.TabStop = True
		Me._cmdMRU_2.Name = "_cmdMRU_2"
		Me._cmdMRU_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdMRU_1.Text = "cmdMRU"
		Me._cmdMRU_1.Size = New System.Drawing.Size(161, 33)
		Me._cmdMRU_1.Location = New System.Drawing.Point(24, 64)
		Me._cmdMRU_1.TabIndex = 2
		Me._cmdMRU_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdMRU_1.BackColor = System.Drawing.SystemColors.Control
		Me._cmdMRU_1.CausesValidation = True
		Me._cmdMRU_1.Enabled = True
		Me._cmdMRU_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdMRU_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdMRU_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdMRU_1.TabStop = True
		Me._cmdMRU_1.Name = "_cmdMRU_1"
		Me._cmdMRU_0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdMRU_0.Text = "cmdMRU"
		Me._cmdMRU_0.Size = New System.Drawing.Size(161, 33)
		Me._cmdMRU_0.Location = New System.Drawing.Point(24, 24)
		Me._cmdMRU_0.TabIndex = 1
		Me._cmdMRU_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdMRU_0.BackColor = System.Drawing.SystemColors.Control
		Me._cmdMRU_0.CausesValidation = True
		Me._cmdMRU_0.Enabled = True
		Me._cmdMRU_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdMRU_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdMRU_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdMRU_0.TabStop = True
		Me._cmdMRU_0.Name = "_cmdMRU_0"
		Me.Controls.Add(cmdNew)
		Me.Controls.Add(cmdOpen)
		Me.Controls.Add(Frame1)
		Me.Frame1.Controls.Add(_cmdMRU_4)
		Me.Frame1.Controls.Add(_cmdMRU_3)
		Me.Frame1.Controls.Add(_cmdMRU_2)
		Me.Frame1.Controls.Add(_cmdMRU_1)
		Me.Frame1.Controls.Add(_cmdMRU_0)
		Me.cmdMRU.SetIndex(_cmdMRU_4, CType(4, Short))
		Me.cmdMRU.SetIndex(_cmdMRU_3, CType(3, Short))
		Me.cmdMRU.SetIndex(_cmdMRU_2, CType(2, Short))
		Me.cmdMRU.SetIndex(_cmdMRU_1, CType(1, Short))
		Me.cmdMRU.SetIndex(_cmdMRU_0, CType(0, Short))
		CType(Me.cmdMRU, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Frame1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class