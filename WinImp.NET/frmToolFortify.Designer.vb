<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmToolFortify
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
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents _Label2_5 As System.Windows.Forms.Label
	Public WithEvents _Label2_4 As System.Windows.Forms.Label
	Public WithEvents _Label2_3 As System.Windows.Forms.Label
	Public WithEvents _Label2_2 As System.Windows.Forms.Label
	Public WithEvents _Label2_1 As System.Windows.Forms.Label
	Public WithEvents _Label2_0 As System.Windows.Forms.Label
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents grid1 As AxMSFlexGridLib.AxMSFlexGrid
	Public WithEvents Label2 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmToolFortify))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me.Frame2 = New System.Windows.Forms.GroupBox
		Me._Label2_5 = New System.Windows.Forms.Label
		Me._Label2_4 = New System.Windows.Forms.Label
		Me._Label2_3 = New System.Windows.Forms.Label
		Me._Label2_2 = New System.Windows.Forms.Label
		Me._Label2_1 = New System.Windows.Forms.Label
		Me._Label2_0 = New System.Windows.Forms.Label
		Me.grid1 = New AxMSFlexGridLib.AxMSFlexGrid
		Me.Label2 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.Frame2.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.grid1, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.Label2, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Fortification Tool"
		Me.ClientSize = New System.Drawing.Size(484, 326)
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
		Me.Name = "frmToolFortify"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(97, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(232, 280)
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
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "Fortify"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Size = New System.Drawing.Size(97, 25)
		Me.cmdOK.Location = New System.Drawing.Point(232, 240)
		Me.cmdOK.TabIndex = 4
		Me.cmdOK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.Frame2.Text = "Game Settings"
		Me.Frame2.Size = New System.Drawing.Size(185, 73)
		Me.Frame2.Location = New System.Drawing.Point(8, 240)
		Me.Frame2.TabIndex = 0
		Me.Frame2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame2.BackColor = System.Drawing.SystemColors.Control
		Me.Frame2.Enabled = True
		Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame2.Visible = True
		Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame2.Name = "Frame2"
		Me._Label2_5.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label2_5.Text = "67"
		Me._Label2_5.Size = New System.Drawing.Size(25, 17)
		Me._Label2_5.Location = New System.Drawing.Point(136, 48)
		Me._Label2_5.TabIndex = 9
		Me._Label2_5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label2_5.BackColor = System.Drawing.SystemColors.Control
		Me._Label2_5.Enabled = True
		Me._Label2_5.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label2_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label2_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label2_5.UseMnemonic = True
		Me._Label2_5.Visible = True
		Me._Label2_5.AutoSize = False
		Me._Label2_5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label2_5.Name = "_Label2_5"
		Me._Label2_4.Text = "Cutoff Level: "
		Me._Label2_4.Size = New System.Drawing.Size(121, 17)
		Me._Label2_4.Location = New System.Drawing.Point(16, 48)
		Me._Label2_4.TabIndex = 8
		Me._Label2_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label2_4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label2_4.BackColor = System.Drawing.SystemColors.Control
		Me._Label2_4.Enabled = True
		Me._Label2_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label2_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label2_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label2_4.UseMnemonic = True
		Me._Label2_4.Visible = True
		Me._Label2_4.AutoSize = False
		Me._Label2_4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label2_4.Name = "_Label2_4"
		Me._Label2_3.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label2_3.Text = "127"
		Me._Label2_3.Size = New System.Drawing.Size(25, 17)
		Me._Label2_3.Location = New System.Drawing.Point(136, 16)
		Me._Label2_3.TabIndex = 7
		Me._Label2_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label2_3.BackColor = System.Drawing.SystemColors.Control
		Me._Label2_3.Enabled = True
		Me._Label2_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label2_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label2_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label2_3.UseMnemonic = True
		Me._Label2_3.Visible = True
		Me._Label2_3.AutoSize = False
		Me._Label2_3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label2_3.Name = "_Label2_3"
		Me._Label2_2.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me._Label2_2.Text = "60"
		Me._Label2_2.Size = New System.Drawing.Size(25, 17)
		Me._Label2_2.Location = New System.Drawing.Point(136, 32)
		Me._Label2_2.TabIndex = 6
		Me._Label2_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label2_2.BackColor = System.Drawing.SystemColors.Control
		Me._Label2_2.Enabled = True
		Me._Label2_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label2_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label2_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label2_2.UseMnemonic = True
		Me._Label2_2.Visible = True
		Me._Label2_2.AutoSize = False
		Me._Label2_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label2_2.Name = "_Label2_2"
		Me._Label2_1.Text = "Maximum Moblity: "
		Me._Label2_1.Size = New System.Drawing.Size(121, 17)
		Me._Label2_1.Location = New System.Drawing.Point(16, 16)
		Me._Label2_1.TabIndex = 2
		Me._Label2_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label2_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label2_1.BackColor = System.Drawing.SystemColors.Control
		Me._Label2_1.Enabled = True
		Me._Label2_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label2_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label2_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label2_1.UseMnemonic = True
		Me._Label2_1.Visible = True
		Me._Label2_1.AutoSize = False
		Me._Label2_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label2_1.Name = "_Label2_1"
		Me._Label2_0.Text = "Mobilty Gain per Update:"
		Me._Label2_0.Size = New System.Drawing.Size(121, 17)
		Me._Label2_0.Location = New System.Drawing.Point(16, 32)
		Me._Label2_0.TabIndex = 1
		Me._Label2_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._Label2_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._Label2_0.BackColor = System.Drawing.SystemColors.Control
		Me._Label2_0.Enabled = True
		Me._Label2_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._Label2_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._Label2_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._Label2_0.UseMnemonic = True
		Me._Label2_0.Visible = True
		Me._Label2_0.AutoSize = False
		Me._Label2_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._Label2_0.Name = "_Label2_0"
		grid1.OcxState = CType(resources.GetObject("grid1.OcxState"), System.Windows.Forms.AxHost.State)
		Me.grid1.Size = New System.Drawing.Size(465, 217)
		Me.grid1.Location = New System.Drawing.Point(8, 8)
		Me.grid1.TabIndex = 3
		Me.grid1.Name = "grid1"
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(Frame2)
		Me.Controls.Add(grid1)
		Me.Frame2.Controls.Add(_Label2_5)
		Me.Frame2.Controls.Add(_Label2_4)
		Me.Frame2.Controls.Add(_Label2_3)
		Me.Frame2.Controls.Add(_Label2_2)
		Me.Frame2.Controls.Add(_Label2_1)
		Me.Frame2.Controls.Add(_Label2_0)
		Me.Label2.SetIndex(_Label2_5, CType(5, Short))
		Me.Label2.SetIndex(_Label2_4, CType(4, Short))
		Me.Label2.SetIndex(_Label2_3, CType(3, Short))
		Me.Label2.SetIndex(_Label2_2, CType(2, Short))
		Me.Label2.SetIndex(_Label2_1, CType(1, Short))
		Me.Label2.SetIndex(_Label2_0, CType(0, Short))
		CType(Me.Label2, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.grid1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Frame2.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class