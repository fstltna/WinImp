<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmServerQuery
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
	Public WithEvents cmdAbort As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents txtRespond As System.Windows.Forms.TextBox
	Public WithEvents lstMsgs As System.Windows.Forms.ListBox
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmServerQuery))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdAbort = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me.txtRespond = New System.Windows.Forms.TextBox
		Me.lstMsgs = New System.Windows.Forms.ListBox
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Server Query"
		Me.ClientSize = New System.Drawing.Size(540, 303)
		Me.Location = New System.Drawing.Point(3, 22)
		Me.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmServerQuery"
		Me.cmdAbort.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdAbort.Text = "Abort"
		Me.cmdAbort.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdAbort.Size = New System.Drawing.Size(81, 33)
		Me.cmdAbort.Location = New System.Drawing.Point(288, 256)
		Me.cmdAbort.TabIndex = 4
		Me.cmdAbort.BackColor = System.Drawing.SystemColors.Control
		Me.cmdAbort.CausesValidation = True
		Me.cmdAbort.Enabled = True
		Me.cmdAbort.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdAbort.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdAbort.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdAbort.TabStop = True
		Me.cmdAbort.Name = "cmdAbort"
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "OK"
		Me.cmdOK.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.Size = New System.Drawing.Size(81, 33)
		Me.cmdOK.Location = New System.Drawing.Point(176, 256)
		Me.cmdOK.TabIndex = 3
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.txtRespond.AutoSize = False
		Me.txtRespond.Size = New System.Drawing.Size(473, 25)
		Me.txtRespond.Location = New System.Drawing.Point(24, 224)
		Me.txtRespond.TabIndex = 1
		Me.txtRespond.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtRespond.AcceptsReturn = True
		Me.txtRespond.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtRespond.BackColor = System.Drawing.SystemColors.Window
		Me.txtRespond.CausesValidation = True
		Me.txtRespond.Enabled = True
		Me.txtRespond.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtRespond.HideSelection = True
		Me.txtRespond.ReadOnly = False
		Me.txtRespond.Maxlength = 0
		Me.txtRespond.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtRespond.MultiLine = False
		Me.txtRespond.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtRespond.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtRespond.TabStop = True
		Me.txtRespond.Visible = True
		Me.txtRespond.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtRespond.Name = "txtRespond"
		Me.lstMsgs.Size = New System.Drawing.Size(497, 167)
		Me.lstMsgs.Location = New System.Drawing.Point(24, 16)
		Me.lstMsgs.TabIndex = 0
		Me.lstMsgs.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstMsgs.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstMsgs.BackColor = System.Drawing.SystemColors.Window
		Me.lstMsgs.CausesValidation = True
		Me.lstMsgs.Enabled = True
		Me.lstMsgs.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstMsgs.IntegralHeight = True
		Me.lstMsgs.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstMsgs.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstMsgs.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstMsgs.Sorted = False
		Me.lstMsgs.TabStop = True
		Me.lstMsgs.Visible = True
		Me.lstMsgs.MultiColumn = False
		Me.lstMsgs.Name = "lstMsgs"
		Me.Label1.Text = "Enter Response"
		Me.Label1.Size = New System.Drawing.Size(137, 17)
		Me.Label1.Location = New System.Drawing.Point(24, 192)
		Me.Label1.TabIndex = 2
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
		Me.Controls.Add(cmdAbort)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(txtRespond)
		Me.Controls.Add(lstMsgs)
		Me.Controls.Add(Label1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class