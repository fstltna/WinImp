<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmEmpCmd
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
	Public WithEvents tmrXMLRPC As System.Windows.Forms.Timer
	Public WithEvents Text1 As System.Windows.Forms.TextBox
	Public WithEvents List1 As System.Windows.Forms.ListBox
	Public WithEvents Winsock As AxMSWinsockLib.AxWinsock
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents imgLights As System.Windows.Forms.PictureBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmEmpCmd))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.tmrXMLRPC = New System.Windows.Forms.Timer(components)
		Me.Text1 = New System.Windows.Forms.TextBox
		Me.List1 = New System.Windows.Forms.ListBox
		Me.Winsock = New AxMSWinsockLib.AxWinsock
		Me.Label1 = New System.Windows.Forms.Label
		Me.imgLights = New System.Windows.Forms.PictureBox
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.Winsock, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.Text = "Command Interface"
		Me.ClientSize = New System.Drawing.Size(583, 337)
		Me.Location = New System.Drawing.Point(4, 23)
		Me.Icon = CType(resources.GetObject("frmEmpCmd.Icon"), System.Drawing.Icon)
		Me.KeyPreview = True
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmEmpCmd"
		Me.tmrXMLRPC.Enabled = False
		Me.tmrXMLRPC.Interval = 1000
		Me.Text1.AutoSize = False
		Me.Text1.Font = New System.Drawing.Font("Courier New", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Text1.Size = New System.Drawing.Size(601, 25)
		Me.Text1.Location = New System.Drawing.Point(0, 312)
		Me.Text1.TabIndex = 1
		Me.Text1.AcceptsReturn = True
		Me.Text1.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.Text1.BackColor = System.Drawing.SystemColors.Window
		Me.Text1.CausesValidation = True
		Me.Text1.Enabled = True
		Me.Text1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Text1.HideSelection = True
		Me.Text1.ReadOnly = False
		Me.Text1.Maxlength = 0
		Me.Text1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.Text1.MultiLine = False
		Me.Text1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text1.TabStop = True
		Me.Text1.Visible = True
		Me.Text1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text1.Name = "Text1"
		Me.List1.Font = New System.Drawing.Font("Courier New", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.List1.Size = New System.Drawing.Size(601, 262)
		Me.List1.Location = New System.Drawing.Point(0, 0)
		Me.List1.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
		Me.List1.TabIndex = 0
		Me.List1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.List1.BackColor = System.Drawing.SystemColors.Window
		Me.List1.CausesValidation = True
		Me.List1.Enabled = True
		Me.List1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.List1.IntegralHeight = True
		Me.List1.Cursor = System.Windows.Forms.Cursors.Default
		Me.List1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.List1.Sorted = False
		Me.List1.TabStop = True
		Me.List1.Visible = True
		Me.List1.MultiColumn = False
		Me.List1.Name = "List1"
		Winsock.OcxState = CType(resources.GetObject("Winsock.OcxState"), System.Windows.Forms.AxHost.State)
		Me.Winsock.Location = New System.Drawing.Point(568, 288)
		Me.Winsock.Name = "Winsock"
		Me.Label1.Text = "Press F2 to execute last command, F9 to select Previous command, Shift F9 to select Next command"
		Me.Label1.Size = New System.Drawing.Size(481, 17)
		Me.Label1.Location = New System.Drawing.Point(56, 288)
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
		Me.imgLights.Size = New System.Drawing.Size(41, 17)
		Me.imgLights.Location = New System.Drawing.Point(8, 288)
		Me.imgLights.Enabled = True
		Me.imgLights.Cursor = System.Windows.Forms.Cursors.Default
		Me.imgLights.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.imgLights.Visible = True
		Me.imgLights.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.imgLights.Name = "imgLights"
		Me.Controls.Add(Text1)
		Me.Controls.Add(List1)
		Me.Controls.Add(Winsock)
		Me.Controls.Add(Label1)
		Me.Controls.Add(imgLights)
		CType(Me.Winsock, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class