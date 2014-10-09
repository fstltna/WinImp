<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmSatellitePath
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
	Public WithEvents txtOrbitalStartLocation As System.Windows.Forms.TextBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents lstPaths As System.Windows.Forms.ListBox
	Public WithEvents cbSatellites As System.Windows.Forms.ComboBox
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmSatellitePath))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.txtOrbitalStartLocation = New System.Windows.Forms.TextBox
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.lstPaths = New System.Windows.Forms.ListBox
		Me.cbSatellites = New System.Windows.Forms.ComboBox
		Me.Label1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "Satellite Path"
		Me.ClientSize = New System.Drawing.Size(178, 215)
		Me.Location = New System.Drawing.Point(4, 30)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Visible = False
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmSatellitePath"
		Me.txtOrbitalStartLocation.AutoSize = False
		Me.txtOrbitalStartLocation.Size = New System.Drawing.Size(81, 19)
		Me.txtOrbitalStartLocation.Location = New System.Drawing.Point(88, 40)
		Me.txtOrbitalStartLocation.TabIndex = 1
		Me.txtOrbitalStartLocation.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtOrbitalStartLocation.AcceptsReturn = True
		Me.txtOrbitalStartLocation.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtOrbitalStartLocation.BackColor = System.Drawing.SystemColors.Window
		Me.txtOrbitalStartLocation.CausesValidation = True
		Me.txtOrbitalStartLocation.Enabled = True
		Me.txtOrbitalStartLocation.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtOrbitalStartLocation.HideSelection = True
		Me.txtOrbitalStartLocation.ReadOnly = False
		Me.txtOrbitalStartLocation.Maxlength = 0
		Me.txtOrbitalStartLocation.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtOrbitalStartLocation.MultiLine = False
		Me.txtOrbitalStartLocation.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtOrbitalStartLocation.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtOrbitalStartLocation.TabStop = True
		Me.txtOrbitalStartLocation.Visible = True
		Me.txtOrbitalStartLocation.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtOrbitalStartLocation.Name = "txtOrbitalStartLocation"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "Close"
		Me.AcceptButton = Me.cmdCancel
		Me.cmdCancel.Size = New System.Drawing.Size(49, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(120, 8)
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
		Me.lstPaths.Size = New System.Drawing.Size(161, 135)
		Me.lstPaths.IntegralHeight = False
		Me.lstPaths.Location = New System.Drawing.Point(8, 72)
		Me.lstPaths.TabIndex = 3
		Me.lstPaths.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstPaths.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstPaths.BackColor = System.Drawing.SystemColors.Window
		Me.lstPaths.CausesValidation = True
		Me.lstPaths.Enabled = True
		Me.lstPaths.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstPaths.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstPaths.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstPaths.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstPaths.Sorted = False
		Me.lstPaths.TabStop = True
		Me.lstPaths.Visible = True
		Me.lstPaths.MultiColumn = False
		Me.lstPaths.Name = "lstPaths"
		Me.cbSatellites.Size = New System.Drawing.Size(105, 21)
		Me.cbSatellites.Location = New System.Drawing.Point(8, 8)
		Me.cbSatellites.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cbSatellites.TabIndex = 0
		Me.cbSatellites.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbSatellites.BackColor = System.Drawing.SystemColors.Window
		Me.cbSatellites.CausesValidation = True
		Me.cbSatellites.Enabled = True
		Me.cbSatellites.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cbSatellites.IntegralHeight = True
		Me.cbSatellites.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbSatellites.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbSatellites.Sorted = False
		Me.cbSatellites.TabStop = True
		Me.cbSatellites.Visible = True
		Me.cbSatellites.Name = "cbSatellites"
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label1.Text = "Orbital Start Location"
		Me.Label1.Size = New System.Drawing.Size(65, 25)
		Me.Label1.Location = New System.Drawing.Point(8, 40)
		Me.Label1.TabIndex = 4
		Me.Label1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.Controls.Add(txtOrbitalStartLocation)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(lstPaths)
		Me.Controls.Add(cbSatellites)
		Me.Controls.Add(Label1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class