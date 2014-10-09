<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmReport
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
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents _tbToolBar_Button1 As System.Windows.Forms.ToolStripButton
	Public WithEvents _tbToolBar_Button2 As System.Windows.Forms.ToolStripButton
	Public WithEvents _tbToolBar_Button3 As System.Windows.Forms.ToolStripButton
	Public WithEvents _tbToolBar_Button4 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents _tbToolBar_Button5 As System.Windows.Forms.ToolStripButton
	Public WithEvents tbToolBar As System.Windows.Forms.ToolStrip
	Public WithEvents Text1 As System.Windows.Forms.TextBox
	Public cd1Open As System.Windows.Forms.OpenFileDialog
	Public cd1Save As System.Windows.Forms.SaveFileDialog
	Public WithEvents imlIcons As System.Windows.Forms.ImageList
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmReport))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdOK = New System.Windows.Forms.Button
		Me.tbToolBar = New System.Windows.Forms.ToolStrip
		Me._tbToolBar_Button1 = New System.Windows.Forms.ToolStripButton
		Me._tbToolBar_Button2 = New System.Windows.Forms.ToolStripButton
		Me._tbToolBar_Button3 = New System.Windows.Forms.ToolStripButton
		Me._tbToolBar_Button4 = New System.Windows.Forms.ToolStripSeparator
		Me._tbToolBar_Button5 = New System.Windows.Forms.ToolStripButton
		Me.Text1 = New System.Windows.Forms.TextBox
		Me.cd1Open = New System.Windows.Forms.OpenFileDialog
		Me.cd1Save = New System.Windows.Forms.SaveFileDialog
		Me.imlIcons = New System.Windows.Forms.ImageList
		Me.tbToolBar.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "Reports"
		Me.ClientSize = New System.Drawing.Size(567, 388)
		Me.Location = New System.Drawing.Point(4, 24)
		Me.Font = New System.Drawing.Font("Courier New", 9.6!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
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
		Me.Name = "frmReport"
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "Command1"
		Me.cmdOK.Size = New System.Drawing.Size(89, 41)
		Me.cmdOK.Location = New System.Drawing.Point(32, 312)
		Me.cmdOK.TabIndex = 2
		Me.cmdOK.Visible = False
		Me.cmdOK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.tbToolBar.ShowItemToolTips = True
		Me.tbToolBar.Dock = System.Windows.Forms.DockStyle.Top
		Me.tbToolBar.Size = New System.Drawing.Size(567, 44)
		Me.tbToolBar.Location = New System.Drawing.Point(0, 0)
		Me.tbToolBar.TabIndex = 0
		Me.tbToolBar.ImageList = imlIcons
		Me.tbToolBar.Name = "tbToolBar"
		Me._tbToolBar_Button1.Size = New System.Drawing.Size(76, 39)
		Me._tbToolBar_Button1.AutoSize = False
		Me._tbToolBar_Button1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._tbToolBar_Button1.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._tbToolBar_Button1.Text = "Save to Disk"
		Me._tbToolBar_Button1.Name = "Save"
		Me._tbToolBar_Button1.ToolTipText = "Save to Disk"
		Me._tbToolBar_Button1.ImageIndex = 0
		Me._tbToolBar_Button1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._tbToolBar_Button2.Size = New System.Drawing.Size(76, 39)
		Me._tbToolBar_Button2.AutoSize = False
		Me._tbToolBar_Button2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._tbToolBar_Button2.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._tbToolBar_Button2.Text = "Copy"
		Me._tbToolBar_Button2.Name = "Copy"
		Me._tbToolBar_Button2.ToolTipText = "Copy to Clipboard"
		Me._tbToolBar_Button2.ImageIndex = 1
		Me._tbToolBar_Button2.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._tbToolBar_Button3.Size = New System.Drawing.Size(76, 39)
		Me._tbToolBar_Button3.AutoSize = False
		Me._tbToolBar_Button3.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._tbToolBar_Button3.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._tbToolBar_Button3.Text = "Save to Log"
		Me._tbToolBar_Button3.Name = "File"
		Me._tbToolBar_Button3.ToolTipText = "Copy to Telegram File"
		Me._tbToolBar_Button3.ImageIndex = 2
		Me._tbToolBar_Button3.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._tbToolBar_Button4.Size = New System.Drawing.Size(76, 39)
		Me._tbToolBar_Button4.AutoSize = False
		Me._tbToolBar_Button4.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._tbToolBar_Button4.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._tbToolBar_Button4.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._tbToolBar_Button5.Size = New System.Drawing.Size(76, 39)
		Me._tbToolBar_Button5.AutoSize = False
		Me._tbToolBar_Button5.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._tbToolBar_Button5.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._tbToolBar_Button5.Text = "Close"
		Me._tbToolBar_Button5.Name = "Exit"
		Me._tbToolBar_Button5.ToolTipText = "Exit "
		Me._tbToolBar_Button5.ImageIndex = 3
		Me._tbToolBar_Button5.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me.Text1.AutoSize = False
		Me.Text1.Size = New System.Drawing.Size(545, 257)
		Me.Text1.Location = New System.Drawing.Point(8, 48)
		Me.Text1.MultiLine = True
		Me.Text1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
		Me.Text1.TabIndex = 1
		Me.Text1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.Text1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text1.TabStop = True
		Me.Text1.Visible = True
		Me.Text1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text1.Name = "Text1"
		Me.imlIcons.ImageSize = New System.Drawing.Size(24, 22)
		Me.imlIcons.TransparentColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me.imlIcons.ImageStream = CType(resources.GetObject("imlIcons.ImageStream"), System.Windows.Forms.ImageListStreamer)
		Me.imlIcons.Images.SetKeyName(0, "")
		Me.imlIcons.Images.SetKeyName(1, "")
		Me.imlIcons.Images.SetKeyName(2, "")
		Me.imlIcons.Images.SetKeyName(3, "")
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(tbToolBar)
		Me.Controls.Add(Text1)
		Me.tbToolBar.Items.Add(_tbToolBar_Button1)
		Me.tbToolBar.Items.Add(_tbToolBar_Button2)
		Me.tbToolBar.Items.Add(_tbToolBar_Button3)
		Me.tbToolBar.Items.Add(_tbToolBar_Button4)
		Me.tbToolBar.Items.Add(_tbToolBar_Button5)
		Me.tbToolBar.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class