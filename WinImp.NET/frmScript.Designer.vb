<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmScript
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
	Public WithEvents _tbToolBar_Button1 As System.Windows.Forms.ToolStripButton
	Public WithEvents _tbToolBar_Button2 As System.Windows.Forms.ToolStripButton
	Public WithEvents _tbToolBar_Button3 As System.Windows.Forms.ToolStripButton
	Public WithEvents _tbToolBar_Button4 As System.Windows.Forms.ToolStripButton
	Public WithEvents _tbToolBar_Button5 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents _tbToolBar_Button6 As System.Windows.Forms.ToolStripButton
	Public WithEvents _tbToolBar_Button7 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents _tbToolBar_Button8 As System.Windows.Forms.ToolStripButton
	Public WithEvents _tbToolBar_Button9 As System.Windows.Forms.ToolStripButton
	Public WithEvents tbToolBar As System.Windows.Forms.ToolStrip
	Public cd1Open As System.Windows.Forms.OpenFileDialog
	Public cd1Save As System.Windows.Forms.SaveFileDialog
	Public WithEvents rt1 As System.Windows.Forms.RichTextBox
	Public WithEvents ImageList1 As System.Windows.Forms.ImageList
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmScript))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.tbToolBar = New System.Windows.Forms.ToolStrip
		Me._tbToolBar_Button1 = New System.Windows.Forms.ToolStripButton
		Me._tbToolBar_Button2 = New System.Windows.Forms.ToolStripButton
		Me._tbToolBar_Button3 = New System.Windows.Forms.ToolStripButton
		Me._tbToolBar_Button4 = New System.Windows.Forms.ToolStripButton
		Me._tbToolBar_Button5 = New System.Windows.Forms.ToolStripSeparator
		Me._tbToolBar_Button6 = New System.Windows.Forms.ToolStripButton
		Me._tbToolBar_Button7 = New System.Windows.Forms.ToolStripSeparator
		Me._tbToolBar_Button8 = New System.Windows.Forms.ToolStripButton
		Me._tbToolBar_Button9 = New System.Windows.Forms.ToolStripButton
		Me.cd1Open = New System.Windows.Forms.OpenFileDialog
		Me.cd1Save = New System.Windows.Forms.SaveFileDialog
		Me.rt1 = New System.Windows.Forms.RichTextBox
		Me.ImageList1 = New System.Windows.Forms.ImageList
		Me.tbToolBar.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.Text = "Script File Editor"
		Me.ClientSize = New System.Drawing.Size(512, 272)
		Me.Location = New System.Drawing.Point(4, 24)
		Me.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Icon = CType(resources.GetObject("frmScript.Icon"), System.Drawing.Icon)
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
		Me.Name = "frmScript"
		Me.tbToolBar.ShowItemToolTips = True
		Me.tbToolBar.Dock = System.Windows.Forms.DockStyle.Top
		Me.tbToolBar.Size = New System.Drawing.Size(512, 43)
		Me.tbToolBar.Location = New System.Drawing.Point(0, 0)
		Me.tbToolBar.TabIndex = 1
		Me.tbToolBar.CanOverflow = False
		Me.tbToolBar.ImageList = ImageList1
		Me.tbToolBar.Name = "tbToolBar"
		Me._tbToolBar_Button1.Size = New System.Drawing.Size(43, 38)
		Me._tbToolBar_Button1.AutoSize = False
		Me._tbToolBar_Button1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._tbToolBar_Button1.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._tbToolBar_Button1.Text = "New"
		Me._tbToolBar_Button1.Name = "New"
		Me._tbToolBar_Button1.ToolTipText = "Start New File"
		Me._tbToolBar_Button1.ImageIndex = 0
		Me._tbToolBar_Button1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._tbToolBar_Button2.Size = New System.Drawing.Size(43, 38)
		Me._tbToolBar_Button2.AutoSize = False
		Me._tbToolBar_Button2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._tbToolBar_Button2.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._tbToolBar_Button2.Text = "Open"
		Me._tbToolBar_Button2.Name = "Open"
		Me._tbToolBar_Button2.ToolTipText = "Open File"
		Me._tbToolBar_Button2.ImageIndex = 1
		Me._tbToolBar_Button2.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._tbToolBar_Button3.Size = New System.Drawing.Size(43, 38)
		Me._tbToolBar_Button3.AutoSize = False
		Me._tbToolBar_Button3.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._tbToolBar_Button3.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._tbToolBar_Button3.Text = "Save"
		Me._tbToolBar_Button3.Name = "Save"
		Me._tbToolBar_Button3.ToolTipText = "Save File "
		Me._tbToolBar_Button3.ImageIndex = 2
		Me._tbToolBar_Button3.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._tbToolBar_Button4.Size = New System.Drawing.Size(43, 38)
		Me._tbToolBar_Button4.AutoSize = False
		Me._tbToolBar_Button4.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._tbToolBar_Button4.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._tbToolBar_Button4.Text = "Exit"
		Me._tbToolBar_Button4.Name = "Exit"
		Me._tbToolBar_Button4.ToolTipText = "Exit Script Editor"
		Me._tbToolBar_Button4.ImageIndex = 3
		Me._tbToolBar_Button4.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._tbToolBar_Button5.Size = New System.Drawing.Size(43, 38)
		Me._tbToolBar_Button5.AutoSize = False
		Me._tbToolBar_Button5.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._tbToolBar_Button5.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._tbToolBar_Button5.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._tbToolBar_Button6.Size = New System.Drawing.Size(43, 38)
		Me._tbToolBar_Button6.AutoSize = False
		Me._tbToolBar_Button6.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._tbToolBar_Button6.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._tbToolBar_Button6.Text = "Rec."
		Me._tbToolBar_Button6.Name = "Record"
		Me._tbToolBar_Button6.ToolTipText = "Start Script Recorder"
		Me._tbToolBar_Button6.ImageIndex = 4
		Me._tbToolBar_Button6.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._tbToolBar_Button7.Size = New System.Drawing.Size(43, 38)
		Me._tbToolBar_Button7.AutoSize = False
		Me._tbToolBar_Button7.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._tbToolBar_Button7.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._tbToolBar_Button7.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._tbToolBar_Button8.Size = New System.Drawing.Size(43, 38)
		Me._tbToolBar_Button8.AutoSize = False
		Me._tbToolBar_Button8.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._tbToolBar_Button8.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._tbToolBar_Button8.Text = "Run"
		Me._tbToolBar_Button8.Name = "Run"
		Me._tbToolBar_Button8.ToolTipText = "Execute Script File"
		Me._tbToolBar_Button8.ImageIndex = 6
		Me._tbToolBar_Button8.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._tbToolBar_Button9.Size = New System.Drawing.Size(43, 38)
		Me._tbToolBar_Button9.AutoSize = False
		Me._tbToolBar_Button9.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._tbToolBar_Button9.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._tbToolBar_Button9.Text = "Timer"
		Me._tbToolBar_Button9.Name = "Timer"
		Me._tbToolBar_Button9.ToolTipText = "Attach to Timer"
		Me._tbToolBar_Button9.ImageIndex = 7
		Me._tbToolBar_Button9.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me.cd1Open.DefaultExt = ".txt"
		Me.cd1Save.DefaultExt = ".txt"
		Me.cd1Open.Title = "Save Script File"
		Me.cd1Save.Title = "Save Script File"
		Me.rt1.Size = New System.Drawing.Size(257, 249)
		Me.rt1.Location = New System.Drawing.Point(0, 48)
		Me.rt1.TabIndex = 0
		Me.rt1.Enabled = True
		Me.rt1.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
		Me.rt1.RTF = resources.GetString("rt1.TextRTF")
		Me.rt1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.rt1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.rt1.Name = "rt1"
		Me.ImageList1.ImageSize = New System.Drawing.Size(16, 15)
		Me.ImageList1.TransparentColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
		Me.ImageList1.Images.SetKeyName(0, "")
		Me.ImageList1.Images.SetKeyName(1, "")
		Me.ImageList1.Images.SetKeyName(2, "")
		Me.ImageList1.Images.SetKeyName(3, "")
		Me.ImageList1.Images.SetKeyName(4, "")
		Me.ImageList1.Images.SetKeyName(5, "")
		Me.ImageList1.Images.SetKeyName(6, "")
		Me.ImageList1.Images.SetKeyName(7, "")
		Me.Controls.Add(tbToolBar)
		Me.Controls.Add(rt1)
		Me.tbToolBar.Items.Add(_tbToolBar_Button1)
		Me.tbToolBar.Items.Add(_tbToolBar_Button2)
		Me.tbToolBar.Items.Add(_tbToolBar_Button3)
		Me.tbToolBar.Items.Add(_tbToolBar_Button4)
		Me.tbToolBar.Items.Add(_tbToolBar_Button5)
		Me.tbToolBar.Items.Add(_tbToolBar_Button6)
		Me.tbToolBar.Items.Add(_tbToolBar_Button7)
		Me.tbToolBar.Items.Add(_tbToolBar_Button8)
		Me.tbToolBar.Items.Add(_tbToolBar_Button9)
		Me.tbToolBar.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class