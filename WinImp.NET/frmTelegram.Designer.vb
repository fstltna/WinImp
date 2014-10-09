<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmTelegram
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
	Public WithEvents imlIcons As System.Windows.Forms.ImageList
	Public WithEvents _tbToolBar_Button1 As System.Windows.Forms.ToolStripButton
	Public WithEvents _tbToolBar_Button2 As System.Windows.Forms.ToolStripButton
	Public WithEvents _tbToolBar_Button3 As System.Windows.Forms.ToolStripButton
	Public WithEvents _tbToolBar_Button4 As System.Windows.Forms.ToolStripButton
	Public WithEvents _tbToolBar_Button5 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents _tbToolBar_Button6 As System.Windows.Forms.ToolStripButton
	Public WithEvents _tbToolBar_Button7 As System.Windows.Forms.ToolStripButton
	Public WithEvents _tbToolBar_Button8 As System.Windows.Forms.ToolStripButton
	Public WithEvents _tbToolBar_Button9 As System.Windows.Forms.ToolStripButton
	Public WithEvents _tbToolBar_Button10 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents _tbToolBar_Button11 As System.Windows.Forms.ToolStripButton
	Public WithEvents _tbToolBar_Button12 As System.Windows.Forms.ToolStripButton
	Public WithEvents _tbToolBar_Button13 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents _tbToolBar_Button14 As System.Windows.Forms.ToolStripSeparator
	Public WithEvents tbToolBar As System.Windows.Forms.ToolStrip
	Public WithEvents tvTreeView As System.Windows.Forms.TreeView
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	Public WithEvents lbNation As System.Windows.Forms.ListBox
	Public WithEvents chkSave As System.Windows.Forms.CheckBox
	Public WithEvents chkInclude As System.Windows.Forms.CheckBox
	Public WithEvents txtBody As System.Windows.Forms.RichTextBox
	Public WithEvents picSplitter As System.Windows.Forms.PictureBox
	Public WithEvents _lblTitle_1 As System.Windows.Forms.Label
	Public WithEvents _lblTitle_0 As System.Windows.Forms.Label
	Public WithEvents picTitles As System.Windows.Forms.Panel
	Public dlgCommonDialogOpen As System.Windows.Forms.OpenFileDialog
	Public dlgCommonDialogSave As System.Windows.Forms.SaveFileDialog
	Public dlgCommonDialogFont As System.Windows.Forms.FontDialog
	Public dlgCommonDialogColor As System.Windows.Forms.ColorDialog
	Public dlgCommonDialogPrint As System.Windows.Forms.PrintDialog
	Public WithEvents _sbStatusBar_Panel1 As System.Windows.Forms.ToolStripStatusLabel
	Public WithEvents _sbStatusBar_Panel2 As System.Windows.Forms.ToolStripStatusLabel
	Public WithEvents _sbStatusBar_Panel3 As System.Windows.Forms.ToolStripStatusLabel
	Public WithEvents sbStatusBar As System.Windows.Forms.StatusStrip
	Public WithEvents optMotd As System.Windows.Forms.RadioButton
	Public WithEvents Option3 As System.Windows.Forms.RadioButton
	Public WithEvents Option2 As System.Windows.Forms.RadioButton
	Public WithEvents Option1 As System.Windows.Forms.RadioButton
	Public WithEvents frameSend As System.Windows.Forms.GroupBox
	Public WithEvents labNation As System.Windows.Forms.Label
	Public WithEvents imgSplitter As System.Windows.Forms.PictureBox
	Public WithEvents lblTitle As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmTelegram))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.imlIcons = New System.Windows.Forms.ImageList
		Me.tbToolBar = New System.Windows.Forms.ToolStrip
		Me._tbToolBar_Button1 = New System.Windows.Forms.ToolStripButton
		Me._tbToolBar_Button2 = New System.Windows.Forms.ToolStripButton
		Me._tbToolBar_Button3 = New System.Windows.Forms.ToolStripButton
		Me._tbToolBar_Button4 = New System.Windows.Forms.ToolStripButton
		Me._tbToolBar_Button5 = New System.Windows.Forms.ToolStripSeparator
		Me._tbToolBar_Button6 = New System.Windows.Forms.ToolStripButton
		Me._tbToolBar_Button7 = New System.Windows.Forms.ToolStripButton
		Me._tbToolBar_Button8 = New System.Windows.Forms.ToolStripButton
		Me._tbToolBar_Button9 = New System.Windows.Forms.ToolStripButton
		Me._tbToolBar_Button10 = New System.Windows.Forms.ToolStripSeparator
		Me._tbToolBar_Button11 = New System.Windows.Forms.ToolStripButton
		Me._tbToolBar_Button12 = New System.Windows.Forms.ToolStripButton
		Me._tbToolBar_Button13 = New System.Windows.Forms.ToolStripSeparator
		Me._tbToolBar_Button14 = New System.Windows.Forms.ToolStripSeparator
		Me.tvTreeView = New System.Windows.Forms.TreeView
		Me.Timer1 = New System.Windows.Forms.Timer(components)
		Me.lbNation = New System.Windows.Forms.ListBox
		Me.chkSave = New System.Windows.Forms.CheckBox
		Me.chkInclude = New System.Windows.Forms.CheckBox
		Me.txtBody = New System.Windows.Forms.RichTextBox
		Me.picSplitter = New System.Windows.Forms.PictureBox
		Me.picTitles = New System.Windows.Forms.Panel
		Me._lblTitle_1 = New System.Windows.Forms.Label
		Me._lblTitle_0 = New System.Windows.Forms.Label
		Me.dlgCommonDialogOpen = New System.Windows.Forms.OpenFileDialog
		Me.dlgCommonDialogSave = New System.Windows.Forms.SaveFileDialog
		Me.dlgCommonDialogFont = New System.Windows.Forms.FontDialog
		Me.dlgCommonDialogColor = New System.Windows.Forms.ColorDialog
		Me.dlgCommonDialogPrint = New System.Windows.Forms.PrintDialog
		Me.sbStatusBar = New System.Windows.Forms.StatusStrip
		Me._sbStatusBar_Panel1 = New System.Windows.Forms.ToolStripStatusLabel
		Me._sbStatusBar_Panel2 = New System.Windows.Forms.ToolStripStatusLabel
		Me._sbStatusBar_Panel3 = New System.Windows.Forms.ToolStripStatusLabel
		Me.frameSend = New System.Windows.Forms.GroupBox
		Me.optMotd = New System.Windows.Forms.RadioButton
		Me.Option3 = New System.Windows.Forms.RadioButton
		Me.Option2 = New System.Windows.Forms.RadioButton
		Me.Option1 = New System.Windows.Forms.RadioButton
		Me.labNation = New System.Windows.Forms.Label
		Me.imgSplitter = New System.Windows.Forms.PictureBox
		Me.lblTitle = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.tbToolBar.SuspendLayout()
		Me.picTitles.SuspendLayout()
		Me.sbStatusBar.SuspendLayout()
		Me.frameSend.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.lblTitle, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.Text = "Telegrams"
		Me.ClientSize = New System.Drawing.Size(573, 441)
		Me.Location = New System.Drawing.Point(11, 30)
		Me.Icon = CType(resources.GetObject("frmTelegram.Icon"), System.Drawing.Icon)
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
		Me.Name = "frmTelegram"
		Me.imlIcons.ImageSize = New System.Drawing.Size(32, 32)
		Me.imlIcons.TransparentColor = System.Drawing.Color.FromARGB(192, 192, 192)
		Me.imlIcons.ImageStream = CType(resources.GetObject("imlIcons.ImageStream"), System.Windows.Forms.ImageListStreamer)
		Me.imlIcons.Images.SetKeyName(0, "")
		Me.imlIcons.Images.SetKeyName(1, "")
		Me.imlIcons.Images.SetKeyName(2, "")
		Me.imlIcons.Images.SetKeyName(3, "")
		Me.imlIcons.Images.SetKeyName(4, "")
		Me.imlIcons.Images.SetKeyName(5, "")
		Me.imlIcons.Images.SetKeyName(6, "")
		Me.imlIcons.Images.SetKeyName(7, "")
		Me.imlIcons.Images.SetKeyName(8, "")
		Me.imlIcons.Images.SetKeyName(9, "")
		Me.tbToolBar.ShowItemToolTips = True
		Me.tbToolBar.Dock = System.Windows.Forms.DockStyle.Top
		Me.tbToolBar.Size = New System.Drawing.Size(573, 64)
		Me.tbToolBar.Location = New System.Drawing.Point(0, 0)
		Me.tbToolBar.TabIndex = 0
		Me.tbToolBar.ImageList = imlIcons
		Me.tbToolBar.Name = "tbToolBar"
		Me._tbToolBar_Button1.Size = New System.Drawing.Size(68, 59)
		Me._tbToolBar_Button1.AutoSize = False
		Me._tbToolBar_Button1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._tbToolBar_Button1.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._tbToolBar_Button1.Text = "First"
		Me._tbToolBar_Button1.Name = "First"
		Me._tbToolBar_Button1.ToolTipText = "Oldest"
		Me._tbToolBar_Button1.ImageIndex = 0
		Me._tbToolBar_Button1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._tbToolBar_Button2.Size = New System.Drawing.Size(68, 59)
		Me._tbToolBar_Button2.AutoSize = False
		Me._tbToolBar_Button2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._tbToolBar_Button2.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._tbToolBar_Button2.Text = "Previous"
		Me._tbToolBar_Button2.Name = "Previous"
		Me._tbToolBar_Button2.ToolTipText = "Previous"
		Me._tbToolBar_Button2.ImageIndex = 1
		Me._tbToolBar_Button2.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._tbToolBar_Button3.Size = New System.Drawing.Size(68, 59)
		Me._tbToolBar_Button3.AutoSize = False
		Me._tbToolBar_Button3.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._tbToolBar_Button3.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._tbToolBar_Button3.Text = "Next"
		Me._tbToolBar_Button3.Name = "Next"
		Me._tbToolBar_Button3.ToolTipText = "Next"
		Me._tbToolBar_Button3.ImageIndex = 2
		Me._tbToolBar_Button3.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._tbToolBar_Button4.Size = New System.Drawing.Size(68, 59)
		Me._tbToolBar_Button4.AutoSize = False
		Me._tbToolBar_Button4.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._tbToolBar_Button4.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._tbToolBar_Button4.Text = "Last"
		Me._tbToolBar_Button4.Name = "Last"
		Me._tbToolBar_Button4.ToolTipText = "Newest"
		Me._tbToolBar_Button4.ImageIndex = 3
		Me._tbToolBar_Button4.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._tbToolBar_Button5.Size = New System.Drawing.Size(68, 59)
		Me._tbToolBar_Button5.AutoSize = False
		Me._tbToolBar_Button5.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._tbToolBar_Button5.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._tbToolBar_Button5.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._tbToolBar_Button6.Size = New System.Drawing.Size(68, 59)
		Me._tbToolBar_Button6.AutoSize = False
		Me._tbToolBar_Button6.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._tbToolBar_Button6.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._tbToolBar_Button6.Text = "New"
		Me._tbToolBar_Button6.Name = "New"
		Me._tbToolBar_Button6.ToolTipText = "New Telegram"
		Me._tbToolBar_Button6.ImageIndex = 4
		Me._tbToolBar_Button6.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._tbToolBar_Button7.Size = New System.Drawing.Size(68, 59)
		Me._tbToolBar_Button7.AutoSize = False
		Me._tbToolBar_Button7.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._tbToolBar_Button7.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._tbToolBar_Button7.Text = "Reply"
		Me._tbToolBar_Button7.Name = "Reply"
		Me._tbToolBar_Button7.ToolTipText = "Reply"
		Me._tbToolBar_Button7.ImageIndex = 5
		Me._tbToolBar_Button7.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._tbToolBar_Button8.Size = New System.Drawing.Size(68, 59)
		Me._tbToolBar_Button8.AutoSize = False
		Me._tbToolBar_Button8.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._tbToolBar_Button8.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._tbToolBar_Button8.Text = "Forward"
		Me._tbToolBar_Button8.Name = "Forward"
		Me._tbToolBar_Button8.ToolTipText = "Forward"
		Me._tbToolBar_Button8.ImageIndex = 6
		Me._tbToolBar_Button8.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._tbToolBar_Button9.Size = New System.Drawing.Size(68, 59)
		Me._tbToolBar_Button9.AutoSize = False
		Me._tbToolBar_Button9.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._tbToolBar_Button9.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._tbToolBar_Button9.Text = "Import"
		Me._tbToolBar_Button9.Name = "Import"
		Me._tbToolBar_Button9.ToolTipText = "Import Intelligence"
		Me._tbToolBar_Button9.ImageIndex = 7
		Me._tbToolBar_Button9.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._tbToolBar_Button10.Size = New System.Drawing.Size(68, 59)
		Me._tbToolBar_Button10.AutoSize = False
		Me._tbToolBar_Button10.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._tbToolBar_Button10.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._tbToolBar_Button10.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._tbToolBar_Button11.Size = New System.Drawing.Size(68, 59)
		Me._tbToolBar_Button11.AutoSize = False
		Me._tbToolBar_Button11.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._tbToolBar_Button11.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._tbToolBar_Button11.Text = "Delete"
		Me._tbToolBar_Button11.Name = "Delete"
		Me._tbToolBar_Button11.ToolTipText = "Delete"
		Me._tbToolBar_Button11.ImageIndex = 8
		Me._tbToolBar_Button11.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._tbToolBar_Button12.Size = New System.Drawing.Size(68, 59)
		Me._tbToolBar_Button12.AutoSize = False
		Me._tbToolBar_Button12.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._tbToolBar_Button12.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._tbToolBar_Button12.Enabled = 0
		Me._tbToolBar_Button12.Text = "Send"
		Me._tbToolBar_Button12.Name = "Send"
		Me._tbToolBar_Button12.ToolTipText = "Send Telegram"
		Me._tbToolBar_Button12.ImageIndex = 9
		Me._tbToolBar_Button12.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._tbToolBar_Button13.Size = New System.Drawing.Size(68, 59)
		Me._tbToolBar_Button13.AutoSize = False
		Me._tbToolBar_Button13.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._tbToolBar_Button13.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._tbToolBar_Button13.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me._tbToolBar_Button14.Size = New System.Drawing.Size(68, 59)
		Me._tbToolBar_Button14.AutoSize = False
		Me._tbToolBar_Button14.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.ImageAndText
		Me._tbToolBar_Button14.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
		Me._tbToolBar_Button14.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
		Me.tvTreeView.LabelEdit = True
		Me.tvTreeView.CausesValidation = True
		Me.tvTreeView.Size = New System.Drawing.Size(135, 312)
		Me.tvTreeView.Location = New System.Drawing.Point(0, 72)
		Me.tvTreeView.TabIndex = 4
		Me.tvTreeView.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.tvTreeView.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.tvTreeView.Name = "tvTreeView"
		Me.Timer1.Enabled = False
		Me.Timer1.Interval = 1
		Me.lbNation.Size = New System.Drawing.Size(169, 124)
		Me.lbNation.Location = New System.Drawing.Point(376, 56)
		Me.lbNation.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
		Me.lbNation.Sorted = True
		Me.lbNation.TabIndex = 13
		Me.lbNation.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lbNation.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lbNation.BackColor = System.Drawing.SystemColors.Window
		Me.lbNation.CausesValidation = True
		Me.lbNation.Enabled = True
		Me.lbNation.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lbNation.IntegralHeight = True
		Me.lbNation.Cursor = System.Windows.Forms.Cursors.Default
		Me.lbNation.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lbNation.TabStop = True
		Me.lbNation.Visible = True
		Me.lbNation.MultiColumn = False
		Me.lbNation.Name = "lbNation"
		Me.chkSave.Text = "Save a copy of sent mail"
		Me.chkSave.Size = New System.Drawing.Size(169, 33)
		Me.chkSave.Location = New System.Drawing.Point(376, 264)
		Me.chkSave.TabIndex = 9
		Me.chkSave.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkSave.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkSave.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkSave.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkSave.BackColor = System.Drawing.SystemColors.Control
		Me.chkSave.CausesValidation = True
		Me.chkSave.Enabled = True
		Me.chkSave.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkSave.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkSave.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkSave.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkSave.TabStop = True
		Me.chkSave.Visible = True
		Me.chkSave.Name = "chkSave"
		Me.chkInclude.Text = "Include text from received telegram in your reply"
		Me.chkInclude.Size = New System.Drawing.Size(169, 41)
		Me.chkInclude.Location = New System.Drawing.Point(376, 224)
		Me.chkInclude.TabIndex = 8
		Me.chkInclude.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkInclude.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkInclude.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkInclude.BackColor = System.Drawing.SystemColors.Control
		Me.chkInclude.CausesValidation = True
		Me.chkInclude.Enabled = True
		Me.chkInclude.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkInclude.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkInclude.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkInclude.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkInclude.TabStop = True
		Me.chkInclude.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkInclude.Visible = True
		Me.chkInclude.Name = "chkInclude"
		Me.txtBody.Size = New System.Drawing.Size(209, 313)
		Me.txtBody.Location = New System.Drawing.Point(136, 72)
		Me.txtBody.TabIndex = 6
		Me.txtBody.Enabled = True
		Me.txtBody.ReadOnly = True
		Me.txtBody.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
		Me.txtBody.RTF = resources.GetString("txtBody.TextRTF")
		Me.txtBody.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtBody.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtBody.Name = "txtBody"
		Me.picSplitter.BackColor = System.Drawing.Color.FromARGB(128, 128, 128)
		Me.picSplitter.Size = New System.Drawing.Size(5, 320)
		Me.picSplitter.Location = New System.Drawing.Point(360, 48)
		Me.picSplitter.TabIndex = 5
		Me.picSplitter.Visible = False
		Me.picSplitter.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picSplitter.Dock = System.Windows.Forms.DockStyle.None
		Me.picSplitter.CausesValidation = True
		Me.picSplitter.Enabled = True
		Me.picSplitter.ForeColor = System.Drawing.SystemColors.ControlText
		Me.picSplitter.Cursor = System.Windows.Forms.Cursors.Default
		Me.picSplitter.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picSplitter.TabStop = True
		Me.picSplitter.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.picSplitter.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picSplitter.Name = "picSplitter"
		Me.picTitles.Dock = System.Windows.Forms.DockStyle.Top
		Me.picTitles.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picTitles.Size = New System.Drawing.Size(573, 28)
		Me.picTitles.Location = New System.Drawing.Point(0, 64)
		Me.picTitles.TabIndex = 1
		Me.picTitles.TabStop = False
		Me.picTitles.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picTitles.BackColor = System.Drawing.SystemColors.Control
		Me.picTitles.CausesValidation = True
		Me.picTitles.Enabled = True
		Me.picTitles.Cursor = System.Windows.Forms.Cursors.Default
		Me.picTitles.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picTitles.Visible = True
		Me.picTitles.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picTitles.Name = "picTitles"
		Me._lblTitle_1.Text = " ListView:"
		Me._lblTitle_1.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblTitle_1.Size = New System.Drawing.Size(214, 28)
		Me._lblTitle_1.Location = New System.Drawing.Point(136, 0)
		Me._lblTitle_1.TabIndex = 3
		Me._lblTitle_1.Tag = " ListView:"
		Me._lblTitle_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblTitle_1.BackColor = System.Drawing.SystemColors.Control
		Me._lblTitle_1.Enabled = True
		Me._lblTitle_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblTitle_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblTitle_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblTitle_1.UseMnemonic = True
		Me._lblTitle_1.Visible = True
		Me._lblTitle_1.AutoSize = False
		Me._lblTitle_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._lblTitle_1.Name = "_lblTitle_1"
		Me._lblTitle_0.Text = "Received Telegrams"
		Me._lblTitle_0.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblTitle_0.Size = New System.Drawing.Size(134, 26)
		Me._lblTitle_0.Location = New System.Drawing.Point(0, 0)
		Me._lblTitle_0.TabIndex = 2
		Me._lblTitle_0.Tag = " TreeView:"
		Me._lblTitle_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblTitle_0.BackColor = System.Drawing.SystemColors.Control
		Me._lblTitle_0.Enabled = True
		Me._lblTitle_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblTitle_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblTitle_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblTitle_0.UseMnemonic = True
		Me._lblTitle_0.Visible = True
		Me._lblTitle_0.AutoSize = False
		Me._lblTitle_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._lblTitle_0.Name = "_lblTitle_0"
		Me.sbStatusBar.Dock = System.Windows.Forms.DockStyle.Bottom
		Me.sbStatusBar.Size = New System.Drawing.Size(573, 18)
		Me.sbStatusBar.Location = New System.Drawing.Point(0, 423)
		Me.sbStatusBar.TabIndex = 7
		Me.sbStatusBar.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.sbStatusBar.Name = "sbStatusBar"
		Me._sbStatusBar_Panel1.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter
		Me._sbStatusBar_Panel1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
		Me._sbStatusBar_Panel1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._sbStatusBar_Panel1.Size = New System.Drawing.Size(280, 18)
		Me._sbStatusBar_Panel1.Text = "xx Telegrams"
		Me._sbStatusBar_Panel1.Spring = True
		Me._sbStatusBar_Panel1.AutoSize = True
		Me._sbStatusBar_Panel1.BorderSides = CType(System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom, System.Windows.Forms.ToolStripStatusLabelBorderSides)
		Me._sbStatusBar_Panel1.Margin = New System.Windows.Forms.Padding(0)
		Me._sbStatusBar_Panel1.AutoSize = False
		Me._sbStatusBar_Panel2.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter
		Me._sbStatusBar_Panel2.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
		Me._sbStatusBar_Panel2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._sbStatusBar_Panel2.Size = New System.Drawing.Size(135, 18)
		Me._sbStatusBar_Panel2.AutoSize = True
		Me._sbStatusBar_Panel2.BorderSides = CType(System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom, System.Windows.Forms.ToolStripStatusLabelBorderSides)
		Me._sbStatusBar_Panel2.Margin = New System.Windows.Forms.Padding(0)
		Me._sbStatusBar_Panel2.AutoSize = False
		Me._sbStatusBar_Panel3.BorderStyle = System.Windows.Forms.Border3DStyle.SunkenOuter
		Me._sbStatusBar_Panel3.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
		Me._sbStatusBar_Panel3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._sbStatusBar_Panel3.Size = New System.Drawing.Size(135, 18)
		Me._sbStatusBar_Panel3.AutoSize = True
		Me._sbStatusBar_Panel3.BorderSides = CType(System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom, System.Windows.Forms.ToolStripStatusLabelBorderSides)
		Me._sbStatusBar_Panel3.Margin = New System.Windows.Forms.Padding(0)
		Me._sbStatusBar_Panel3.AutoSize = False
		Me.frameSend.Text = "What to Send"
		Me.frameSend.Size = New System.Drawing.Size(145, 121)
		Me.frameSend.Location = New System.Drawing.Point(376, 296)
		Me.frameSend.TabIndex = 10
		Me.frameSend.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frameSend.BackColor = System.Drawing.SystemColors.Control
		Me.frameSend.Enabled = True
		Me.frameSend.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frameSend.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frameSend.Visible = True
		Me.frameSend.Padding = New System.Windows.Forms.Padding(0)
		Me.frameSend.Name = "frameSend"
		Me.optMotd.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optMotd.Text = "Motd"
		Me.optMotd.Size = New System.Drawing.Size(49, 17)
		Me.optMotd.Location = New System.Drawing.Point(24, 96)
		Me.optMotd.TabIndex = 16
		Me.optMotd.Visible = False
		Me.optMotd.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optMotd.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optMotd.BackColor = System.Drawing.SystemColors.Control
		Me.optMotd.CausesValidation = True
		Me.optMotd.Enabled = True
		Me.optMotd.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optMotd.Cursor = System.Windows.Forms.Cursors.Default
		Me.optMotd.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optMotd.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optMotd.TabStop = True
		Me.optMotd.Checked = False
		Me.optMotd.Name = "optMotd"
		Me.Option3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option3.Text = "Flash"
		Me.Option3.Size = New System.Drawing.Size(65, 17)
		Me.Option3.Location = New System.Drawing.Point(24, 72)
		Me.Option3.TabIndex = 15
		Me.Option3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Option3.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option3.BackColor = System.Drawing.SystemColors.Control
		Me.Option3.CausesValidation = True
		Me.Option3.Enabled = True
		Me.Option3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Option3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Option3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Option3.Appearance = System.Windows.Forms.Appearance.Normal
		Me.Option3.TabStop = True
		Me.Option3.Checked = False
		Me.Option3.Visible = True
		Me.Option3.Name = "Option3"
		Me.Option2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option2.Text = "Announcement"
		Me.Option2.Size = New System.Drawing.Size(97, 17)
		Me.Option2.Location = New System.Drawing.Point(24, 48)
		Me.Option2.TabIndex = 12
		Me.Option2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Option2.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option2.BackColor = System.Drawing.SystemColors.Control
		Me.Option2.CausesValidation = True
		Me.Option2.Enabled = True
		Me.Option2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Option2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Option2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Option2.Appearance = System.Windows.Forms.Appearance.Normal
		Me.Option2.TabStop = True
		Me.Option2.Checked = False
		Me.Option2.Visible = True
		Me.Option2.Name = "Option2"
		Me.Option1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option1.Text = "Telegram"
		Me.Option1.Size = New System.Drawing.Size(97, 17)
		Me.Option1.Location = New System.Drawing.Point(24, 24)
		Me.Option1.TabIndex = 11
		Me.Option1.Checked = True
		Me.Option1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Option1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option1.BackColor = System.Drawing.SystemColors.Control
		Me.Option1.CausesValidation = True
		Me.Option1.Enabled = True
		Me.Option1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Option1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Option1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Option1.Appearance = System.Windows.Forms.Appearance.Normal
		Me.Option1.TabStop = True
		Me.Option1.Visible = True
		Me.Option1.Name = "Option1"
		Me.labNation.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.labNation.Text = "<< ANNOUNCEMENT >>"
		Me.labNation.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.labNation.Size = New System.Drawing.Size(169, 25)
		Me.labNation.Location = New System.Drawing.Point(376, 192)
		Me.labNation.TabIndex = 14
		Me.labNation.BackColor = System.Drawing.SystemColors.Control
		Me.labNation.Enabled = True
		Me.labNation.ForeColor = System.Drawing.SystemColors.ControlText
		Me.labNation.Cursor = System.Windows.Forms.Cursors.Default
		Me.labNation.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.labNation.UseMnemonic = True
		Me.labNation.Visible = True
		Me.labNation.AutoSize = False
		Me.labNation.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.labNation.Name = "labNation"
		Me.imgSplitter.Size = New System.Drawing.Size(10, 320)
		Me.imgSplitter.Location = New System.Drawing.Point(131, 47)
		Me.imgSplitter.Cursor = System.Windows.Forms.Cursors.SizeWE
		Me.imgSplitter.Enabled = True
		Me.imgSplitter.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.imgSplitter.Visible = True
		Me.imgSplitter.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.imgSplitter.Name = "imgSplitter"
		Me.Controls.Add(tbToolBar)
		Me.Controls.Add(tvTreeView)
		Me.Controls.Add(lbNation)
		Me.Controls.Add(chkSave)
		Me.Controls.Add(chkInclude)
		Me.Controls.Add(txtBody)
		Me.Controls.Add(picSplitter)
		Me.Controls.Add(picTitles)
		Me.Controls.Add(sbStatusBar)
		Me.Controls.Add(frameSend)
		Me.Controls.Add(labNation)
		Me.Controls.Add(imgSplitter)
		Me.tbToolBar.Items.Add(_tbToolBar_Button1)
		Me.tbToolBar.Items.Add(_tbToolBar_Button2)
		Me.tbToolBar.Items.Add(_tbToolBar_Button3)
		Me.tbToolBar.Items.Add(_tbToolBar_Button4)
		Me.tbToolBar.Items.Add(_tbToolBar_Button5)
		Me.tbToolBar.Items.Add(_tbToolBar_Button6)
		Me.tbToolBar.Items.Add(_tbToolBar_Button7)
		Me.tbToolBar.Items.Add(_tbToolBar_Button8)
		Me.tbToolBar.Items.Add(_tbToolBar_Button9)
		Me.tbToolBar.Items.Add(_tbToolBar_Button10)
		Me.tbToolBar.Items.Add(_tbToolBar_Button11)
		Me.tbToolBar.Items.Add(_tbToolBar_Button12)
		Me.tbToolBar.Items.Add(_tbToolBar_Button13)
		Me.tbToolBar.Items.Add(_tbToolBar_Button14)
		Me.picTitles.Controls.Add(_lblTitle_1)
		Me.picTitles.Controls.Add(_lblTitle_0)
		Me.sbStatusBar.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me._sbStatusBar_Panel1})
		Me.sbStatusBar.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me._sbStatusBar_Panel2})
		Me.sbStatusBar.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me._sbStatusBar_Panel3})
		Me.frameSend.Controls.Add(optMotd)
		Me.frameSend.Controls.Add(Option3)
		Me.frameSend.Controls.Add(Option2)
		Me.frameSend.Controls.Add(Option1)
		Me.lblTitle.SetIndex(_lblTitle_1, CType(1, Short))
		Me.lblTitle.SetIndex(_lblTitle_0, CType(0, Short))
		CType(Me.lblTitle, System.ComponentModel.ISupportInitialize).EndInit()
		Me.tbToolBar.ResumeLayout(False)
		Me.picTitles.ResumeLayout(False)
		Me.sbStatusBar.ResumeLayout(False)
		Me.frameSend.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class