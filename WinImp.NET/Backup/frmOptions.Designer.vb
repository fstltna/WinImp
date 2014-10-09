<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmOptions
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
	Public WithEvents txtPathLength As System.Windows.Forms.TextBox
	Public WithEvents frmPathLength As System.Windows.Forms.GroupBox
	Public WithEvents optFriendlyAlly As System.Windows.Forms.RadioButton
	Public WithEvents optFriendlyNeutral As System.Windows.Forms.RadioButton
	Public WithEvents frmFriendly As System.Windows.Forms.GroupBox
	Public WithEvents txtDefenseScaling As System.Windows.Forms.TextBox
	Public WithEvents lblDefenseScaling As System.Windows.Forms.Label
	Public WithEvents frmAutoAttack As System.Windows.Forms.GroupBox
	Public WithEvents chkDumpHeaders As System.Windows.Forms.CheckBox
	Public WithEvents chkCapital As System.Windows.Forms.CheckBox
	Public WithEvents chkLocalHelp As System.Windows.Forms.CheckBox
	Public WithEvents chkSound As System.Windows.Forms.CheckBox
	Public WithEvents chkNoIron As System.Windows.Forms.CheckBox
	Public WithEvents chkRetreat As System.Windows.Forms.CheckBox
	Public WithEvents chkMaxProd As System.Windows.Forms.CheckBox
	Public WithEvents fraGame As System.Windows.Forms.GroupBox
	Public WithEvents picPlaying As System.Windows.Forms.Panel
	Public WithEvents chkShortName As System.Windows.Forms.CheckBox
	Public WithEvents chkClear As System.Windows.Forms.CheckBox
	Public WithEvents chkToolbar As System.Windows.Forms.CheckBox
	Public WithEvents chkStatusBar As System.Windows.Forms.CheckBox
	Public WithEvents chkWarship As System.Windows.Forms.CheckBox
	Public WithEvents fraView As System.Windows.Forms.GroupBox
	Public WithEvents chkFlashLogging As System.Windows.Forms.CheckBox
	Public WithEvents chkFlashPlayerColors As System.Windows.Forms.CheckBox
	Public WithEvents chkFlashChat As System.Windows.Forms.CheckBox
	Public WithEvents frmFlash As System.Windows.Forms.GroupBox
	Public WithEvents picBorderColor As System.Windows.Forms.PictureBox
	Public WithEvents cmdImage As System.Windows.Forms.Button
	Public WithEvents txtImageName As System.Windows.Forms.TextBox
	Public WithEvents lblBorderColor As System.Windows.Forms.Label
	Public WithEvents frmImage As System.Windows.Forms.GroupBox
	Public WithEvents picView As System.Windows.Forms.Panel
	Public WithEvents cmbPlayer As System.Windows.Forms.ComboBox
	Public WithEvents fraPlayer As System.Windows.Forms.GroupBox
	Public WithEvents chkFullDumpwithoutSea As System.Windows.Forms.CheckBox
	Public WithEvents frmDeity As System.Windows.Forms.GroupBox
	Public WithEvents chkSuppressUnitDamageReports As System.Windows.Forms.CheckBox
	Public WithEvents chkSuppressTelegramNotification As System.Windows.Forms.CheckBox
	Public WithEvents chkSuppressUpdateCommands As System.Windows.Forms.CheckBox
	Public WithEvents chkSuppressOilContentAtSea As System.Windows.Forms.CheckBox
	Public WithEvents chkCommodityRatio As System.Windows.Forms.CheckBox
	Public WithEvents chkRead As System.Windows.Forms.CheckBox
	Public WithEvents chkSuppressBmap As System.Windows.Forms.CheckBox
	Public WithEvents chkSuppressStrength As System.Windows.Forms.CheckBox
	Public WithEvents chkUpdate As System.Windows.Forms.CheckBox
	Public WithEvents fraUpdate As System.Windows.Forms.GroupBox
	Public WithEvents picUpdate As System.Windows.Forms.Panel
	Public WithEvents chkTTSFlashes As System.Windows.Forms.CheckBox
	Public WithEvents chkTTSBulletins As System.Windows.Forms.CheckBox
	Public WithEvents chkTTSTelegrams As System.Windows.Forms.CheckBox
	Public WithEvents chkTTSReportWindow As System.Windows.Forms.CheckBox
	Public WithEvents frmTTS As System.Windows.Forms.GroupBox
	Public WithEvents picTTS As System.Windows.Forms.Panel
	Public WithEvents _optThemeGames_7 As System.Windows.Forms.RadioButton
	Public WithEvents _optThemeGames_6 As System.Windows.Forms.RadioButton
	Public WithEvents _optThemeGames_5 As System.Windows.Forms.RadioButton
	Public WithEvents _optThemeGames_4 As System.Windows.Forms.RadioButton
	Public WithEvents _optThemeGames_3 As System.Windows.Forms.RadioButton
	Public WithEvents _optThemeGames_2 As System.Windows.Forms.RadioButton
	Public WithEvents _optThemeGames_1 As System.Windows.Forms.RadioButton
	Public WithEvents _optThemeGames_0 As System.Windows.Forms.RadioButton
	Public WithEvents fraThemeGame As System.Windows.Forms.GroupBox
	Public WithEvents picThemeGame As System.Windows.Forms.Panel
	Public WithEvents optCustomScript As System.Windows.Forms.RadioButton
	Public WithEvents optJumpPoint As System.Windows.Forms.RadioButton
	Public WithEvents cmbJumpPoint As System.Windows.Forms.ComboBox
	Public WithEvents cmdScriptImageBrowse As System.Windows.Forms.Button
	Public WithEvents txtScriptImageFileName As System.Windows.Forms.TextBox
	Public WithEvents chkScriptReport As System.Windows.Forms.CheckBox
	Public WithEvents txtScriptFileName As System.Windows.Forms.TextBox
	Public WithEvents cmdScriptBrowse As System.Windows.Forms.Button
	Public WithEvents lblScriptImage As System.Windows.Forms.Label
	Public WithEvents lblScriptFileName As System.Windows.Forms.Label
	Public WithEvents fraScriptButton As System.Windows.Forms.GroupBox
	Public WithEvents cmbScriptButton As System.Windows.Forms.ComboBox
	Public WithEvents picScriptButton As System.Windows.Forms.Panel
	Public WithEvents txtItemDistributionPanelLabel As System.Windows.Forms.TextBox
	Public WithEvents txtItemSectorPanelLabel As System.Windows.Forms.TextBox
	Public WithEvents txtItemFormLabel As System.Windows.Forms.TextBox
	Public WithEvents lblItemDistributionPanelLabel As System.Windows.Forms.Label
	Public WithEvents lblItemSectorPanelLabel As System.Windows.Forms.Label
	Public WithEvents lblItemFormLabel As System.Windows.Forms.Label
	Public WithEvents frmItemLabels As System.Windows.Forms.GroupBox
	Public WithEvents cbItem As System.Windows.Forms.ComboBox
	Public WithEvents txtItemConditionalName As System.Windows.Forms.TextBox
	Public WithEvents txtItemIntelligenceNames As System.Windows.Forms.TextBox
	Public WithEvents txtItemDBName As System.Windows.Forms.TextBox
	Public WithEvents txtItemFormName As System.Windows.Forms.TextBox
	Public WithEvents lblItemConditionalName As System.Windows.Forms.Label
	Public WithEvents lblItemIntelligenceNames As System.Windows.Forms.Label
	Public WithEvents lblItemDBName As System.Windows.Forms.Label
	Public WithEvents lblItemFormName As System.Windows.Forms.Label
	Public WithEvents frmItem As System.Windows.Forms.GroupBox
	Public WithEvents picItem As System.Windows.Forms.Panel
	Public WithEvents _cmbTelegramSound_2 As System.Windows.Forms.ComboBox
	Public WithEvents _txtTelegram_2 As System.Windows.Forms.TextBox
	Public WithEvents _chkTelegram_2 As System.Windows.Forms.CheckBox
	Public WithEvents _lblTelegram_2 As System.Windows.Forms.Label
	Public WithEvents _fraTelegram_2 As System.Windows.Forms.GroupBox
	Public WithEvents _cmbTelegramSound_1 As System.Windows.Forms.ComboBox
	Public WithEvents _txtTelegram_1 As System.Windows.Forms.TextBox
	Public WithEvents Text1 As System.Windows.Forms.TextBox
	Public WithEvents _chkTelegram_1 As System.Windows.Forms.CheckBox
	Public WithEvents _lblTelegram_1 As System.Windows.Forms.Label
	Public WithEvents _fraTelegram_1 As System.Windows.Forms.GroupBox
	Public WithEvents cmbTelegram As System.Windows.Forms.ComboBox
	Public WithEvents _cmbTelegramSound_0 As System.Windows.Forms.ComboBox
	Public WithEvents _txtTelegram_0 As System.Windows.Forms.TextBox
	Public WithEvents _chkTelegram_0 As System.Windows.Forms.CheckBox
	Public WithEvents _lblTelegram_0 As System.Windows.Forms.Label
	Public WithEvents _fraTelegram_0 As System.Windows.Forms.GroupBox
	Public WithEvents picTelegram As System.Windows.Forms.Panel
	Public WithEvents lstToolbar As System.Windows.Forms.CheckedListBox
	Public WithEvents fraToolbar As System.Windows.Forms.GroupBox
	Public WithEvents picToolbar As System.Windows.Forms.Panel
	Public WithEvents tbsOptions As AxMSComctlLib.AxTabStrip
	Public WithEvents cmdApply As System.Windows.Forms.Button
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents chkTelegram As Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray
	Public WithEvents cmbTelegramSound As Microsoft.VisualBasic.Compatibility.VB6.ComboBoxArray
	Public WithEvents fraTelegram As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
	Public WithEvents lblTelegram As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents optThemeGames As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
	Public WithEvents txtTelegram As Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmOptions))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.picPlaying = New System.Windows.Forms.Panel
		Me.frmPathLength = New System.Windows.Forms.GroupBox
		Me.txtPathLength = New System.Windows.Forms.TextBox
		Me.frmFriendly = New System.Windows.Forms.GroupBox
		Me.optFriendlyAlly = New System.Windows.Forms.RadioButton
		Me.optFriendlyNeutral = New System.Windows.Forms.RadioButton
		Me.frmAutoAttack = New System.Windows.Forms.GroupBox
		Me.txtDefenseScaling = New System.Windows.Forms.TextBox
		Me.lblDefenseScaling = New System.Windows.Forms.Label
		Me.fraGame = New System.Windows.Forms.GroupBox
		Me.chkDumpHeaders = New System.Windows.Forms.CheckBox
		Me.chkCapital = New System.Windows.Forms.CheckBox
		Me.chkLocalHelp = New System.Windows.Forms.CheckBox
		Me.chkSound = New System.Windows.Forms.CheckBox
		Me.chkNoIron = New System.Windows.Forms.CheckBox
		Me.chkRetreat = New System.Windows.Forms.CheckBox
		Me.chkMaxProd = New System.Windows.Forms.CheckBox
		Me.picView = New System.Windows.Forms.Panel
		Me.fraView = New System.Windows.Forms.GroupBox
		Me.chkShortName = New System.Windows.Forms.CheckBox
		Me.chkClear = New System.Windows.Forms.CheckBox
		Me.chkToolbar = New System.Windows.Forms.CheckBox
		Me.chkStatusBar = New System.Windows.Forms.CheckBox
		Me.chkWarship = New System.Windows.Forms.CheckBox
		Me.frmFlash = New System.Windows.Forms.GroupBox
		Me.chkFlashLogging = New System.Windows.Forms.CheckBox
		Me.chkFlashPlayerColors = New System.Windows.Forms.CheckBox
		Me.chkFlashChat = New System.Windows.Forms.CheckBox
		Me.frmImage = New System.Windows.Forms.GroupBox
		Me.picBorderColor = New System.Windows.Forms.PictureBox
		Me.cmdImage = New System.Windows.Forms.Button
		Me.txtImageName = New System.Windows.Forms.TextBox
		Me.lblBorderColor = New System.Windows.Forms.Label
		Me.picUpdate = New System.Windows.Forms.Panel
		Me.fraPlayer = New System.Windows.Forms.GroupBox
		Me.cmbPlayer = New System.Windows.Forms.ComboBox
		Me.frmDeity = New System.Windows.Forms.GroupBox
		Me.chkFullDumpwithoutSea = New System.Windows.Forms.CheckBox
		Me.fraUpdate = New System.Windows.Forms.GroupBox
		Me.chkSuppressUnitDamageReports = New System.Windows.Forms.CheckBox
		Me.chkSuppressTelegramNotification = New System.Windows.Forms.CheckBox
		Me.chkSuppressUpdateCommands = New System.Windows.Forms.CheckBox
		Me.chkSuppressOilContentAtSea = New System.Windows.Forms.CheckBox
		Me.chkCommodityRatio = New System.Windows.Forms.CheckBox
		Me.chkRead = New System.Windows.Forms.CheckBox
		Me.chkSuppressBmap = New System.Windows.Forms.CheckBox
		Me.chkSuppressStrength = New System.Windows.Forms.CheckBox
		Me.chkUpdate = New System.Windows.Forms.CheckBox
		Me.picTTS = New System.Windows.Forms.Panel
		Me.frmTTS = New System.Windows.Forms.GroupBox
		Me.chkTTSFlashes = New System.Windows.Forms.CheckBox
		Me.chkTTSBulletins = New System.Windows.Forms.CheckBox
		Me.chkTTSTelegrams = New System.Windows.Forms.CheckBox
		Me.chkTTSReportWindow = New System.Windows.Forms.CheckBox
		Me.picThemeGame = New System.Windows.Forms.Panel
		Me.fraThemeGame = New System.Windows.Forms.GroupBox
		Me._optThemeGames_7 = New System.Windows.Forms.RadioButton
		Me._optThemeGames_6 = New System.Windows.Forms.RadioButton
		Me._optThemeGames_5 = New System.Windows.Forms.RadioButton
		Me._optThemeGames_4 = New System.Windows.Forms.RadioButton
		Me._optThemeGames_3 = New System.Windows.Forms.RadioButton
		Me._optThemeGames_2 = New System.Windows.Forms.RadioButton
		Me._optThemeGames_1 = New System.Windows.Forms.RadioButton
		Me._optThemeGames_0 = New System.Windows.Forms.RadioButton
		Me.picScriptButton = New System.Windows.Forms.Panel
		Me.optCustomScript = New System.Windows.Forms.RadioButton
		Me.optJumpPoint = New System.Windows.Forms.RadioButton
		Me.fraScriptButton = New System.Windows.Forms.GroupBox
		Me.cmbJumpPoint = New System.Windows.Forms.ComboBox
		Me.cmdScriptImageBrowse = New System.Windows.Forms.Button
		Me.txtScriptImageFileName = New System.Windows.Forms.TextBox
		Me.chkScriptReport = New System.Windows.Forms.CheckBox
		Me.txtScriptFileName = New System.Windows.Forms.TextBox
		Me.cmdScriptBrowse = New System.Windows.Forms.Button
		Me.lblScriptImage = New System.Windows.Forms.Label
		Me.lblScriptFileName = New System.Windows.Forms.Label
		Me.cmbScriptButton = New System.Windows.Forms.ComboBox
		Me.picItem = New System.Windows.Forms.Panel
		Me.frmItemLabels = New System.Windows.Forms.GroupBox
		Me.txtItemDistributionPanelLabel = New System.Windows.Forms.TextBox
		Me.txtItemSectorPanelLabel = New System.Windows.Forms.TextBox
		Me.txtItemFormLabel = New System.Windows.Forms.TextBox
		Me.lblItemDistributionPanelLabel = New System.Windows.Forms.Label
		Me.lblItemSectorPanelLabel = New System.Windows.Forms.Label
		Me.lblItemFormLabel = New System.Windows.Forms.Label
		Me.cbItem = New System.Windows.Forms.ComboBox
		Me.frmItem = New System.Windows.Forms.GroupBox
		Me.txtItemConditionalName = New System.Windows.Forms.TextBox
		Me.txtItemIntelligenceNames = New System.Windows.Forms.TextBox
		Me.txtItemDBName = New System.Windows.Forms.TextBox
		Me.txtItemFormName = New System.Windows.Forms.TextBox
		Me.lblItemConditionalName = New System.Windows.Forms.Label
		Me.lblItemIntelligenceNames = New System.Windows.Forms.Label
		Me.lblItemDBName = New System.Windows.Forms.Label
		Me.lblItemFormName = New System.Windows.Forms.Label
		Me.picTelegram = New System.Windows.Forms.Panel
		Me._fraTelegram_2 = New System.Windows.Forms.GroupBox
		Me._cmbTelegramSound_2 = New System.Windows.Forms.ComboBox
		Me._txtTelegram_2 = New System.Windows.Forms.TextBox
		Me._chkTelegram_2 = New System.Windows.Forms.CheckBox
		Me._lblTelegram_2 = New System.Windows.Forms.Label
		Me._fraTelegram_1 = New System.Windows.Forms.GroupBox
		Me._cmbTelegramSound_1 = New System.Windows.Forms.ComboBox
		Me._txtTelegram_1 = New System.Windows.Forms.TextBox
		Me.Text1 = New System.Windows.Forms.TextBox
		Me._chkTelegram_1 = New System.Windows.Forms.CheckBox
		Me._lblTelegram_1 = New System.Windows.Forms.Label
		Me.cmbTelegram = New System.Windows.Forms.ComboBox
		Me._fraTelegram_0 = New System.Windows.Forms.GroupBox
		Me._cmbTelegramSound_0 = New System.Windows.Forms.ComboBox
		Me._txtTelegram_0 = New System.Windows.Forms.TextBox
		Me._chkTelegram_0 = New System.Windows.Forms.CheckBox
		Me._lblTelegram_0 = New System.Windows.Forms.Label
		Me.picToolbar = New System.Windows.Forms.Panel
		Me.fraToolbar = New System.Windows.Forms.GroupBox
		Me.lstToolbar = New System.Windows.Forms.CheckedListBox
		Me.tbsOptions = New AxMSComctlLib.AxTabStrip
		Me.cmdApply = New System.Windows.Forms.Button
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me.chkTelegram = New Microsoft.VisualBasic.Compatibility.VB6.CheckBoxArray(components)
		Me.cmbTelegramSound = New Microsoft.VisualBasic.Compatibility.VB6.ComboBoxArray(components)
		Me.fraTelegram = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(components)
		Me.lblTelegram = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
		Me.optThemeGames = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(components)
		Me.txtTelegram = New Microsoft.VisualBasic.Compatibility.VB6.TextBoxArray(components)
		Me.picPlaying.SuspendLayout()
		Me.frmPathLength.SuspendLayout()
		Me.frmFriendly.SuspendLayout()
		Me.frmAutoAttack.SuspendLayout()
		Me.fraGame.SuspendLayout()
		Me.picView.SuspendLayout()
		Me.fraView.SuspendLayout()
		Me.frmFlash.SuspendLayout()
		Me.frmImage.SuspendLayout()
		Me.picUpdate.SuspendLayout()
		Me.fraPlayer.SuspendLayout()
		Me.frmDeity.SuspendLayout()
		Me.fraUpdate.SuspendLayout()
		Me.picTTS.SuspendLayout()
		Me.frmTTS.SuspendLayout()
		Me.picThemeGame.SuspendLayout()
		Me.fraThemeGame.SuspendLayout()
		Me.picScriptButton.SuspendLayout()
		Me.fraScriptButton.SuspendLayout()
		Me.picItem.SuspendLayout()
		Me.frmItemLabels.SuspendLayout()
		Me.frmItem.SuspendLayout()
		Me.picTelegram.SuspendLayout()
		Me._fraTelegram_2.SuspendLayout()
		Me._fraTelegram_1.SuspendLayout()
		Me._fraTelegram_0.SuspendLayout()
		Me.picToolbar.SuspendLayout()
		Me.fraToolbar.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.tbsOptions, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.chkTelegram, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.cmbTelegramSound, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.fraTelegram, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.lblTelegram, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.optThemeGames, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.txtTelegram, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Text = "Options"
		Me.ClientSize = New System.Drawing.Size(410, 223)
		Me.Location = New System.Drawing.Point(171, 100)
		Me.KeyPreview = True
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmOptions"
		Me.picPlaying.Size = New System.Drawing.Size(379, 148)
		Me.picPlaying.Location = New System.Drawing.Point(16, 32)
		Me.picPlaying.TabIndex = 18
		Me.picPlaying.TabStop = False
		Me.picPlaying.Visible = False
		Me.picPlaying.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picPlaying.Dock = System.Windows.Forms.DockStyle.None
		Me.picPlaying.BackColor = System.Drawing.SystemColors.Control
		Me.picPlaying.CausesValidation = True
		Me.picPlaying.Enabled = True
		Me.picPlaying.ForeColor = System.Drawing.SystemColors.ControlText
		Me.picPlaying.Cursor = System.Windows.Forms.Cursors.Default
		Me.picPlaying.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picPlaying.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picPlaying.Name = "picPlaying"
		Me.frmPathLength.Text = "Path Length"
		Me.frmPathLength.Size = New System.Drawing.Size(73, 41)
		Me.frmPathLength.Location = New System.Drawing.Point(304, 104)
		Me.frmPathLength.TabIndex = 118
		Me.frmPathLength.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmPathLength.BackColor = System.Drawing.SystemColors.Control
		Me.frmPathLength.Enabled = True
		Me.frmPathLength.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmPathLength.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmPathLength.Visible = True
		Me.frmPathLength.Padding = New System.Windows.Forms.Padding(0)
		Me.frmPathLength.Name = "frmPathLength"
		Me.txtPathLength.AutoSize = False
		Me.txtPathLength.Size = New System.Drawing.Size(41, 19)
		Me.txtPathLength.Location = New System.Drawing.Point(16, 16)
		Me.txtPathLength.TabIndex = 119
		Me.txtPathLength.Text = "20"
		Me.ToolTip1.SetToolTip(Me.txtPathLength, "Maximum Path Length Checked")
		Me.txtPathLength.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtPathLength.AcceptsReturn = True
		Me.txtPathLength.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtPathLength.BackColor = System.Drawing.SystemColors.Window
		Me.txtPathLength.CausesValidation = True
		Me.txtPathLength.Enabled = True
		Me.txtPathLength.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtPathLength.HideSelection = True
		Me.txtPathLength.ReadOnly = False
		Me.txtPathLength.Maxlength = 0
		Me.txtPathLength.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtPathLength.MultiLine = False
		Me.txtPathLength.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtPathLength.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtPathLength.TabStop = True
		Me.txtPathLength.Visible = True
		Me.txtPathLength.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtPathLength.Name = "txtPathLength"
		Me.frmFriendly.Text = "Friendly Relation"
		Me.frmFriendly.Size = New System.Drawing.Size(129, 41)
		Me.frmFriendly.Location = New System.Drawing.Point(0, 104)
		Me.frmFriendly.TabIndex = 24
		Me.frmFriendly.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmFriendly.BackColor = System.Drawing.SystemColors.Control
		Me.frmFriendly.Enabled = True
		Me.frmFriendly.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmFriendly.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmFriendly.Visible = True
		Me.frmFriendly.Padding = New System.Windows.Forms.Padding(0)
		Me.frmFriendly.Name = "frmFriendly"
		Me.optFriendlyAlly.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optFriendlyAlly.Text = "Allied"
		Me.optFriendlyAlly.Size = New System.Drawing.Size(49, 17)
		Me.optFriendlyAlly.Location = New System.Drawing.Point(8, 16)
		Me.optFriendlyAlly.TabIndex = 26
		Me.optFriendlyAlly.Checked = True
		Me.optFriendlyAlly.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optFriendlyAlly.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optFriendlyAlly.BackColor = System.Drawing.SystemColors.Control
		Me.optFriendlyAlly.CausesValidation = True
		Me.optFriendlyAlly.Enabled = True
		Me.optFriendlyAlly.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optFriendlyAlly.Cursor = System.Windows.Forms.Cursors.Default
		Me.optFriendlyAlly.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optFriendlyAlly.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optFriendlyAlly.TabStop = True
		Me.optFriendlyAlly.Visible = True
		Me.optFriendlyAlly.Name = "optFriendlyAlly"
		Me.optFriendlyNeutral.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optFriendlyNeutral.Text = "Neutral"
		Me.optFriendlyNeutral.Size = New System.Drawing.Size(57, 17)
		Me.optFriendlyNeutral.Location = New System.Drawing.Point(64, 16)
		Me.optFriendlyNeutral.TabIndex = 25
		Me.optFriendlyNeutral.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optFriendlyNeutral.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optFriendlyNeutral.BackColor = System.Drawing.SystemColors.Control
		Me.optFriendlyNeutral.CausesValidation = True
		Me.optFriendlyNeutral.Enabled = True
		Me.optFriendlyNeutral.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optFriendlyNeutral.Cursor = System.Windows.Forms.Cursors.Default
		Me.optFriendlyNeutral.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optFriendlyNeutral.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optFriendlyNeutral.TabStop = True
		Me.optFriendlyNeutral.Checked = False
		Me.optFriendlyNeutral.Visible = True
		Me.optFriendlyNeutral.Name = "optFriendlyNeutral"
		Me.frmAutoAttack.Text = "Auto Attack"
		Me.frmAutoAttack.Size = New System.Drawing.Size(161, 41)
		Me.frmAutoAttack.Location = New System.Drawing.Point(136, 104)
		Me.frmAutoAttack.TabIndex = 56
		Me.frmAutoAttack.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmAutoAttack.BackColor = System.Drawing.SystemColors.Control
		Me.frmAutoAttack.Enabled = True
		Me.frmAutoAttack.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmAutoAttack.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmAutoAttack.Visible = True
		Me.frmAutoAttack.Padding = New System.Windows.Forms.Padding(0)
		Me.frmAutoAttack.Name = "frmAutoAttack"
		Me.txtDefenseScaling.AutoSize = False
		Me.txtDefenseScaling.Size = New System.Drawing.Size(41, 19)
		Me.txtDefenseScaling.Location = New System.Drawing.Point(112, 16)
		Me.txtDefenseScaling.TabIndex = 57
		Me.txtDefenseScaling.Text = "20"
		Me.txtDefenseScaling.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtDefenseScaling.AcceptsReturn = True
		Me.txtDefenseScaling.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtDefenseScaling.BackColor = System.Drawing.SystemColors.Window
		Me.txtDefenseScaling.CausesValidation = True
		Me.txtDefenseScaling.Enabled = True
		Me.txtDefenseScaling.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtDefenseScaling.HideSelection = True
		Me.txtDefenseScaling.ReadOnly = False
		Me.txtDefenseScaling.Maxlength = 0
		Me.txtDefenseScaling.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtDefenseScaling.MultiLine = False
		Me.txtDefenseScaling.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtDefenseScaling.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtDefenseScaling.TabStop = True
		Me.txtDefenseScaling.Visible = True
		Me.txtDefenseScaling.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtDefenseScaling.Name = "txtDefenseScaling"
		Me.lblDefenseScaling.Text = "DefenseScaling (%)"
		Me.lblDefenseScaling.Size = New System.Drawing.Size(105, 17)
		Me.lblDefenseScaling.Location = New System.Drawing.Point(8, 16)
		Me.lblDefenseScaling.TabIndex = 58
		Me.lblDefenseScaling.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblDefenseScaling.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblDefenseScaling.BackColor = System.Drawing.SystemColors.Control
		Me.lblDefenseScaling.Enabled = True
		Me.lblDefenseScaling.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblDefenseScaling.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblDefenseScaling.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblDefenseScaling.UseMnemonic = True
		Me.lblDefenseScaling.Visible = True
		Me.lblDefenseScaling.AutoSize = False
		Me.lblDefenseScaling.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblDefenseScaling.Name = "lblDefenseScaling"
		Me.fraGame.Text = "Game"
		Me.fraGame.Size = New System.Drawing.Size(377, 99)
		Me.fraGame.Location = New System.Drawing.Point(0, 0)
		Me.fraGame.TabIndex = 19
		Me.fraGame.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraGame.BackColor = System.Drawing.SystemColors.Control
		Me.fraGame.Enabled = True
		Me.fraGame.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraGame.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraGame.Visible = True
		Me.fraGame.Padding = New System.Windows.Forms.Padding(0)
		Me.fraGame.Name = "fraGame"
		Me.chkDumpHeaders.Text = "Display Dump Headers"
		Me.chkDumpHeaders.Size = New System.Drawing.Size(145, 25)
		Me.chkDumpHeaders.Location = New System.Drawing.Point(136, 72)
		Me.chkDumpHeaders.TabIndex = 116
		Me.chkDumpHeaders.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkDumpHeaders.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkDumpHeaders.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkDumpHeaders.BackColor = System.Drawing.SystemColors.Control
		Me.chkDumpHeaders.CausesValidation = True
		Me.chkDumpHeaders.Enabled = True
		Me.chkDumpHeaders.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkDumpHeaders.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkDumpHeaders.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkDumpHeaders.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkDumpHeaders.TabStop = True
		Me.chkDumpHeaders.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkDumpHeaders.Visible = True
		Me.chkDumpHeaders.Name = "chkDumpHeaders"
		Me.chkCapital.Text = "Use Capital as Start"
		Me.chkCapital.Size = New System.Drawing.Size(113, 13)
		Me.chkCapital.Location = New System.Drawing.Point(8, 76)
		Me.chkCapital.TabIndex = 65
		Me.chkCapital.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkCapital.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkCapital.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkCapital.BackColor = System.Drawing.SystemColors.Control
		Me.chkCapital.CausesValidation = True
		Me.chkCapital.Enabled = True
		Me.chkCapital.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkCapital.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkCapital.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkCapital.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkCapital.TabStop = True
		Me.chkCapital.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkCapital.Visible = True
		Me.chkCapital.Name = "chkCapital"
		Me.chkLocalHelp.Text = "Use Local Help"
		Me.chkLocalHelp.Size = New System.Drawing.Size(105, 13)
		Me.chkLocalHelp.Location = New System.Drawing.Point(8, 56)
		Me.chkLocalHelp.TabIndex = 45
		Me.chkLocalHelp.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkLocalHelp.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkLocalHelp.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkLocalHelp.BackColor = System.Drawing.SystemColors.Control
		Me.chkLocalHelp.CausesValidation = True
		Me.chkLocalHelp.Enabled = True
		Me.chkLocalHelp.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkLocalHelp.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkLocalHelp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkLocalHelp.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkLocalHelp.TabStop = True
		Me.chkLocalHelp.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkLocalHelp.Visible = True
		Me.chkLocalHelp.Name = "chkLocalHelp"
		Me.chkSound.Text = "Use Sound"
		Me.chkSound.Size = New System.Drawing.Size(89, 13)
		Me.chkSound.Location = New System.Drawing.Point(8, 16)
		Me.chkSound.TabIndex = 23
		Me.chkSound.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkSound.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkSound.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkSound.BackColor = System.Drawing.SystemColors.Control
		Me.chkSound.CausesValidation = True
		Me.chkSound.Enabled = True
		Me.chkSound.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkSound.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkSound.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkSound.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkSound.TabStop = True
		Me.chkSound.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkSound.Visible = True
		Me.chkSound.Name = "chkSound"
		Me.chkNoIron.Text = "No Iron Game"
		Me.chkNoIron.Size = New System.Drawing.Size(89, 13)
		Me.chkNoIron.Location = New System.Drawing.Point(8, 36)
		Me.chkNoIron.TabIndex = 22
		Me.chkNoIron.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkNoIron.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkNoIron.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkNoIron.BackColor = System.Drawing.SystemColors.Control
		Me.chkNoIron.CausesValidation = True
		Me.chkNoIron.Enabled = True
		Me.chkNoIron.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkNoIron.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkNoIron.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkNoIron.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkNoIron.TabStop = True
		Me.chkNoIron.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkNoIron.Visible = True
		Me.chkNoIron.Name = "chkNoIron"
		Me.chkRetreat.Text = "Set Retreat on Nav/March"
		Me.chkRetreat.Size = New System.Drawing.Size(153, 13)
		Me.chkRetreat.Location = New System.Drawing.Point(136, 56)
		Me.chkRetreat.TabIndex = 21
		Me.chkRetreat.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkRetreat.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkRetreat.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkRetreat.BackColor = System.Drawing.SystemColors.Control
		Me.chkRetreat.CausesValidation = True
		Me.chkRetreat.Enabled = True
		Me.chkRetreat.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkRetreat.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkRetreat.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkRetreat.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkRetreat.TabStop = True
		Me.chkRetreat.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkRetreat.Visible = True
		Me.chkRetreat.Name = "chkRetreat"
		Me.chkMaxProd.Text = "Max. Prod. instead of Pop."
		Me.chkMaxProd.Size = New System.Drawing.Size(153, 13)
		Me.chkMaxProd.Location = New System.Drawing.Point(136, 37)
		Me.chkMaxProd.TabIndex = 20
		Me.chkMaxProd.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkMaxProd.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkMaxProd.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkMaxProd.BackColor = System.Drawing.SystemColors.Control
		Me.chkMaxProd.CausesValidation = True
		Me.chkMaxProd.Enabled = True
		Me.chkMaxProd.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkMaxProd.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkMaxProd.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkMaxProd.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkMaxProd.TabStop = True
		Me.chkMaxProd.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkMaxProd.Visible = True
		Me.chkMaxProd.Name = "chkMaxProd"
		Me.picView.Size = New System.Drawing.Size(379, 148)
		Me.picView.Location = New System.Drawing.Point(16, 32)
		Me.picView.TabIndex = 5
		Me.picView.TabStop = False
		Me.picView.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picView.Dock = System.Windows.Forms.DockStyle.None
		Me.picView.BackColor = System.Drawing.SystemColors.Control
		Me.picView.CausesValidation = True
		Me.picView.Enabled = True
		Me.picView.ForeColor = System.Drawing.SystemColors.ControlText
		Me.picView.Cursor = System.Windows.Forms.Cursors.Default
		Me.picView.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picView.Visible = True
		Me.picView.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picView.Name = "picView"
		Me.fraView.Text = "Viewing"
		Me.fraView.Size = New System.Drawing.Size(249, 88)
		Me.fraView.Location = New System.Drawing.Point(0, 8)
		Me.fraView.TabIndex = 6
		Me.fraView.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraView.BackColor = System.Drawing.SystemColors.Control
		Me.fraView.Enabled = True
		Me.fraView.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraView.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraView.Visible = True
		Me.fraView.Padding = New System.Windows.Forms.Padding(0)
		Me.fraView.Name = "fraView"
		Me.chkShortName.Text = "Use Short Names in Unit Grid"
		Me.chkShortName.Size = New System.Drawing.Size(161, 17)
		Me.chkShortName.Location = New System.Drawing.Point(80, 16)
		Me.chkShortName.TabIndex = 117
		Me.chkShortName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkShortName.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkShortName.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkShortName.BackColor = System.Drawing.SystemColors.Control
		Me.chkShortName.CausesValidation = True
		Me.chkShortName.Enabled = True
		Me.chkShortName.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkShortName.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkShortName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkShortName.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkShortName.TabStop = True
		Me.chkShortName.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkShortName.Visible = True
		Me.chkShortName.Name = "chkShortName"
		Me.chkClear.Text = "Clear Result Box on New Command"
		Me.chkClear.Size = New System.Drawing.Size(193, 25)
		Me.chkClear.Location = New System.Drawing.Point(8, 56)
		Me.chkClear.TabIndex = 17
		Me.chkClear.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkClear.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkClear.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkClear.BackColor = System.Drawing.SystemColors.Control
		Me.chkClear.CausesValidation = True
		Me.chkClear.Enabled = True
		Me.chkClear.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkClear.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkClear.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkClear.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkClear.TabStop = True
		Me.chkClear.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkClear.Visible = True
		Me.chkClear.Name = "chkClear"
		Me.chkToolbar.Text = "Toolbar"
		Me.chkToolbar.Size = New System.Drawing.Size(81, 17)
		Me.chkToolbar.Location = New System.Drawing.Point(8, 16)
		Me.chkToolbar.TabIndex = 9
		Me.chkToolbar.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkToolbar.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkToolbar.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkToolbar.BackColor = System.Drawing.SystemColors.Control
		Me.chkToolbar.CausesValidation = True
		Me.chkToolbar.Enabled = True
		Me.chkToolbar.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkToolbar.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkToolbar.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkToolbar.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkToolbar.TabStop = True
		Me.chkToolbar.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkToolbar.Visible = True
		Me.chkToolbar.Name = "chkToolbar"
		Me.chkStatusBar.Text = "Status Bar"
		Me.chkStatusBar.Size = New System.Drawing.Size(73, 17)
		Me.chkStatusBar.Location = New System.Drawing.Point(8, 36)
		Me.chkStatusBar.TabIndex = 8
		Me.chkStatusBar.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkStatusBar.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkStatusBar.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkStatusBar.BackColor = System.Drawing.SystemColors.Control
		Me.chkStatusBar.CausesValidation = True
		Me.chkStatusBar.Enabled = True
		Me.chkStatusBar.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkStatusBar.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkStatusBar.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkStatusBar.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkStatusBar.TabStop = True
		Me.chkStatusBar.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkStatusBar.Visible = True
		Me.chkStatusBar.Name = "chkStatusBar"
		Me.chkWarship.Text = "Show Only Warships"
		Me.chkWarship.Size = New System.Drawing.Size(129, 17)
		Me.chkWarship.Location = New System.Drawing.Point(80, 36)
		Me.chkWarship.TabIndex = 7
		Me.chkWarship.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkWarship.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkWarship.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkWarship.BackColor = System.Drawing.SystemColors.Control
		Me.chkWarship.CausesValidation = True
		Me.chkWarship.Enabled = True
		Me.chkWarship.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkWarship.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkWarship.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkWarship.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkWarship.TabStop = True
		Me.chkWarship.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkWarship.Visible = True
		Me.chkWarship.Name = "chkWarship"
		Me.frmFlash.Text = "Flash"
		Me.frmFlash.Size = New System.Drawing.Size(377, 41)
		Me.frmFlash.Location = New System.Drawing.Point(0, 104)
		Me.frmFlash.TabIndex = 59
		Me.frmFlash.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmFlash.BackColor = System.Drawing.SystemColors.Control
		Me.frmFlash.Enabled = True
		Me.frmFlash.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmFlash.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmFlash.Visible = True
		Me.frmFlash.Padding = New System.Windows.Forms.Padding(0)
		Me.frmFlash.Name = "frmFlash"
		Me.chkFlashLogging.Text = "Logging"
		Me.chkFlashLogging.Size = New System.Drawing.Size(65, 17)
		Me.chkFlashLogging.Location = New System.Drawing.Point(280, 16)
		Me.chkFlashLogging.TabIndex = 104
		Me.chkFlashLogging.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkFlashLogging.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkFlashLogging.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkFlashLogging.BackColor = System.Drawing.SystemColors.Control
		Me.chkFlashLogging.CausesValidation = True
		Me.chkFlashLogging.Enabled = True
		Me.chkFlashLogging.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkFlashLogging.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkFlashLogging.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkFlashLogging.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkFlashLogging.TabStop = True
		Me.chkFlashLogging.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkFlashLogging.Visible = True
		Me.chkFlashLogging.Name = "chkFlashLogging"
		Me.chkFlashPlayerColors.Text = "Use Player Colors"
		Me.chkFlashPlayerColors.Size = New System.Drawing.Size(113, 17)
		Me.chkFlashPlayerColors.Location = New System.Drawing.Point(144, 16)
		Me.chkFlashPlayerColors.TabIndex = 61
		Me.chkFlashPlayerColors.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkFlashPlayerColors.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkFlashPlayerColors.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkFlashPlayerColors.BackColor = System.Drawing.SystemColors.Control
		Me.chkFlashPlayerColors.CausesValidation = True
		Me.chkFlashPlayerColors.Enabled = True
		Me.chkFlashPlayerColors.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkFlashPlayerColors.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkFlashPlayerColors.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkFlashPlayerColors.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkFlashPlayerColors.TabStop = True
		Me.chkFlashPlayerColors.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkFlashPlayerColors.Visible = True
		Me.chkFlashPlayerColors.Name = "chkFlashPlayerColors"
		Me.chkFlashChat.Text = "Window Enabled"
		Me.chkFlashChat.Size = New System.Drawing.Size(113, 17)
		Me.chkFlashChat.Location = New System.Drawing.Point(8, 16)
		Me.chkFlashChat.TabIndex = 60
		Me.chkFlashChat.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkFlashChat.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkFlashChat.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkFlashChat.BackColor = System.Drawing.SystemColors.Control
		Me.chkFlashChat.CausesValidation = True
		Me.chkFlashChat.Enabled = True
		Me.chkFlashChat.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkFlashChat.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkFlashChat.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkFlashChat.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkFlashChat.TabStop = True
		Me.chkFlashChat.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkFlashChat.Visible = True
		Me.chkFlashChat.Name = "chkFlashChat"
		Me.frmImage.Text = "Background Image"
		Me.frmImage.Size = New System.Drawing.Size(121, 89)
		Me.frmImage.Location = New System.Drawing.Point(256, 8)
		Me.frmImage.TabIndex = 50
		Me.frmImage.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmImage.BackColor = System.Drawing.SystemColors.Control
		Me.frmImage.Enabled = True
		Me.frmImage.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmImage.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmImage.Visible = True
		Me.frmImage.Padding = New System.Windows.Forms.Padding(0)
		Me.frmImage.Name = "frmImage"
		Me.picBorderColor.Size = New System.Drawing.Size(33, 17)
		Me.picBorderColor.Location = New System.Drawing.Point(80, 64)
		Me.picBorderColor.TabIndex = 53
		Me.picBorderColor.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picBorderColor.Dock = System.Windows.Forms.DockStyle.None
		Me.picBorderColor.BackColor = System.Drawing.SystemColors.Control
		Me.picBorderColor.CausesValidation = True
		Me.picBorderColor.Enabled = True
		Me.picBorderColor.ForeColor = System.Drawing.SystemColors.ControlText
		Me.picBorderColor.Cursor = System.Windows.Forms.Cursors.Default
		Me.picBorderColor.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picBorderColor.TabStop = True
		Me.picBorderColor.Visible = True
		Me.picBorderColor.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.picBorderColor.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.picBorderColor.Name = "picBorderColor"
		Me.cmdImage.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdImage.Text = "Browse"
		Me.cmdImage.Size = New System.Drawing.Size(57, 17)
		Me.cmdImage.Location = New System.Drawing.Point(32, 40)
		Me.cmdImage.TabIndex = 52
		Me.cmdImage.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdImage.BackColor = System.Drawing.SystemColors.Control
		Me.cmdImage.CausesValidation = True
		Me.cmdImage.Enabled = True
		Me.cmdImage.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdImage.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdImage.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdImage.TabStop = True
		Me.cmdImage.Name = "cmdImage"
		Me.txtImageName.AutoSize = False
		Me.txtImageName.Size = New System.Drawing.Size(97, 21)
		Me.txtImageName.Location = New System.Drawing.Point(16, 16)
		Me.txtImageName.TabIndex = 51
		Me.txtImageName.Text = "Text1"
		Me.txtImageName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtImageName.AcceptsReturn = True
		Me.txtImageName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtImageName.BackColor = System.Drawing.SystemColors.Window
		Me.txtImageName.CausesValidation = True
		Me.txtImageName.Enabled = True
		Me.txtImageName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtImageName.HideSelection = True
		Me.txtImageName.ReadOnly = False
		Me.txtImageName.Maxlength = 0
		Me.txtImageName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtImageName.MultiLine = False
		Me.txtImageName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtImageName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtImageName.TabStop = True
		Me.txtImageName.Visible = True
		Me.txtImageName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtImageName.Name = "txtImageName"
		Me.lblBorderColor.Text = "Border Color"
		Me.lblBorderColor.Size = New System.Drawing.Size(65, 17)
		Me.lblBorderColor.Location = New System.Drawing.Point(16, 64)
		Me.lblBorderColor.TabIndex = 54
		Me.lblBorderColor.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblBorderColor.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblBorderColor.BackColor = System.Drawing.SystemColors.Control
		Me.lblBorderColor.Enabled = True
		Me.lblBorderColor.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblBorderColor.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblBorderColor.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblBorderColor.UseMnemonic = True
		Me.lblBorderColor.Visible = True
		Me.lblBorderColor.AutoSize = False
		Me.lblBorderColor.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblBorderColor.Name = "lblBorderColor"
		Me.picUpdate.Size = New System.Drawing.Size(379, 148)
		Me.picUpdate.Location = New System.Drawing.Point(16, 32)
		Me.picUpdate.TabIndex = 4
		Me.picUpdate.TabStop = False
		Me.picUpdate.Visible = False
		Me.picUpdate.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picUpdate.Dock = System.Windows.Forms.DockStyle.None
		Me.picUpdate.BackColor = System.Drawing.SystemColors.Control
		Me.picUpdate.CausesValidation = True
		Me.picUpdate.Enabled = True
		Me.picUpdate.ForeColor = System.Drawing.SystemColors.ControlText
		Me.picUpdate.Cursor = System.Windows.Forms.Cursors.Default
		Me.picUpdate.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picUpdate.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picUpdate.Name = "picUpdate"
		Me.fraPlayer.Text = "Display Other Players"
		Me.fraPlayer.Size = New System.Drawing.Size(169, 41)
		Me.fraPlayer.Location = New System.Drawing.Point(200, 104)
		Me.fraPlayer.TabIndex = 15
		Me.fraPlayer.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraPlayer.BackColor = System.Drawing.SystemColors.Control
		Me.fraPlayer.Enabled = True
		Me.fraPlayer.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraPlayer.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraPlayer.Visible = True
		Me.fraPlayer.Padding = New System.Windows.Forms.Padding(0)
		Me.fraPlayer.Name = "fraPlayer"
		Me.cmbPlayer.Size = New System.Drawing.Size(153, 21)
		Me.cmbPlayer.Location = New System.Drawing.Point(8, 16)
		Me.cmbPlayer.Items.AddRange(New Object(){"Disabled", "Every 1 Minute", "Every 2 Minutes", "Every 5 Minutes", "Every 10 Minutes"})
		Me.cmbPlayer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cmbPlayer.TabIndex = 16
		Me.cmbPlayer.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbPlayer.BackColor = System.Drawing.SystemColors.Window
		Me.cmbPlayer.CausesValidation = True
		Me.cmbPlayer.Enabled = True
		Me.cmbPlayer.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbPlayer.IntegralHeight = True
		Me.cmbPlayer.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbPlayer.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbPlayer.Sorted = False
		Me.cmbPlayer.TabStop = True
		Me.cmbPlayer.Visible = True
		Me.cmbPlayer.Name = "cmbPlayer"
		Me.frmDeity.Text = "Deity"
		Me.frmDeity.Size = New System.Drawing.Size(169, 41)
		Me.frmDeity.Location = New System.Drawing.Point(0, 104)
		Me.frmDeity.TabIndex = 62
		Me.frmDeity.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmDeity.BackColor = System.Drawing.SystemColors.Control
		Me.frmDeity.Enabled = True
		Me.frmDeity.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmDeity.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmDeity.Visible = True
		Me.frmDeity.Padding = New System.Windows.Forms.Padding(0)
		Me.frmDeity.Name = "frmDeity"
		Me.chkFullDumpwithoutSea.Text = "Full Dump without Sea"
		Me.chkFullDumpwithoutSea.Size = New System.Drawing.Size(137, 13)
		Me.chkFullDumpwithoutSea.Location = New System.Drawing.Point(16, 16)
		Me.chkFullDumpwithoutSea.TabIndex = 63
		Me.chkFullDumpwithoutSea.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkFullDumpwithoutSea.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkFullDumpwithoutSea.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkFullDumpwithoutSea.BackColor = System.Drawing.SystemColors.Control
		Me.chkFullDumpwithoutSea.CausesValidation = True
		Me.chkFullDumpwithoutSea.Enabled = True
		Me.chkFullDumpwithoutSea.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkFullDumpwithoutSea.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkFullDumpwithoutSea.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkFullDumpwithoutSea.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkFullDumpwithoutSea.TabStop = True
		Me.chkFullDumpwithoutSea.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkFullDumpwithoutSea.Visible = True
		Me.chkFullDumpwithoutSea.Name = "chkFullDumpwithoutSea"
		Me.fraUpdate.Text = "Updating"
		Me.fraUpdate.Size = New System.Drawing.Size(369, 100)
		Me.fraUpdate.Location = New System.Drawing.Point(0, 0)
		Me.fraUpdate.TabIndex = 10
		Me.fraUpdate.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraUpdate.BackColor = System.Drawing.SystemColors.Control
		Me.fraUpdate.Enabled = True
		Me.fraUpdate.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraUpdate.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraUpdate.Visible = True
		Me.fraUpdate.Padding = New System.Windows.Forms.Padding(0)
		Me.fraUpdate.Name = "fraUpdate"
		Me.chkSuppressUnitDamageReports.Text = "Suppress Unit Damage Reports"
		Me.chkSuppressUnitDamageReports.Size = New System.Drawing.Size(177, 17)
		Me.chkSuppressUnitDamageReports.Location = New System.Drawing.Point(8, 64)
		Me.chkSuppressUnitDamageReports.TabIndex = 101
		Me.ToolTip1.SetToolTip(Me.chkSuppressUnitDamageReports, "Suppress the displaying of Telegram Notification with INORM OFF")
		Me.chkSuppressUnitDamageReports.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkSuppressUnitDamageReports.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkSuppressUnitDamageReports.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkSuppressUnitDamageReports.BackColor = System.Drawing.SystemColors.Control
		Me.chkSuppressUnitDamageReports.CausesValidation = True
		Me.chkSuppressUnitDamageReports.Enabled = True
		Me.chkSuppressUnitDamageReports.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkSuppressUnitDamageReports.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkSuppressUnitDamageReports.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkSuppressUnitDamageReports.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkSuppressUnitDamageReports.TabStop = True
		Me.chkSuppressUnitDamageReports.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkSuppressUnitDamageReports.Visible = True
		Me.chkSuppressUnitDamageReports.Name = "chkSuppressUnitDamageReports"
		Me.chkSuppressTelegramNotification.Text = "Suppress Telegrams Notifiations"
		Me.chkSuppressTelegramNotification.Size = New System.Drawing.Size(177, 17)
		Me.chkSuppressTelegramNotification.Location = New System.Drawing.Point(8, 80)
		Me.chkSuppressTelegramNotification.TabIndex = 100
		Me.ToolTip1.SetToolTip(Me.chkSuppressTelegramNotification, "Suppress the displaying of Telegram Notification with INORM OFF")
		Me.chkSuppressTelegramNotification.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkSuppressTelegramNotification.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkSuppressTelegramNotification.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkSuppressTelegramNotification.BackColor = System.Drawing.SystemColors.Control
		Me.chkSuppressTelegramNotification.CausesValidation = True
		Me.chkSuppressTelegramNotification.Enabled = True
		Me.chkSuppressTelegramNotification.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkSuppressTelegramNotification.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkSuppressTelegramNotification.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkSuppressTelegramNotification.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkSuppressTelegramNotification.TabStop = True
		Me.chkSuppressTelegramNotification.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkSuppressTelegramNotification.Visible = True
		Me.chkSuppressTelegramNotification.Name = "chkSuppressTelegramNotification"
		Me.chkSuppressUpdateCommands.Text = "Suppress Update Commands"
		Me.chkSuppressUpdateCommands.Size = New System.Drawing.Size(161, 17)
		Me.chkSuppressUpdateCommands.Location = New System.Drawing.Point(184, 64)
		Me.chkSuppressUpdateCommands.TabIndex = 99
		Me.ToolTip1.SetToolTip(Me.chkSuppressUpdateCommands, "Suppress the displaying of the update commands")
		Me.chkSuppressUpdateCommands.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkSuppressUpdateCommands.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkSuppressUpdateCommands.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkSuppressUpdateCommands.BackColor = System.Drawing.SystemColors.Control
		Me.chkSuppressUpdateCommands.CausesValidation = True
		Me.chkSuppressUpdateCommands.Enabled = True
		Me.chkSuppressUpdateCommands.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkSuppressUpdateCommands.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkSuppressUpdateCommands.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkSuppressUpdateCommands.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkSuppressUpdateCommands.TabStop = True
		Me.chkSuppressUpdateCommands.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkSuppressUpdateCommands.Visible = True
		Me.chkSuppressUpdateCommands.Name = "chkSuppressUpdateCommands"
		Me.chkSuppressOilContentAtSea.Text = "Suppress Oil Content Refresh"
		Me.chkSuppressOilContentAtSea.Size = New System.Drawing.Size(161, 17)
		Me.chkSuppressOilContentAtSea.Location = New System.Drawing.Point(184, 16)
		Me.chkSuppressOilContentAtSea.TabIndex = 98
		Me.ToolTip1.SetToolTip(Me.chkSuppressOilContentAtSea, "Suppress Oil Content Refresh For Sea Sectors")
		Me.chkSuppressOilContentAtSea.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkSuppressOilContentAtSea.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkSuppressOilContentAtSea.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkSuppressOilContentAtSea.BackColor = System.Drawing.SystemColors.Control
		Me.chkSuppressOilContentAtSea.CausesValidation = True
		Me.chkSuppressOilContentAtSea.Enabled = True
		Me.chkSuppressOilContentAtSea.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkSuppressOilContentAtSea.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkSuppressOilContentAtSea.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkSuppressOilContentAtSea.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkSuppressOilContentAtSea.TabStop = True
		Me.chkSuppressOilContentAtSea.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkSuppressOilContentAtSea.Visible = True
		Me.chkSuppressOilContentAtSea.Name = "chkSuppressOilContentAtSea"
		Me.chkCommodityRatio.Text = "Show Commodity Ratio"
		Me.chkCommodityRatio.Size = New System.Drawing.Size(137, 17)
		Me.chkCommodityRatio.Location = New System.Drawing.Point(8, 48)
		Me.chkCommodityRatio.TabIndex = 97
		Me.chkCommodityRatio.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkCommodityRatio.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkCommodityRatio.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkCommodityRatio.BackColor = System.Drawing.SystemColors.Control
		Me.chkCommodityRatio.CausesValidation = True
		Me.chkCommodityRatio.Enabled = True
		Me.chkCommodityRatio.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkCommodityRatio.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkCommodityRatio.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkCommodityRatio.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkCommodityRatio.TabStop = True
		Me.chkCommodityRatio.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkCommodityRatio.Visible = True
		Me.chkCommodityRatio.Name = "chkCommodityRatio"
		Me.chkRead.Text = "Automatic Read"
		Me.chkRead.Size = New System.Drawing.Size(153, 17)
		Me.chkRead.Location = New System.Drawing.Point(8, 32)
		Me.chkRead.TabIndex = 14
		Me.chkRead.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkRead.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkRead.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkRead.BackColor = System.Drawing.SystemColors.Control
		Me.chkRead.CausesValidation = True
		Me.chkRead.Enabled = True
		Me.chkRead.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkRead.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkRead.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkRead.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkRead.TabStop = True
		Me.chkRead.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkRead.Visible = True
		Me.chkRead.Name = "chkRead"
		Me.chkSuppressBmap.Text = "Suppress Bmap Refresh"
		Me.chkSuppressBmap.Size = New System.Drawing.Size(137, 17)
		Me.chkSuppressBmap.Location = New System.Drawing.Point(184, 48)
		Me.chkSuppressBmap.TabIndex = 13
		Me.chkSuppressBmap.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkSuppressBmap.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkSuppressBmap.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkSuppressBmap.BackColor = System.Drawing.SystemColors.Control
		Me.chkSuppressBmap.CausesValidation = True
		Me.chkSuppressBmap.Enabled = True
		Me.chkSuppressBmap.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkSuppressBmap.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkSuppressBmap.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkSuppressBmap.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkSuppressBmap.TabStop = True
		Me.chkSuppressBmap.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkSuppressBmap.Visible = True
		Me.chkSuppressBmap.Name = "chkSuppressBmap"
		Me.chkSuppressStrength.Text = "Suppress Strength Refresh"
		Me.chkSuppressStrength.Size = New System.Drawing.Size(153, 17)
		Me.chkSuppressStrength.Location = New System.Drawing.Point(184, 32)
		Me.chkSuppressStrength.TabIndex = 12
		Me.chkSuppressStrength.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkSuppressStrength.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkSuppressStrength.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkSuppressStrength.BackColor = System.Drawing.SystemColors.Control
		Me.chkSuppressStrength.CausesValidation = True
		Me.chkSuppressStrength.Enabled = True
		Me.chkSuppressStrength.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkSuppressStrength.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkSuppressStrength.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkSuppressStrength.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkSuppressStrength.TabStop = True
		Me.chkSuppressStrength.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkSuppressStrength.Visible = True
		Me.chkSuppressStrength.Name = "chkSuppressStrength"
		Me.chkUpdate.Text = "Automatic Update"
		Me.chkUpdate.Size = New System.Drawing.Size(153, 17)
		Me.chkUpdate.Location = New System.Drawing.Point(8, 16)
		Me.chkUpdate.TabIndex = 11
		Me.chkUpdate.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkUpdate.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkUpdate.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkUpdate.BackColor = System.Drawing.SystemColors.Control
		Me.chkUpdate.CausesValidation = True
		Me.chkUpdate.Enabled = True
		Me.chkUpdate.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkUpdate.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkUpdate.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkUpdate.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkUpdate.TabStop = True
		Me.chkUpdate.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkUpdate.Visible = True
		Me.chkUpdate.Name = "chkUpdate"
		Me.picTTS.Size = New System.Drawing.Size(379, 148)
		Me.picTTS.Location = New System.Drawing.Point(16, 32)
		Me.picTTS.TabIndex = 110
		Me.picTTS.TabStop = False
		Me.picTTS.Visible = False
		Me.picTTS.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picTTS.Dock = System.Windows.Forms.DockStyle.None
		Me.picTTS.BackColor = System.Drawing.SystemColors.Control
		Me.picTTS.CausesValidation = True
		Me.picTTS.Enabled = True
		Me.picTTS.ForeColor = System.Drawing.SystemColors.ControlText
		Me.picTTS.Cursor = System.Windows.Forms.Cursors.Default
		Me.picTTS.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picTTS.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picTTS.Name = "picTTS"
		Me.frmTTS.Text = "Text To Speech"
		Me.frmTTS.Size = New System.Drawing.Size(137, 136)
		Me.frmTTS.Location = New System.Drawing.Point(8, 8)
		Me.frmTTS.TabIndex = 111
		Me.frmTTS.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmTTS.BackColor = System.Drawing.SystemColors.Control
		Me.frmTTS.Enabled = True
		Me.frmTTS.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmTTS.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmTTS.Visible = True
		Me.frmTTS.Padding = New System.Windows.Forms.Padding(0)
		Me.frmTTS.Name = "frmTTS"
		Me.chkTTSFlashes.Text = "Flashes"
		Me.chkTTSFlashes.Size = New System.Drawing.Size(105, 17)
		Me.chkTTSFlashes.Location = New System.Drawing.Point(16, 96)
		Me.chkTTSFlashes.TabIndex = 115
		Me.ToolTip1.SetToolTip(Me.chkTTSFlashes, "Enable TTS for Flashes")
		Me.chkTTSFlashes.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkTTSFlashes.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkTTSFlashes.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkTTSFlashes.BackColor = System.Drawing.SystemColors.Control
		Me.chkTTSFlashes.CausesValidation = True
		Me.chkTTSFlashes.Enabled = True
		Me.chkTTSFlashes.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkTTSFlashes.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkTTSFlashes.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkTTSFlashes.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkTTSFlashes.TabStop = True
		Me.chkTTSFlashes.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkTTSFlashes.Visible = True
		Me.chkTTSFlashes.Name = "chkTTSFlashes"
		Me.chkTTSBulletins.Text = "Bulletins"
		Me.chkTTSBulletins.Size = New System.Drawing.Size(105, 17)
		Me.chkTTSBulletins.Location = New System.Drawing.Point(16, 72)
		Me.chkTTSBulletins.TabIndex = 114
		Me.ToolTip1.SetToolTip(Me.chkTTSBulletins, "Enable TTS for Bulletins")
		Me.chkTTSBulletins.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkTTSBulletins.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkTTSBulletins.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkTTSBulletins.BackColor = System.Drawing.SystemColors.Control
		Me.chkTTSBulletins.CausesValidation = True
		Me.chkTTSBulletins.Enabled = True
		Me.chkTTSBulletins.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkTTSBulletins.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkTTSBulletins.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkTTSBulletins.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkTTSBulletins.TabStop = True
		Me.chkTTSBulletins.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkTTSBulletins.Visible = True
		Me.chkTTSBulletins.Name = "chkTTSBulletins"
		Me.chkTTSTelegrams.Text = "Telegram"
		Me.chkTTSTelegrams.Size = New System.Drawing.Size(105, 17)
		Me.chkTTSTelegrams.Location = New System.Drawing.Point(16, 48)
		Me.chkTTSTelegrams.TabIndex = 113
		Me.ToolTip1.SetToolTip(Me.chkTTSTelegrams, "Enable TTS for Telegrams")
		Me.chkTTSTelegrams.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkTTSTelegrams.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkTTSTelegrams.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkTTSTelegrams.BackColor = System.Drawing.SystemColors.Control
		Me.chkTTSTelegrams.CausesValidation = True
		Me.chkTTSTelegrams.Enabled = True
		Me.chkTTSTelegrams.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkTTSTelegrams.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkTTSTelegrams.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkTTSTelegrams.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkTTSTelegrams.TabStop = True
		Me.chkTTSTelegrams.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkTTSTelegrams.Visible = True
		Me.chkTTSTelegrams.Name = "chkTTSTelegrams"
		Me.chkTTSReportWindow.Text = "Report Window"
		Me.chkTTSReportWindow.Size = New System.Drawing.Size(105, 17)
		Me.chkTTSReportWindow.Location = New System.Drawing.Point(16, 24)
		Me.chkTTSReportWindow.TabIndex = 112
		Me.ToolTip1.SetToolTip(Me.chkTTSReportWindow, "Enable TTS for Report Window")
		Me.chkTTSReportWindow.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkTTSReportWindow.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkTTSReportWindow.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkTTSReportWindow.BackColor = System.Drawing.SystemColors.Control
		Me.chkTTSReportWindow.CausesValidation = True
		Me.chkTTSReportWindow.Enabled = True
		Me.chkTTSReportWindow.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkTTSReportWindow.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkTTSReportWindow.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkTTSReportWindow.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkTTSReportWindow.TabStop = True
		Me.chkTTSReportWindow.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkTTSReportWindow.Visible = True
		Me.chkTTSReportWindow.Name = "chkTTSReportWindow"
		Me.picThemeGame.Size = New System.Drawing.Size(379, 148)
		Me.picThemeGame.Location = New System.Drawing.Point(16, 32)
		Me.picThemeGame.TabIndex = 46
		Me.picThemeGame.TabStop = False
		Me.picThemeGame.Visible = False
		Me.picThemeGame.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picThemeGame.Dock = System.Windows.Forms.DockStyle.None
		Me.picThemeGame.BackColor = System.Drawing.SystemColors.Control
		Me.picThemeGame.CausesValidation = True
		Me.picThemeGame.Enabled = True
		Me.picThemeGame.ForeColor = System.Drawing.SystemColors.ControlText
		Me.picThemeGame.Cursor = System.Windows.Forms.Cursors.Default
		Me.picThemeGame.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picThemeGame.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picThemeGame.Name = "picThemeGame"
		Me.fraThemeGame.Text = "Theme Game"
		Me.fraThemeGame.Size = New System.Drawing.Size(257, 136)
		Me.fraThemeGame.Location = New System.Drawing.Point(8, 8)
		Me.fraThemeGame.TabIndex = 47
		Me.fraThemeGame.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraThemeGame.BackColor = System.Drawing.SystemColors.Control
		Me.fraThemeGame.Enabled = True
		Me.fraThemeGame.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraThemeGame.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraThemeGame.Visible = True
		Me.fraThemeGame.Padding = New System.Windows.Forms.Padding(0)
		Me.fraThemeGame.Name = "fraThemeGame"
		Me._optThemeGames_7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optThemeGames_7.Text = "Heavy Metal"
		Me._optThemeGames_7.Size = New System.Drawing.Size(97, 17)
		Me._optThemeGames_7.Location = New System.Drawing.Point(112, 64)
		Me._optThemeGames_7.TabIndex = 109
		Me.ToolTip1.SetToolTip(Me._optThemeGames_7, "South Pacific: The Search for Atlantis")
		Me._optThemeGames_7.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optThemeGames_7.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optThemeGames_7.BackColor = System.Drawing.SystemColors.Control
		Me._optThemeGames_7.CausesValidation = True
		Me._optThemeGames_7.Enabled = True
		Me._optThemeGames_7.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optThemeGames_7.Cursor = System.Windows.Forms.Cursors.Default
		Me._optThemeGames_7.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optThemeGames_7.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optThemeGames_7.TabStop = True
		Me._optThemeGames_7.Checked = False
		Me._optThemeGames_7.Visible = True
		Me._optThemeGames_7.Name = "_optThemeGames_7"
		Me._optThemeGames_6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optThemeGames_6.Text = "SP: Search for Atlantis"
		Me._optThemeGames_6.Size = New System.Drawing.Size(137, 17)
		Me._optThemeGames_6.Location = New System.Drawing.Point(112, 40)
		Me._optThemeGames_6.TabIndex = 105
		Me.ToolTip1.SetToolTip(Me._optThemeGames_6, "South Pacific: The Search for Atlantis")
		Me._optThemeGames_6.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optThemeGames_6.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optThemeGames_6.BackColor = System.Drawing.SystemColors.Control
		Me._optThemeGames_6.CausesValidation = True
		Me._optThemeGames_6.Enabled = True
		Me._optThemeGames_6.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optThemeGames_6.Cursor = System.Windows.Forms.Cursors.Default
		Me._optThemeGames_6.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optThemeGames_6.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optThemeGames_6.TabStop = True
		Me._optThemeGames_6.Checked = False
		Me._optThemeGames_6.Visible = True
		Me._optThemeGames_6.Name = "_optThemeGames_6"
		Me._optThemeGames_5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optThemeGames_5.Text = "SP 2005"
		Me._optThemeGames_5.Size = New System.Drawing.Size(73, 17)
		Me._optThemeGames_5.Location = New System.Drawing.Point(112, 16)
		Me._optThemeGames_5.TabIndex = 103
		Me.ToolTip1.SetToolTip(Me._optThemeGames_5, "South Pacific 2005")
		Me._optThemeGames_5.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optThemeGames_5.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optThemeGames_5.BackColor = System.Drawing.SystemColors.Control
		Me._optThemeGames_5.CausesValidation = True
		Me._optThemeGames_5.Enabled = True
		Me._optThemeGames_5.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optThemeGames_5.Cursor = System.Windows.Forms.Cursors.Default
		Me._optThemeGames_5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optThemeGames_5.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optThemeGames_5.TabStop = True
		Me._optThemeGames_5.Checked = False
		Me._optThemeGames_5.Visible = True
		Me._optThemeGames_5.Name = "_optThemeGames_5"
		Me._optThemeGames_4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optThemeGames_4.Text = "EE9"
		Me._optThemeGames_4.Size = New System.Drawing.Size(97, 17)
		Me._optThemeGames_4.Location = New System.Drawing.Point(16, 112)
		Me._optThemeGames_4.TabIndex = 102
		Me._optThemeGames_4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optThemeGames_4.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optThemeGames_4.BackColor = System.Drawing.SystemColors.Control
		Me._optThemeGames_4.CausesValidation = True
		Me._optThemeGames_4.Enabled = True
		Me._optThemeGames_4.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optThemeGames_4.Cursor = System.Windows.Forms.Cursors.Default
		Me._optThemeGames_4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optThemeGames_4.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optThemeGames_4.TabStop = True
		Me._optThemeGames_4.Checked = False
		Me._optThemeGames_4.Visible = True
		Me._optThemeGames_4.Name = "_optThemeGames_4"
		Me._optThemeGames_3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optThemeGames_3.Text = "2K4"
		Me._optThemeGames_3.Size = New System.Drawing.Size(97, 17)
		Me._optThemeGames_3.Location = New System.Drawing.Point(16, 88)
		Me._optThemeGames_3.TabIndex = 64
		Me._optThemeGames_3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optThemeGames_3.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optThemeGames_3.BackColor = System.Drawing.SystemColors.Control
		Me._optThemeGames_3.CausesValidation = True
		Me._optThemeGames_3.Enabled = True
		Me._optThemeGames_3.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optThemeGames_3.Cursor = System.Windows.Forms.Cursors.Default
		Me._optThemeGames_3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optThemeGames_3.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optThemeGames_3.TabStop = True
		Me._optThemeGames_3.Checked = False
		Me._optThemeGames_3.Visible = True
		Me._optThemeGames_3.Name = "_optThemeGames_3"
		Me._optThemeGames_2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optThemeGames_2.Text = "Retro"
		Me._optThemeGames_2.Size = New System.Drawing.Size(97, 17)
		Me._optThemeGames_2.Location = New System.Drawing.Point(16, 64)
		Me._optThemeGames_2.TabIndex = 55
		Me._optThemeGames_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optThemeGames_2.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optThemeGames_2.BackColor = System.Drawing.SystemColors.Control
		Me._optThemeGames_2.CausesValidation = True
		Me._optThemeGames_2.Enabled = True
		Me._optThemeGames_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optThemeGames_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._optThemeGames_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optThemeGames_2.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optThemeGames_2.TabStop = True
		Me._optThemeGames_2.Checked = False
		Me._optThemeGames_2.Visible = True
		Me._optThemeGames_2.Name = "_optThemeGames_2"
		Me._optThemeGames_1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optThemeGames_1.Text = "St@r W@rs"
		Me._optThemeGames_1.Size = New System.Drawing.Size(97, 17)
		Me._optThemeGames_1.Location = New System.Drawing.Point(16, 40)
		Me._optThemeGames_1.TabIndex = 49
		Me._optThemeGames_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optThemeGames_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optThemeGames_1.BackColor = System.Drawing.SystemColors.Control
		Me._optThemeGames_1.CausesValidation = True
		Me._optThemeGames_1.Enabled = True
		Me._optThemeGames_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optThemeGames_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._optThemeGames_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optThemeGames_1.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optThemeGames_1.TabStop = True
		Me._optThemeGames_1.Checked = False
		Me._optThemeGames_1.Visible = True
		Me._optThemeGames_1.Name = "_optThemeGames_1"
		Me._optThemeGames_0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optThemeGames_0.Text = "None"
		Me._optThemeGames_0.Size = New System.Drawing.Size(97, 17)
		Me._optThemeGames_0.Location = New System.Drawing.Point(16, 16)
		Me._optThemeGames_0.TabIndex = 48
		Me._optThemeGames_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._optThemeGames_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._optThemeGames_0.BackColor = System.Drawing.SystemColors.Control
		Me._optThemeGames_0.CausesValidation = True
		Me._optThemeGames_0.Enabled = True
		Me._optThemeGames_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._optThemeGames_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._optThemeGames_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._optThemeGames_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._optThemeGames_0.TabStop = True
		Me._optThemeGames_0.Checked = False
		Me._optThemeGames_0.Visible = True
		Me._optThemeGames_0.Name = "_optThemeGames_0"
		Me.picScriptButton.Size = New System.Drawing.Size(379, 148)
		Me.picScriptButton.Location = New System.Drawing.Point(16, 32)
		Me.picScriptButton.TabIndex = 66
		Me.picScriptButton.TabStop = False
		Me.picScriptButton.Visible = False
		Me.picScriptButton.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picScriptButton.Dock = System.Windows.Forms.DockStyle.None
		Me.picScriptButton.BackColor = System.Drawing.SystemColors.Control
		Me.picScriptButton.CausesValidation = True
		Me.picScriptButton.Enabled = True
		Me.picScriptButton.ForeColor = System.Drawing.SystemColors.ControlText
		Me.picScriptButton.Cursor = System.Windows.Forms.Cursors.Default
		Me.picScriptButton.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picScriptButton.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picScriptButton.Name = "picScriptButton"
		Me.optCustomScript.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optCustomScript.Text = "Custom Script"
		Me.optCustomScript.Size = New System.Drawing.Size(65, 29)
		Me.optCustomScript.Location = New System.Drawing.Point(312, 4)
		Me.optCustomScript.TabIndex = 107
		Me.optCustomScript.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optCustomScript.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optCustomScript.BackColor = System.Drawing.SystemColors.Control
		Me.optCustomScript.CausesValidation = True
		Me.optCustomScript.Enabled = True
		Me.optCustomScript.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optCustomScript.Cursor = System.Windows.Forms.Cursors.Default
		Me.optCustomScript.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optCustomScript.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optCustomScript.TabStop = True
		Me.optCustomScript.Checked = False
		Me.optCustomScript.Visible = True
		Me.optCustomScript.Name = "optCustomScript"
		Me.optJumpPoint.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optJumpPoint.Text = "Jump Point"
		Me.optJumpPoint.Size = New System.Drawing.Size(49, 29)
		Me.optJumpPoint.Location = New System.Drawing.Point(256, 4)
		Me.optJumpPoint.TabIndex = 106
		Me.optJumpPoint.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optJumpPoint.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optJumpPoint.BackColor = System.Drawing.SystemColors.Control
		Me.optJumpPoint.CausesValidation = True
		Me.optJumpPoint.Enabled = True
		Me.optJumpPoint.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optJumpPoint.Cursor = System.Windows.Forms.Cursors.Default
		Me.optJumpPoint.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optJumpPoint.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optJumpPoint.TabStop = True
		Me.optJumpPoint.Checked = False
		Me.optJumpPoint.Visible = True
		Me.optJumpPoint.Name = "optJumpPoint"
		Me.fraScriptButton.Text = "Custom Script Button"
		Me.fraScriptButton.Size = New System.Drawing.Size(377, 105)
		Me.fraScriptButton.Location = New System.Drawing.Point(0, 40)
		Me.fraScriptButton.TabIndex = 68
		Me.fraScriptButton.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraScriptButton.BackColor = System.Drawing.SystemColors.Control
		Me.fraScriptButton.Enabled = True
		Me.fraScriptButton.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraScriptButton.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraScriptButton.Visible = True
		Me.fraScriptButton.Padding = New System.Windows.Forms.Padding(0)
		Me.fraScriptButton.Name = "fraScriptButton"
		Me.cmbJumpPoint.Size = New System.Drawing.Size(97, 21)
		Me.cmbJumpPoint.Location = New System.Drawing.Point(96, 16)
		Me.cmbJumpPoint.Items.AddRange(New Object(){"Jump Point 1", "Jump Point 2", "Jump Point 3", "Jump Point 4", "Jump Point 5"})
		Me.cmbJumpPoint.TabIndex = 108
		Me.cmbJumpPoint.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbJumpPoint.BackColor = System.Drawing.SystemColors.Window
		Me.cmbJumpPoint.CausesValidation = True
		Me.cmbJumpPoint.Enabled = True
		Me.cmbJumpPoint.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbJumpPoint.IntegralHeight = True
		Me.cmbJumpPoint.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbJumpPoint.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbJumpPoint.Sorted = False
		Me.cmbJumpPoint.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cmbJumpPoint.TabStop = True
		Me.cmbJumpPoint.Visible = True
		Me.cmbJumpPoint.Name = "cmbJumpPoint"
		Me.cmdScriptImageBrowse.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdScriptImageBrowse.Text = "Image Browse"
		Me.cmdScriptImageBrowse.Size = New System.Drawing.Size(89, 25)
		Me.cmdScriptImageBrowse.Location = New System.Drawing.Point(280, 72)
		Me.cmdScriptImageBrowse.TabIndex = 75
		Me.cmdScriptImageBrowse.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdScriptImageBrowse.BackColor = System.Drawing.SystemColors.Control
		Me.cmdScriptImageBrowse.CausesValidation = True
		Me.cmdScriptImageBrowse.Enabled = True
		Me.cmdScriptImageBrowse.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdScriptImageBrowse.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdScriptImageBrowse.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdScriptImageBrowse.TabStop = True
		Me.cmdScriptImageBrowse.Name = "cmdScriptImageBrowse"
		Me.txtScriptImageFileName.AutoSize = False
		Me.txtScriptImageFileName.Size = New System.Drawing.Size(193, 19)
		Me.txtScriptImageFileName.Location = New System.Drawing.Point(80, 72)
		Me.txtScriptImageFileName.TabIndex = 74
		Me.txtScriptImageFileName.Text = "Text1"
		Me.txtScriptImageFileName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtScriptImageFileName.AcceptsReturn = True
		Me.txtScriptImageFileName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtScriptImageFileName.BackColor = System.Drawing.SystemColors.Window
		Me.txtScriptImageFileName.CausesValidation = True
		Me.txtScriptImageFileName.Enabled = True
		Me.txtScriptImageFileName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtScriptImageFileName.HideSelection = True
		Me.txtScriptImageFileName.ReadOnly = False
		Me.txtScriptImageFileName.Maxlength = 0
		Me.txtScriptImageFileName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtScriptImageFileName.MultiLine = False
		Me.txtScriptImageFileName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtScriptImageFileName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtScriptImageFileName.TabStop = True
		Me.txtScriptImageFileName.Visible = True
		Me.txtScriptImageFileName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtScriptImageFileName.Name = "txtScriptImageFileName"
		Me.chkScriptReport.Text = "Send Output to a Report Window"
		Me.chkScriptReport.Size = New System.Drawing.Size(185, 17)
		Me.chkScriptReport.Location = New System.Drawing.Point(16, 40)
		Me.chkScriptReport.TabIndex = 72
		Me.chkScriptReport.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkScriptReport.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkScriptReport.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkScriptReport.BackColor = System.Drawing.SystemColors.Control
		Me.chkScriptReport.CausesValidation = True
		Me.chkScriptReport.Enabled = True
		Me.chkScriptReport.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkScriptReport.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkScriptReport.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkScriptReport.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkScriptReport.TabStop = True
		Me.chkScriptReport.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkScriptReport.Visible = True
		Me.chkScriptReport.Name = "chkScriptReport"
		Me.txtScriptFileName.AutoSize = False
		Me.txtScriptFileName.Size = New System.Drawing.Size(273, 19)
		Me.txtScriptFileName.Location = New System.Drawing.Point(96, 16)
		Me.txtScriptFileName.TabIndex = 70
		Me.txtScriptFileName.Text = "Text1"
		Me.txtScriptFileName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtScriptFileName.AcceptsReturn = True
		Me.txtScriptFileName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtScriptFileName.BackColor = System.Drawing.SystemColors.Window
		Me.txtScriptFileName.CausesValidation = True
		Me.txtScriptFileName.Enabled = True
		Me.txtScriptFileName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtScriptFileName.HideSelection = True
		Me.txtScriptFileName.ReadOnly = False
		Me.txtScriptFileName.Maxlength = 0
		Me.txtScriptFileName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtScriptFileName.MultiLine = False
		Me.txtScriptFileName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtScriptFileName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtScriptFileName.TabStop = True
		Me.txtScriptFileName.Visible = True
		Me.txtScriptFileName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtScriptFileName.Name = "txtScriptFileName"
		Me.cmdScriptBrowse.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdScriptBrowse.Text = "File Name Browse"
		Me.cmdScriptBrowse.Size = New System.Drawing.Size(105, 25)
		Me.cmdScriptBrowse.Location = New System.Drawing.Point(264, 40)
		Me.cmdScriptBrowse.TabIndex = 69
		Me.cmdScriptBrowse.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdScriptBrowse.BackColor = System.Drawing.SystemColors.Control
		Me.cmdScriptBrowse.CausesValidation = True
		Me.cmdScriptBrowse.Enabled = True
		Me.cmdScriptBrowse.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdScriptBrowse.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdScriptBrowse.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdScriptBrowse.TabStop = True
		Me.cmdScriptBrowse.Name = "cmdScriptBrowse"
		Me.lblScriptImage.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblScriptImage.Text = "Button Image"
		Me.lblScriptImage.Size = New System.Drawing.Size(65, 17)
		Me.lblScriptImage.Location = New System.Drawing.Point(8, 72)
		Me.lblScriptImage.TabIndex = 73
		Me.lblScriptImage.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblScriptImage.BackColor = System.Drawing.SystemColors.Control
		Me.lblScriptImage.Enabled = True
		Me.lblScriptImage.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblScriptImage.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblScriptImage.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblScriptImage.UseMnemonic = True
		Me.lblScriptImage.Visible = True
		Me.lblScriptImage.AutoSize = False
		Me.lblScriptImage.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblScriptImage.Name = "lblScriptImage"
		Me.lblScriptFileName.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblScriptFileName.Text = "Scrit File Name"
		Me.lblScriptFileName.Size = New System.Drawing.Size(81, 17)
		Me.lblScriptFileName.Location = New System.Drawing.Point(8, 16)
		Me.lblScriptFileName.TabIndex = 71
		Me.lblScriptFileName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblScriptFileName.BackColor = System.Drawing.SystemColors.Control
		Me.lblScriptFileName.Enabled = True
		Me.lblScriptFileName.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblScriptFileName.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblScriptFileName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblScriptFileName.UseMnemonic = True
		Me.lblScriptFileName.Visible = True
		Me.lblScriptFileName.AutoSize = False
		Me.lblScriptFileName.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblScriptFileName.Name = "lblScriptFileName"
		Me.cmbScriptButton.Size = New System.Drawing.Size(249, 21)
		Me.cmbScriptButton.Location = New System.Drawing.Point(0, 8)
		Me.cmbScriptButton.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cmbScriptButton.TabIndex = 67
		Me.cmbScriptButton.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbScriptButton.BackColor = System.Drawing.SystemColors.Window
		Me.cmbScriptButton.CausesValidation = True
		Me.cmbScriptButton.Enabled = True
		Me.cmbScriptButton.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbScriptButton.IntegralHeight = True
		Me.cmbScriptButton.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbScriptButton.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbScriptButton.Sorted = False
		Me.cmbScriptButton.TabStop = True
		Me.cmbScriptButton.Visible = True
		Me.cmbScriptButton.Name = "cmbScriptButton"
		Me.picItem.Size = New System.Drawing.Size(379, 148)
		Me.picItem.Location = New System.Drawing.Point(16, 32)
		Me.picItem.TabIndex = 27
		Me.picItem.TabStop = False
		Me.picItem.Visible = False
		Me.picItem.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picItem.Dock = System.Windows.Forms.DockStyle.None
		Me.picItem.BackColor = System.Drawing.SystemColors.Control
		Me.picItem.CausesValidation = True
		Me.picItem.Enabled = True
		Me.picItem.ForeColor = System.Drawing.SystemColors.ControlText
		Me.picItem.Cursor = System.Windows.Forms.Cursors.Default
		Me.picItem.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picItem.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picItem.Name = "picItem"
		Me.frmItemLabels.Text = "Labels"
		Me.frmItemLabels.Size = New System.Drawing.Size(177, 89)
		Me.frmItemLabels.Location = New System.Drawing.Point(8, 40)
		Me.frmItemLabels.TabIndex = 38
		Me.frmItemLabels.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmItemLabels.BackColor = System.Drawing.SystemColors.Control
		Me.frmItemLabels.Enabled = True
		Me.frmItemLabels.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmItemLabels.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmItemLabels.Visible = True
		Me.frmItemLabels.Padding = New System.Windows.Forms.Padding(0)
		Me.frmItemLabels.Name = "frmItemLabels"
		Me.txtItemDistributionPanelLabel.AutoSize = False
		Me.txtItemDistributionPanelLabel.Size = New System.Drawing.Size(73, 19)
		Me.txtItemDistributionPanelLabel.Location = New System.Drawing.Point(96, 64)
		Me.txtItemDistributionPanelLabel.TabIndex = 43
		Me.txtItemDistributionPanelLabel.Text = "Text1"
		Me.txtItemDistributionPanelLabel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtItemDistributionPanelLabel.AcceptsReturn = True
		Me.txtItemDistributionPanelLabel.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtItemDistributionPanelLabel.BackColor = System.Drawing.SystemColors.Window
		Me.txtItemDistributionPanelLabel.CausesValidation = True
		Me.txtItemDistributionPanelLabel.Enabled = True
		Me.txtItemDistributionPanelLabel.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtItemDistributionPanelLabel.HideSelection = True
		Me.txtItemDistributionPanelLabel.ReadOnly = False
		Me.txtItemDistributionPanelLabel.Maxlength = 0
		Me.txtItemDistributionPanelLabel.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtItemDistributionPanelLabel.MultiLine = False
		Me.txtItemDistributionPanelLabel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtItemDistributionPanelLabel.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtItemDistributionPanelLabel.TabStop = True
		Me.txtItemDistributionPanelLabel.Visible = True
		Me.txtItemDistributionPanelLabel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtItemDistributionPanelLabel.Name = "txtItemDistributionPanelLabel"
		Me.txtItemSectorPanelLabel.AutoSize = False
		Me.txtItemSectorPanelLabel.Size = New System.Drawing.Size(73, 19)
		Me.txtItemSectorPanelLabel.Location = New System.Drawing.Point(96, 40)
		Me.txtItemSectorPanelLabel.TabIndex = 41
		Me.txtItemSectorPanelLabel.Text = "Text1"
		Me.txtItemSectorPanelLabel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtItemSectorPanelLabel.AcceptsReturn = True
		Me.txtItemSectorPanelLabel.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtItemSectorPanelLabel.BackColor = System.Drawing.SystemColors.Window
		Me.txtItemSectorPanelLabel.CausesValidation = True
		Me.txtItemSectorPanelLabel.Enabled = True
		Me.txtItemSectorPanelLabel.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtItemSectorPanelLabel.HideSelection = True
		Me.txtItemSectorPanelLabel.ReadOnly = False
		Me.txtItemSectorPanelLabel.Maxlength = 0
		Me.txtItemSectorPanelLabel.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtItemSectorPanelLabel.MultiLine = False
		Me.txtItemSectorPanelLabel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtItemSectorPanelLabel.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtItemSectorPanelLabel.TabStop = True
		Me.txtItemSectorPanelLabel.Visible = True
		Me.txtItemSectorPanelLabel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtItemSectorPanelLabel.Name = "txtItemSectorPanelLabel"
		Me.txtItemFormLabel.AutoSize = False
		Me.txtItemFormLabel.Size = New System.Drawing.Size(73, 19)
		Me.txtItemFormLabel.Location = New System.Drawing.Point(96, 16)
		Me.txtItemFormLabel.TabIndex = 39
		Me.txtItemFormLabel.Text = "Text1"
		Me.txtItemFormLabel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtItemFormLabel.AcceptsReturn = True
		Me.txtItemFormLabel.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtItemFormLabel.BackColor = System.Drawing.SystemColors.Window
		Me.txtItemFormLabel.CausesValidation = True
		Me.txtItemFormLabel.Enabled = True
		Me.txtItemFormLabel.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtItemFormLabel.HideSelection = True
		Me.txtItemFormLabel.ReadOnly = False
		Me.txtItemFormLabel.Maxlength = 0
		Me.txtItemFormLabel.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtItemFormLabel.MultiLine = False
		Me.txtItemFormLabel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtItemFormLabel.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtItemFormLabel.TabStop = True
		Me.txtItemFormLabel.Visible = True
		Me.txtItemFormLabel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtItemFormLabel.Name = "txtItemFormLabel"
		Me.lblItemDistributionPanelLabel.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblItemDistributionPanelLabel.Text = "Distibution Panel"
		Me.lblItemDistributionPanelLabel.Size = New System.Drawing.Size(81, 17)
		Me.lblItemDistributionPanelLabel.Location = New System.Drawing.Point(8, 64)
		Me.lblItemDistributionPanelLabel.TabIndex = 44
		Me.lblItemDistributionPanelLabel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblItemDistributionPanelLabel.BackColor = System.Drawing.SystemColors.Control
		Me.lblItemDistributionPanelLabel.Enabled = True
		Me.lblItemDistributionPanelLabel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblItemDistributionPanelLabel.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblItemDistributionPanelLabel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblItemDistributionPanelLabel.UseMnemonic = True
		Me.lblItemDistributionPanelLabel.Visible = True
		Me.lblItemDistributionPanelLabel.AutoSize = False
		Me.lblItemDistributionPanelLabel.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblItemDistributionPanelLabel.Name = "lblItemDistributionPanelLabel"
		Me.lblItemSectorPanelLabel.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblItemSectorPanelLabel.Text = "Sector Panel"
		Me.lblItemSectorPanelLabel.Size = New System.Drawing.Size(65, 17)
		Me.lblItemSectorPanelLabel.Location = New System.Drawing.Point(24, 40)
		Me.lblItemSectorPanelLabel.TabIndex = 42
		Me.lblItemSectorPanelLabel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblItemSectorPanelLabel.BackColor = System.Drawing.SystemColors.Control
		Me.lblItemSectorPanelLabel.Enabled = True
		Me.lblItemSectorPanelLabel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblItemSectorPanelLabel.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblItemSectorPanelLabel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblItemSectorPanelLabel.UseMnemonic = True
		Me.lblItemSectorPanelLabel.Visible = True
		Me.lblItemSectorPanelLabel.AutoSize = False
		Me.lblItemSectorPanelLabel.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblItemSectorPanelLabel.Name = "lblItemSectorPanelLabel"
		Me.lblItemFormLabel.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblItemFormLabel.Text = "Form"
		Me.lblItemFormLabel.Size = New System.Drawing.Size(49, 17)
		Me.lblItemFormLabel.Location = New System.Drawing.Point(40, 16)
		Me.lblItemFormLabel.TabIndex = 40
		Me.lblItemFormLabel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblItemFormLabel.BackColor = System.Drawing.SystemColors.Control
		Me.lblItemFormLabel.Enabled = True
		Me.lblItemFormLabel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblItemFormLabel.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblItemFormLabel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblItemFormLabel.UseMnemonic = True
		Me.lblItemFormLabel.Visible = True
		Me.lblItemFormLabel.AutoSize = False
		Me.lblItemFormLabel.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblItemFormLabel.Name = "lblItemFormLabel"
		Me.cbItem.Size = New System.Drawing.Size(177, 21)
		Me.cbItem.Location = New System.Drawing.Point(8, 8)
		Me.cbItem.TabIndex = 29
		Me.cbItem.Text = "Combo1"
		Me.cbItem.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbItem.BackColor = System.Drawing.SystemColors.Window
		Me.cbItem.CausesValidation = True
		Me.cbItem.Enabled = True
		Me.cbItem.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cbItem.IntegralHeight = True
		Me.cbItem.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbItem.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbItem.Sorted = False
		Me.cbItem.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cbItem.TabStop = True
		Me.cbItem.Visible = True
		Me.cbItem.Name = "cbItem"
		Me.frmItem.Text = "Names"
		Me.frmItem.Size = New System.Drawing.Size(161, 119)
		Me.frmItem.Location = New System.Drawing.Point(208, 8)
		Me.frmItem.TabIndex = 28
		Me.frmItem.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.frmItem.BackColor = System.Drawing.SystemColors.Control
		Me.frmItem.Enabled = True
		Me.frmItem.ForeColor = System.Drawing.SystemColors.ControlText
		Me.frmItem.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.frmItem.Visible = True
		Me.frmItem.Padding = New System.Windows.Forms.Padding(0)
		Me.frmItem.Name = "frmItem"
		Me.txtItemConditionalName.AutoSize = False
		Me.txtItemConditionalName.Size = New System.Drawing.Size(73, 19)
		Me.txtItemConditionalName.Location = New System.Drawing.Point(80, 64)
		Me.txtItemConditionalName.TabIndex = 36
		Me.txtItemConditionalName.Text = "Text1"
		Me.txtItemConditionalName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtItemConditionalName.AcceptsReturn = True
		Me.txtItemConditionalName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtItemConditionalName.BackColor = System.Drawing.SystemColors.Window
		Me.txtItemConditionalName.CausesValidation = True
		Me.txtItemConditionalName.Enabled = True
		Me.txtItemConditionalName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtItemConditionalName.HideSelection = True
		Me.txtItemConditionalName.ReadOnly = False
		Me.txtItemConditionalName.Maxlength = 0
		Me.txtItemConditionalName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtItemConditionalName.MultiLine = False
		Me.txtItemConditionalName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtItemConditionalName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtItemConditionalName.TabStop = True
		Me.txtItemConditionalName.Visible = True
		Me.txtItemConditionalName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtItemConditionalName.Name = "txtItemConditionalName"
		Me.txtItemIntelligenceNames.AutoSize = False
		Me.txtItemIntelligenceNames.Size = New System.Drawing.Size(73, 19)
		Me.txtItemIntelligenceNames.Location = New System.Drawing.Point(80, 88)
		Me.txtItemIntelligenceNames.TabIndex = 34
		Me.txtItemIntelligenceNames.Text = "Text1"
		Me.txtItemIntelligenceNames.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtItemIntelligenceNames.AcceptsReturn = True
		Me.txtItemIntelligenceNames.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtItemIntelligenceNames.BackColor = System.Drawing.SystemColors.Window
		Me.txtItemIntelligenceNames.CausesValidation = True
		Me.txtItemIntelligenceNames.Enabled = True
		Me.txtItemIntelligenceNames.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtItemIntelligenceNames.HideSelection = True
		Me.txtItemIntelligenceNames.ReadOnly = False
		Me.txtItemIntelligenceNames.Maxlength = 0
		Me.txtItemIntelligenceNames.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtItemIntelligenceNames.MultiLine = False
		Me.txtItemIntelligenceNames.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtItemIntelligenceNames.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtItemIntelligenceNames.TabStop = True
		Me.txtItemIntelligenceNames.Visible = True
		Me.txtItemIntelligenceNames.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtItemIntelligenceNames.Name = "txtItemIntelligenceNames"
		Me.txtItemDBName.AutoSize = False
		Me.txtItemDBName.Size = New System.Drawing.Size(73, 19)
		Me.txtItemDBName.Location = New System.Drawing.Point(80, 40)
		Me.txtItemDBName.TabIndex = 32
		Me.txtItemDBName.Text = "Text1"
		Me.txtItemDBName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtItemDBName.AcceptsReturn = True
		Me.txtItemDBName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtItemDBName.BackColor = System.Drawing.SystemColors.Window
		Me.txtItemDBName.CausesValidation = True
		Me.txtItemDBName.Enabled = True
		Me.txtItemDBName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtItemDBName.HideSelection = True
		Me.txtItemDBName.ReadOnly = False
		Me.txtItemDBName.Maxlength = 0
		Me.txtItemDBName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtItemDBName.MultiLine = False
		Me.txtItemDBName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtItemDBName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtItemDBName.TabStop = True
		Me.txtItemDBName.Visible = True
		Me.txtItemDBName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtItemDBName.Name = "txtItemDBName"
		Me.txtItemFormName.AutoSize = False
		Me.txtItemFormName.Size = New System.Drawing.Size(73, 19)
		Me.txtItemFormName.Location = New System.Drawing.Point(80, 16)
		Me.txtItemFormName.TabIndex = 30
		Me.txtItemFormName.Text = "Text1"
		Me.txtItemFormName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtItemFormName.AcceptsReturn = True
		Me.txtItemFormName.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtItemFormName.BackColor = System.Drawing.SystemColors.Window
		Me.txtItemFormName.CausesValidation = True
		Me.txtItemFormName.Enabled = True
		Me.txtItemFormName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtItemFormName.HideSelection = True
		Me.txtItemFormName.ReadOnly = False
		Me.txtItemFormName.Maxlength = 0
		Me.txtItemFormName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtItemFormName.MultiLine = False
		Me.txtItemFormName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtItemFormName.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtItemFormName.TabStop = True
		Me.txtItemFormName.Visible = True
		Me.txtItemFormName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtItemFormName.Name = "txtItemFormName"
		Me.lblItemConditionalName.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblItemConditionalName.Text = "Conditional"
		Me.lblItemConditionalName.Size = New System.Drawing.Size(57, 17)
		Me.lblItemConditionalName.Location = New System.Drawing.Point(16, 64)
		Me.lblItemConditionalName.TabIndex = 37
		Me.lblItemConditionalName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblItemConditionalName.BackColor = System.Drawing.SystemColors.Control
		Me.lblItemConditionalName.Enabled = True
		Me.lblItemConditionalName.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblItemConditionalName.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblItemConditionalName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblItemConditionalName.UseMnemonic = True
		Me.lblItemConditionalName.Visible = True
		Me.lblItemConditionalName.AutoSize = False
		Me.lblItemConditionalName.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblItemConditionalName.Name = "lblItemConditionalName"
		Me.lblItemIntelligenceNames.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblItemIntelligenceNames.Text = "Intelligence"
		Me.lblItemIntelligenceNames.Size = New System.Drawing.Size(57, 17)
		Me.lblItemIntelligenceNames.Location = New System.Drawing.Point(16, 88)
		Me.lblItemIntelligenceNames.TabIndex = 35
		Me.lblItemIntelligenceNames.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblItemIntelligenceNames.BackColor = System.Drawing.SystemColors.Control
		Me.lblItemIntelligenceNames.Enabled = True
		Me.lblItemIntelligenceNames.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblItemIntelligenceNames.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblItemIntelligenceNames.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblItemIntelligenceNames.UseMnemonic = True
		Me.lblItemIntelligenceNames.Visible = True
		Me.lblItemIntelligenceNames.AutoSize = False
		Me.lblItemIntelligenceNames.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblItemIntelligenceNames.Name = "lblItemIntelligenceNames"
		Me.lblItemDBName.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblItemDBName.Text = "Database"
		Me.lblItemDBName.Size = New System.Drawing.Size(49, 17)
		Me.lblItemDBName.Location = New System.Drawing.Point(24, 40)
		Me.lblItemDBName.TabIndex = 33
		Me.lblItemDBName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblItemDBName.BackColor = System.Drawing.SystemColors.Control
		Me.lblItemDBName.Enabled = True
		Me.lblItemDBName.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblItemDBName.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblItemDBName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblItemDBName.UseMnemonic = True
		Me.lblItemDBName.Visible = True
		Me.lblItemDBName.AutoSize = False
		Me.lblItemDBName.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblItemDBName.Name = "lblItemDBName"
		Me.lblItemFormName.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblItemFormName.Text = "Form"
		Me.lblItemFormName.Size = New System.Drawing.Size(25, 17)
		Me.lblItemFormName.Location = New System.Drawing.Point(48, 16)
		Me.lblItemFormName.TabIndex = 31
		Me.lblItemFormName.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblItemFormName.BackColor = System.Drawing.SystemColors.Control
		Me.lblItemFormName.Enabled = True
		Me.lblItemFormName.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblItemFormName.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblItemFormName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblItemFormName.UseMnemonic = True
		Me.lblItemFormName.Visible = True
		Me.lblItemFormName.AutoSize = False
		Me.lblItemFormName.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblItemFormName.Name = "lblItemFormName"
		Me.picTelegram.Size = New System.Drawing.Size(379, 148)
		Me.picTelegram.Location = New System.Drawing.Point(16, 32)
		Me.picTelegram.TabIndex = 79
		Me.picTelegram.TabStop = False
		Me.picTelegram.Visible = False
		Me.picTelegram.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picTelegram.Dock = System.Windows.Forms.DockStyle.None
		Me.picTelegram.BackColor = System.Drawing.SystemColors.Control
		Me.picTelegram.CausesValidation = True
		Me.picTelegram.Enabled = True
		Me.picTelegram.ForeColor = System.Drawing.SystemColors.ControlText
		Me.picTelegram.Cursor = System.Windows.Forms.Cursors.Default
		Me.picTelegram.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picTelegram.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picTelegram.Name = "picTelegram"
		Me._fraTelegram_2.Text = "Hard Limit"
		Me._fraTelegram_2.Size = New System.Drawing.Size(121, 117)
		Me._fraTelegram_2.Location = New System.Drawing.Point(256, 32)
		Me._fraTelegram_2.TabIndex = 83
		Me.ToolTip1.SetToolTip(Me._fraTelegram_2, "Set the column where automatic line where new line is added")
		Me._fraTelegram_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._fraTelegram_2.BackColor = System.Drawing.SystemColors.Control
		Me._fraTelegram_2.Enabled = True
		Me._fraTelegram_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._fraTelegram_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._fraTelegram_2.Visible = True
		Me._fraTelegram_2.Padding = New System.Windows.Forms.Padding(0)
		Me._fraTelegram_2.Name = "_fraTelegram_2"
		Me._cmbTelegramSound_2.Size = New System.Drawing.Size(105, 21)
		Me._cmbTelegramSound_2.Location = New System.Drawing.Point(8, 80)
		Me._cmbTelegramSound_2.Items.AddRange(New Object(){"No Sound", "1 Beep", "2 Beeps", "3 Beeps"})
		Me._cmbTelegramSound_2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me._cmbTelegramSound_2.TabIndex = 96
		Me._cmbTelegramSound_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmbTelegramSound_2.BackColor = System.Drawing.SystemColors.Window
		Me._cmbTelegramSound_2.CausesValidation = True
		Me._cmbTelegramSound_2.Enabled = True
		Me._cmbTelegramSound_2.ForeColor = System.Drawing.SystemColors.WindowText
		Me._cmbTelegramSound_2.IntegralHeight = True
		Me._cmbTelegramSound_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmbTelegramSound_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmbTelegramSound_2.Sorted = False
		Me._cmbTelegramSound_2.TabStop = True
		Me._cmbTelegramSound_2.Visible = True
		Me._cmbTelegramSound_2.Name = "_cmbTelegramSound_2"
		Me._txtTelegram_2.AutoSize = False
		Me._txtTelegram_2.Size = New System.Drawing.Size(57, 19)
		Me._txtTelegram_2.Location = New System.Drawing.Point(56, 48)
		Me._txtTelegram_2.TabIndex = 95
		Me._txtTelegram_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtTelegram_2.AcceptsReturn = True
		Me._txtTelegram_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtTelegram_2.BackColor = System.Drawing.SystemColors.Window
		Me._txtTelegram_2.CausesValidation = True
		Me._txtTelegram_2.Enabled = True
		Me._txtTelegram_2.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtTelegram_2.HideSelection = True
		Me._txtTelegram_2.ReadOnly = False
		Me._txtTelegram_2.Maxlength = 0
		Me._txtTelegram_2.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtTelegram_2.MultiLine = False
		Me._txtTelegram_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtTelegram_2.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtTelegram_2.TabStop = True
		Me._txtTelegram_2.Visible = True
		Me._txtTelegram_2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtTelegram_2.Name = "_txtTelegram_2"
		Me._chkTelegram_2.Text = "Enabled"
		Me._chkTelegram_2.Size = New System.Drawing.Size(65, 13)
		Me._chkTelegram_2.Location = New System.Drawing.Point(8, 24)
		Me._chkTelegram_2.TabIndex = 86
		Me._chkTelegram_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkTelegram_2.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkTelegram_2.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkTelegram_2.BackColor = System.Drawing.SystemColors.Control
		Me._chkTelegram_2.CausesValidation = True
		Me._chkTelegram_2.Enabled = True
		Me._chkTelegram_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkTelegram_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkTelegram_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkTelegram_2.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkTelegram_2.TabStop = True
		Me._chkTelegram_2.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._chkTelegram_2.Visible = True
		Me._chkTelegram_2.Name = "_chkTelegram_2"
		Me._lblTelegram_2.Text = "Column"
		Me._lblTelegram_2.Size = New System.Drawing.Size(41, 17)
		Me._lblTelegram_2.Location = New System.Drawing.Point(8, 48)
		Me._lblTelegram_2.TabIndex = 89
		Me._lblTelegram_2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblTelegram_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblTelegram_2.BackColor = System.Drawing.SystemColors.Control
		Me._lblTelegram_2.Enabled = True
		Me._lblTelegram_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblTelegram_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblTelegram_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblTelegram_2.UseMnemonic = True
		Me._lblTelegram_2.Visible = True
		Me._lblTelegram_2.AutoSize = False
		Me._lblTelegram_2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblTelegram_2.Name = "_lblTelegram_2"
		Me._fraTelegram_1.Text = "Soft Limit"
		Me._fraTelegram_1.Size = New System.Drawing.Size(121, 117)
		Me._fraTelegram_1.Location = New System.Drawing.Point(128, 32)
		Me._fraTelegram_1.TabIndex = 82
		Me.ToolTip1.SetToolTip(Me._fraTelegram_1, "Set the column where the automatic new line")
		Me._fraTelegram_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._fraTelegram_1.BackColor = System.Drawing.SystemColors.Control
		Me._fraTelegram_1.Enabled = True
		Me._fraTelegram_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._fraTelegram_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._fraTelegram_1.Visible = True
		Me._fraTelegram_1.Padding = New System.Windows.Forms.Padding(0)
		Me._fraTelegram_1.Name = "_fraTelegram_1"
		Me._cmbTelegramSound_1.Size = New System.Drawing.Size(105, 21)
		Me._cmbTelegramSound_1.Location = New System.Drawing.Point(8, 80)
		Me._cmbTelegramSound_1.Items.AddRange(New Object(){"No Sound", "1 Beep", "2 Beeps", "3 Beeps"})
		Me._cmbTelegramSound_1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me._cmbTelegramSound_1.TabIndex = 94
		Me._cmbTelegramSound_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmbTelegramSound_1.BackColor = System.Drawing.SystemColors.Window
		Me._cmbTelegramSound_1.CausesValidation = True
		Me._cmbTelegramSound_1.Enabled = True
		Me._cmbTelegramSound_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._cmbTelegramSound_1.IntegralHeight = True
		Me._cmbTelegramSound_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmbTelegramSound_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmbTelegramSound_1.Sorted = False
		Me._cmbTelegramSound_1.TabStop = True
		Me._cmbTelegramSound_1.Visible = True
		Me._cmbTelegramSound_1.Name = "_cmbTelegramSound_1"
		Me._txtTelegram_1.AutoSize = False
		Me._txtTelegram_1.Size = New System.Drawing.Size(57, 19)
		Me._txtTelegram_1.Location = New System.Drawing.Point(56, 48)
		Me._txtTelegram_1.TabIndex = 93
		Me._txtTelegram_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtTelegram_1.AcceptsReturn = True
		Me._txtTelegram_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtTelegram_1.BackColor = System.Drawing.SystemColors.Window
		Me._txtTelegram_1.CausesValidation = True
		Me._txtTelegram_1.Enabled = True
		Me._txtTelegram_1.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtTelegram_1.HideSelection = True
		Me._txtTelegram_1.ReadOnly = False
		Me._txtTelegram_1.Maxlength = 0
		Me._txtTelegram_1.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtTelegram_1.MultiLine = False
		Me._txtTelegram_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtTelegram_1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtTelegram_1.TabStop = True
		Me._txtTelegram_1.Visible = True
		Me._txtTelegram_1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtTelegram_1.Name = "_txtTelegram_1"
		Me.Text1.AutoSize = False
		Me.Text1.Size = New System.Drawing.Size(41, 19)
		Me.Text1.Location = New System.Drawing.Point(136, 8)
		Me.Text1.TabIndex = 91
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
		Me.Text1.MultiLine = False
		Me.Text1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Text1.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.Text1.TabStop = True
		Me.Text1.Visible = True
		Me.Text1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Text1.Name = "Text1"
		Me._chkTelegram_1.Text = "Enabled"
		Me._chkTelegram_1.Size = New System.Drawing.Size(65, 13)
		Me._chkTelegram_1.Location = New System.Drawing.Point(8, 24)
		Me._chkTelegram_1.TabIndex = 85
		Me._chkTelegram_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkTelegram_1.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkTelegram_1.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkTelegram_1.BackColor = System.Drawing.SystemColors.Control
		Me._chkTelegram_1.CausesValidation = True
		Me._chkTelegram_1.Enabled = True
		Me._chkTelegram_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkTelegram_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkTelegram_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkTelegram_1.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkTelegram_1.TabStop = True
		Me._chkTelegram_1.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._chkTelegram_1.Visible = True
		Me._chkTelegram_1.Name = "_chkTelegram_1"
		Me._lblTelegram_1.Text = "Column"
		Me._lblTelegram_1.Size = New System.Drawing.Size(41, 17)
		Me._lblTelegram_1.Location = New System.Drawing.Point(8, 48)
		Me._lblTelegram_1.TabIndex = 88
		Me._lblTelegram_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblTelegram_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblTelegram_1.BackColor = System.Drawing.SystemColors.Control
		Me._lblTelegram_1.Enabled = True
		Me._lblTelegram_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblTelegram_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblTelegram_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblTelegram_1.UseMnemonic = True
		Me._lblTelegram_1.Visible = True
		Me._lblTelegram_1.AutoSize = False
		Me._lblTelegram_1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblTelegram_1.Name = "_lblTelegram_1"
		Me.cmbTelegram.Size = New System.Drawing.Size(377, 21)
		Me.cmbTelegram.Location = New System.Drawing.Point(0, 8)
		Me.cmbTelegram.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cmbTelegram.TabIndex = 81
		Me.cmbTelegram.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmbTelegram.BackColor = System.Drawing.SystemColors.Window
		Me.cmbTelegram.CausesValidation = True
		Me.cmbTelegram.Enabled = True
		Me.cmbTelegram.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cmbTelegram.IntegralHeight = True
		Me.cmbTelegram.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmbTelegram.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmbTelegram.Sorted = False
		Me.cmbTelegram.TabStop = True
		Me.cmbTelegram.Visible = True
		Me.cmbTelegram.Name = "cmbTelegram"
		Me._fraTelegram_0.Text = "Warning"
		Me._fraTelegram_0.Size = New System.Drawing.Size(121, 117)
		Me._fraTelegram_0.Location = New System.Drawing.Point(0, 32)
		Me._fraTelegram_0.TabIndex = 80
		Me.ToolTip1.SetToolTip(Me._fraTelegram_0, "Set the column where the warning starts")
		Me._fraTelegram_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._fraTelegram_0.BackColor = System.Drawing.SystemColors.Control
		Me._fraTelegram_0.Enabled = True
		Me._fraTelegram_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._fraTelegram_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._fraTelegram_0.Visible = True
		Me._fraTelegram_0.Padding = New System.Windows.Forms.Padding(0)
		Me._fraTelegram_0.Name = "_fraTelegram_0"
		Me._cmbTelegramSound_0.Size = New System.Drawing.Size(105, 21)
		Me._cmbTelegramSound_0.Location = New System.Drawing.Point(8, 80)
		Me._cmbTelegramSound_0.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me._cmbTelegramSound_0.TabIndex = 92
		Me._cmbTelegramSound_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmbTelegramSound_0.BackColor = System.Drawing.SystemColors.Window
		Me._cmbTelegramSound_0.CausesValidation = True
		Me._cmbTelegramSound_0.Enabled = True
		Me._cmbTelegramSound_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._cmbTelegramSound_0.IntegralHeight = True
		Me._cmbTelegramSound_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmbTelegramSound_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmbTelegramSound_0.Sorted = False
		Me._cmbTelegramSound_0.TabStop = True
		Me._cmbTelegramSound_0.Visible = True
		Me._cmbTelegramSound_0.Name = "_cmbTelegramSound_0"
		Me._txtTelegram_0.AutoSize = False
		Me._txtTelegram_0.Size = New System.Drawing.Size(57, 19)
		Me._txtTelegram_0.Location = New System.Drawing.Point(56, 48)
		Me._txtTelegram_0.TabIndex = 90
		Me._txtTelegram_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._txtTelegram_0.AcceptsReturn = True
		Me._txtTelegram_0.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me._txtTelegram_0.BackColor = System.Drawing.SystemColors.Window
		Me._txtTelegram_0.CausesValidation = True
		Me._txtTelegram_0.Enabled = True
		Me._txtTelegram_0.ForeColor = System.Drawing.SystemColors.WindowText
		Me._txtTelegram_0.HideSelection = True
		Me._txtTelegram_0.ReadOnly = False
		Me._txtTelegram_0.Maxlength = 0
		Me._txtTelegram_0.Cursor = System.Windows.Forms.Cursors.IBeam
		Me._txtTelegram_0.MultiLine = False
		Me._txtTelegram_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._txtTelegram_0.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me._txtTelegram_0.TabStop = True
		Me._txtTelegram_0.Visible = True
		Me._txtTelegram_0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me._txtTelegram_0.Name = "_txtTelegram_0"
		Me._chkTelegram_0.Text = "Enabled"
		Me._chkTelegram_0.Size = New System.Drawing.Size(65, 13)
		Me._chkTelegram_0.Location = New System.Drawing.Point(8, 24)
		Me._chkTelegram_0.TabIndex = 84
		Me._chkTelegram_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._chkTelegram_0.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me._chkTelegram_0.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me._chkTelegram_0.BackColor = System.Drawing.SystemColors.Control
		Me._chkTelegram_0.CausesValidation = True
		Me._chkTelegram_0.Enabled = True
		Me._chkTelegram_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._chkTelegram_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._chkTelegram_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._chkTelegram_0.Appearance = System.Windows.Forms.Appearance.Normal
		Me._chkTelegram_0.TabStop = True
		Me._chkTelegram_0.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me._chkTelegram_0.Visible = True
		Me._chkTelegram_0.Name = "_chkTelegram_0"
		Me._lblTelegram_0.Text = "Column"
		Me._lblTelegram_0.Size = New System.Drawing.Size(41, 17)
		Me._lblTelegram_0.Location = New System.Drawing.Point(8, 48)
		Me._lblTelegram_0.TabIndex = 87
		Me._lblTelegram_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblTelegram_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me._lblTelegram_0.BackColor = System.Drawing.SystemColors.Control
		Me._lblTelegram_0.Enabled = True
		Me._lblTelegram_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._lblTelegram_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblTelegram_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblTelegram_0.UseMnemonic = True
		Me._lblTelegram_0.Visible = True
		Me._lblTelegram_0.AutoSize = False
		Me._lblTelegram_0.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me._lblTelegram_0.Name = "_lblTelegram_0"
		Me.picToolbar.Size = New System.Drawing.Size(379, 148)
		Me.picToolbar.Location = New System.Drawing.Point(16, 32)
		Me.picToolbar.TabIndex = 76
		Me.picToolbar.TabStop = False
		Me.picToolbar.Visible = False
		Me.picToolbar.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picToolbar.Dock = System.Windows.Forms.DockStyle.None
		Me.picToolbar.BackColor = System.Drawing.SystemColors.Control
		Me.picToolbar.CausesValidation = True
		Me.picToolbar.Enabled = True
		Me.picToolbar.ForeColor = System.Drawing.SystemColors.ControlText
		Me.picToolbar.Cursor = System.Windows.Forms.Cursors.Default
		Me.picToolbar.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picToolbar.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picToolbar.Name = "picToolbar"
		Me.fraToolbar.Text = "Toolbar Button Visiblity"
		Me.fraToolbar.Size = New System.Drawing.Size(377, 148)
		Me.fraToolbar.Location = New System.Drawing.Point(0, 0)
		Me.fraToolbar.TabIndex = 77
		Me.fraToolbar.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraToolbar.BackColor = System.Drawing.SystemColors.Control
		Me.fraToolbar.Enabled = True
		Me.fraToolbar.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraToolbar.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraToolbar.Visible = True
		Me.fraToolbar.Padding = New System.Windows.Forms.Padding(0)
		Me.fraToolbar.Name = "fraToolbar"
		Me.lstToolbar.Size = New System.Drawing.Size(361, 109)
		Me.lstToolbar.Location = New System.Drawing.Point(8, 16)
		Me.lstToolbar.TabIndex = 78
		Me.lstToolbar.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstToolbar.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstToolbar.BackColor = System.Drawing.SystemColors.Window
		Me.lstToolbar.CausesValidation = True
		Me.lstToolbar.Enabled = True
		Me.lstToolbar.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstToolbar.IntegralHeight = True
		Me.lstToolbar.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstToolbar.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstToolbar.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstToolbar.Sorted = False
		Me.lstToolbar.TabStop = True
		Me.lstToolbar.Visible = True
		Me.lstToolbar.MultiColumn = True
		Me.lstToolbar.ColumnWidth = 181
		Me.lstToolbar.Name = "lstToolbar"
		tbsOptions.OcxState = CType(resources.GetObject("tbsOptions.OcxState"), System.Windows.Forms.AxHost.State)
		Me.tbsOptions.Size = New System.Drawing.Size(393, 179)
		Me.tbsOptions.Location = New System.Drawing.Point(7, 8)
		Me.tbsOptions.TabIndex = 0
		Me.tbsOptions.Name = "tbsOptions"
		Me.cmdApply.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdApply.Text = "Apply"
		Me.cmdApply.Size = New System.Drawing.Size(73, 25)
		Me.cmdApply.Location = New System.Drawing.Point(328, 193)
		Me.cmdApply.TabIndex = 3
		Me.cmdApply.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdApply.BackColor = System.Drawing.SystemColors.Control
		Me.cmdApply.CausesValidation = True
		Me.cmdApply.Enabled = True
		Me.cmdApply.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdApply.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdApply.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdApply.TabStop = True
		Me.cmdApply.Name = "cmdApply"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(73, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(248, 193)
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
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "OK"
		Me.cmdOK.Size = New System.Drawing.Size(73, 25)
		Me.cmdOK.Location = New System.Drawing.Point(168, 193)
		Me.cmdOK.TabIndex = 1
		Me.cmdOK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.Controls.Add(picPlaying)
		Me.Controls.Add(picView)
		Me.Controls.Add(picUpdate)
		Me.Controls.Add(picTTS)
		Me.Controls.Add(picThemeGame)
		Me.Controls.Add(picScriptButton)
		Me.Controls.Add(picItem)
		Me.Controls.Add(picTelegram)
		Me.Controls.Add(picToolbar)
		Me.Controls.Add(tbsOptions)
		Me.Controls.Add(cmdApply)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOK)
		Me.picPlaying.Controls.Add(frmPathLength)
		Me.picPlaying.Controls.Add(frmFriendly)
		Me.picPlaying.Controls.Add(frmAutoAttack)
		Me.picPlaying.Controls.Add(fraGame)
		Me.frmPathLength.Controls.Add(txtPathLength)
		Me.frmFriendly.Controls.Add(optFriendlyAlly)
		Me.frmFriendly.Controls.Add(optFriendlyNeutral)
		Me.frmAutoAttack.Controls.Add(txtDefenseScaling)
		Me.frmAutoAttack.Controls.Add(lblDefenseScaling)
		Me.fraGame.Controls.Add(chkDumpHeaders)
		Me.fraGame.Controls.Add(chkCapital)
		Me.fraGame.Controls.Add(chkLocalHelp)
		Me.fraGame.Controls.Add(chkSound)
		Me.fraGame.Controls.Add(chkNoIron)
		Me.fraGame.Controls.Add(chkRetreat)
		Me.fraGame.Controls.Add(chkMaxProd)
		Me.picView.Controls.Add(fraView)
		Me.picView.Controls.Add(frmFlash)
		Me.picView.Controls.Add(frmImage)
		Me.fraView.Controls.Add(chkShortName)
		Me.fraView.Controls.Add(chkClear)
		Me.fraView.Controls.Add(chkToolbar)
		Me.fraView.Controls.Add(chkStatusBar)
		Me.fraView.Controls.Add(chkWarship)
		Me.frmFlash.Controls.Add(chkFlashLogging)
		Me.frmFlash.Controls.Add(chkFlashPlayerColors)
		Me.frmFlash.Controls.Add(chkFlashChat)
		Me.frmImage.Controls.Add(picBorderColor)
		Me.frmImage.Controls.Add(cmdImage)
		Me.frmImage.Controls.Add(txtImageName)
		Me.frmImage.Controls.Add(lblBorderColor)
		Me.picUpdate.Controls.Add(fraPlayer)
		Me.picUpdate.Controls.Add(frmDeity)
		Me.picUpdate.Controls.Add(fraUpdate)
		Me.fraPlayer.Controls.Add(cmbPlayer)
		Me.frmDeity.Controls.Add(chkFullDumpwithoutSea)
		Me.fraUpdate.Controls.Add(chkSuppressUnitDamageReports)
		Me.fraUpdate.Controls.Add(chkSuppressTelegramNotification)
		Me.fraUpdate.Controls.Add(chkSuppressUpdateCommands)
		Me.fraUpdate.Controls.Add(chkSuppressOilContentAtSea)
		Me.fraUpdate.Controls.Add(chkCommodityRatio)
		Me.fraUpdate.Controls.Add(chkRead)
		Me.fraUpdate.Controls.Add(chkSuppressBmap)
		Me.fraUpdate.Controls.Add(chkSuppressStrength)
		Me.fraUpdate.Controls.Add(chkUpdate)
		Me.picTTS.Controls.Add(frmTTS)
		Me.frmTTS.Controls.Add(chkTTSFlashes)
		Me.frmTTS.Controls.Add(chkTTSBulletins)
		Me.frmTTS.Controls.Add(chkTTSTelegrams)
		Me.frmTTS.Controls.Add(chkTTSReportWindow)
		Me.picThemeGame.Controls.Add(fraThemeGame)
		Me.fraThemeGame.Controls.Add(_optThemeGames_7)
		Me.fraThemeGame.Controls.Add(_optThemeGames_6)
		Me.fraThemeGame.Controls.Add(_optThemeGames_5)
		Me.fraThemeGame.Controls.Add(_optThemeGames_4)
		Me.fraThemeGame.Controls.Add(_optThemeGames_3)
		Me.fraThemeGame.Controls.Add(_optThemeGames_2)
		Me.fraThemeGame.Controls.Add(_optThemeGames_1)
		Me.fraThemeGame.Controls.Add(_optThemeGames_0)
		Me.picScriptButton.Controls.Add(optCustomScript)
		Me.picScriptButton.Controls.Add(optJumpPoint)
		Me.picScriptButton.Controls.Add(fraScriptButton)
		Me.picScriptButton.Controls.Add(cmbScriptButton)
		Me.fraScriptButton.Controls.Add(cmbJumpPoint)
		Me.fraScriptButton.Controls.Add(cmdScriptImageBrowse)
		Me.fraScriptButton.Controls.Add(txtScriptImageFileName)
		Me.fraScriptButton.Controls.Add(chkScriptReport)
		Me.fraScriptButton.Controls.Add(txtScriptFileName)
		Me.fraScriptButton.Controls.Add(cmdScriptBrowse)
		Me.fraScriptButton.Controls.Add(lblScriptImage)
		Me.fraScriptButton.Controls.Add(lblScriptFileName)
		Me.picItem.Controls.Add(frmItemLabels)
		Me.picItem.Controls.Add(cbItem)
		Me.picItem.Controls.Add(frmItem)
		Me.frmItemLabels.Controls.Add(txtItemDistributionPanelLabel)
		Me.frmItemLabels.Controls.Add(txtItemSectorPanelLabel)
		Me.frmItemLabels.Controls.Add(txtItemFormLabel)
		Me.frmItemLabels.Controls.Add(lblItemDistributionPanelLabel)
		Me.frmItemLabels.Controls.Add(lblItemSectorPanelLabel)
		Me.frmItemLabels.Controls.Add(lblItemFormLabel)
		Me.frmItem.Controls.Add(txtItemConditionalName)
		Me.frmItem.Controls.Add(txtItemIntelligenceNames)
		Me.frmItem.Controls.Add(txtItemDBName)
		Me.frmItem.Controls.Add(txtItemFormName)
		Me.frmItem.Controls.Add(lblItemConditionalName)
		Me.frmItem.Controls.Add(lblItemIntelligenceNames)
		Me.frmItem.Controls.Add(lblItemDBName)
		Me.frmItem.Controls.Add(lblItemFormName)
		Me.picTelegram.Controls.Add(_fraTelegram_2)
		Me.picTelegram.Controls.Add(_fraTelegram_1)
		Me.picTelegram.Controls.Add(cmbTelegram)
		Me.picTelegram.Controls.Add(_fraTelegram_0)
		Me._fraTelegram_2.Controls.Add(_cmbTelegramSound_2)
		Me._fraTelegram_2.Controls.Add(_txtTelegram_2)
		Me._fraTelegram_2.Controls.Add(_chkTelegram_2)
		Me._fraTelegram_2.Controls.Add(_lblTelegram_2)
		Me._fraTelegram_1.Controls.Add(_cmbTelegramSound_1)
		Me._fraTelegram_1.Controls.Add(_txtTelegram_1)
		Me._fraTelegram_1.Controls.Add(Text1)
		Me._fraTelegram_1.Controls.Add(_chkTelegram_1)
		Me._fraTelegram_1.Controls.Add(_lblTelegram_1)
		Me._fraTelegram_0.Controls.Add(_cmbTelegramSound_0)
		Me._fraTelegram_0.Controls.Add(_txtTelegram_0)
		Me._fraTelegram_0.Controls.Add(_chkTelegram_0)
		Me._fraTelegram_0.Controls.Add(_lblTelegram_0)
		Me.picToolbar.Controls.Add(fraToolbar)
		Me.fraToolbar.Controls.Add(lstToolbar)
		Me.chkTelegram.SetIndex(_chkTelegram_2, CType(2, Short))
		Me.chkTelegram.SetIndex(_chkTelegram_1, CType(1, Short))
		Me.chkTelegram.SetIndex(_chkTelegram_0, CType(0, Short))
		Me.cmbTelegramSound.SetIndex(_cmbTelegramSound_2, CType(2, Short))
		Me.cmbTelegramSound.SetIndex(_cmbTelegramSound_1, CType(1, Short))
		Me.cmbTelegramSound.SetIndex(_cmbTelegramSound_0, CType(0, Short))
		Me.fraTelegram.SetIndex(_fraTelegram_2, CType(2, Short))
		Me.fraTelegram.SetIndex(_fraTelegram_1, CType(1, Short))
		Me.fraTelegram.SetIndex(_fraTelegram_0, CType(0, Short))
		Me.lblTelegram.SetIndex(_lblTelegram_2, CType(2, Short))
		Me.lblTelegram.SetIndex(_lblTelegram_1, CType(1, Short))
		Me.lblTelegram.SetIndex(_lblTelegram_0, CType(0, Short))
		Me.optThemeGames.SetIndex(_optThemeGames_7, CType(7, Short))
		Me.optThemeGames.SetIndex(_optThemeGames_6, CType(6, Short))
		Me.optThemeGames.SetIndex(_optThemeGames_5, CType(5, Short))
		Me.optThemeGames.SetIndex(_optThemeGames_4, CType(4, Short))
		Me.optThemeGames.SetIndex(_optThemeGames_3, CType(3, Short))
		Me.optThemeGames.SetIndex(_optThemeGames_2, CType(2, Short))
		Me.optThemeGames.SetIndex(_optThemeGames_1, CType(1, Short))
		Me.optThemeGames.SetIndex(_optThemeGames_0, CType(0, Short))
		Me.txtTelegram.SetIndex(_txtTelegram_2, CType(2, Short))
		Me.txtTelegram.SetIndex(_txtTelegram_1, CType(1, Short))
		Me.txtTelegram.SetIndex(_txtTelegram_0, CType(0, Short))
		CType(Me.txtTelegram, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.optThemeGames, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.lblTelegram, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.fraTelegram, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.cmbTelegramSound, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.chkTelegram, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.tbsOptions, System.ComponentModel.ISupportInitialize).EndInit()
		Me.picPlaying.ResumeLayout(False)
		Me.frmPathLength.ResumeLayout(False)
		Me.frmFriendly.ResumeLayout(False)
		Me.frmAutoAttack.ResumeLayout(False)
		Me.fraGame.ResumeLayout(False)
		Me.picView.ResumeLayout(False)
		Me.fraView.ResumeLayout(False)
		Me.frmFlash.ResumeLayout(False)
		Me.frmImage.ResumeLayout(False)
		Me.picUpdate.ResumeLayout(False)
		Me.fraPlayer.ResumeLayout(False)
		Me.frmDeity.ResumeLayout(False)
		Me.fraUpdate.ResumeLayout(False)
		Me.picTTS.ResumeLayout(False)
		Me.frmTTS.ResumeLayout(False)
		Me.picThemeGame.ResumeLayout(False)
		Me.fraThemeGame.ResumeLayout(False)
		Me.picScriptButton.ResumeLayout(False)
		Me.fraScriptButton.ResumeLayout(False)
		Me.picItem.ResumeLayout(False)
		Me.frmItemLabels.ResumeLayout(False)
		Me.frmItem.ResumeLayout(False)
		Me.picTelegram.ResumeLayout(False)
		Me._fraTelegram_2.ResumeLayout(False)
		Me._fraTelegram_1.ResumeLayout(False)
		Me._fraTelegram_0.ResumeLayout(False)
		Me.picToolbar.ResumeLayout(False)
		Me.fraToolbar.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class