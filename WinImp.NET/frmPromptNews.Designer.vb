<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPromptNews
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
	Public WithEvents optNone As System.Windows.Forms.RadioButton
	Public WithEvents optHourly As System.Windows.Forms.RadioButton
	Public WithEvents optNewsPaper As System.Windows.Forms.RadioButton
	Public WithEvents fraOutput As System.Windows.Forms.GroupBox
	Public WithEvents cmdHelp As System.Windows.Forms.Button
	Public WithEvents chkTelegramCount As System.Windows.Forms.CheckBox
	Public WithEvents cbVictim As System.Windows.Forms.ComboBox
	Public WithEvents cbActor As System.Windows.Forms.ComboBox
	Public WithEvents chkVictim As System.Windows.Forms.CheckBox
	Public WithEvents chkActor As System.Windows.Forms.CheckBox
	Public WithEvents chkHeadlines As System.Windows.Forms.CheckBox
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents cbDays As System.Windows.Forms.ComboBox
	Public WithEvents Option2 As System.Windows.Forms.RadioButton
	Public WithEvents Option1 As System.Windows.Forms.RadioButton
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents Label2 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPromptNews))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.fraOutput = New System.Windows.Forms.GroupBox
		Me.optNone = New System.Windows.Forms.RadioButton
		Me.optHourly = New System.Windows.Forms.RadioButton
		Me.optNewsPaper = New System.Windows.Forms.RadioButton
		Me.cmdHelp = New System.Windows.Forms.Button
		Me.Frame2 = New System.Windows.Forms.GroupBox
		Me.chkTelegramCount = New System.Windows.Forms.CheckBox
		Me.cbVictim = New System.Windows.Forms.ComboBox
		Me.cbActor = New System.Windows.Forms.ComboBox
		Me.chkVictim = New System.Windows.Forms.CheckBox
		Me.chkActor = New System.Windows.Forms.CheckBox
		Me.chkHeadlines = New System.Windows.Forms.CheckBox
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.cbDays = New System.Windows.Forms.ComboBox
		Me.Option2 = New System.Windows.Forms.RadioButton
		Me.Option1 = New System.Windows.Forms.RadioButton
		Me.Label1 = New System.Windows.Forms.Label
		Me.cmdOK = New System.Windows.Forms.Button
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.Label2 = New System.Windows.Forms.Label
		Me.fraOutput.SuspendLayout()
		Me.Frame2.SuspendLayout()
		Me.Frame1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.ControlBox = False
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.ClientSize = New System.Drawing.Size(512, 175)
		Me.Location = New System.Drawing.Point(3, 3)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmPromptNews"
		Me.fraOutput.Text = "Report Output"
		Me.fraOutput.Size = New System.Drawing.Size(265, 41)
		Me.fraOutput.Location = New System.Drawing.Point(240, 128)
		Me.fraOutput.TabIndex = 16
		Me.fraOutput.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraOutput.BackColor = System.Drawing.SystemColors.Control
		Me.fraOutput.Enabled = True
		Me.fraOutput.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraOutput.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraOutput.Visible = True
		Me.fraOutput.Padding = New System.Windows.Forms.Padding(0)
		Me.fraOutput.Name = "fraOutput"
		Me.optNone.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optNone.Text = "None"
		Me.optNone.Size = New System.Drawing.Size(49, 17)
		Me.optNone.Location = New System.Drawing.Point(200, 16)
		Me.optNone.TabIndex = 19
		Me.optNone.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optNone.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optNone.BackColor = System.Drawing.SystemColors.Control
		Me.optNone.CausesValidation = True
		Me.optNone.Enabled = True
		Me.optNone.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optNone.Cursor = System.Windows.Forms.Cursors.Default
		Me.optNone.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optNone.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optNone.TabStop = True
		Me.optNone.Checked = False
		Me.optNone.Visible = True
		Me.optNone.Name = "optNone"
		Me.optHourly.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optHourly.Text = "Hourly Activity"
		Me.optHourly.Size = New System.Drawing.Size(89, 17)
		Me.optHourly.Location = New System.Drawing.Point(96, 16)
		Me.optHourly.TabIndex = 18
		Me.optHourly.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optHourly.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optHourly.BackColor = System.Drawing.SystemColors.Control
		Me.optHourly.CausesValidation = True
		Me.optHourly.Enabled = True
		Me.optHourly.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optHourly.Cursor = System.Windows.Forms.Cursors.Default
		Me.optHourly.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optHourly.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optHourly.TabStop = True
		Me.optHourly.Checked = False
		Me.optHourly.Visible = True
		Me.optHourly.Name = "optHourly"
		Me.optNewsPaper.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optNewsPaper.Text = "Newspaper"
		Me.optNewsPaper.Size = New System.Drawing.Size(81, 17)
		Me.optNewsPaper.Location = New System.Drawing.Point(8, 16)
		Me.optNewsPaper.TabIndex = 17
		Me.optNewsPaper.Checked = True
		Me.optNewsPaper.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optNewsPaper.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optNewsPaper.BackColor = System.Drawing.SystemColors.Control
		Me.optNewsPaper.CausesValidation = True
		Me.optNewsPaper.Enabled = True
		Me.optNewsPaper.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optNewsPaper.Cursor = System.Windows.Forms.Cursors.Default
		Me.optNewsPaper.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optNewsPaper.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optNewsPaper.TabStop = True
		Me.optNewsPaper.Visible = True
		Me.optNewsPaper.Name = "optNewsPaper"
		Me.cmdHelp.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdHelp.Text = "?"
		Me.cmdHelp.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdHelp.Size = New System.Drawing.Size(25, 25)
		Me.cmdHelp.Location = New System.Drawing.Point(488, 0)
		Me.cmdHelp.TabIndex = 14
		Me.ToolTip1.SetToolTip(Me.cmdHelp, "Click for Help")
		Me.cmdHelp.BackColor = System.Drawing.SystemColors.Control
		Me.cmdHelp.CausesValidation = True
		Me.cmdHelp.Enabled = True
		Me.cmdHelp.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdHelp.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdHelp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdHelp.TabStop = True
		Me.cmdHelp.Name = "cmdHelp"
		Me.Frame2.Text = "Options"
		Me.Frame2.Size = New System.Drawing.Size(217, 137)
		Me.Frame2.Location = New System.Drawing.Point(8, 32)
		Me.Frame2.TabIndex = 9
		Me.Frame2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame2.BackColor = System.Drawing.SystemColors.Control
		Me.Frame2.Enabled = True
		Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame2.Visible = True
		Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame2.Name = "Frame2"
		Me.chkTelegramCount.Text = "Analyse Telegram Activity (Show Relations Grid)"
		Me.chkTelegramCount.Size = New System.Drawing.Size(145, 25)
		Me.chkTelegramCount.Location = New System.Drawing.Point(8, 104)
		Me.chkTelegramCount.TabIndex = 15
		Me.chkTelegramCount.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkTelegramCount.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkTelegramCount.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkTelegramCount.BackColor = System.Drawing.SystemColors.Control
		Me.chkTelegramCount.CausesValidation = True
		Me.chkTelegramCount.Enabled = True
		Me.chkTelegramCount.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkTelegramCount.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkTelegramCount.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkTelegramCount.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkTelegramCount.TabStop = True
		Me.chkTelegramCount.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkTelegramCount.Visible = True
		Me.chkTelegramCount.Name = "chkTelegramCount"
		Me.cbVictim.Size = New System.Drawing.Size(145, 21)
		Me.cbVictim.Location = New System.Drawing.Point(64, 56)
		Me.cbVictim.Sorted = True
		Me.cbVictim.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cbVictim.TabIndex = 12
		Me.cbVictim.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbVictim.BackColor = System.Drawing.SystemColors.Window
		Me.cbVictim.CausesValidation = True
		Me.cbVictim.Enabled = True
		Me.cbVictim.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cbVictim.IntegralHeight = True
		Me.cbVictim.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbVictim.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbVictim.TabStop = True
		Me.cbVictim.Visible = True
		Me.cbVictim.Name = "cbVictim"
		Me.cbActor.Size = New System.Drawing.Size(145, 21)
		Me.cbActor.Location = New System.Drawing.Point(64, 24)
		Me.cbActor.Sorted = True
		Me.cbActor.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cbActor.TabIndex = 11
		Me.cbActor.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbActor.BackColor = System.Drawing.SystemColors.Window
		Me.cbActor.CausesValidation = True
		Me.cbActor.Enabled = True
		Me.cbActor.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cbActor.IntegralHeight = True
		Me.cbActor.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbActor.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbActor.TabStop = True
		Me.cbActor.Visible = True
		Me.cbActor.Name = "cbActor"
		Me.chkVictim.Text = "Victim"
		Me.chkVictim.Size = New System.Drawing.Size(49, 17)
		Me.chkVictim.Location = New System.Drawing.Point(8, 56)
		Me.chkVictim.TabIndex = 7
		Me.chkVictim.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkVictim.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkVictim.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkVictim.BackColor = System.Drawing.SystemColors.Control
		Me.chkVictim.CausesValidation = True
		Me.chkVictim.Enabled = True
		Me.chkVictim.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkVictim.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkVictim.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkVictim.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkVictim.TabStop = True
		Me.chkVictim.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkVictim.Visible = True
		Me.chkVictim.Name = "chkVictim"
		Me.chkActor.Text = "Actor"
		Me.chkActor.Size = New System.Drawing.Size(49, 17)
		Me.chkActor.Location = New System.Drawing.Point(8, 24)
		Me.chkActor.TabIndex = 6
		Me.chkActor.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkActor.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkActor.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkActor.BackColor = System.Drawing.SystemColors.Control
		Me.chkActor.CausesValidation = True
		Me.chkActor.Enabled = True
		Me.chkActor.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkActor.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkActor.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkActor.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkActor.TabStop = True
		Me.chkActor.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkActor.Visible = True
		Me.chkActor.Name = "chkActor"
		Me.chkHeadlines.Text = "Show Headlines only"
		Me.chkHeadlines.Size = New System.Drawing.Size(137, 17)
		Me.chkHeadlines.Location = New System.Drawing.Point(8, 80)
		Me.chkHeadlines.TabIndex = 8
		Me.chkHeadlines.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkHeadlines.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkHeadlines.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkHeadlines.BackColor = System.Drawing.SystemColors.Control
		Me.chkHeadlines.CausesValidation = True
		Me.chkHeadlines.Enabled = True
		Me.chkHeadlines.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkHeadlines.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkHeadlines.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkHeadlines.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkHeadlines.TabStop = True
		Me.chkHeadlines.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkHeadlines.Visible = True
		Me.chkHeadlines.Name = "chkHeadlines"
		Me.Frame1.Text = "View News Entries"
		Me.Frame1.Size = New System.Drawing.Size(177, 81)
		Me.Frame1.Location = New System.Drawing.Point(240, 32)
		Me.Frame1.TabIndex = 2
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.cbDays.Size = New System.Drawing.Size(41, 21)
		Me.cbDays.Location = New System.Drawing.Point(88, 48)
		Me.cbDays.TabIndex = 5
		Me.cbDays.Text = "cbDays"
		Me.cbDays.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbDays.BackColor = System.Drawing.SystemColors.Window
		Me.cbDays.CausesValidation = True
		Me.cbDays.Enabled = True
		Me.cbDays.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cbDays.IntegralHeight = True
		Me.cbDays.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbDays.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbDays.Sorted = False
		Me.cbDays.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
		Me.cbDays.TabStop = True
		Me.cbDays.Visible = True
		Me.cbDays.Name = "cbDays"
		Me.Option2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option2.Text = "for the last"
		Me.Option2.Size = New System.Drawing.Size(81, 17)
		Me.Option2.Location = New System.Drawing.Point(8, 48)
		Me.Option2.TabIndex = 4
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
		Me.Option1.Text = "Since Last Time I checked"
		Me.Option1.Size = New System.Drawing.Size(153, 17)
		Me.Option1.Location = New System.Drawing.Point(8, 24)
		Me.Option1.TabIndex = 3
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
		Me.Label1.Text = "day(s)"
		Me.Label1.Size = New System.Drawing.Size(33, 17)
		Me.Label1.Location = New System.Drawing.Point(136, 48)
		Me.Label1.TabIndex = 10
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
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "OK"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Size = New System.Drawing.Size(73, 33)
		Me.cmdOK.Location = New System.Drawing.Point(432, 40)
		Me.cmdOK.TabIndex = 0
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
		Me.cmdCancel.Size = New System.Drawing.Size(73, 33)
		Me.cmdCancel.Location = New System.Drawing.Point(432, 80)
		Me.cmdCancel.TabIndex = 1
		Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.Label2.Text = "Newspaper"
		Me.Label2.Font = New System.Drawing.Font("Arial Black", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(105, 17)
		Me.Label2.Location = New System.Drawing.Point(8, 8)
		Me.Label2.TabIndex = 13
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
		Me.Controls.Add(fraOutput)
		Me.Controls.Add(cmdHelp)
		Me.Controls.Add(Frame2)
		Me.Controls.Add(Frame1)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(Label2)
		Me.fraOutput.Controls.Add(optNone)
		Me.fraOutput.Controls.Add(optHourly)
		Me.fraOutput.Controls.Add(optNewsPaper)
		Me.Frame2.Controls.Add(chkTelegramCount)
		Me.Frame2.Controls.Add(cbVictim)
		Me.Frame2.Controls.Add(cbActor)
		Me.Frame2.Controls.Add(chkVictim)
		Me.Frame2.Controls.Add(chkActor)
		Me.Frame2.Controls.Add(chkHeadlines)
		Me.Frame1.Controls.Add(cbDays)
		Me.Frame1.Controls.Add(Option2)
		Me.Frame1.Controls.Add(Option1)
		Me.Frame1.Controls.Add(Label1)
		Me.fraOutput.ResumeLayout(False)
		Me.Frame2.ResumeLayout(False)
		Me.Frame1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class