<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmChat
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
	Public WithEvents mnuCopyReportWindow As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents mnuBaseMenu As System.Windows.Forms.ToolStripMenuItem
	Public WithEvents MainMenu1 As System.Windows.Forms.MenuStrip
	Public WithEvents cbAllies As System.Windows.Forms.ComboBox
	Public WithEvents chkBeep As System.Windows.Forms.CheckBox
	Public WithEvents chkTop As System.Windows.Forms.CheckBox
	Public WithEvents txtUser As System.Windows.Forms.TextBox
	Public WithEvents optUser As System.Windows.Forms.RadioButton
	Public WithEvents optAll As System.Windows.Forms.RadioButton
	Public WithEvents txtMessage As System.Windows.Forms.TextBox
	Public WithEvents rtbText As System.Windows.Forms.RichTextBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmChat))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.MainMenu1 = New System.Windows.Forms.MenuStrip
		Me.mnuBaseMenu = New System.Windows.Forms.ToolStripMenuItem
		Me.mnuCopyReportWindow = New System.Windows.Forms.ToolStripMenuItem
		Me.cbAllies = New System.Windows.Forms.ComboBox
		Me.chkBeep = New System.Windows.Forms.CheckBox
		Me.chkTop = New System.Windows.Forms.CheckBox
		Me.txtUser = New System.Windows.Forms.TextBox
		Me.optUser = New System.Windows.Forms.RadioButton
		Me.optAll = New System.Windows.Forms.RadioButton
		Me.txtMessage = New System.Windows.Forms.TextBox
		Me.rtbText = New System.Windows.Forms.RichTextBox
		Me.MainMenu1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "Chat (EmpireHub)"
		Me.ClientSize = New System.Drawing.Size(345, 241)
		Me.Location = New System.Drawing.Point(4, 30)
		Me.KeyPreview = True
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.WindowState = System.Windows.Forms.FormWindowState.Minimized
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
		Me.Name = "frmChat"
		Me.mnuBaseMenu.Name = "mnuBaseMenu"
		Me.mnuBaseMenu.Text = "Base Menu"
		Me.mnuBaseMenu.Visible = False
		Me.mnuBaseMenu.Checked = False
		Me.mnuBaseMenu.Enabled = True
		Me.mnuCopyReportWindow.Name = "mnuCopyReportWindow"
		Me.mnuCopyReportWindow.Text = "Copy to Report Window"
		Me.mnuCopyReportWindow.Checked = False
		Me.mnuCopyReportWindow.Enabled = True
		Me.mnuCopyReportWindow.Visible = True
		Me.cbAllies.Size = New System.Drawing.Size(73, 21)
		Me.cbAllies.Location = New System.Drawing.Point(112, 216)
		Me.cbAllies.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cbAllies.TabIndex = 7
		Me.cbAllies.Visible = False
		Me.cbAllies.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbAllies.BackColor = System.Drawing.SystemColors.Window
		Me.cbAllies.CausesValidation = True
		Me.cbAllies.Enabled = True
		Me.cbAllies.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cbAllies.IntegralHeight = True
		Me.cbAllies.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbAllies.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbAllies.Sorted = False
		Me.cbAllies.TabStop = True
		Me.cbAllies.Name = "cbAllies"
		Me.chkBeep.Text = "Beep"
		Me.chkBeep.Size = New System.Drawing.Size(49, 17)
		Me.chkBeep.Location = New System.Drawing.Point(296, 216)
		Me.chkBeep.TabIndex = 6
		Me.chkBeep.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkBeep.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkBeep.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkBeep.BackColor = System.Drawing.SystemColors.Control
		Me.chkBeep.CausesValidation = True
		Me.chkBeep.Enabled = True
		Me.chkBeep.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkBeep.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkBeep.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkBeep.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkBeep.TabStop = True
		Me.chkBeep.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkBeep.Visible = True
		Me.chkBeep.Name = "chkBeep"
		Me.chkTop.Text = "Always on Top"
		Me.chkTop.Size = New System.Drawing.Size(97, 17)
		Me.chkTop.Location = New System.Drawing.Point(192, 216)
		Me.chkTop.TabIndex = 5
		Me.chkTop.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkTop.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkTop.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkTop.BackColor = System.Drawing.SystemColors.Control
		Me.chkTop.CausesValidation = True
		Me.chkTop.Enabled = True
		Me.chkTop.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkTop.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkTop.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkTop.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkTop.TabStop = True
		Me.chkTop.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkTop.Visible = True
		Me.chkTop.Name = "chkTop"
		Me.txtUser.AutoSize = False
		Me.txtUser.Size = New System.Drawing.Size(81, 19)
		Me.txtUser.Location = New System.Drawing.Point(104, 216)
		Me.txtUser.TabIndex = 4
		Me.txtUser.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtUser.AcceptsReturn = True
		Me.txtUser.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtUser.BackColor = System.Drawing.SystemColors.Window
		Me.txtUser.CausesValidation = True
		Me.txtUser.Enabled = True
		Me.txtUser.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtUser.HideSelection = True
		Me.txtUser.ReadOnly = False
		Me.txtUser.Maxlength = 0
		Me.txtUser.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtUser.MultiLine = False
		Me.txtUser.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtUser.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtUser.TabStop = True
		Me.txtUser.Visible = True
		Me.txtUser.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtUser.Name = "txtUser"
		Me.optUser.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optUser.Text = "User"
		Me.optUser.Size = New System.Drawing.Size(57, 21)
		Me.optUser.Location = New System.Drawing.Point(48, 216)
		Me.optUser.TabIndex = 3
		Me.optUser.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optUser.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optUser.BackColor = System.Drawing.SystemColors.Control
		Me.optUser.CausesValidation = True
		Me.optUser.Enabled = True
		Me.optUser.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optUser.Cursor = System.Windows.Forms.Cursors.Default
		Me.optUser.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optUser.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optUser.TabStop = True
		Me.optUser.Checked = False
		Me.optUser.Visible = True
		Me.optUser.Name = "optUser"
		Me.optAll.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optAll.Text = "All"
		Me.optAll.Size = New System.Drawing.Size(33, 21)
		Me.optAll.Location = New System.Drawing.Point(8, 216)
		Me.optAll.TabIndex = 2
		Me.optAll.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optAll.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optAll.BackColor = System.Drawing.SystemColors.Control
		Me.optAll.CausesValidation = True
		Me.optAll.Enabled = True
		Me.optAll.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optAll.Cursor = System.Windows.Forms.Cursors.Default
		Me.optAll.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optAll.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optAll.TabStop = True
		Me.optAll.Checked = False
		Me.optAll.Visible = True
		Me.optAll.Name = "optAll"
		Me.txtMessage.AutoSize = False
		Me.txtMessage.Size = New System.Drawing.Size(331, 33)
		Me.txtMessage.Location = New System.Drawing.Point(8, 176)
		Me.txtMessage.Maxlength = 1000
		Me.txtMessage.MultiLine = True
		Me.txtMessage.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
		Me.txtMessage.TabIndex = 1
		Me.txtMessage.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtMessage.AcceptsReturn = True
		Me.txtMessage.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtMessage.BackColor = System.Drawing.SystemColors.Window
		Me.txtMessage.CausesValidation = True
		Me.txtMessage.Enabled = True
		Me.txtMessage.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtMessage.HideSelection = True
		Me.txtMessage.ReadOnly = False
		Me.txtMessage.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtMessage.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtMessage.TabStop = True
		Me.txtMessage.Visible = True
		Me.txtMessage.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtMessage.Name = "txtMessage"
		Me.rtbText.Size = New System.Drawing.Size(331, 137)
		Me.rtbText.Location = New System.Drawing.Point(8, 32)
		Me.rtbText.TabIndex = 0
		Me.rtbText.TabStop = 0
		Me.rtbText.Enabled = True
		Me.rtbText.ReadOnly = True
		Me.rtbText.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
		Me.rtbText.RTF = resources.GetString("rtbText.TextRTF")
		Me.rtbText.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.rtbText.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.rtbText.Name = "rtbText"
		Me.Controls.Add(cbAllies)
		Me.Controls.Add(chkBeep)
		Me.Controls.Add(chkTop)
		Me.Controls.Add(txtUser)
		Me.Controls.Add(optUser)
		Me.Controls.Add(optAll)
		Me.Controls.Add(txtMessage)
		Me.Controls.Add(rtbText)
		MainMenu1.Items.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuBaseMenu})
		mnuBaseMenu.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem(){Me.mnuCopyReportWindow})
		Me.Controls.Add(MainMenu1)
		Me.MainMenu1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class