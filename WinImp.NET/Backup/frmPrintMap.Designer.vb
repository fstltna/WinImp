<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPrintMap
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
	Public WithEvents cmdBmp As System.Windows.Forms.Button
	Public WithEvents chkSea As System.Windows.Forms.CheckBox
	Public WithEvents txtOrigin As System.Windows.Forms.TextBox
	Public WithEvents Option1 As System.Windows.Forms.RadioButton
	Public WithEvents optPortrait As System.Windows.Forms.RadioButton
	Public WithEvents fraOrientation As System.Windows.Forms.GroupBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdPrint As System.Windows.Forms.Button
	Public WithEvents cmdPreview As System.Windows.Forms.Button
	Public WithEvents txtScale As System.Windows.Forms.TextBox
	Public WithEvents cbPrinter As System.Windows.Forms.ComboBox
	Public WithEvents lblCenter As System.Windows.Forms.Label
	Public WithEvents lblScale As System.Windows.Forms.Label
	Public WithEvents lblPrinter As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPrintMap))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdBmp = New System.Windows.Forms.Button
		Me.chkSea = New System.Windows.Forms.CheckBox
		Me.txtOrigin = New System.Windows.Forms.TextBox
		Me.fraOrientation = New System.Windows.Forms.GroupBox
		Me.Option1 = New System.Windows.Forms.RadioButton
		Me.optPortrait = New System.Windows.Forms.RadioButton
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdPrint = New System.Windows.Forms.Button
		Me.cmdPreview = New System.Windows.Forms.Button
		Me.txtScale = New System.Windows.Forms.TextBox
		Me.cbPrinter = New System.Windows.Forms.ComboBox
		Me.lblCenter = New System.Windows.Forms.Label
		Me.lblScale = New System.Windows.Forms.Label
		Me.lblPrinter = New System.Windows.Forms.Label
		Me.fraOrientation.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "Print Map"
		Me.ClientSize = New System.Drawing.Size(266, 129)
		Me.Location = New System.Drawing.Point(4, 30)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
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
		Me.Name = "frmPrintMap"
		Me.cmdBmp.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdBmp.Text = "bmp"
		Me.AcceptButton = Me.cmdBmp
		Me.cmdBmp.Size = New System.Drawing.Size(57, 25)
		Me.cmdBmp.Location = New System.Drawing.Point(72, 96)
		Me.cmdBmp.TabIndex = 13
		Me.cmdBmp.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdBmp.BackColor = System.Drawing.SystemColors.Control
		Me.cmdBmp.CausesValidation = True
		Me.cmdBmp.Enabled = True
		Me.cmdBmp.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdBmp.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdBmp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdBmp.TabStop = True
		Me.cmdBmp.Name = "cmdBmp"
		Me.chkSea.Text = "Show Known Sea Sectors"
		Me.chkSea.Size = New System.Drawing.Size(145, 17)
		Me.chkSea.Location = New System.Drawing.Point(8, 72)
		Me.chkSea.TabIndex = 12
		Me.chkSea.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkSea.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chkSea.FlatStyle = System.Windows.Forms.FlatStyle.Standard
		Me.chkSea.BackColor = System.Drawing.SystemColors.Control
		Me.chkSea.CausesValidation = True
		Me.chkSea.Enabled = True
		Me.chkSea.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chkSea.Cursor = System.Windows.Forms.Cursors.Default
		Me.chkSea.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chkSea.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chkSea.TabStop = True
		Me.chkSea.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chkSea.Visible = True
		Me.chkSea.Name = "chkSea"
		Me.txtOrigin.AutoSize = False
		Me.txtOrigin.Size = New System.Drawing.Size(41, 19)
		Me.txtOrigin.Location = New System.Drawing.Point(120, 40)
		Me.txtOrigin.TabIndex = 11
		Me.txtOrigin.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtOrigin.AcceptsReturn = True
		Me.txtOrigin.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtOrigin.BackColor = System.Drawing.SystemColors.Window
		Me.txtOrigin.CausesValidation = True
		Me.txtOrigin.Enabled = True
		Me.txtOrigin.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtOrigin.HideSelection = True
		Me.txtOrigin.ReadOnly = False
		Me.txtOrigin.Maxlength = 0
		Me.txtOrigin.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtOrigin.MultiLine = False
		Me.txtOrigin.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtOrigin.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtOrigin.TabStop = True
		Me.txtOrigin.Visible = True
		Me.txtOrigin.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtOrigin.Name = "txtOrigin"
		Me.fraOrientation.Text = "Orientation"
		Me.fraOrientation.Size = New System.Drawing.Size(89, 57)
		Me.fraOrientation.Location = New System.Drawing.Point(168, 32)
		Me.fraOrientation.TabIndex = 7
		Me.fraOrientation.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.fraOrientation.BackColor = System.Drawing.SystemColors.Control
		Me.fraOrientation.Enabled = True
		Me.fraOrientation.ForeColor = System.Drawing.SystemColors.ControlText
		Me.fraOrientation.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.fraOrientation.Visible = True
		Me.fraOrientation.Padding = New System.Windows.Forms.Padding(0)
		Me.fraOrientation.Name = "fraOrientation"
		Me.Option1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Option1.Text = "Landscape"
		Me.Option1.Size = New System.Drawing.Size(73, 13)
		Me.Option1.Location = New System.Drawing.Point(8, 32)
		Me.Option1.TabIndex = 9
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
		Me.Option1.Checked = False
		Me.Option1.Visible = True
		Me.Option1.Name = "Option1"
		Me.optPortrait.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optPortrait.Text = "Portrait"
		Me.optPortrait.Size = New System.Drawing.Size(57, 13)
		Me.optPortrait.Location = New System.Drawing.Point(8, 16)
		Me.optPortrait.TabIndex = 8
		Me.optPortrait.Checked = True
		Me.optPortrait.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optPortrait.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.optPortrait.BackColor = System.Drawing.SystemColors.Control
		Me.optPortrait.CausesValidation = True
		Me.optPortrait.Enabled = True
		Me.optPortrait.ForeColor = System.Drawing.SystemColors.ControlText
		Me.optPortrait.Cursor = System.Windows.Forms.Cursors.Default
		Me.optPortrait.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.optPortrait.Appearance = System.Windows.Forms.Appearance.Normal
		Me.optPortrait.TabStop = True
		Me.optPortrait.Visible = True
		Me.optPortrait.Name = "optPortrait"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(57, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(8, 96)
		Me.cmdCancel.TabIndex = 6
		Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.cmdPrint.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdPrint.Text = "Print"
		Me.cmdPrint.Size = New System.Drawing.Size(57, 25)
		Me.cmdPrint.Location = New System.Drawing.Point(200, 96)
		Me.cmdPrint.TabIndex = 5
		Me.cmdPrint.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdPrint.BackColor = System.Drawing.SystemColors.Control
		Me.cmdPrint.CausesValidation = True
		Me.cmdPrint.Enabled = True
		Me.cmdPrint.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdPrint.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdPrint.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdPrint.TabStop = True
		Me.cmdPrint.Name = "cmdPrint"
		Me.cmdPreview.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdPreview.Text = "Preview"
		Me.cmdPreview.Size = New System.Drawing.Size(57, 25)
		Me.cmdPreview.Location = New System.Drawing.Point(136, 96)
		Me.cmdPreview.TabIndex = 4
		Me.cmdPreview.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdPreview.BackColor = System.Drawing.SystemColors.Control
		Me.cmdPreview.CausesValidation = True
		Me.cmdPreview.Enabled = True
		Me.cmdPreview.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdPreview.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdPreview.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdPreview.TabStop = True
		Me.cmdPreview.Name = "cmdPreview"
		Me.txtScale.AutoSize = False
		Me.txtScale.Size = New System.Drawing.Size(25, 19)
		Me.txtScale.Location = New System.Drawing.Point(48, 40)
		Me.txtScale.TabIndex = 3
		Me.txtScale.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtScale.AcceptsReturn = True
		Me.txtScale.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtScale.BackColor = System.Drawing.SystemColors.Window
		Me.txtScale.CausesValidation = True
		Me.txtScale.Enabled = True
		Me.txtScale.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtScale.HideSelection = True
		Me.txtScale.ReadOnly = False
		Me.txtScale.Maxlength = 0
		Me.txtScale.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtScale.MultiLine = False
		Me.txtScale.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtScale.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtScale.TabStop = True
		Me.txtScale.Visible = True
		Me.txtScale.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtScale.Name = "txtScale"
		Me.cbPrinter.Size = New System.Drawing.Size(209, 21)
		Me.cbPrinter.Location = New System.Drawing.Point(48, 8)
		Me.cbPrinter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cbPrinter.TabIndex = 0
		Me.cbPrinter.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbPrinter.BackColor = System.Drawing.SystemColors.Window
		Me.cbPrinter.CausesValidation = True
		Me.cbPrinter.Enabled = True
		Me.cbPrinter.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cbPrinter.IntegralHeight = True
		Me.cbPrinter.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbPrinter.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbPrinter.Sorted = False
		Me.cbPrinter.TabStop = True
		Me.cbPrinter.Visible = True
		Me.cbPrinter.Name = "cbPrinter"
		Me.lblCenter.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblCenter.Text = "Center"
		Me.lblCenter.Size = New System.Drawing.Size(33, 17)
		Me.lblCenter.Location = New System.Drawing.Point(80, 40)
		Me.lblCenter.TabIndex = 10
		Me.lblCenter.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblCenter.BackColor = System.Drawing.SystemColors.Control
		Me.lblCenter.Enabled = True
		Me.lblCenter.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblCenter.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblCenter.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblCenter.UseMnemonic = True
		Me.lblCenter.Visible = True
		Me.lblCenter.AutoSize = False
		Me.lblCenter.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblCenter.Name = "lblCenter"
		Me.lblScale.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblScale.Text = "Scale"
		Me.lblScale.Size = New System.Drawing.Size(33, 17)
		Me.lblScale.Location = New System.Drawing.Point(8, 40)
		Me.lblScale.TabIndex = 2
		Me.lblScale.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblScale.BackColor = System.Drawing.SystemColors.Control
		Me.lblScale.Enabled = True
		Me.lblScale.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblScale.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblScale.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblScale.UseMnemonic = True
		Me.lblScale.Visible = True
		Me.lblScale.AutoSize = False
		Me.lblScale.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblScale.Name = "lblScale"
		Me.lblPrinter.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblPrinter.Text = "Printer"
		Me.lblPrinter.Size = New System.Drawing.Size(33, 17)
		Me.lblPrinter.Location = New System.Drawing.Point(8, 8)
		Me.lblPrinter.TabIndex = 1
		Me.lblPrinter.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblPrinter.BackColor = System.Drawing.SystemColors.Control
		Me.lblPrinter.Enabled = True
		Me.lblPrinter.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblPrinter.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblPrinter.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblPrinter.UseMnemonic = True
		Me.lblPrinter.Visible = True
		Me.lblPrinter.AutoSize = False
		Me.lblPrinter.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblPrinter.Name = "lblPrinter"
		Me.Controls.Add(cmdBmp)
		Me.Controls.Add(chkSea)
		Me.Controls.Add(txtOrigin)
		Me.Controls.Add(fraOrientation)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdPrint)
		Me.Controls.Add(cmdPreview)
		Me.Controls.Add(txtScale)
		Me.Controls.Add(cbPrinter)
		Me.Controls.Add(lblCenter)
		Me.Controls.Add(lblScale)
		Me.Controls.Add(lblPrinter)
		Me.fraOrientation.Controls.Add(Option1)
		Me.fraOrientation.Controls.Add(optPortrait)
		Me.fraOrientation.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class