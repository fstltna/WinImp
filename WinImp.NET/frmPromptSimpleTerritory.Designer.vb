<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPromptSimpleTerritory
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
	Public WithEvents cmdTool As System.Windows.Forms.Button
	Public WithEvents txtTerrNumber As System.Windows.Forms.TextBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents cbField As System.Windows.Forms.ComboBox
	Public WithEvents txtMultOrigin As System.Windows.Forms.TextBox
	Public WithEvents cmdHelp As System.Windows.Forms.Button
	Public WithEvents lblTerrNumber As System.Windows.Forms.Label
	Public WithEvents lblSector As System.Windows.Forms.Label
	Public WithEvents lblTerr As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPromptSimpleTerritory))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdTool = New System.Windows.Forms.Button
		Me.txtTerrNumber = New System.Windows.Forms.TextBox
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me.cbField = New System.Windows.Forms.ComboBox
		Me.txtMultOrigin = New System.Windows.Forms.TextBox
		Me.cmdHelp = New System.Windows.Forms.Button
		Me.lblTerrNumber = New System.Windows.Forms.Label
		Me.lblSector = New System.Windows.Forms.Label
		Me.lblTerr = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.ControlBox = False
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.ClientSize = New System.Drawing.Size(273, 129)
		Me.Location = New System.Drawing.Point(1, 1)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmPromptSimpleTerritory"
		Me.cmdTool.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdTool
		Me.cmdTool.Text = "Tool"
		Me.cmdTool.Size = New System.Drawing.Size(65, 25)
		Me.cmdTool.Location = New System.Drawing.Point(200, 96)
		Me.cmdTool.TabIndex = 10
		Me.cmdTool.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdTool.BackColor = System.Drawing.SystemColors.Control
		Me.cmdTool.CausesValidation = True
		Me.cmdTool.Enabled = True
		Me.cmdTool.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdTool.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdTool.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdTool.TabStop = True
		Me.cmdTool.Name = "cmdTool"
		Me.txtTerrNumber.AutoSize = False
		Me.txtTerrNumber.Size = New System.Drawing.Size(89, 19)
		Me.txtTerrNumber.Location = New System.Drawing.Point(104, 96)
		Me.txtTerrNumber.TabIndex = 9
		Me.txtTerrNumber.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtTerrNumber.AcceptsReturn = True
		Me.txtTerrNumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtTerrNumber.BackColor = System.Drawing.SystemColors.Window
		Me.txtTerrNumber.CausesValidation = True
		Me.txtTerrNumber.Enabled = True
		Me.txtTerrNumber.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtTerrNumber.HideSelection = True
		Me.txtTerrNumber.ReadOnly = False
		Me.txtTerrNumber.Maxlength = 0
		Me.txtTerrNumber.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtTerrNumber.MultiLine = False
		Me.txtTerrNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtTerrNumber.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtTerrNumber.TabStop = True
		Me.txtTerrNumber.Visible = True
		Me.txtTerrNumber.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtTerrNumber.Name = "txtTerrNumber"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(65, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(200, 64)
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
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "OK"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Size = New System.Drawing.Size(65, 25)
		Me.cmdOK.Location = New System.Drawing.Point(200, 32)
		Me.cmdOK.TabIndex = 5
		Me.cmdOK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.cbField.Size = New System.Drawing.Size(89, 21)
		Me.cbField.Location = New System.Drawing.Point(104, 64)
		Me.cbField.Items.AddRange(New Object(){"Terr(0)", "Terr1", "Terr2", "Terr3"})
		Me.cbField.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cbField.TabIndex = 3
		Me.cbField.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbField.BackColor = System.Drawing.SystemColors.Window
		Me.cbField.CausesValidation = True
		Me.cbField.Enabled = True
		Me.cbField.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cbField.IntegralHeight = True
		Me.cbField.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbField.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbField.Sorted = False
		Me.cbField.TabStop = True
		Me.cbField.Visible = True
		Me.cbField.Name = "cbField"
		Me.txtMultOrigin.AutoSize = False
		Me.txtMultOrigin.Size = New System.Drawing.Size(89, 19)
		Me.txtMultOrigin.Location = New System.Drawing.Point(104, 32)
		Me.txtMultOrigin.TabIndex = 1
		Me.txtMultOrigin.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtMultOrigin.AcceptsReturn = True
		Me.txtMultOrigin.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtMultOrigin.BackColor = System.Drawing.SystemColors.Window
		Me.txtMultOrigin.CausesValidation = True
		Me.txtMultOrigin.Enabled = True
		Me.txtMultOrigin.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtMultOrigin.HideSelection = True
		Me.txtMultOrigin.ReadOnly = False
		Me.txtMultOrigin.Maxlength = 0
		Me.txtMultOrigin.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtMultOrigin.MultiLine = False
		Me.txtMultOrigin.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtMultOrigin.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtMultOrigin.TabStop = True
		Me.txtMultOrigin.Visible = True
		Me.txtMultOrigin.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtMultOrigin.Name = "txtMultOrigin"
		Me.cmdHelp.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdHelp.Text = "?"
		Me.cmdHelp.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdHelp.Size = New System.Drawing.Size(25, 25)
		Me.cmdHelp.Location = New System.Drawing.Point(248, 0)
		Me.cmdHelp.TabIndex = 0
		Me.ToolTip1.SetToolTip(Me.cmdHelp, "Click for Help")
		Me.cmdHelp.BackColor = System.Drawing.SystemColors.Control
		Me.cmdHelp.CausesValidation = True
		Me.cmdHelp.Enabled = True
		Me.cmdHelp.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdHelp.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdHelp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdHelp.TabStop = True
		Me.cmdHelp.Name = "cmdHelp"
		Me.lblTerrNumber.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblTerrNumber.Text = "Territory Number"
		Me.lblTerrNumber.Size = New System.Drawing.Size(81, 17)
		Me.lblTerrNumber.Location = New System.Drawing.Point(16, 96)
		Me.lblTerrNumber.TabIndex = 8
		Me.lblTerrNumber.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblTerrNumber.BackColor = System.Drawing.SystemColors.Control
		Me.lblTerrNumber.Enabled = True
		Me.lblTerrNumber.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblTerrNumber.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblTerrNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblTerrNumber.UseMnemonic = True
		Me.lblTerrNumber.Visible = True
		Me.lblTerrNumber.AutoSize = False
		Me.lblTerrNumber.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblTerrNumber.Name = "lblTerrNumber"
		Me.lblSector.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblSector.Text = "Sectors"
		Me.lblSector.Size = New System.Drawing.Size(41, 17)
		Me.lblSector.Location = New System.Drawing.Point(56, 32)
		Me.lblSector.TabIndex = 7
		Me.lblSector.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblSector.BackColor = System.Drawing.SystemColors.Control
		Me.lblSector.Enabled = True
		Me.lblSector.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblSector.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblSector.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblSector.UseMnemonic = True
		Me.lblSector.Visible = True
		Me.lblSector.AutoSize = False
		Me.lblSector.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblSector.Name = "lblSector"
		Me.lblTerr.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.lblTerr.Text = "Territory Field"
		Me.lblTerr.Size = New System.Drawing.Size(65, 17)
		Me.lblTerr.Location = New System.Drawing.Point(32, 64)
		Me.lblTerr.TabIndex = 4
		Me.lblTerr.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblTerr.BackColor = System.Drawing.SystemColors.Control
		Me.lblTerr.Enabled = True
		Me.lblTerr.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblTerr.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblTerr.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblTerr.UseMnemonic = True
		Me.lblTerr.Visible = True
		Me.lblTerr.AutoSize = False
		Me.lblTerr.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblTerr.Name = "lblTerr"
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
		Me.Label2.Text = "Territory"
		Me.Label2.Font = New System.Drawing.Font("Arial Black", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(65, 17)
		Me.Label2.Location = New System.Drawing.Point(32, 8)
		Me.Label2.TabIndex = 2
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
		Me.Controls.Add(cmdTool)
		Me.Controls.Add(txtTerrNumber)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(cbField)
		Me.Controls.Add(txtMultOrigin)
		Me.Controls.Add(cmdHelp)
		Me.Controls.Add(lblTerrNumber)
		Me.Controls.Add(lblSector)
		Me.Controls.Add(lblTerr)
		Me.Controls.Add(Label2)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class