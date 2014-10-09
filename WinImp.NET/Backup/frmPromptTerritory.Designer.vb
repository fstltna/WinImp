<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPromptTerritory
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
	Public WithEvents cbField As System.Windows.Forms.ComboBox
	Public WithEvents cmdHelp As System.Windows.Forms.Button
	Public WithEvents cmdAddNew As System.Windows.Forms.Button
	Public WithEvents cbTerr As System.Windows.Forms.ComboBox
	Public WithEvents lstSectors As System.Windows.Forms.ListBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents lblField As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPromptTerritory))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cbField = New System.Windows.Forms.ComboBox
		Me.cmdHelp = New System.Windows.Forms.Button
		Me.cmdAddNew = New System.Windows.Forms.Button
		Me.cbTerr = New System.Windows.Forms.ComboBox
		Me.lstSectors = New System.Windows.Forms.ListBox
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdOK = New System.Windows.Forms.Button
		Me.lblField = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.ControlBox = False
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.ClientSize = New System.Drawing.Size(401, 186)
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
		Me.Name = "frmPromptTerritory"
		Me.cbField.Size = New System.Drawing.Size(73, 21)
		Me.cbField.Location = New System.Drawing.Point(104, 72)
		Me.cbField.Items.AddRange(New Object(){"Terr(0)", "Terr1", "Terr2", "Terr3"})
		Me.cbField.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cbField.TabIndex = 10
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
		Me.cmdHelp.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdHelp.Text = "?"
		Me.cmdHelp.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdHelp.Size = New System.Drawing.Size(25, 25)
		Me.cmdHelp.Location = New System.Drawing.Point(376, 0)
		Me.cmdHelp.TabIndex = 9
		Me.ToolTip1.SetToolTip(Me.cmdHelp, "Click for Help")
		Me.cmdHelp.BackColor = System.Drawing.SystemColors.Control
		Me.cmdHelp.CausesValidation = True
		Me.cmdHelp.Enabled = True
		Me.cmdHelp.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdHelp.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdHelp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdHelp.TabStop = True
		Me.cmdHelp.Name = "cmdHelp"
		Me.cmdAddNew.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdAddNew
		Me.cmdAddNew.Text = "Add New Territory"
		Me.cmdAddNew.Size = New System.Drawing.Size(145, 25)
		Me.cmdAddNew.Location = New System.Drawing.Point(24, 104)
		Me.cmdAddNew.TabIndex = 8
		Me.cmdAddNew.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdAddNew.BackColor = System.Drawing.SystemColors.Control
		Me.cmdAddNew.CausesValidation = True
		Me.cmdAddNew.Enabled = True
		Me.cmdAddNew.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdAddNew.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdAddNew.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdAddNew.TabStop = True
		Me.cmdAddNew.Name = "cmdAddNew"
		Me.cbTerr.Size = New System.Drawing.Size(73, 21)
		Me.cbTerr.Location = New System.Drawing.Point(104, 40)
		Me.cbTerr.Sorted = True
		Me.cbTerr.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cbTerr.TabIndex = 5
		Me.cbTerr.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cbTerr.BackColor = System.Drawing.SystemColors.Window
		Me.cbTerr.CausesValidation = True
		Me.cbTerr.Enabled = True
		Me.cbTerr.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cbTerr.IntegralHeight = True
		Me.cbTerr.Cursor = System.Windows.Forms.Cursors.Default
		Me.cbTerr.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cbTerr.TabStop = True
		Me.cbTerr.Visible = True
		Me.cbTerr.Name = "cbTerr"
		Me.lstSectors.Size = New System.Drawing.Size(193, 150)
		Me.lstSectors.Location = New System.Drawing.Point(200, 32)
		Me.lstSectors.TabIndex = 4
		Me.lstSectors.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstSectors.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstSectors.BackColor = System.Drawing.SystemColors.Window
		Me.lstSectors.CausesValidation = True
		Me.lstSectors.Enabled = True
		Me.lstSectors.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstSectors.IntegralHeight = True
		Me.lstSectors.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstSectors.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstSectors.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstSectors.Sorted = False
		Me.lstSectors.TabStop = True
		Me.lstSectors.Visible = True
		Me.lstSectors.MultiColumn = False
		Me.lstSectors.Name = "lstSectors"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(81, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(8, 144)
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
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "OK"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Size = New System.Drawing.Size(81, 25)
		Me.cmdOK.Location = New System.Drawing.Point(104, 144)
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
		Me.lblField.Text = "Select Territory Field"
		Me.lblField.Size = New System.Drawing.Size(81, 25)
		Me.lblField.Location = New System.Drawing.Point(16, 72)
		Me.lblField.TabIndex = 11
		Me.lblField.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblField.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblField.BackColor = System.Drawing.SystemColors.Control
		Me.lblField.Enabled = True
		Me.lblField.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblField.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblField.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblField.UseMnemonic = True
		Me.lblField.Visible = True
		Me.lblField.AutoSize = False
		Me.lblField.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblField.Name = "lblField"
		Me.Label4.Text = "Edit Territory"
		Me.Label4.Size = New System.Drawing.Size(81, 17)
		Me.Label4.Location = New System.Drawing.Point(16, 40)
		Me.Label4.TabIndex = 7
		Me.Label4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label4.BackColor = System.Drawing.SystemColors.Control
		Me.Label4.Enabled = True
		Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label4.UseMnemonic = True
		Me.Label4.Visible = True
		Me.Label4.AutoSize = False
		Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label4.Name = "Label4"
		Me.Label3.Text = "Sector List"
		Me.Label3.Size = New System.Drawing.Size(137, 17)
		Me.Label3.Location = New System.Drawing.Point(200, 16)
		Me.Label3.TabIndex = 6
		Me.Label3.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label3.BackColor = System.Drawing.SystemColors.Control
		Me.Label3.Enabled = True
		Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label3.UseMnemonic = True
		Me.Label3.Visible = True
		Me.Label3.AutoSize = False
		Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label3.Name = "Label3"
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label1.Text = "Set"
		Me.Label1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Size = New System.Drawing.Size(33, 17)
		Me.Label1.Location = New System.Drawing.Point(8, 16)
		Me.Label1.TabIndex = 3
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
		Me.Label2.Text = "Territory"
		Me.Label2.Font = New System.Drawing.Font("Arial Black", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Size = New System.Drawing.Size(73, 17)
		Me.Label2.Location = New System.Drawing.Point(48, 16)
		Me.Label2.TabIndex = 2
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
		Me.Controls.Add(cbField)
		Me.Controls.Add(cmdHelp)
		Me.Controls.Add(cmdAddNew)
		Me.Controls.Add(cbTerr)
		Me.Controls.Add(lstSectors)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(lblField)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label1)
		Me.Controls.Add(Label2)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class