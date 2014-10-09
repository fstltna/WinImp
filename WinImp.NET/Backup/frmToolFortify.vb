Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmToolFortify
	Inherits System.Windows.Forms.Form
	
	'Change Log
	'180206 rjk: Replace ldump with GetLandUnits.
	
	'UPGRADE_WARNING: Arrays in structure rs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Dim rs As DAO.Recordset
	Dim MaxMob As Short
	Dim MobPerUpdate As Short
	Dim strSect As String
	Dim mobavail As Single
	Dim mobused As Single
	Dim fortfact As Single
	Dim cmd() As String
	' Dim ary() As String    8/2003 efj  removed
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim n As Short
		Dim strCmd As String
		
		BuildCommands()
		frmEmpCmd.SubmitEmpireCommand("bf1", False) 'run as batch file
		For n = 1 To UBound(cmd, 2)
			strCmd = "fortify " & VB.Left(cmd(2, n), Len(cmd(2, n)) - 1) & " " & cmd(1, n)
			frmEmpCmd.SubmitEmpireCommand(strCmd, True) 'sectors
		Next n
		frmEmpCmd.SubmitEmpireCommand("bf2", False) 'run as batch file
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		GetLandUnits()
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
		Me.Close()
	End Sub
	
	Private Sub frmToolFortify_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Set always on top
		' Dim success As Long    8/2003 efj  removed
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
		cmdOK.Enabled = False
		
		If Not (rsVersion.BOF And rsVersion.EOF) Then
			MobPerUpdate = rsVersion.Fields("Mob gain - Land").Value
			MaxMob = rsVersion.Fields("Max mob - Land").Value
			rs = rsLand.Clone
			rs.Index = "PrimaryKey"
			FillGrid()
		End If
	End Sub
	
	Public Sub FillGrid()
		
		' Dim n As Integer    8/2003 efj  removed
		' Dim var1 As Variant    8/2003 efj  removed
		' Dim totmobavail As Single    8/2003 efj  removed
		
		'Setup grid
		grid1.Visible = False
		grid1.Clear()
		grid1.WordWrap = True
		grid1.Rows = 2
		grid1.Cols = 8
		grid1.FixedCols = 1
		grid1.set_RowHeight(0, -1) 'return to default
		grid1.set_RowHeight(0, grid1.get_RowHeight(0) * 2)
		
		grid1.set_ColWidth(0, -1) 'return to default
		grid1.set_ColWidth(-1, grid1.get_ColWidth(0))
		grid1.set_ColWidth(0, grid1.get_ColWidth(0) * 1.5)
		grid1.set_ColWidth(1, grid1.get_ColWidth(0) / 2#)
		grid1.set_ColWidth(2, grid1.get_ColWidth(1) / 1.25)
		grid1.set_ColWidth(3, grid1.get_ColWidth(2))
		grid1.set_ColWidth(4, grid1.get_ColWidth(2))
		grid1.set_ColWidth(5, grid1.get_ColWidth(2))
		grid1.set_ColWidth(6, grid1.get_ColWidth(2))
		grid1.set_ColWidth(7, grid1.get_ColWidth(2) * 2.5)
		grid1.set_ColAlignment(0, MSFlexGridLib.AlignmentSettings.flexAlignLeftBottom)
		grid1.set_ColAlignment(1, MSFlexGridLib.AlignmentSettings.flexAlignCenterBottom)
		grid1.set_ColAlignment(2, MSFlexGridLib.AlignmentSettings.flexAlignRightBottom)
		grid1.set_ColAlignment(3, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter)
		grid1.set_ColAlignment(4, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter)
		grid1.set_ColAlignment(5, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter)
		grid1.set_ColAlignment(6, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter)
		grid1.set_ColAlignment(7, MSFlexGridLib.AlignmentSettings.flexAlignLeftBottom)
		
		'Fill in row headers
		grid1.row = 0
		grid1.col = 0
		grid1.CellFontBold = True
		grid1.Text = "Unit ID"
		grid1.col = 1
		grid1.CellFontBold = True
		grid1.Text = "Loc"
		grid1.col = 2
		grid1.CellFontBold = True
		grid1.Text = "Eff"
		grid1.col = 3
		grid1.CellFontBold = True
		grid1.Text = "Curr Mob"
		grid1.col = 4
		grid1.CellFontBold = True
		grid1.Text = "Curr Fort"
		grid1.col = 5
		grid1.CellFontBold = True
		grid1.Text = "Mob used"
		grid1.col = 6
		grid1.CellFontBold = True
		grid1.Text = "New Fort"
		grid1.col = 7
		grid1.CellFontBold = True
		grid1.Text = "  Action"
		
		If rs.EOF And rs.BOF Then
			Exit Sub
		End If
		
		'Find out where the engineers are
		strSect = vbNullString
		rs.MoveFirst()
		While Not rs.EOF
			If rs.Fields(1).Value = "eng" Then
				strSect = strSect & " " & SectorString(rs.Fields("x").Value, rs.Fields("y").Value) & " "
			End If
			rs.MoveNext()
		End While
		
		rs.MoveFirst()
		While Not rs.EOF
			grid1.row = grid1.Rows - 1
			grid1.Rows = grid1.Rows + 1
			
			'Get the unit description
			grid1.col = 0
			grid1.Text = rs.Fields(1).Value ' Set in case lookup fails
			rsBuild.Seek("=", "l", rs.Fields(1))
			If Not rsBuild.NoMatch Then
				If Len(rsBuild.Fields("desc").Value) > 0 Then
					grid1.Text = rsBuild.Fields("desc").Value
				End If
			End If
			grid1.Text = grid1.Text & " #" & CStr(rs.Fields(0).Value)
			
			grid1.col = 1
			grid1.Text = SectorString(rs.Fields("x").Value, rs.Fields("y").Value)
			
			'check for an engineer in the same sector
			If InStr(strSect, grid1.Text) > 0 Then
				fortfact = 1.5
			Else
				fortfact = 1#
			End If
			
			grid1.col = 2
			grid1.Text = CStr(rs.Fields("eff").Value)
			grid1.col = 3
			grid1.Text = CStr(rs.Fields("mob").Value)
			grid1.col = 4
			grid1.Text = CStr(rs.Fields("fort").Value)
			mobused = CShort(0.5 + (MaxMob - rs.Fields("fort").Value) / fortfact)
			mobavail = rs.Fields("mob").Value - (MaxMob - MobPerUpdate)
			If mobavail < 0 Then mobavail = 0
			If mobused > mobavail Then mobused = mobavail
			grid1.col = 5
			grid1.Text = CStr(mobused)
			grid1.col = 6
			grid1.Text = (CStr(rs.Fields("fort").Value + CShort(mobused * fortfact - 0.3)))
			grid1.col = 7
			If MaxMob <= rs.Fields("fort").Value Then
				grid1.Text = "  Fully Fort."
			ElseIf mobused = 0 Then 
				grid1.Text = "  None"
			Else
				grid1.Text = "  Use Excess"
			End If
			rs.MoveNext()
		End While
		
		grid1.Visible = True
		cmdOK.Enabled = True
	End Sub
	
	
	Private Sub frmToolFortify_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	Private Sub grid1_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles grid1.ClickEvent
		
		' Dim n As Integer    8/2003 efj  removed
		Dim id As Short
		' Dim sx As Integer    8/2003 efj  removed
		' Dim sy As Integer    8/2003 efj  removed
		Dim saverow As Short
		
		'If you clicked on a header, resort the grid
		saverow = grid1.MouseRow
		If saverow = 0 Then
			
			'Otherwise, mark the column
		Else
			
			grid1.row = saverow
			grid1.col = 0
			id = CShort(Trim(Mid(grid1.Text, InStr(grid1.Text, "#") + 1)))
			rs.Seek("=", id)
			If rs.NoMatch Then
				Exit Sub
			End If
			
			'check for an engineer in the same sector
			grid1.col = 1
			If InStr(strSect, grid1.Text) > 0 Then
				fortfact = 1.5
			Else
				fortfact = 1#
			End If
			
			grid1.col = 7
			
			'Toggle setting
			Select Case grid1.Text
				Case "  None"
					grid1.Text = "  Use Excess"
					mobavail = rs.Fields("mob").Value - (MaxMob - MobPerUpdate)
					If mobavail < 0 Then
						grid1.Text = "  Use Req"
					End If
				Case "  Use Excess"
					grid1.Text = "  Use Req"
				Case "  Use Req"
					grid1.Text = "  Use All"
				Case "  Use All"
					grid1.Text = "  None"
			End Select
			
			'process new one
			Select Case grid1.Text
				Case "  None"
					mobused = 0
				Case "  Use Excess"
					mobused = CShort(0.5 + (MaxMob - rs.Fields("fort").Value) / fortfact)
					mobavail = rs.Fields("mob").Value - (MaxMob - MobPerUpdate)
					If mobavail < 0 Then mobavail = 0
					If mobused > mobavail Then mobused = mobavail
				Case "  Use Req"
					mobused = CShort(0.5 + (MaxMob - rs.Fields("fort").Value) / fortfact)
					mobavail = rs.Fields("mob").Value
					If mobused >= mobavail Then
						mobused = mobavail
						grid1.Text = "  Use All"
					End If
					
				Case "  Use All"
					mobused = CShort(0.5 + (MaxMob - rs.Fields("fort").Value) / fortfact)
					mobavail = rs.Fields("mob").Value
					If mobused < mobavail Then
						grid1.Text = "  None"
						mobused = 0
					End If
			End Select
			grid1.col = 5
			grid1.Text = CStr(mobused)
			grid1.col = 6
			grid1.Text = CStr(CShort(rs.Fields("fort").Value + CShort(mobused * fortfact - 0.3)))
		End If
		
	End Sub
	
	
	Public Sub BuildCommands()
		
		Dim n As Short
		Dim m As Short
		Dim found As Boolean
		
		
		ReDim cmd(2, 0)
		For n = 1 To grid1.Rows - 1
			grid1.row = n
			grid1.col = 5
			m = 1
			found = False
			If Val(grid1.Text) > 0 Then
				Do While m <= UBound(cmd, 2)
					If cmd(1, m) = grid1.Text Then
						grid1.col = 0
						cmd(2, m) = cmd(2, m) & Trim(Mid(grid1.Text, InStr(grid1.Text, "#") + 1)) & "/"
						found = True
						Exit Do
					End If
					m = m + 1
				Loop 
				If Not found Then
					m = UBound(cmd, 2) + 1
					ReDim Preserve cmd(2, m)
					grid1.col = 5
					cmd(1, m) = grid1.Text
					grid1.col = 0
					cmd(2, m) = cmd(2, m) & Trim(Mid(grid1.Text, InStr(grid1.Text, "#") + 1)) & "/"
				End If
			End If
		Next n
		
	End Sub
End Class