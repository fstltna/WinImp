Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmToolShow
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'051203 rjk: removed the special case for SHOW_CAP in determining what tech. level to display
	'            now all reports will be at a consisent level.
	'121303 rjk: Added show item. Fixed sorting for SHOW_CAPACITY in SHOW_SECTOR
	'250104 rjk: Added uw as cargo for St@r W@rs
	'270304 rjk: Switched over to DeleteAllRecords for clearing a table
	'261105 rjk: Do not compute avail use the value from the server.
	'281205 rjk: Switch Infrastructure from text to rsBuild table and present in a table.
	'180206 rjk: Add research column for show nukes.
	'120306 rjk: Add support for xdump in Refresh actions.
	'180606 rjk: Updated sector mobility costs for 4.3.6 server changes.
	'140706 rjk: Fixed FillGridItem, it was showing one too many rows.
	
	Dim SortCol As Short
	Dim SortDecending As Boolean
	Dim UnitType As String
	Dim ReportType As String
	Dim ReportTech As Short
	
	Enum sort_type
		TextSort = 0
		NumericSort = 1
	End Enum
	
	Const SHOW_SHIP As String = "s"
	Const SHOW_PLANE As String = "p"
	Const SHOW_LAND As String = "l"
	Const SHOW_NUKE As String = "n"
	Const SHOW_SECTOR As String = "c"
	Const SHOW_BRIDGE As String = "b"
	Const SHOW_TOWER As String = "t"
	Const SHOW_ITEM As String = "i"
	
	Const SHOW_BUILD As String = "b"
	Const SHOW_STAT As String = "s"
	Const SHOW_CAPACITY As String = "c"
	
	Private Sub FillTitle()
		Dim strTemp As String
		
		strTemp = "Show "
		Select Case UnitType
			Case SHOW_SHIP
				strTemp = strTemp & "Ship "
			Case SHOW_PLANE
				strTemp = strTemp & "Plane "
			Case SHOW_LAND
				strTemp = strTemp & "Land "
			Case SHOW_NUKE
				strTemp = strTemp & "Nuke "
			Case SHOW_SECTOR
				strTemp = strTemp & "Sector "
			Case SHOW_BRIDGE
				strTemp = strTemp & "Bridge/Bridge Tower "
			Case SHOW_ITEM
				strTemp = strTemp & "Item "
		End Select
		
		Select Case ReportType
			Case SHOW_BUILD
				strTemp = strTemp & "Build"
			Case SHOW_STAT
				strTemp = strTemp & "Statistics"
			Case SHOW_CAPACITY
				strTemp = strTemp & "Capacities"
		End Select
		
		Me.Text = strTemp
	End Sub
	
	Public Sub FillGrid()
		FillTitle()
		SetReportTech()
		Select Case ReportType
			Case SHOW_BUILD
				Select Case UnitType
					Case SHOW_SHIP, SHOW_LAND, SHOW_PLANE, SHOW_NUKE
						FillGridBuild()
					Case SHOW_SECTOR
						FillGridBuildSector()
					Case SHOW_BRIDGE
						FillTextBox()
					Case SHOW_ITEM
						FillGridItem()
				End Select
			Case SHOW_STAT
				Select Case UnitType
					Case SHOW_SHIP, SHOW_LAND, SHOW_PLANE, SHOW_NUKE
						FillGridStat()
					Case SHOW_BRIDGE
						FillTextBox()
					Case SHOW_SECTOR
						FillGridSector()
					Case SHOW_ITEM
						FillGridItem()
				End Select
			Case SHOW_CAPACITY
				Select Case UnitType
					Case SHOW_SHIP, SHOW_LAND, SHOW_PLANE
						FillGridCap()
					Case SHOW_NUKE
						FillGridStat()
					Case SHOW_BRIDGE
						FillTextBox()
					Case SHOW_SECTOR
						FillGridSector()
					Case SHOW_ITEM
						FillGridItem()
				End Select
		End Select
	End Sub
	
	Private Sub FillGridBuild()
		Dim n As Short
		Dim l As Integer
		Dim MaxColumn As Short
		Dim AvailCol As Short
		Dim hldIndex As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hldIndex = rsBuild.Index
		rsBuild.Index = "tech"
		
		'Reset sort variables
		SortCol = -1
		SortDecending = True
		
		TextBox1.Visible = False
		grid1.Visible = False
		grid1.Clear()
		grid1.Rows = 2
		grid1.Cols = 6
		grid1.FixedCols = 1
		grid1.FixedRows = 1
		
		Select Case UnitType
			Case SHOW_SHIP
				MaxColumn = 6
				AvailCol = 4
			Case SHOW_LAND
				MaxColumn = 7
				AvailCol = 5
				grid1.set_TextMatrix(0, AvailCol - 1, "guns")
			Case SHOW_PLANE
				MaxColumn = 7
				AvailCol = 5
				grid1.set_TextMatrix(0, AvailCol - 1, "crew")
			Case SHOW_NUKE
				MaxColumn = 9
				AvailCol = 6
				grid1.set_TextMatrix(0, AvailCol - 2, "oil")
				grid1.set_TextMatrix(0, AvailCol - 1, "rad")
		End Select
		
		grid1.Cols = MaxColumn + 1
		'Fill headers
		grid1.set_TextMatrix(0, 0, "type")
		grid1.set_TextMatrix(0, 1, "description")
		grid1.set_TextMatrix(0, 2, "lcm")
		grid1.set_TextMatrix(0, 3, "hcm")
		grid1.set_TextMatrix(0, AvailCol, "avail")
		grid1.set_TextMatrix(0, AvailCol + 1, "tech")
		If UnitType = SHOW_NUKE Then
			grid1.set_TextMatrix(0, AvailCol + 2, "res")
			grid1.set_TextMatrix(0, AvailCol + 3, "cost")
		Else
			grid1.set_TextMatrix(0, AvailCol + 2, "cost")
		End If
		
		grid1.set_ColWidth(-1, -1)
		l = grid1.get_ColWidth(0) / 2
		grid1.set_ColWidth(0, l)
		grid1.set_ColWidth(1, l * 3)
		grid1.set_ColData(0, sort_type.TextSort)
		grid1.set_ColData(1, sort_type.TextSort)
		grid1.row = 0
		For n = 2 To MaxColumn
			grid1.set_ColWidth(n, l)
			grid1.set_ColData(n, sort_type.NumericSort)
			grid1.col = n
			grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
		Next n
		grid1.set_ColWidth(MaxColumn, l * 1.5)
		'Fill in rows
		n = 1
		If Not (rsBuild.BOF And rsBuild.EOF) Then
			rsBuild.MoveFirst()
			While Not rsBuild.EOF
				If rsBuild.Fields("type").Value = UnitType Then
					grid1.set_TextMatrix(n, 0, rsBuild.Fields("id").Value)
					grid1.set_TextMatrix(n, 1, rsBuild.Fields("desc").Value)
					grid1.set_TextMatrix(n, 2, rsBuild.Fields("lcm").Value)
					grid1.set_TextMatrix(n, 3, rsBuild.Fields("hcm").Value)
					grid1.set_TextMatrix(n, AvailCol, rsBuild.Fields("avail").Value)
					grid1.set_TextMatrix(n, AvailCol + 1, rsBuild.Fields("tech").Value)
					grid1.set_TextMatrix(n, AvailCol + 2, rsBuild.Fields("cost").Value)
					If UnitType = SHOW_LAND Then
						grid1.set_TextMatrix(n, AvailCol - 1, rsBuild.Fields("gun").Value)
					ElseIf UnitType = SHOW_PLANE Then 
						grid1.set_TextMatrix(n, AvailCol - 1, rsBuild.Fields("mil").Value)
					ElseIf UnitType = SHOW_NUKE Then 
						grid1.set_TextMatrix(n, AvailCol - 1, rsBuild.Fields("rad").Value)
						grid1.set_TextMatrix(n, AvailCol - 2, rsBuild.Fields("oil").Value)
						grid1.set_TextMatrix(n, AvailCol + 2, rsBuild.Fields("stat5").Value)
						grid1.set_TextMatrix(n, AvailCol + 3, rsBuild.Fields("cost").Value)
					End If
					
					grid1.row = n
					If rsBuild.Fields("tech").Value > TechLevel Then
						grid1.row = n
						For l = 0 To MaxColumn
							grid1.col = l
							grid1.CellForeColor = System.Drawing.Color.Red
						Next 
					End If
					
					For l = 2 To MaxColumn
						grid1.col = l
						grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
					Next l
					
					n = n + 1
					grid1.Rows = n + 1
				End If
				rsBuild.MoveNext()
			End While
		End If
		grid1.Visible = True
		
		rsShowText.Seek("=", UnitType, "b", 1)
		If Not rsShowText.NoMatch Then
			sb1.Text = "Current tech level is '" & CStr(TechLevel) & "'.   " + rsShowText.Fields("Text").Value
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsBuild.Index = hldIndex
	End Sub
	
	Private Sub FillGridStat()
		Dim n As Short
		Dim l As Integer
		Dim MaxColumn As Short
		Dim hldIndex As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hldIndex = rsBuild.Index
		rsBuild.Index = "tech"
		
		'Reset sort variables
		SortCol = -1
		SortDecending = True
		
		TextBox1.Visible = False
		grid1.Visible = False
		grid1.Clear()
		grid1.Rows = 2
		grid1.FixedCols = 1
		grid1.FixedRows = 1
		
		Select Case UnitType
			Case SHOW_SHIP
				MaxColumn = 11
				grid1.Cols = MaxColumn + 1
				grid1.set_TextMatrix(0, 2, "def.")
				grid1.set_TextMatrix(0, 3, "spd")
				grid1.set_TextMatrix(0, 4, "visib.")
				grid1.set_TextMatrix(0, 5, "spy")
				grid1.set_TextMatrix(0, 6, "range")
				grid1.set_TextMatrix(0, 7, "fire")
				grid1.set_TextMatrix(0, 8, "land")
				grid1.set_TextMatrix(0, 9, "plane")
				grid1.set_TextMatrix(0, 10, "helo")
				grid1.set_TextMatrix(0, 11, "xplane")
			Case SHOW_LAND
				MaxColumn = 17
				grid1.Cols = MaxColumn + 1
				grid1.set_TextMatrix(0, 2, "att.")
				grid1.set_TextMatrix(0, 3, "def.")
				grid1.set_TextMatrix(0, 4, "vul.")
				grid1.set_TextMatrix(0, 5, "spd")
				grid1.set_TextMatrix(0, 6, "visib.")
				grid1.set_TextMatrix(0, 7, "spy")
				grid1.set_TextMatrix(0, 8, "radius")
				grid1.set_TextMatrix(0, 9, "range")
				grid1.set_TextMatrix(0, 10, "acc.")
				grid1.set_TextMatrix(0, 11, "fire")
				grid1.set_TextMatrix(0, 12, "ammo")
				grid1.set_TextMatrix(0, 13, "anti-air")
				grid1.set_TextMatrix(0, 14, "f cap")
				grid1.set_TextMatrix(0, 15, "f use")
				grid1.set_TextMatrix(0, 16, "xplane")
				grid1.set_TextMatrix(0, 17, "land")
			Case SHOW_PLANE
				MaxColumn = 8
				grid1.Cols = MaxColumn + 1
				grid1.set_TextMatrix(0, 2, "acc.")
				grid1.set_TextMatrix(0, 3, "load")
				grid1.set_TextMatrix(0, 4, "att.")
				grid1.set_TextMatrix(0, 5, "def.")
				grid1.set_TextMatrix(0, 6, "range")
				grid1.set_TextMatrix(0, 7, "fuel")
				grid1.set_TextMatrix(0, 8, "stlth%")
			Case SHOW_NUKE
				MaxColumn = 4
				grid1.Cols = MaxColumn + 5
				grid1.set_TextMatrix(0, 2, "blast")
				grid1.set_TextMatrix(0, 3, "dam.")
				grid1.set_TextMatrix(0, 4, "lbs")
				grid1.set_TextMatrix(0, 5, "tech")
				grid1.set_TextMatrix(0, 6, "res")
				grid1.set_TextMatrix(0, 7, "cost")
				grid1.set_TextMatrix(0, 8, "abilities")
		End Select
		'Fill headers
		grid1.set_TextMatrix(0, 0, "type")
		grid1.set_TextMatrix(0, 1, "description")
		
		grid1.set_ColWidth(-1, -1)
		l = grid1.get_ColWidth(0) / 2
		grid1.set_ColWidth(0, l)
		grid1.set_ColWidth(1, l * 3)
		grid1.set_ColData(0, sort_type.TextSort)
		grid1.set_ColData(1, sort_type.TextSort)
		grid1.row = 0
		For n = 2 To MaxColumn
			grid1.set_ColWidth(n, l * 1.25)
			grid1.set_ColData(n, sort_type.NumericSort)
			grid1.col = n
			grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
		Next n
		If UnitType = SHOW_NUKE Then
			grid1.set_ColWidth(MaxColumn + 1, l * 1.25)
			grid1.col = MaxColumn + 1
			grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
			grid1.set_ColWidth(MaxColumn + 2, l * 1.25)
			grid1.col = MaxColumn + 2
			grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
			grid1.col = MaxColumn + 3
			grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
			grid1.col = MaxColumn + 4
			grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
			grid1.set_ColData(MaxColumn + 4, sort_type.TextSort)
		End If
		
		'Fill in rows
		n = 1
		If Not (rsBuild.BOF And rsBuild.EOF) Then
			rsBuild.MoveFirst()
			While Not rsBuild.EOF
				If rsBuild.Fields("type").Value = UnitType Then
					grid1.set_TextMatrix(n, 0, rsBuild.Fields("id").Value)
					grid1.set_TextMatrix(n, 1, rsBuild.Fields("desc").Value)
					For l = 2 To MaxColumn
						If l < 4 And UnitType = SHOW_LAND Then
							grid1.set_TextMatrix(n, l, VB6.Format(rsBuild.Fields("stat" & CStr(l - 1)).Value, "###0.00"))
						Else
							grid1.set_TextMatrix(n, l, CStr(rsBuild.Fields("stat" & CStr(l - 1)).Value))
						End If
					Next l
					
					If UnitType = SHOW_NUKE Then
						grid1.set_TextMatrix(n, MaxColumn + 1, rsBuild.Fields("tech").Value)
						grid1.col = MaxColumn + 1
						grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
						grid1.set_TextMatrix(n, MaxColumn + 2, rsBuild.Fields("stat5").Value)
						grid1.col = MaxColumn + 2
						grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
						grid1.set_TextMatrix(n, MaxColumn + 3, rsBuild.Fields("cost").Value)
						grid1.col = MaxColumn + 3
						grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
						grid1.set_TextMatrix(n, MaxColumn + 4, rsBuild.Fields("cap1").Value)
						grid1.col = MaxColumn + 4
						grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
					End If
					
					grid1.row = n
					If rsBuild.Fields("tech").Value > TechLevel Then
						For l = 0 To MaxColumn
							grid1.col = l
							grid1.CellForeColor = System.Drawing.Color.Red
						Next 
						If UnitType = SHOW_NUKE Then
							For l = MaxColumn + 1 To MaxColumn + 4
								grid1.col = l
								grid1.CellForeColor = System.Drawing.Color.Red
							Next 
						End If
					End If
					
					For l = 2 To MaxColumn
						grid1.col = l
						grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
					Next l
					
					n = n + 1
					grid1.Rows = n + 1
				End If
				rsBuild.MoveNext()
			End While
		End If
		grid1.Visible = True
		
		rsShowText.Seek("=", UnitType, ReportType, 1)
		If Not rsShowText.NoMatch Then
			sb1.Text = "Current tech level is '" & CStr(TechLevel) & "'.   " + rsShowText.Fields("Text").Value
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsBuild.Index = hldIndex
	End Sub
	
	Private Sub FillGridCap()
		Dim n As Short
		Dim l As Integer
		Dim MaxColumn As Short
		' Dim AvailCol As Integer    8/2003 efj  removed
		Dim strTemp As Object 'drk@unxsoft.com 8/11/02.  Seems to be required to get stuff to co-operate below...
		Dim strType As String
		Dim CargoColumn As Short
		Dim hldIndex As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hldIndex = rsBuild.Index
		rsBuild.Index = "tech"
		
		'Reset sort variables
		SortCol = -1
		SortDecending = True
		
		TextBox1.Visible = False
		grid1.Visible = False
		grid1.Clear()
		grid1.Rows = 2
		grid1.FixedCols = 1
		grid1.FixedRows = 1
		
		Select Case UnitType
			Case SHOW_SHIP
				MaxColumn = 9
				grid1.Cols = MaxColumn + 1
				CargoColumn = MaxColumn - 1
				grid1.set_TextMatrix(0, 2, "mil")
				grid1.set_TextMatrix(0, 3, "gun")
				grid1.set_TextMatrix(0, 4, "shell")
				grid1.set_TextMatrix(0, 5, "food")
				grid1.set_TextMatrix(0, 6, "civ")
				grid1.set_TextMatrix(0, 7, "uw")
				grid1.set_TextMatrix(0, 8, "cargo")
				grid1.set_TextMatrix(0, 9, "abilities")
			Case SHOW_LAND
				MaxColumn = 6
				grid1.Cols = MaxColumn + 1
				CargoColumn = MaxColumn
				grid1.set_TextMatrix(0, 2, "mil")
				grid1.set_TextMatrix(0, 3, "gun")
				grid1.set_TextMatrix(0, 4, "shell")
				grid1.set_TextMatrix(0, 5, "food")
				grid1.set_TextMatrix(0, 6, "cargo/abilities")
			Case SHOW_PLANE
				MaxColumn = 2
				CargoColumn = 2
				grid1.Cols = MaxColumn + 1
				grid1.set_TextMatrix(0, 2, "abilities")
		End Select
		'Fill headers
		grid1.set_TextMatrix(0, 0, "type")
		grid1.set_TextMatrix(0, 1, "description")
		
		grid1.set_ColWidth(-1, -1)
		l = grid1.get_ColWidth(0) / 2
		grid1.set_ColWidth(0, l)
		grid1.set_ColWidth(1, l * 3)
		grid1.set_ColData(0, sort_type.TextSort)
		grid1.set_ColData(1, sort_type.TextSort)
		grid1.row = 0
		For n = 2 To MaxColumn
			grid1.set_ColWidth(n, l)
			If n < CargoColumn Then
				grid1.set_ColData(n, sort_type.NumericSort)
			Else
				grid1.set_ColData(n, sort_type.TextSort)
			End If
			grid1.col = n
			grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
		Next n
		grid1.set_ColWidth(MaxColumn, l * 10)
		grid1.col = MaxColumn
		grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
		
		If UnitType = SHOW_SHIP Then
			grid1.set_ColWidth(CargoColumn, l * 3)
		End If
		grid1.col = CargoColumn
		grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
		'Fill in rows
		n = 1
		If Not (rsBuild.BOF And rsBuild.EOF) Then
			rsBuild.MoveFirst()
			While Not rsBuild.EOF
				If rsBuild.Fields("type").Value = UnitType Then
					grid1.set_TextMatrix(n, 0, rsBuild.Fields("id").Value)
					grid1.set_TextMatrix(n, 1, rsBuild.Fields("desc").Value)
					For l = 2 To CargoColumn - 1
						grid1.set_TextMatrix(n, l, "0")
					Next l
					For l = 1 To 20
						'UPGRADE_WARNING: Couldn't resolve default property of object strTemp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						strTemp = rsBuild.Fields("cap" & CStr(l)).Value
						If Len(strTemp) > 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object strTemp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If VB.Left(strTemp, 1) >= "1" And VB.Left(strTemp, 1) <= "9" Then
								'UPGRADE_WARNING: Couldn't resolve default property of object strTemp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								strType = VB.Right(strTemp, 1)
								Select Case strType
									Case "m"
										'UPGRADE_WARNING: Couldn't resolve default property of object strTemp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										grid1.set_TextMatrix(n, 2, VB.Left(strTemp, Len(strTemp) - 1))
									Case "g"
										'UPGRADE_WARNING: Couldn't resolve default property of object strTemp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										grid1.set_TextMatrix(n, 3, VB.Left(strTemp, Len(strTemp) - 1))
									Case "s"
										'UPGRADE_WARNING: Couldn't resolve default property of object strTemp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										grid1.set_TextMatrix(n, 4, VB.Left(strTemp, Len(strTemp) - 1))
									Case "f"
										'UPGRADE_WARNING: Couldn't resolve default property of object strTemp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										grid1.set_TextMatrix(n, 5, VB.Left(strTemp, Len(strTemp) - 1))
									Case "c"
										If UnitType = SHOW_LAND Then '090903 rjk: LOTRII Civilians are allowed as cargo
											'UPGRADE_WARNING: Couldn't resolve default property of object strTemp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											grid1.set_TextMatrix(n, CargoColumn, grid1.get_TextMatrix(n, CargoColumn) + strTemp + " ")
										Else
											'UPGRADE_WARNING: Couldn't resolve default property of object strTemp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											grid1.set_TextMatrix(n, 6, VB.Left(strTemp, Len(strTemp) - 1))
										End If
									Case "u"
										If UnitType = SHOW_LAND Then '240104 rjk: St@r W@rs UW are allowed as cargo
											'UPGRADE_WARNING: Couldn't resolve default property of object strTemp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											grid1.set_TextMatrix(n, CargoColumn, grid1.get_TextMatrix(n, CargoColumn) + strTemp + " ")
										Else
											'UPGRADE_WARNING: Couldn't resolve default property of object strTemp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											grid1.set_TextMatrix(n, 7, VB.Left(strTemp, Len(strTemp) - 1))
										End If
									Case Else
										'UPGRADE_WARNING: Couldn't resolve default property of object strTemp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										grid1.set_TextMatrix(n, CargoColumn, grid1.get_TextMatrix(n, CargoColumn) + strTemp + " ")
								End Select
							Else
								'UPGRADE_WARNING: Couldn't resolve default property of object strTemp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								grid1.set_TextMatrix(n, MaxColumn, grid1.get_TextMatrix(n, MaxColumn) + strTemp + " ")
							End If
						End If
					Next l
					
					grid1.row = n
					If rsBuild.Fields("tech").Value > TechLevel Then
						For l = 0 To MaxColumn
							grid1.col = l
							grid1.CellForeColor = System.Drawing.Color.Red
						Next 
					End If
					
					For l = 2 To MaxColumn - 1
						grid1.col = l
						grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
					Next l
					grid1.col = CargoColumn
					grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
					grid1.col = MaxColumn
					grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
					
					n = n + 1
					grid1.Rows = n + 1
				End If
				rsBuild.MoveNext()
			End While
		End If
		grid1.Visible = True
		
		rsShowText.Seek("=", UnitType, "c", 1)
		If Not rsShowText.NoMatch Then
			sb1.Text = "Current tech level is '" & CStr(TechLevel) & "'.   " + rsShowText.Fields("Text").Value
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsBuild.Index = hldIndex
	End Sub
	
	Private Sub cmdClear_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClear.Click
		Dim Result As Short
		
		Result = MsgBox("Do you want to clear ALL unit data " & "stored for this game, including higher tech defaults?", MsgBoxStyle.YesNo, "Clear Confirmation")
		
		If Result = MsgBoxResult.Yes Then
			DeleteAllRecords(rsBuild)
			
			If Not (rsShowText.EOF And rsShowText.BOF) Then
				rsShowText.MoveFirst()
			End If
			
			While Not rsShowText.EOF
				If InStr("spln", rsShowText.Fields("unittype").Value) > 0 Then
					If rsShowText.Fields("linenumber").Value = 1 Then
						rsShowText.Edit()
						rsShowText.Fields("text").Value = "No data available - use Refresh"
						rsShowText.Update()
					Else
						rsShowText.Delete()
					End If
				End If
				rsShowText.MoveNext()
			End While
		End If
		
		Result = MsgBox("Do you want to clear sector and bridge data " & "stored for this game.  If you choose Yes, you will need to Refresh sector " & " data for all of WinACE's functions to work correctly", MsgBoxStyle.YesNo, "Clear Confirmation")
		
		If Result = MsgBoxResult.Yes Then
			DeleteAllRecords(rsSectorType)
			
			If Not (rsShowText.EOF And rsShowText.BOF) Then
				rsShowText.MoveFirst()
			End If
			
			While Not rsShowText.EOF
				If InStr("cbt", rsShowText.Fields("unittype").Value) > 0 Then
					If rsShowText.Fields("linenumber").Value = 1 Then
						rsShowText.Edit()
						rsShowText.Fields("text").Value = "No data available - use Refresh"
						rsShowText.Update()
					Else
						rsShowText.Delete()
					End If
				End If
				rsShowText.MoveNext()
			End While
		End If
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Me.Close()
	End Sub
	
	Private Function BuildCommand(ByRef unit As String, ByRef report As String) As String
		Select Case unit
			Case SHOW_LAND
				BuildCommand = "land"
			Case SHOW_PLANE
				BuildCommand = "plane"
			Case SHOW_SHIP
				BuildCommand = "ship"
			Case SHOW_BRIDGE
				BuildCommand = "bridge"
			Case SHOW_TOWER
				BuildCommand = "tower"
			Case SHOW_NUKE
				BuildCommand = "nuke"
			Case SHOW_SECTOR
				BuildCommand = "sector"
			Case SHOW_ITEM
				BuildCommand = "item"
		End Select
		BuildCommand = "show " & BuildCommand & " " & report & " "
		If Len(txtTech.Text) > 0 Then
			BuildCommand = BuildCommand & " " & txtTech.Text
			' ElseIf report = SHOW_STAT Then 'rjk 05/12/03: removed this special case, so all reports are displayed at a consisent tech. level.
			'    BuildCommand = BuildCommand + " " + CStr(TechLevel)
		End If
	End Function
	
	Private Function BuildXdumpCommand(ByRef unit As String) As String
		Select Case unit
			Case SHOW_LAND
				BuildXdumpCommand = "land-chr"
			Case SHOW_PLANE
				BuildXdumpCommand = "plane-chr"
			Case SHOW_SHIP
				BuildXdumpCommand = "ship-chr"
			Case SHOW_BRIDGE
				BuildXdumpCommand = "bridge-chr"
			Case SHOW_TOWER
				BuildXdumpCommand = "tower"
			Case SHOW_NUKE
				BuildXdumpCommand = "nuke-chr"
			Case SHOW_SECTOR
				BuildXdumpCommand = "sector-chr"
			Case SHOW_BRIDGE, SHOW_TOWER
				BuildXdumpCommand = "ver"
			Case SHOW_ITEM
				BuildXdumpCommand = "item"
		End Select
		BuildXdumpCommand = "xdump " & BuildXdumpCommand & " *"
	End Function
	
	Private Sub cmdRefresh_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRefresh.Click
		If VersionCheck(4, 3, 0) = WinAceRoutines.enumVersion.VER_LESS Then
			frmEmpCmd.SubmitEmpireCommand(BuildCommand(UnitType, ReportType), True)
			If UnitType = SHOW_BRIDGE Then
				frmEmpCmd.SubmitEmpireCommand(BuildCommand(SHOW_TOWER, ReportType), True)
			End If
		Else
			frmEmpCmd.SubmitEmpireCommand(BuildXdumpCommand(UnitType), True)
		End If
	End Sub
	
	Public Sub cmdRefreshAll_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRefreshAll.Click
		If VersionCheck(4, 3, 0) = WinAceRoutines.enumVersion.VER_LESS Then
			cmdRefreshUnits_Click(cmdRefreshUnits, New System.EventArgs())
			frmEmpCmd.SubmitEmpireCommand(BuildCommand("n", "b"), False)
			frmEmpCmd.SubmitEmpireCommand(BuildCommand("n", "s"), False)
			frmEmpCmd.SubmitEmpireCommand(BuildCommand("n", "c"), False)
			frmEmpCmd.SubmitEmpireCommand(BuildCommand("c", "b"), False)
			frmEmpCmd.SubmitEmpireCommand(BuildCommand("c", "s"), False)
			frmEmpCmd.SubmitEmpireCommand(BuildCommand("c", "c"), False)
			frmEmpCmd.SubmitEmpireCommand(BuildCommand("b", "b"), False)
			frmEmpCmd.SubmitEmpireCommand(BuildCommand("t", "b"), False)
			frmEmpCmd.SubmitEmpireCommand(BuildCommand("i", "b"), False)
		Else
			RequestConfigurationTables()
		End If
	End Sub
	
	Private Sub cmdRefreshUnits_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRefreshUnits.Click
		If VersionCheck(4, 3, 0) = WinAceRoutines.enumVersion.VER_LESS Then
			frmEmpCmd.SubmitEmpireCommand(BuildCommand("s", "b"), False)
			frmEmpCmd.SubmitEmpireCommand(BuildCommand("s", "s"), False)
			frmEmpCmd.SubmitEmpireCommand(BuildCommand("s", "c"), False)
			frmEmpCmd.SubmitEmpireCommand(BuildCommand("l", "b"), False)
			frmEmpCmd.SubmitEmpireCommand(BuildCommand("l", "s"), False)
			frmEmpCmd.SubmitEmpireCommand(BuildCommand("l", "c"), False)
			frmEmpCmd.SubmitEmpireCommand(BuildCommand("p", "b"), False)
			frmEmpCmd.SubmitEmpireCommand(BuildCommand("p", "s"), False)
			frmEmpCmd.SubmitEmpireCommand(BuildCommand("p", "c"), False)
		Else
			RequestUnitConfigurationTables()
		End If
	End Sub
	
	Private Sub frmToolShow_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		'Set always on top
		' Dim success As Long    8/2003 efj  removed
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
		
		'Restore Toolbar
		'UPGRADE_ISSUE: MSComctlLib.Toolbar method Toolbar1.RestoreToolbar was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		Toolbar1.RestoreToolBar(APPNAME, "Layout", "ToolShow Toolbar")
		
		SetBounds(VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2), VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		
		
		UnitType = SHOW_SHIP
		ReportType = SHOW_BUILD
		txtTech.Text = vbNullString
		FillGrid()
		
	End Sub
	
	'UPGRADE_WARNING: Event frmToolShow.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmToolShow_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		
		If Me.WindowState = System.Windows.Forms.FormWindowState.Minimized Then Exit Sub
		
		If VB6.PixelsToTwipsX(Me.Width) < VB6.PixelsToTwipsX(cmdOK.Width) * 5 Then
			Me.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(cmdOK.Width) * 5)
		End If
		
		If VB6.PixelsToTwipsY(Me.Height) < VB6.PixelsToTwipsY(cmdOK.Top) + VB6.PixelsToTwipsY(cmdOK.Height) * 3 Then
			Me.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(cmdOK.Top) + VB6.PixelsToTwipsY(cmdOK.Height) * 3)
		End If
		
		With Me
			grid1.SetBounds(0, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Toolbar1.Top) + VB6.PixelsToTwipsY(Toolbar1.Height)), VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Width) - VB6.PixelsToTwipsX(cmdOK.Width) * 1.5), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(.sb1.Top) - VB6.PixelsToTwipsY(Toolbar1.Height)))
			'UPGRADE_NOTE: Refresh was upgraded to CtlRefresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
			grid1.CtlRefresh()
			Toolbar1.Width = .Width
			cmdOK.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(.Width) - VB6.PixelsToTwipsX(cmdOK.Width) * 1.25)
			cmdRefresh.Left = cmdOK.Left
			cmdRefreshUnits.Left = cmdOK.Left
			cmdRefreshAll.Left = cmdOK.Left
			cmdClear.Left = cmdOK.Left
			Label1.Left = cmdOK.Left
			txtTech.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(cmdOK.Left) + VB6.PixelsToTwipsX(cmdOK.Width) - VB6.PixelsToTwipsX(txtTech.Width))
		End With
		
		With grid1
			TextBox1.SetBounds(.Left, .Top, .Width, .Height)
		End With
		TextBox1.Visible = False
		
		
	End Sub
	
	Private Sub frmToolShow_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_ISSUE: MSComctlLib.Toolbar method Toolbar1.SaveToolbar was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		Toolbar1.SaveToolBar(APPNAME, "Layout", "ToolShow Toolbar")
	End Sub
	
	Private Sub grid1_MouseUpEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSFlexGridLib.DMSFlexGridEvents_MouseUpEvent) Handles grid1.MouseUpEvent
		'If clicking in top row, resort current box
		If grid1.MouseRow = 0 Then
			SortGrid((grid1.MouseCol))
			Exit Sub
		End If
	End Sub
	
	Private Sub Toolbar1_ButtonClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _Toolbar1_Button1.Click, _Toolbar1_Button2.Click, _Toolbar1_Button3.Click, _Toolbar1_Button4.Click, _Toolbar1_Button5.Click, _Toolbar1_Button6.Click, _Toolbar1_Button7.Click, _Toolbar1_Button8.Click, _Toolbar1_Button9.Click, _Toolbar1_Button10.Click, _Toolbar1_Button11.Click, _Toolbar1_Button12.Click, _Toolbar1_Button13.Click, _Toolbar1_Button14.Click
		Dim Button As System.Windows.Forms.ToolStripItem = CType(eventSender, System.Windows.Forms.ToolStripItem)
		Select Case Button.Name
			Case "Ship"
				UnitType = SHOW_SHIP
			Case "Plane"
				UnitType = SHOW_PLANE
			Case "Land"
				UnitType = SHOW_LAND
			Case "Nuke"
				UnitType = SHOW_NUKE
			Case "Bridge"
				UnitType = SHOW_BRIDGE
			Case "Sector"
				UnitType = SHOW_SECTOR
			Case "Item"
				UnitType = SHOW_ITEM
			Case "Build"
				ReportType = SHOW_BUILD
			Case "Stat"
				ReportType = SHOW_STAT
			Case "Cap"
				ReportType = SHOW_CAPACITY
		End Select
		
		FillGrid()
	End Sub
	
	Private Sub FillTextBox()
		
		' Dim n As Integer    8/2003 efj  removed
		' Dim l As Long    8/2003 efj  removed
		' Dim MaxColumn As Integer    8/2003 efj  removed
		' Dim AvailCol As Integer    8/2003 efj  removed
		
		'Reset sort variables
		SortCol = -1
		SortDecending = True
		
		TextBox1.Visible = False
		grid1.Visible = False
		TextBox1.Text = vbNullString
		
		If Not (rsShowText.BOF And rsShowText.EOF) Then
			rsShowText.MoveFirst()
		End If
		
		While Not rsShowText.EOF
			If (rsShowText.Fields("UnitType").Value = UnitType Or UnitType = SHOW_BRIDGE And rsShowText.Fields("UnitType").Value = "t") And rsShowText.Fields("reporttype").Value = SHOW_BUILD And rsShowText.Fields("Linenumber").Value > 1 Then
				TextBox1.Text = TextBox1.Text + rsShowText.Fields("Text").Value + vbCrLf
			End If
			rsShowText.MoveNext()
		End While
		
		TextBox1.Visible = True
	End Sub
	
	Private Sub FillGridBuildSector()
		Dim n As Short
		Dim l As Integer
		Dim MaxColumn As Short
		Dim AvailCol As Short
		Dim hldIndex As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hldIndex = rsBuild.Index
		rsBuild.Index = "tech"
		
		'Reset sort variables
		SortCol = -1
		SortDecending = True
		
		TextBox1.Visible = False
		grid1.Visible = False
		grid1.Clear()
		grid1.Rows = 1
		grid1.Cols = 6
		grid1.FixedCols = 1
		'grid1.FixedRows = 1
		
		'sector type    cost to des    cost for 1% eff   lcms for 1%    hcms for 1%
		'Fill First header
		grid1.set_TextMatrix(0, 0, "sector type")
		grid1.set_TextMatrix(0, 1, "$$$ - des")
		grid1.set_TextMatrix(0, 2, "$$$ - 1%")
		grid1.set_TextMatrix(0, 3, "lcms - 1%")
		grid1.set_TextMatrix(0, 4, "hcms - 1%")
		grid1.set_TextMatrix(0, 5, "mobility - 1%")
		
		grid1.set_ColWidth(-1, -1)
		grid1.set_ColData(0, sort_type.TextSort)
		grid1.col = 0
		grid1.row = 0
		grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
		For n = 1 To grid1.Cols - 1
			grid1.set_ColData(n, sort_type.NumericSort)
			grid1.col = n
			grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
		Next n
		'Fill in rows
		n = 1
		If Not (rsBuild.BOF And rsBuild.EOF) Then
			rsBuild.MoveFirst()
			While Not rsBuild.EOF
				If rsBuild.Fields("type").Value = "i" Then
					grid1.Rows = grid1.Rows + 1
					grid1.set_TextMatrix(n, 0, rsBuild.Fields("id").Value)
					If Len(grid1.get_TextMatrix(n, 0)) = 1 Then
						grid1.set_TextMatrix(n, 1, rsBuild.Fields("stat1").Value)
					Else
						grid1.set_TextMatrix(n, 1, "0")
					End If
					grid1.set_TextMatrix(n, 2, rsBuild.Fields("cost").Value)
					grid1.set_TextMatrix(n, 3, rsBuild.Fields("lcm").Value)
					grid1.set_TextMatrix(n, 4, rsBuild.Fields("hcm").Value)
					If Len(grid1.get_TextMatrix(n, 0)) > 1 Then
						grid1.set_TextMatrix(n, 5, rsBuild.Fields("stat2").Value)
					Else
						grid1.set_TextMatrix(n, 5, "0")
					End If
					n = n + 1
				End If
				rsBuild.MoveNext()
			End While
		End If
		grid1.Visible = True
		
		rsShowText.Seek("=", UnitType, "b", 1)
		If Not rsShowText.NoMatch Then
			sb1.Text = "Current tech level is '" & CStr(TechLevel) & "'.   " + rsShowText.Fields("Text").Value
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsBuild.Index = hldIndex
	End Sub
	
	Private Sub FillGridSector()
		Dim n As Short
		Dim l As Integer
		Dim MaxColumn As Short
		
		'Reset sort variables
		SortCol = -1
		SortDecending = True
		
		TextBox1.Visible = False
		grid1.Visible = False
		grid1.Clear()
		grid1.Rows = 2
		grid1.FixedCols = 1
		grid1.FixedRows = 1
		
		Select Case ReportType
			Case SHOW_CAPACITY
				MaxColumn = 12
				grid1.Cols = MaxColumn + 1
				grid1.set_TextMatrix(0, 2, "product")
				grid1.set_TextMatrix(0, 3, "use1")
				grid1.set_TextMatrix(0, 4, "use2")
				grid1.set_TextMatrix(0, 5, "use3")
				grid1.set_TextMatrix(0, 6, "level")
				grid1.set_TextMatrix(0, 7, "min")
				grid1.set_TextMatrix(0, 8, "lag")
				grid1.set_TextMatrix(0, 9, "eff%")
				grid1.set_TextMatrix(0, 10, "$$$")
				grid1.set_TextMatrix(0, 11, "dep")
				grid1.set_TextMatrix(0, 12, "code")
			Case SHOW_STAT
				MaxColumn = 10
				grid1.Cols = MaxColumn + 1
				If VersionCheck(4, 3, 6) <> WinAceRoutines.enumVersion.VER_LESS Then
					MaxColumn = 11
					grid1.Cols = MaxColumn + 1
					grid1.set_TextMatrix(0, 2, "mcost 0%")
					grid1.set_TextMatrix(0, 3, "mcost 100%")
					grid1.set_TextMatrix(0, 4, "max off")
					grid1.set_TextMatrix(0, 5, "max def")
					grid1.set_TextMatrix(0, 6, "mil")
					grid1.set_TextMatrix(0, 7, "uw")
					grid1.set_TextMatrix(0, 8, "civ")
					grid1.set_TextMatrix(0, 9, "bar")
					grid1.set_TextMatrix(0, 10, "other")
					grid1.set_TextMatrix(0, 11, "maxpop")
				Else
					MaxColumn = 10
					grid1.Cols = MaxColumn + 1
					grid1.set_TextMatrix(0, 2, "mcost")
					grid1.set_TextMatrix(0, 3, "max off")
					grid1.set_TextMatrix(0, 4, "max def")
					grid1.set_TextMatrix(0, 5, "mil")
					grid1.set_TextMatrix(0, 6, "uw")
					grid1.set_TextMatrix(0, 7, "civ")
					grid1.set_TextMatrix(0, 8, "bar")
					grid1.set_TextMatrix(0, 9, "other")
					grid1.set_TextMatrix(0, 10, "maxpop")
				End If
		End Select
		
		'Fill headers
		grid1.set_TextMatrix(0, 0, "des")
		grid1.set_TextMatrix(0, 1, "description")
		grid1.row = 0
		grid1.col = 0
		grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
		grid1.col = 1
		grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
		
		grid1.set_ColWidth(-1, -1)
		l = grid1.get_ColWidth(0) / 2
		grid1.set_ColWidth(0, l)
		grid1.set_ColWidth(1, l * 3)
		grid1.set_ColData(0, sort_type.TextSort)
		grid1.set_ColData(1, sort_type.TextSort)
		For n = 2 To MaxColumn
			grid1.set_ColWidth(n, l * 1.1)
			'121303 rjk: Fixed sorting for SHOW_CAPACITY
			If ReportType = SHOW_STAT Or ((n < 7 Or n = 12) And ReportType = SHOW_CAPACITY) Then
				grid1.set_ColData(n, sort_type.TextSort)
			Else
				grid1.set_ColData(n, sort_type.NumericSort)
			End If
			grid1.col = n
			grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
		Next n
		
		Select Case ReportType
			Case SHOW_STAT
				If VersionCheck(4, 3, 6) <> WinAceRoutines.enumVersion.VER_LESS Then
					grid1.set_ColWidth(2, l * 1.6)
					grid1.set_ColWidth(3, l * 2#)
					grid1.set_ColWidth(4, l * 1.3)
					grid1.set_ColWidth(5, l * 1.4)
					grid1.set_ColWidth(11, l * 1.3)
				Else
					grid1.set_ColWidth(3, l * 1.5)
					grid1.set_ColWidth(4, l * 1.5)
					grid1.set_ColWidth(10, l * 1.3)
				End If
			Case SHOW_CAPACITY
				grid1.set_ColWidth(2, l * 1.5)
				grid1.col = 2
				grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
				grid1.col = 6
				grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
				grid1.col = 12
				grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
				For n = 7 To 11
					grid1.set_ColData(n, sort_type.NumericSort)
				Next n
		End Select
		'Fill in rows
		n = 1
		If Not (rsSectorType.BOF And rsSectorType.EOF) Then
			rsSectorType.MoveFirst()
		End If
		
		While Not rsSectorType.EOF
			grid1.set_TextMatrix(n, 0, rsSectorType.Fields("des").Value)
			grid1.set_TextMatrix(n, 1, rsSectorType.Fields("desc").Value)
			grid1.row = n
			grid1.col = 0
			grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
			grid1.CellFontBold = True
			
			For l = 2 To MaxColumn
				grid1.col = l
				grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
			Next l
			Select Case ReportType
				Case SHOW_CAPACITY
					grid1.set_TextMatrix(n, 2, rsSectorType.Fields("product").Value)
					If rsSectorType.Fields("use1n").Value > 0 Then
						grid1.set_TextMatrix(n, 3, CStr(rsSectorType.Fields("use1n").Value) & " " + rsSectorType.Fields("use1s").Value)
					End If
					If rsSectorType.Fields("use2n").Value > 0 Then
						grid1.set_TextMatrix(n, 4, CStr(rsSectorType.Fields("use2n").Value) & " " + rsSectorType.Fields("use2s").Value)
					End If
					If rsSectorType.Fields("use3n").Value > 0 Then
						grid1.set_TextMatrix(n, 5, CStr(rsSectorType.Fields("use3n").Value) & " " + rsSectorType.Fields("use3s").Value)
					End If
					
					grid1.set_TextMatrix(n, 6, rsSectorType.Fields("level").Value)
					grid1.set_TextMatrix(n, 7, rsSectorType.Fields("min").Value)
					grid1.set_TextMatrix(n, 8, rsSectorType.Fields("lag").Value)
					grid1.set_TextMatrix(n, 9, rsSectorType.Fields("eff").Value)
					grid1.set_TextMatrix(n, 10, rsSectorType.Fields("cost").Value)
					grid1.set_TextMatrix(n, 11, rsSectorType.Fields("dep").Value)
					grid1.set_TextMatrix(n, 12, rsSectorType.Fields("pcode").Value)
					grid1.col = 2
					grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
					grid1.col = 3
					grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
					grid1.col = 4
					grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
					grid1.col = 5
					grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
					grid1.col = 6
					grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
					grid1.col = 12
					grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
				Case SHOW_STAT
					If VersionCheck(4, 3, 6) <> WinAceRoutines.enumVersion.VER_LESS Then
						grid1.set_TextMatrix(n, 2, VB6.Format(rsSectorType.Fields("mcost").Value, "##0.000"))
						grid1.set_TextMatrix(n, 3, VB6.Format(rsSectorType.Fields("mcost100").Value, "##0.000"))
						grid1.set_TextMatrix(n, 4, VB6.Format(rsSectorType.Fields("offense").Value, "##0.00"))
						grid1.set_TextMatrix(n, 5, VB6.Format(rsSectorType.Fields("defense").Value, "##0.00"))
						grid1.set_TextMatrix(n, 6, rsSectorType.Fields("pack_mil").Value)
						grid1.set_TextMatrix(n, 7, rsSectorType.Fields("pack_uw").Value)
						grid1.set_TextMatrix(n, 8, rsSectorType.Fields("pack_civ").Value)
						grid1.set_TextMatrix(n, 9, rsSectorType.Fields("pack_bar").Value)
						grid1.set_TextMatrix(n, 10, rsSectorType.Fields("pack_other").Value)
						grid1.set_TextMatrix(n, 11, rsSectorType.Fields("maxpop").Value)
					Else
						grid1.set_TextMatrix(n, 2, rsSectorType.Fields("mcost").Value)
						grid1.set_TextMatrix(n, 3, VB6.Format(rsSectorType.Fields("offense").Value, "##0.00"))
						grid1.set_TextMatrix(n, 4, VB6.Format(rsSectorType.Fields("defense").Value, "##0.00"))
						grid1.set_TextMatrix(n, 5, rsSectorType.Fields("pack_mil").Value)
						grid1.set_TextMatrix(n, 6, rsSectorType.Fields("pack_uw").Value)
						grid1.set_TextMatrix(n, 7, rsSectorType.Fields("pack_civ").Value)
						grid1.set_TextMatrix(n, 8, rsSectorType.Fields("pack_bar").Value)
						grid1.set_TextMatrix(n, 9, rsSectorType.Fields("pack_other").Value)
						grid1.set_TextMatrix(n, 10, rsSectorType.Fields("maxpop").Value)
					End If
			End Select
			
			n = n + 1
			grid1.Rows = n + 1
			rsSectorType.MoveNext()
		End While
		grid1.Visible = True
		
		rsShowText.Seek("=", "c", ReportType, 1)
		If Not rsShowText.NoMatch Then
			sb1.Text = "Current tech level is '" & CStr(TechLevel) & "'.   " + rsShowText.Fields("Text").Value
		End If
	End Sub
	
	Private Sub SetReportTech()
		Dim n As Short
		Dim n2 As Short
		
		rsShowText.Seek("=", UnitType, ReportType, 1)
		If rsShowText.NoMatch Then
			ReportTech = TechLevel
			Exit Sub
		End If
		
		n = InStr(rsShowText.Fields("Text").Value, "'")
		If n = 0 Then
			ReportTech = TechLevel
			Exit Sub
		End If
		
		n2 = InStr(n + 1, rsShowText.Fields("Text").Value, "'")
		
		If n2 <= n Then
			ReportTech = TechLevel
			Exit Sub
		End If
		
		ReportTech = CShort(Mid(rsShowText.Fields("Text").Value, n + 1, n2 - n - 1))
	End Sub
	
	Private Sub SortGrid(ByRef col As Short)
		'This does a simple bubble sort of the grid in place.  Due to the small size
		'of the sample, a more efficent sort is usually not necessary.
		Dim n As Short
		Dim row As Short
		Dim var1 As Object
		Dim var2 As Object
		Dim sng1 As Single
		Dim sng2 As Single
		
		'If clicked multiple times, change the sort order
		If SortCol = col Then
			SortDecending = Not SortDecending
		End If
		SortCol = col
		
		grid1.Visible = False
		
		n = grid1.Rows - 3
		
		If grid1.get_ColData(col) = sort_type.TextSort Then
			'text sort
			While n > 0
				For row = 1 To n
					'UPGRADE_WARNING: Couldn't resolve default property of object var1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					var1 = grid1.get_TextMatrix(row, col)
					'UPGRADE_WARNING: Couldn't resolve default property of object var2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					var2 = grid1.get_TextMatrix(row + 1, col)
					
					If SortDecending Then
						'UPGRADE_WARNING: Couldn't resolve default property of object var2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object var1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If var1 > var2 Then
							grid1.set_RowPosition(row + 1, row)
						End If
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object var2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object var1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If var1 < var2 Then
							grid1.set_RowPosition(row + 1, row)
						End If
					End If
				Next row
				n = n - 1
			End While
		Else
			'numeric sort
			While n > 0
				For row = 1 To n
					sng1 = CSng(grid1.get_TextMatrix(row, col))
					sng2 = CSng(grid1.get_TextMatrix(row + 1, col))
					
					If SortDecending Then
						If sng1 > sng2 Then
							grid1.set_RowPosition(row + 1, row)
						End If
					Else
						If sng1 < sng2 Then
							grid1.set_RowPosition(row + 1, row)
						End If
					End If
				Next row
				n = n - 1
			End While
		End If
		grid1.Visible = True
	End Sub
	
	Private Sub FillGridItem()
		Dim n As Short
		Dim l As Integer
		'Reset sort variables
		SortCol = -1
		SortDecending = True
		
		TextBox1.Visible = False
		grid1.Visible = False
		grid1.Clear()
		grid1.Rows = Items.Count + 1
		grid1.FixedCols = 1
		grid1.FixedRows = 1
		
		grid1.Cols = 9
		grid1.row = 0
		grid1.set_ColWidth(-1, -1)
		l = grid1.get_ColWidth(0) / 2
		grid1.col = 0
		grid1.set_ColWidth(-1, l)
		grid1.set_TextMatrix(0, 0, "letter")
		grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter
		grid1.set_ColData(0, sort_type.TextSort)
		grid1.col = 1
		grid1.set_TextMatrix(0, 1, "value")
		grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
		grid1.set_ColData(1, sort_type.NumericSort)
		grid1.col = 2
		grid1.set_TextMatrix(0, 2, "sell")
		grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
		grid1.set_ColData(2, sort_type.TextSort)
		grid1.col = 3
		grid1.set_TextMatrix(0, 3, "lbs")
		grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
		grid1.set_ColData(3, sort_type.NumericSort)
		grid1.col = 4
		grid1.set_TextMatrix(0, 4, "rg pk")
		grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
		grid1.set_ColData(4, sort_type.NumericSort)
		grid1.set_ColWidth(4, l * 1.1)
		grid1.col = 5
		grid1.set_TextMatrix(0, 5, "wh pk")
		grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
		grid1.set_ColData(5, sort_type.NumericSort)
		grid1.set_ColWidth(5, l * 1.1)
		grid1.col = 6
		grid1.set_TextMatrix(0, 6, "ur pk")
		grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
		grid1.set_ColData(6, sort_type.NumericSort)
		grid1.set_ColWidth(6, l * 1.1)
		grid1.col = 7
		grid1.set_TextMatrix(0, 7, "bk pk")
		grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
		grid1.set_ColData(7, sort_type.NumericSort)
		grid1.set_ColWidth(7, l * 1.1)
		grid1.col = 8
		grid1.set_ColWidth(8, l * 5)
		grid1.set_TextMatrix(0, 8, "name")
		grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
		grid1.set_ColData(8, sort_type.TextSort)
		'Fill in rows
		For n = 1 To Items.Count
			grid1.set_TextMatrix(n, 0, Items(n).Letter)
			grid1.set_TextMatrix(n, 1, CStr(Items(n).Value))
			If Items(n).Sellable Then
				grid1.set_TextMatrix(n, 2, "yes")
			Else
				grid1.set_TextMatrix(n, 2, "no")
			End If
			grid1.set_TextMatrix(n, 3, CStr(Items(n).Weight))
			grid1.set_TextMatrix(n, 4, CStr(Items(n).Packing(EmpItem.enumItemPacking.PACKING_REGULAR)))
			grid1.set_TextMatrix(n, 5, CStr(Items(n).Packing(EmpItem.enumItemPacking.PACKING_WAREHOUSE)))
			grid1.set_TextMatrix(n, 6, CStr(Items(n).Packing(EmpItem.enumItemPacking.PACKING_URBAN)))
			grid1.set_TextMatrix(n, 7, CStr(Items(n).Packing(EmpItem.enumItemPacking.PACKING_BANK)))
			grid1.set_TextMatrix(n, 8, Items(n).ItemName)
		Next n
		
		grid1.Visible = True
		rsShowText.Seek("=", "c", ReportType, 1)
		If Not rsShowText.NoMatch Then
			sb1.Text = "Current tech level is '" & CStr(TechLevel) & "'.   " + rsShowText.Fields("Text").Value
		End If
	End Sub
End Class