Option Strict Off
Option Explicit On
Friend Class frmToolFeeder
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'171003 rjk: Added Strength fields to Sector database
	'201103 rjk: Added option to control strength updates
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	
	Dim PackDes As String
	'UPGRADE_WARNING: Arrays in structure rs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
	Dim rs As DAO.Recordset
	Dim InSetup As Boolean
	Dim SortCol As Short
	Dim SortDecending As Boolean
	Dim SectorMob As Short
	Dim SectorFood As Short
	
	'drk 5/25/03:  made general to work with more commodities than just food (e.g. maybe we want to move civs from areas of overpopulation to areas of shortage)
	Dim comm_name As String
	Dim comm_dist As String
	
	Private Structure record 'drk 5/24/03 arrays of variants suck squid.  Removed that and added a proper user defined type
		Dim secx As Short
		Dim secy As Short
		Dim comm_reqd As Object
		Dim comm_shortage As Short
		Dim mcost As Single
		Dim pcost As Single
		Dim checked As Boolean
	End Structure
	
	Dim aryValues() As record
	
	'UPGRADE_WARNING: Event CheckSetThresh.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub CheckSetThresh_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CheckSetThresh.CheckStateChanged
		Label5.Enabled = CheckSetThresh.CheckState
		Label4.Enabled = CheckSetThresh.CheckState
		Text2.Enabled = CheckSetThresh.CheckState
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim n As Short
		Dim var1 As record
		Dim mobavail As Single
		Dim foodavail As Short
		Dim strCmd As String
		
		mobavail = SectorMob - Val(Text1(0).Text)
		foodavail = SectorFood - Val(Text1(1).Text)
		
		frmEmpCmd.SubmitEmpireCommand("bf1", False) 'run as batch file
		Dim newthresh As Short
		For n = 1 To UBound(aryValues)
			'UPGRADE_WARNING: Couldn't resolve default property of object var1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			var1 = aryValues(n)
			If var1.checked = True Then
				If mobavail > var1.mcost And foodavail > var1.comm_shortage Then
					foodavail = foodavail - var1.comm_shortage
					mobavail = mobavail - var1.mcost
					strCmd = "move " & comm_name & " " & txtOrigin.Text & " " & CStr(var1.comm_shortage) & " " & SectorString(var1.secx, var1.secy)
					frmEmpCmd.SubmitEmpireCommand(strCmd, False) 'sectors
					
					'drk@unxsoft.com 6/23/02
					If CheckSetThresh.CheckState Then
						'UPGRADE_WARNING: Couldn't resolve default property of object var1.comm_reqd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						newthresh = Int(0.5 + (var1.comm_reqd * (1 + Val(Text2.Text) / 100)))
						rsSectors.Seek("=", var1.secx, var1.secy)
						'drk 7/20/02: this can get very BTU intensive for large countries.  Save some by only submitting the thresh command if we're actually increasing the thresh
						If (newthresh > rsSectors.Fields(comm_dist).Value) Then
							strCmd = "thresh " & comm_name & " " & SectorString(var1.secx, var1.secy) & " " & newthresh
							frmEmpCmd.SubmitEmpireCommand(strCmd, False)
						End If
					End If
					
				End If
			End If
		Next n
		frmEmpCmd.SubmitEmpireCommand("bf2", False) 'run as batch file
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		GetSectors()
		'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
		GetCurrentStrength(tsSectors)
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
		Me.Close()
	End Sub
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		frmEmpCmd.SubmitEmpireCommand("bf1", False) 'run as batch file
		frmEmpCmd.SubmitEmpireCommand("starvation *", True) 'sectors
		frmEmpCmd.SubmitEmpireCommand("bf2", False) 'run as batch file
		frmEmpCmd.SubmitEmpireCommand("db2", False) 'force map redraw
	End Sub
	
	Private Sub frmToolFeeder_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Set always on top
		' Dim success As Long    8/2003 efj  removed
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
		InSetup = True
		
		lblError.Text = vbNullString
		Label2(0).Text = vbNullString
		Label2(1).Text = vbNullString
		Text1(2).Text = CStr(99)
		rs = rsSectors.Clone
		rs.Index = "PrimaryKey"
		SortDecending = True
		SortCol = 2
		ReDim aryValues(0)
		
		comm_name = "food"
		comm_dist = "f_dist"
	End Sub
	
	Public Sub FillGrid(ByRef ClearFirst As Boolean)
		
		Dim n As Short
		Dim var1 As record
		Dim totmobavail As Single
		Dim mobavail As Single
		Dim mobused As Single
		Dim mobtotal As Single
		Dim totfoodavail As Short
		Dim foodavail As Short
		Dim foodused As Short
		Dim foodtotal As Short
		Dim labels As Object
		
		If InSetup Then
			Exit Sub
		End If
		
		'Setup grid
		grid1.Visible = False
		If ClearFirst Then
			grid1.Clear()
			
			grid1.Rows = 2
			
			grid1.Cols = 7
			grid1.FixedCols = 2
			grid1.set_ColWidth(0, -1) 'return to default
			grid1.set_ColWidth(0, grid1.get_ColWidth(0) / 2.5)
			grid1.set_ColWidth(1, grid1.get_ColWidth(0))
			grid1.set_ColWidth(2, grid1.get_ColWidth(0) * 2.1)
			grid1.set_ColWidth(3, grid1.get_ColWidth(0) * 2.1)
			grid1.set_ColWidth(6, grid1.get_ColWidth(0) * 1.9)
			grid1.set_ColAlignment(0, MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter)
			grid1.set_ColAlignment(1, MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter)
			grid1.set_ColAlignment(2, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter)
			grid1.set_ColAlignment(3, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter)
			grid1.set_ColAlignment(4, MSFlexGridLib.AlignmentSettings.flexAlignRightCenter)
			grid1.set_ColAlignment(6, MSFlexGridLib.AlignmentSettings.flexAlignCenterCenter)
			
			'Fill in row headers
			grid1.col = 0
			grid1.row = 0
			
			'UPGRADE_WARNING: Array has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			'UPGRADE_WARNING: Couldn't resolve default property of object labels. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			labels = New Object(){"X", "Y", comm_name & " req.", "shortage", "path cost", "mob. cost", "Select"}
			For n = 0 To 6
				grid1.col = n
				grid1.CellFontBold = True
				'UPGRADE_WARNING: Couldn't resolve default property of object labels(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				grid1.Text = labels(n)
			Next n
			
		End If
		
		totmobavail = SectorMob - Val(Text1(0).Text)
		totfoodavail = SectorFood - Val(Text1(1).Text)
		mobavail = totmobavail
		foodavail = totfoodavail
		mobused = 0
		foodused = 0
		
		For n = 1 To UBound(aryValues)
			'UPGRADE_WARNING: Couldn't resolve default property of object var1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			var1 = aryValues(n)
			
			If ClearFirst Then
				grid1.row = grid1.Rows - 1
				grid1.Rows = grid1.Rows + 1
			Else
				grid1.row = n
			End If
			
			grid1.col = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object record_nth_elt(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			grid1.Text = CStr(record_nth_elt(var1, (grid1.col)))
			grid1.col = 1
			'UPGRADE_WARNING: Couldn't resolve default property of object record_nth_elt(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			grid1.Text = CStr(record_nth_elt(var1, (grid1.col)))
			
			'Set the check value
			grid1.col = 6 '
			If var1.checked = True Then '
				grid1.CellForeColor = System.Drawing.Color.Black '
				grid1.Text = "Yes" '
				
				mobtotal = mobtotal + var1.mcost
				foodtotal = foodtotal + var1.comm_shortage
				If mobavail > var1.mcost And foodavail > var1.comm_shortage Then
					mobused = mobused + var1.mcost
					foodused = foodused + var1.comm_shortage
				End If
				
				grid1.col = 2
				grid1.CellForeColor = System.Drawing.Color.Black
				'UPGRADE_WARNING: Couldn't resolve default property of object record_nth_elt(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				grid1.Text = CStr(record_nth_elt(var1, (grid1.col))) & "   "
				
				grid1.col = 3
				'UPGRADE_WARNING: Couldn't resolve default property of object var1.comm_reqd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If foodavail > var1.comm_reqd Then
					'UPGRADE_WARNING: Couldn't resolve default property of object var1.comm_reqd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					foodavail = foodavail - var1.comm_reqd
					grid1.CellForeColor = System.Drawing.Color.Black
				Else
					grid1.CellForeColor = System.Drawing.Color.Red
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object record_nth_elt(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				grid1.Text = CStr(record_nth_elt(var1, (grid1.col))) & "   "
				
				
				grid1.col = 4
				grid1.CellForeColor = System.Drawing.Color.Black
				'UPGRADE_WARNING: Couldn't resolve default property of object record_nth_elt(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				grid1.Text = VB6.Format(CSng(record_nth_elt(var1, (grid1.col))), "###0.0000") & "   "
				grid1.col = 5
				If mobavail > var1.mcost Then
					mobavail = mobavail - var1.mcost
					grid1.CellForeColor = System.Drawing.Color.Black
				Else
					grid1.CellForeColor = System.Drawing.Color.Red
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object record_nth_elt(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				grid1.Text = VB6.Format(CSng(record_nth_elt(var1, (grid1.col))), "###0.0000") & "   "
			Else
				grid1.col = 2
				grid1.CellForeColor = System.Drawing.Color.Blue
				'UPGRADE_WARNING: Couldn't resolve default property of object record_nth_elt(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				grid1.Text = CStr(record_nth_elt(var1, (grid1.col))) & "   "
				grid1.col = 3
				grid1.CellForeColor = System.Drawing.Color.Blue
				'UPGRADE_WARNING: Couldn't resolve default property of object record_nth_elt(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				grid1.Text = CStr(record_nth_elt(var1, (grid1.col))) & "   "
				grid1.col = 4
				grid1.CellForeColor = System.Drawing.Color.Blue
				'UPGRADE_WARNING: Couldn't resolve default property of object record_nth_elt(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				grid1.Text = VB6.Format(CSng(record_nth_elt(var1, (grid1.col))), "###0.0000") & "   "
				grid1.col = 5
				grid1.CellForeColor = System.Drawing.Color.Blue
				'UPGRADE_WARNING: Couldn't resolve default property of object record_nth_elt(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				grid1.Text = VB6.Format(CSng(record_nth_elt(var1, (grid1.col))), "###0.0000") & "   "
				grid1.col = 6
				grid1.CellForeColor = System.Drawing.Color.Blue
				grid1.Text = "No"
			End If
			
		Next n
		
		Label2(0).Text = CStr(foodused) & " " & comm_name & ", " & CStr(Int(mobused)) & " mob."
		Label2(1).Text = "Shortage: "
		If foodtotal > totfoodavail Then
			Label2(1).Text = Label2(1).Text & CStr(foodtotal - totfoodavail) & " " & comm_name & " "
		End If
		If mobtotal > totmobavail Then
			Label2(1).Text = Label2(1).Text & CStr(Int(mobtotal - totmobavail)) & " mob."
		End If
		
		If Len(Label2(1).Text) = 10 Then
			Label2(1).Text = Label2(1).Text & "None"
		End If
		
		grid1.Visible = True
	End Sub
	
	
	Private Sub frmToolFeeder_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	Private Sub grid1_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles grid1.ClickEvent
		
		Dim n As Short
		Dim var1 As record
		Dim sx As Short
		Dim sy As Short
		' Dim saverow As Integer    8/2003 efj  removed
		
		'If you clicked on a header, resort the grid
		If grid1.MouseRow = 0 Then
			If SortCol = grid1.MouseCol Then
				SortDecending = Not SortDecending
			End If
			SortCol = grid1.MouseCol
			SortValues()
			FillGrid(False)
			
			'Otherwise, mark the column
		ElseIf grid1.MouseRow < grid1.Rows Then  'drk@unxsoft.com, bugfix. would crash on certain pathological conditions
			If grid1.get_TextMatrix(grid1.MouseRow, 0) <> vbNullString Then 'ditto
				'        saverow = grid1.MouseRow    8/2003 efj  removed
				grid1.row = grid1.MouseRow
				grid1.col = 0
				sx = CShort(grid1.Text)
				grid1.col = 1
				sy = CShort(grid1.Text)
				
				grid1.col = 6
				For n = 1 To UBound(aryValues)
					'UPGRADE_WARNING: Couldn't resolve default property of object var1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					var1 = aryValues(n)
					If sx = var1.secx And sy = var1.secy Then
						If grid1.Text = "Yes" Then
							grid1.Text = "No"
							var1.checked = 0
						Else
							grid1.Text = "Yes"
							var1.checked = 1
						End If
						'UPGRADE_WARNING: Couldn't resolve default property of object aryValues(n). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						aryValues(n) = var1
					End If
				Next n
				
				FillGrid(False)
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Option1.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option1_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option1.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = Option1.GetIndex(eventSender)
			LoadValues()
			SortValues()
			FillGrid(True)
		End If
	End Sub
	
	Private Sub Text1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text1.Leave
		Dim Index As Short = Text1.GetIndex(eventSender)
		LoadValues()
		SortValues()
		FillGrid(True)
	End Sub
	
	'UPGRADE_WARNING: Event txtOrigin.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtOrigin_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOrigin.TextChanged
		
		Dim sx As Short
		Dim sy As Short
		
		If ParseSectors(sx, sy, (txtOrigin.Text)) Then
			FillSectorInfo(sx, sy)
			InSetup = False
			LoadValues()
			SortValues()
			FillGrid(True)
		Else
			HandleError("Not a valid Sector")
		End If
		
	End Sub
	
	
	Private Sub FillSectorInfo(ByRef secx As Short, ByRef secy As Short)
		
		'get sector record
		rs.Seek("=", secx, secy)
		If rs.NoMatch Then
			HandleError("Sector not found")
			Exit Sub
		End If
		
		SectorMob = rs.Fields("mob").Value
		SectorFood = rs.Fields(comm_name).Value
		
		lblError.Text = CStr(SectorFood) & " " & comm_name & ", " & CStr(SectorMob) & " mob."
		
		'set sector type
		If rs.Fields("eff").Value < 60 Then
			PackDes = vbNullString
		Else
			PackDes = rs.Fields("des").Value
		End If
		
		Text1(0).Text = CStr(Int(SectorMob / 2))
		Text1(1).Text = CStr(FoodRequired(rs))
		
	End Sub
	Private Sub HandleError(ByRef ErrMsg As String)
		lblError.Text = ErrMsg
		Text1(0).Text = vbNullString
		Text1(1).Text = vbNullString
		Text1(2).Text = vbNullString
		Label2(0).Text = vbNullString
		Label2(1).Text = vbNullString
		grid1.Clear()
	End Sub
	Private Function PathCost() As Single
		
		On Error Resume Next
		Dim pvar As Object
		'Dim pcost As Single 'drk 5/24/03 duh
		
		'Compute the path cost
		'UPGRADE_WARNING: Couldn't resolve default property of object BestPath(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object pvar. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		pvar = BestPath((txtOrigin.Text), SectorString((rs.Fields("x").Value), (rs.Fields("y").Value)), EmpCommon.enumMobType.MT_COMMODITY)
		'pcost = pvar(2)
		'UPGRADE_WARNING: Couldn't resolve default property of object pvar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Len(CStr(pvar(1))) = 0 Then
			PathCost = 0
			Exit Function
		End If
		
		'Compute the mobility cost
		'PathCost = pcost
		'UPGRADE_WARNING: Couldn't resolve default property of object pvar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PathCost = pvar(2)
		
	End Function
	
	Private Sub LoadValues()
		
		Dim StartX As Short
		Dim StartY As Short
		Dim MaxRange As Short
		Dim n As Short
		Dim fval As record
		Dim StrMsg As String
		
		If InSetup Then
			Exit Sub
		End If
		
		ReDim aryValues(0)
		
		If Not ParseSectors(StartX, StartY, (txtOrigin.Text)) Then
			Exit Sub
		End If
		
		MaxRange = Val(Text1(2).Text)
		If MaxRange <= 0 Then
			Exit Sub
		End If
		
		n = 1
		'Run through all the sector checking food requirements
		rs.MoveFirst()
		While Not rs.EOF
			If SectorDistance((rs.Fields("x").Value), (rs.Fields("y").Value), StartX, StartY) <= MaxRange Then
				If comm_name = "food" Then
					
					'UPGRADE_WARNING: Couldn't resolve default property of object fval.comm_reqd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					fval.comm_reqd = FoodRequired(rs)
					
					If Option1(1).Checked Then
						'base on calculated food
						fval.comm_shortage = FoodRequired(rs) - rs.Fields("food").Value
					Else
						'base on Starvation event markers
						fval.comm_shortage = 0
						StrMsg = EventMarkers.Find((rs.Fields("x").Value), (rs.Fields("y").Value))
						If InStr(StrMsg, "Starvation") > 0 Then
							fval.comm_shortage = Val(StringBetween(StrMsg, "p/", "f"))
						End If
					End If
				ElseIf comm_name = "civ" Then 
					'UPGRADE_WARNING: Couldn't resolve default property of object fval.comm_reqd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					fval.comm_reqd = 768 'fixme
					
				End If
				
				If fval.comm_shortage > 0 Then
					fval.pcost = PathCost
					If fval.pcost > 0 Then
						fval.mcost = fval.pcost * MobilityWeight("f", fval.comm_shortage, PackDes)
						'save the values
						fval.secx = rs.Fields("x").Value
						fval.secy = rs.Fields("y").Value
						fval.checked = True
						ReDim Preserve aryValues(n)
						'UPGRADE_WARNING: Couldn't resolve default property of object aryValues(n). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						aryValues(n) = fval
						n = n + 1
					End If
				End If
			End If
			rs.MoveNext()
		End While
		
	End Sub
	
	
	Private Sub SortValues()
		
		'This does a simple bubble sort of the array in place.  Due to the small size
		'of the sample, a more efficent sort is usually not necessary.
		
		Dim n As Short
		Dim row As Short
		Dim var1 As record
		Dim var2 As record
		Dim var1c As Single
		Dim var2c As Single
		
		If InSetup Then
			Exit Sub
		End If
		
		n = UBound(aryValues) - 1
		While n > 0
			For row = 1 To n
				'UPGRADE_WARNING: Couldn't resolve default property of object var1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				var1 = aryValues(row)
				'UPGRADE_WARNING: Couldn't resolve default property of object record_nth_elt(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				var1c = CSng(record_nth_elt(var1, SortCol))
				'UPGRADE_WARNING: Couldn't resolve default property of object var2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				var2 = aryValues(row + 1)
				'UPGRADE_WARNING: Couldn't resolve default property of object record_nth_elt(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				var2c = CSng(record_nth_elt(var2, SortCol))
				
				If SortDecending Then
					If var1c > var2c Then
						'UPGRADE_WARNING: Couldn't resolve default property of object aryValues(row). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						aryValues(row) = var2
						'UPGRADE_WARNING: Couldn't resolve default property of object aryValues(row + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						aryValues(row + 1) = var1
					End If
				Else
					If var1c < var2c Then
						'UPGRADE_WARNING: Couldn't resolve default property of object aryValues(row). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						aryValues(row) = var2
						'UPGRADE_WARNING: Couldn't resolve default property of object aryValues(row + 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						aryValues(row + 1) = var1
					End If
				End If
			Next row
			n = n - 1
		End While
		
	End Sub
	
	'drk 5/24/03: previously, we had stored things in variant arrays, so it was easy to pick out
	'the nth element.  It is preferable to store things in user-defined types, but now we
	'need a helper function to pull out the nth element
	Private Function record_nth_elt(ByRef r As record, ByRef S As Short) As Object
		Select Case S
			Case 0
				'UPGRADE_WARNING: Couldn't resolve default property of object record_nth_elt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				record_nth_elt = r.secx
			Case 1
				'UPGRADE_WARNING: Couldn't resolve default property of object record_nth_elt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				record_nth_elt = r.secy
			Case 2
				'UPGRADE_WARNING: Couldn't resolve default property of object r.comm_reqd. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object record_nth_elt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				record_nth_elt = r.comm_reqd
			Case 3
				'UPGRADE_WARNING: Couldn't resolve default property of object record_nth_elt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				record_nth_elt = r.comm_shortage
			Case 4
				'UPGRADE_WARNING: Couldn't resolve default property of object record_nth_elt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				record_nth_elt = r.pcost
			Case 5
				'UPGRADE_WARNING: Couldn't resolve default property of object record_nth_elt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				record_nth_elt = r.mcost
			Case 6
				'UPGRADE_WARNING: Couldn't resolve default property of object record_nth_elt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				record_nth_elt = r.checked
			Case Else
				System.Diagnostics.Debug.Assert(False, "")
		End Select
	End Function
End Class