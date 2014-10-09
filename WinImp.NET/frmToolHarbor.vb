Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmToolHarbor
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'280903 rjk: switched consts to enum and add unknown.
	'            general reformatting.
	'031003 rjk: Added tooltip to cmdBackShort and cmdNextShort.
	'171003 rjk: Added Strength fields to Sector database
	'241003 rjk: Fixed logic cmdNextShort and cmdBackShort so NextBase is called
	'251003 rjk: Changed "ETU per undate" to "ETU per update" - code cleanup
	'161103 rjk: Changed the name TxtSector to txtSector, code cleanup.
	'201103 rjk: Added option to control strength updates
	'130304 rjk: Only add a row if there is efficiency gain
	'130304 rjk: Fixed FillGridRow to accomodate large avail requirements from St@r W@rs
	'180304 rjk: Added Green color if the unit will be at max. eff. gain but not 100%.
	'030404 rjk: Switched to use the avail from the server instead of calculating it.
	'080404 rjk: Fill in simulated build into Label 4
	'080404 rjk: Switched "Left:" to show Not Used: if "Left:" is empty.
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	'090405 rjk: Added to "Add Update Avail" checkbox to the "Use Current Avail".
	'            This can be used to estimate the production for this and the next update.
	'200505 rjk: Added capability to forecast upto 3 updates ahead for both avail modes.
	'291205 rjk: Added unit workforce to building calculation.
	'301205 rjk: Added Theme Game SP 2005, builds the units with military/shells/guns.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	'140506 rjk: Use shells for lcms for South Pacific Atlantis
	'041006 rjk: Prevent stopped units from appearing in the build tool.
	'090607 rjk: Remove the Rollover Avail Scaling by Sector efficiency.
	'            Fix Rollover avail calculations for production of avail.
	
	Private ItemType As enumHarborType
	Dim secx As Short
	Dim secy As Short
	Dim strBuilds() As String
	' Dim cb As Integer    8/2003 efj  removed
	Dim OriginChange As Boolean
	Dim Costs() As Single
	Dim CostItems As Short
	Dim Header As String
	'UPGRADE_NOTE: Short was upgraded to Short_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Dim Short_Renamed As Boolean
	Dim iMaxLeft As Short
	
	Private Enum enumHarborType
		FH_UNKNOWN
		FH_BUILD_SHIPS
		FH_BUILD_PLANES
		FH_BUILD_LANDS
		FH_IMPROVE_LANDS
	End Enum
	
	Const FH_BEST_CASE As Short = 1
	Const FH_WORST_CASE As Short = 2
	Const FH_AVERAGE_CASE As Short = 3
	
	'UPGRADE_WARNING: Event cmbUpdates.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbUpdates_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbUpdates.SelectedIndexChanged
		If Not OriginChange Then
			txtOrigin_TextChanged(txtOrigin, New System.EventArgs())
		End If
	End Sub
	
	Private Sub cmdCenter_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCenter.Click
		frmDrawMap.MoveMap(secx, secy)
	End Sub
	
	Private Sub cmdNext_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdNext.Click
		NextBase(True, False)
	End Sub
	
	Private Sub cmdBack_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBack.Click
		NextBase(False, False)
	End Sub
	
	'UPGRADE_WARNING: Event cmdNextShort.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmdNextShort_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdNextShort.CheckedChanged
		If eventSender.Checked Then
			If cmdNextShort.Checked Then
				cmdNextShort.Checked = False
				'Else 102303 rjk: removed so NextBase is called
				NextBase(True, True)
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event cmdBackShort.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmdBackShort_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBackShort.CheckedChanged
		If eventSender.Checked Then
			If cmdBackShort.Checked Then
				cmdBackShort.Checked = False
				'Else 102303 rjk: removed so NextBase is called
				NextBase(False, True)
			End If
		End If
	End Sub
	
	Private Sub NextBase(ByRef GoingForward As Boolean, ByRef ShortOnly As Boolean)
		Dim Wrap As Boolean
		'UPGRADE_WARNING: Arrays in structure rs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim rs As DAO.Recordset
		
		'Go to the current Sector
		rsSectors.Seek("=", secx, secy)
		If rsSectors.NoMatch Then
			HandleError("Sector not found")
			Exit Sub
		End If
		
		'move to the next sector
		If GoingForward Then
			If Not rsSectors.EOF Then
				rsSectors.MoveNext()
			End If
		Else
			If Not rsSectors.BOF Then
				rsSectors.MovePrevious()
			End If
		End If
		
		'Search for the next build sector, wraping once if necessary
		Wrap = False
		While Not Wrap
			If GoingForward And Not Wrap And rsSectors.EOF Then
				rsSectors.MoveFirst()
				Wrap = True
			End If
			If Not GoingForward And Not Wrap And rsSectors.BOF Then
				rsSectors.MoveLast()
				Wrap = True
			End If
			While Not ((GoingForward And rsSectors.EOF) Or (Not GoingForward And rsSectors.BOF))
				If InStr("h!*f", rsSectors.Fields("des").Value) > 0 Then
					
					'check for units being built
					Select Case rsSectors.Fields("des").Value
						Case "h"
							rs = rsShip
						Case "!", "f"
							rs = rsLand
						Case "*"
							rs = rsPlane
					End Select
					
					'Scan units
					If Not (rs.BOF And rs.EOF) Then
						rs.MoveFirst()
					End If
					Do While Not rs.EOF
						'Check to see if it is in this hex
						If rsSectors.Fields("x").Value = rs.Fields("x").Value And rsSectors.Fields("y").Value = rs.Fields("y").Value Then
							'check to see that it is not 100% already
							If rs.Fields("eff").Value < 100 Then
								txtOrigin.Text = SectorString(rsSectors.Fields("x").Value, rsSectors.Fields("y").Value)
								If Not (ShortOnly And Not Short_Renamed) Then
									Exit Sub
								Else
									Exit Do
								End If
							End If
						End If
						rs.MoveNext()
					Loop 
				End If
				If GoingForward Then
					rsSectors.MoveNext()
				Else
					rsSectors.MovePrevious()
				End If
			End While
		End While
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Me.Close()
	End Sub
	
	Private Sub cmdPull_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPull.Click
		'Fill resource shorted string
		Dim n As Short
		Dim i As Short
		Dim ShortAmt As Short
		Dim ShortItem As String
		Dim DistX As Short
		Dim DistY As Short
		Dim strCmd As String
		
		'get sector record
		rsSectors.Seek("=", secx, secy)
		If rsSectors.NoMatch Then
			HandleError("Sector not found")
			Exit Sub
		End If
		DistX = rsSectors.Fields("dist_x").Value
		DistY = rsSectors.Fields("dist_y").Value
		
		'get sector record
		rsSectors.Seek("=", DistX, DistY)
		If rsSectors.NoMatch Then
			HandleError("Sector not found")
			Exit Sub
		End If
		
		i = 0
		For n = 3 To CostItems
			If Costs(6, n) > Costs(5, n) Then
				ShortAmt = Costs(6, n) - Costs(5, n)
				ShortItem = VB.Left(Trim(Mid(Header, (n + 2) * 5 + 1, 5)), 3)
				i = i + 1
				If rsSectors.Fields(ShortItem).Value >= ShortAmt Then
					strCmd = "move " & ShortItem & " " & SectorString(DistX, DistY) & " " & Str(ShortAmt) & " " & SectorString(secx, secy)
					frmEmpCmd.SubmitEmpireCommand(strCmd, True)
				End If
			End If
		Next n
		
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		GetSectors()
		'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
		GetCurrentStrength(tsSectors)
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
		'reset to sector record
		rsSectors.Seek("=", secx, secy)
		If rsSectors.NoMatch Then
			HandleError("Sector not found")
			Exit Sub
		End If
	End Sub
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim BuildType As String
		Dim i As Short
		
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
		Load(frmToolHarborQueue)
		Select Case ItemType
			Case enumHarborType.FH_BUILD_SHIPS
				BuildType = "s"
			Case enumHarborType.FH_BUILD_LANDS, enumHarborType.FH_IMPROVE_LANDS
				BuildType = "l"
			Case enumHarborType.FH_BUILD_PLANES
				BuildType = "p"
		End Select
		
		'Load Build Queue
		On Error GoTo SkipQueue
		For i = LBound(strBuilds) + 1 To UBound(strBuilds)
			frmToolHarborQueue.List1.Items.Add(strBuilds(i))
		Next i
		
		frmToolHarborQueue.LoadTypebox(BuildType)
		
		frmToolHarborQueue.strSector = txtOrigin.Text
		
SkipQueue: On Error Resume Next
		Me.Visible = False
		frmToolHarborQueue.ShowDialog()
		
		'Simulated Builds:  1dd, 3bbc, 6awc, 121pt
		Label4.Text = vbNullString
		
		'Save Build Queue
		ReDim strBuilds(frmToolHarborQueue.List1.Items.Count)
		Dim n As Short
		Dim n1 As Short
		For i = 1 To frmToolHarborQueue.List1.Items.Count
			
			strBuilds(i) = VB6.GetItemString(frmToolHarborQueue.List1, i - 1)
			n = InStr(strBuilds(i), "#")
			If n > 0 Then
				n1 = InStr(n, strBuilds(i), " ")
				If n1 > 0 Then
					Label4.Text = Label4.Text & " " & Trim(VB.Left(strBuilds(i), n - 1)) & Mid(strBuilds(i), n, n1 - n)
				End If
			End If
		Next i
		
		If Len(Label4.Text) > 0 Then
			Label4.Text = "Simulated Builds:" & Label4.Text
		End If
		
		frmToolHarborQueue.Close()
		If ItemType <> enumHarborType.FH_UNKNOWN Then
			FillGrid()
		End If
		Me.Visible = True
	End Sub
	
	Private Sub frmToolHarbor_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Set always on top
		' Dim success As Long    8/2003 efj  removed
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
		
		lblError.Text = vbNullString
		Label2(0).Text = vbNullString
		Label2(1).Text = vbNullString
		Label2(2).Text = vbNullString
		Label4.Text = vbNullString
		
		cmbUpdates.SelectedIndex = 0
		If VersionCheck(4, 3, 6) <> WinAceRoutines.enumVersion.VER_LESS Then
			Option2(1).Checked = True
			Option2(0).Enabled = False
		End If
		
		ReDim strBuilds(0)
	End Sub
	
	Public Sub FillGrid()
		Dim Assumptions As Short
		'UPGRADE_WARNING: Arrays in structure rs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim rs As DAO.Recordset
		Dim BuildType As String
		Dim n As Short
		Dim fieldcnt As Short
		Dim MaxEffGain As Short
		Dim strID As String
		Dim InitialEff As Short
		Dim ETU_factor As Single
		Dim civneeded As Single
		Dim iWF As Short
		Dim iOnShip As Short
		Dim iMil As Short
		Dim iIndex As Short
		Dim strField As String
		
		On Error Resume Next
		
		'Reset error string
		lblError.Text = vbNullString
		
		Short_Renamed = False
		ETU_factor = rsVersion.Fields("ETU per update").Value
		
		'Determmine which type of unit we are displaying
		Header = "unit curr new  cost availlcms hcms "
		Select Case ItemType
			Case enumHarborType.FH_BUILD_SHIPS
				rs = rsShip
				rs.Requery()
				BuildType = "s"
				fieldcnt = 7
				CostItems = 4
				MaxEffGain = rsVersion.Fields("Eff gain - Ship").Value
				InitialEff = 20
				Header = Header & "wf   "
				
			Case enumHarborType.FH_BUILD_LANDS, enumHarborType.FH_IMPROVE_LANDS
				rs = rsLand
				rs.Requery()
				BuildType = "l"
				fieldcnt = 8
				CostItems = 5
				MaxEffGain = rsVersion.Fields("Eff gain - Land").Value
				InitialEff = 10
				Header = Header & "guns "
				
			Case enumHarborType.FH_BUILD_PLANES
				rs = rsPlane
				rs.Requery()
				BuildType = "p"
				CostItems = 5
				fieldcnt = 8
				MaxEffGain = rsVersion.Fields("Eff gain - Plane").Value
				InitialEff = 10
				Header = Header & "mil  "
		End Select
		
		'UPGRADE_WARNING: Lower bound of array Costs was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim Costs(7, CostItems)
		
		'Costs array
		'First Dim
		'1) holds full cost
		'2) holds non rounded cost for current unit
		'3) holds rounded (integer) cost for current unit)
		'4) holds cumulative remander (used in average case)
		'5) holds starting goods
		'6) holds the required goods
		'7) holds the expended goods
		
		'Second Dim
		'1) Cost
		'2) Avail
		'3) lcm
		'4) hcm
		'5) mil or guns
		
		'Setup grid
		grid1.Visible = False
		grid1.Clear()
		grid1.Rows = 2
		grid1.Cols = fieldcnt
		grid1.set_ColWidth(0, -1) 'return to default
		grid1.set_ColWidth(-1, grid1.get_ColWidth(0) * 0.6)
		grid1.set_ColWidth(0, grid1.get_ColWidth(0) * 1.5)
		
		'Fill in row headers
		grid1.col = 0
		grid1.row = 0
		
		For n = 0 To fieldcnt - 1
			grid1.col = n
			grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
			grid1.CellFontBold = True
			grid1.Text = Trim(Mid(Header, n * 5 + 1, 5))
		Next n
		
		'Clear result labels
		Label2(0).Text = vbNullString
		Label2(1).Text = vbNullString
		Label2(2).Text = vbNullString
		
		'Initial resources
		Costs(5, 1) = 9999999 'Assume enough money
		Costs(5, 2) = Val(Text1(0).Text) 'avail
		Costs(5, 3) = Val(Text1(2).Text) 'lcm
		Costs(5, 4) = Val(Text1(3).Text) 'hcm
		Costs(5, 5) = Val(Text1(1).Text) 'guns or mil
		For n = 1 To CostItems
			Costs(4, n) = 0.5
		Next n
		
		'get assumptions
		If Option1(0).Checked Then
			Assumptions = FH_BEST_CASE
		ElseIf Option1(1).Checked Then 
			Assumptions = FH_AVERAGE_CASE
		Else
			Assumptions = FH_WORST_CASE
		End If
		
		'Fill in rows with new builds
		For n = LBound(strBuilds) + 1 To UBound(strBuilds)
			strID = VB.Left(strBuilds(n), InStr(strBuilds(n), "(") - 1)
			FillGridRow(strID, Costs, BuildType, Trim(VB.Left(strBuilds(n), 5)), Assumptions, 0, InitialEff, 0)
		Next n
		
		'Fill in rows with actual
		rs.MoveFirst()
		While Not rs.EOF
			
			'Check to see if it is in this hex
			If secx = rs.Fields("x").Value And secy = rs.Fields("y").Value Then
				'check to see that it is not 100% already
				If rs.Fields("eff").Value < 100 And (VersionCheck(4, 3, 6) = WinAceRoutines.enumVersion.VER_LESS Or rs.Fields("off").Value = 0) Then
					If BuildType = "p" Then
						iOnShip = rs.Fields("land").Value
					Else
						iOnShip = -1
					End If
					iWF = WorkForce(BuildType, rs.Fields("type").Value, rs.Fields("civ").Value, rs.Fields("mil").Value, iOnShip) * ETU_factor
					FillGridRow(rs.Fields("type").Value + " #" + CStr(rs.Fields("id").Value), Costs, BuildType, rs.Fields("type").Value, Assumptions, rs.Fields("eff").Value, MaxEffGain, iWF)
				End If
			End If
			rs.MoveNext()
		End While
		
		'Finish new builds
		For n = LBound(strBuilds) + 1 To UBound(strBuilds)
			strID = VB.Left(strBuilds(n), InStr(strBuilds(n), "(") - 1)
			iWF = 0
			If frmOptions.bSP2005 And BuildType = "s" Then
				rsBuild.Seek("=", BuildType, Trim(VB.Left(strBuilds(n), 5)))
				If Not rsBuild.NoMatch Then
					iMil = 0
					iIndex = 1
					Do 
						strField = rsBuild.Fields("cap" & CStr(iIndex)).Value
						If VB.Right(strField, 1) = "m" Or (VB.Left(strField, 1) >= "0" And VB.Left(strField, 1) <= "9") Then
							iMil = CShort(VB.Left(strField, Len(strField) - 1))
							Exit Do
						End If
						iIndex = iIndex + 1
					Loop While rs.Fields("cap" & CStr(iIndex)).Value <> "" And iIndex < 10
					iWF = WorkForce(BuildType, Trim(VB.Left(strBuilds(n), 5)), 0, iMil, False) * ETU_factor
				End If
			End If
			FillGridRow(strID, Costs, BuildType, Trim(VB.Left(strBuilds(n), 5)), Assumptions, InitialEff, MaxEffGain, iWF)
		Next n
		
		'Fill resource expended string
		For n = 2 To CostItems
			If Costs(7, n) > 0 Then
				Label2(0).Text = Label2(0).Text & CStr(Costs(7, n)) & " " & Trim(Mid(Header, (n + 2) * 5 + 1, 5)) & " "
			End If
		Next n
		
		If Len(Label2(0).Text) = 0 Then 'make sure something was built
			Label2(0).Text = "Nothing Built"
		End If
		'Fill cost string
		Label2(2).Text = "Cost: $" & CStr(Costs(7, 1))
		
		'Fill resource shorted string
		For n = 2 To CostItems
			If Costs(6, n) > Costs(5, n) Then
				Label2(1).Text = Label2(1).Text & CStr(Costs(6, n) - Costs(5, n)) & " " & Trim(Mid(Header, (n + 2) * 5 + 1, 5)) & " "
				Short_Renamed = True
				If n = 2 And HarborAvail Then 'for avail - calc civs needed
					civneeded = (Costs(6, n) - Costs(5, n)) / (ETU_factor / 100)
					civneeded = civneeded / (1 + (ETU_factor * rsVersion.Fields("Civ Birthrate").Value / 1000))
					Label2(1).Text = Label2(1).Text & "(" & CStr(Int(civneeded + 0.9)) & " civs) "
				End If
			End If
		Next n
		If Len(Label2(1).Text) = 0 Then
			Label2(1).Text = "None"
		Else
			Label2(2).Text = Label2(2).Text & " ($" & CStr(Costs(6, 1)) & " for full build)"
		End If
		Label2(1).Text = "Short: " & Label2(1).Text
		
		'Fill resource left string
		Label2(3).Text = vbNullString
		For n = 2 To CostItems
			If Costs(6, n) > Costs(7, n) Then
				Label2(3).Text = Label2(3).Text & CStr(Costs(6, n) - Costs(7, n)) & " " & Trim(Mid(Header, (n + 2) * 5 + 1, 5)) & " "
			End If
		Next n
		If Len(Label2(3).Text) = 0 Then
			For n = 2 To CostItems
				If Costs(5, n) > Costs(7, n) Then
					If n = 2 And iMaxLeft < CDbl(CStr(Costs(5, n) - Costs(7, n))) Then
						Label2(3).Text = Label2(3).Text & CStr(iMaxLeft)
					Else
						Label2(3).Text = Label2(3).Text & CStr(Costs(5, n) - Costs(7, n))
					End If
					Label2(3).Text = Label2(3).Text & " " & Trim(Mid(Header, (n + 2) * 5 + 1, 5)) & " "
				End If
			Next n
			Label2(3).Text = "Not used: " & Label2(3).Text
		Else
			Label2(3).Text = "Left: " & Label2(3).Text
		End If
		
Error_Exit: 
		'Reshow the grid
		grid1.Visible = True
	End Sub
	
	Private Sub FillGridRow(ByRef UnitDes As String, ByRef Costs() As Single, ByRef BuildType As String, ByRef UnitType As String, ByRef Assumptions As Short, ByVal CurrEff As Short, ByVal MaxEffGain As Short, ByRef iWF As Short)
		Dim OutofResources As Boolean
		Dim NewEff As Short
		Dim n As Short
		Dim X As Integer '130304 rjk: Changed to Long to Large avail requirements from St@r W@rs
		Dim EffGain As Single
		'UPGRADE_WARNING: Lower bound of array hldcosts was changed from 1,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		Dim hldcosts(2, 5) As Single
		Dim cRowColor As Integer
		Dim fWF As Single
		
		fWF = CSng(iWF) / 100#
		'get the build requirements
		rsBuild.Seek("=", BuildType, UnitType)
		If rsBuild.NoMatch Then
			lblError.Text = "Type " & UnitType & " unknown"
		Else
			'Compute efficieny gain
			NewEff = CurrEff + MaxEffGain
			If NewEff > 100 Then
				NewEff = 100
			End If
			EffGain = (NewEff - CurrEff) / 100
			
			'load costs
			Costs(1, 1) = rsBuild.Fields("cost").Value
			'030404 rjk: Switched to use the avail from the server instead of calculating it,
			'            in case the calculation in the server is changed.
			'            20 + rsBuild.Fields("lcm") + rsBuild.Fields("hcm") * 2
			Costs(1, 2) = rsBuild.Fields("avail").Value
			Costs(1, 3) = rsBuild.Fields("lcm").Value
			Costs(1, 4) = rsBuild.Fields("hcm").Value
			Select Case ItemType
				Case enumHarborType.FH_BUILD_SHIPS
				Case enumHarborType.FH_BUILD_PLANES
					Costs(1, 5) = rsBuild.Fields("mil").Value
				Case enumHarborType.FH_BUILD_LANDS, enumHarborType.FH_IMPROVE_LANDS
					Costs(1, 5) = rsBuild.Fields("gun").Value
			End Select
			
			'compute actual costs
			ComputeCosts(Costs, EffGain, Assumptions, 0)
			
			'check for overages (not enough resouces)
			OutofResources = False
			For n = 2 To CostItems
				X = Costs(5, n) - Costs(7, n)
				If n = 2 Then
					X = X + fWF
				End If
				If X <= 0 And Costs(3, n) > 0 Then
					OutofResources = True
				End If
			Next n
			
			'if we are not out of anything, check that we have enough
			'for a full build
			If Not OutofResources Then
				For n = 2 To CostItems
					X = Costs(5, n) - Costs(7, n)
					If n = 2 Then
						X = X + fWF
					End If
					If Costs(3, n) > X And Costs(3, n) > 0 Then
						OutofResources = True
						If (X / Costs(1, n)) < EffGain Then
							EffGain = Int((X * 100) / Costs(1, n)) / 100
						End If
					End If
				Next n
				
				'If this entry caused you to go over
				If OutofResources Then
					'Save full costs that will be spent
					For n = 1 To CostItems
						hldcosts(1, n) = Costs(3, n)
						hldcosts(2, n) = Costs(4, n)
					Next n
					
					'recompute how much can be build without going over
					ComputeCosts(Costs, EffGain, Assumptions, fWF)
					
					If EffGain > 0 Then
						
						'Write out a special row
						grid1.row = grid1.row + 1
						
						'write out the description
						grid1.col = 0
						grid1.CellForeColor = System.Drawing.Color.Black
						grid1.Text = UnitDes
						grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
						
						'old efficiency
						grid1.col = 1
						grid1.Text = CStr(CurrEff)
						
						'new efficency
						grid1.col = 2
						grid1.Text = CStr(CurrEff + CShort(EffGain * 100))
						'write out values
						For n = 1 To CostItems
							'Save actual costs that will be spent
							Costs(7, n) = Costs(7, n) + Costs(3, n)
							Costs(6, n) = Costs(6, n) + Costs(3, n)
							grid1.col = n + 2
							grid1.Text = CStr(CShort(Costs(3, n)))
							If grid1.Text = "0" Then
								grid1.Text = vbNullString
							End If
							grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
						Next n
						grid1.Rows = grid1.Rows + 1 '130304 rjk: Only add a row if there is efficiency gain
					End If
					
					'now reset current eff and costs
					CurrEff = CurrEff + CShort(EffGain * 100)
					EffGain = (NewEff - CurrEff) / 100
					For n = 1 To CostItems
						Costs(3, n) = hldcosts(1, n) - Costs(3, n)
						Costs(4, n) = hldcosts(2, n)
					Next n
				End If
			End If
			grid1.row = grid1.row + 1
			
			If (CurrEff + CShort(EffGain * 100)) >= 100 Then
				cRowColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue)
			Else
				cRowColor = vbMediumGreen
			End If
			If OutofResources Then
				cRowColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
			End If
			
			'write out the description
			grid1.col = 0
			grid1.CellForeColor = System.Drawing.ColorTranslator.FromOle(cRowColor)
			
			grid1.Text = UnitDes ' Set in case lookup fails
			grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignLeftCenter
			
			'old efficiency
			grid1.col = 1
			grid1.CellForeColor = System.Drawing.ColorTranslator.FromOle(cRowColor)
			grid1.Text = CStr(CurrEff)
			
			'new efficency
			grid1.col = 2
			grid1.CellForeColor = System.Drawing.ColorTranslator.FromOle(cRowColor)
			grid1.Text = CStr(NewEff)
			
			'write out values
			For n = 1 To CostItems
				grid1.col = n + 2
				Costs(6, n) = Costs(6, n) + Costs(3, n)
				If Not OutofResources Then
					Costs(7, n) = Costs(7, n) + Costs(3, n)
				End If
				grid1.CellForeColor = System.Drawing.ColorTranslator.FromOle(cRowColor)
				'negatives are in red
				If Costs(6, n) > Costs(5, n) Then
					grid1.CellForeColor = System.Drawing.Color.Red
				ElseIf OutofResources Then 
					grid1.CellForeColor = System.Drawing.ColorTranslator.FromOle(RGB(127, 127, 127))
				End If
				grid1.Text = CStr(CShort(Costs(3, n)))
				If grid1.Text = "0" Then
					grid1.Text = vbNullString
				End If
				grid1.CellAlignment = MSFlexGridLib.AlignmentSettings.flexAlignRightCenter
			Next n
			grid1.Rows = grid1.Rows + 1
		End If
	End Sub
	
	Private Sub ComputeCosts(ByRef Costs() As Single, ByRef EffGain As Single, ByRef Assumption As Short, ByRef fWF As Single)
		Dim n As Short
		
		For n = 1 To CostItems
			If n = 2 And fWF > 0# Then
				If EffGain > (fWF / Costs(1, n)) Then
					Costs(2, n) = Costs(1, n) * (EffGain - (fWF / Costs(1, n)))
				Else
					Costs(2, n) = 0
				End If
			Else
				Costs(2, n) = Costs(1, n) * EffGain
			End If
			If Assumption = FH_BEST_CASE Then
				Costs(3, n) = CSng(Int(Costs(2, n)))
			ElseIf Assumption = FH_WORST_CASE Then 
				Costs(3, n) = CSng(Int(Costs(2, n) + 0.995))
			Else
				Costs(3, n) = CSng(Int(Costs(2, n)))
				Costs(4, n) = Costs(4, n) + Costs(2, n) - Costs(3, n)
				If Costs(4, n) > 1 Then
					Costs(3, n) = Costs(3, n) + 1
					Costs(4, n) = Costs(4, n) - 1
				End If
			End If
		Next n
	End Sub
	
	Private Sub frmToolHarbor_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	'UPGRADE_WARNING: Event Option1.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option1_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option1.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = Option1.GetIndex(eventSender)
			If ItemType <> enumHarborType.FH_UNKNOWN Then
				FillGrid()
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Option2.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option2_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option2.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = Option2.GetIndex(eventSender)
			If Not OriginChange Then
				'Toggle the avail settings
				Select Case ItemType
					Case enumHarborType.FH_BUILD_SHIPS
						HarborAvail = Option2(1).Checked
					Case enumHarborType.FH_BUILD_LANDS
						HQAvail = Option2(1).Checked
					Case enumHarborType.FH_IMPROVE_LANDS
						FortAvail = Option2(1).Checked
					Case enumHarborType.FH_BUILD_PLANES
						AirportAvail = Option2(1).Checked
				End Select
				
				SaveNationVariables()
				txtOrigin_TextChanged(txtOrigin, New System.EventArgs())
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Text1.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Text1_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text1.TextChanged
		Dim Index As Short = Text1.GetIndex(eventSender)
		If ItemType <> enumHarborType.FH_UNKNOWN Then
			FillGrid()
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtOrigin.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtOrigin_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOrigin.TextChanged
		Dim sx As Short
		Dim sy As Short
		
		'use origin change to avoid firing cascading events
		OriginChange = True
		
		ReDim strBuilds(0)
		Label4.Text = vbNullString
		
		If ParseSectors(sx, sy, (txtOrigin.Text)) Then
			secx = sx
			secy = sy
			FillSectorInfo()
			If ItemType <> enumHarborType.FH_UNKNOWN Then
				FillGrid()
			End If
		Else
			HandleError("Not a valid Sector")
		End If
		OriginChange = False
	End Sub
	
	Private Sub FillSectorInfo()
		Dim iCurrentAvail As Short
		Dim iRolloverAvail As Short
		Dim iUpdateAvail As Short
		Dim v As productionDataType
		
		ItemType = enumHarborType.FH_UNKNOWN
		
		'get sector record
		rsSectors.Seek("=", secx, secy)
		If rsSectors.NoMatch Then
			HandleError("Sector not found")
			Exit Sub
		End If
		
		'set sector type
		Select Case rsSectors.Fields("des").Value
			Case "h"
				ItemType = enumHarborType.FH_BUILD_SHIPS
				Me.Text = "Harbor at " & SectorString(secx, secy)
				Text1(1).Visible = False
				Label1(1).Visible = False
				Command1.Enabled = True
				Option2(0).Checked = Not HarborAvail
				Option2(1).Checked = HarborAvail
				
			Case "!"
				ItemType = enumHarborType.FH_BUILD_LANDS
				Me.Text = "Headquarters at " & SectorString(secx, secy)
				Text1(1).Visible = True
				Label1(1).Visible = True
				Text1(1).Text = CStr(rsSectors.Fields("gun").Value)
				Label1(1).Text = "guns"
				Command1.Enabled = True
				Option2(0).Checked = Not HQAvail
				Option2(1).Checked = HQAvail
				
			Case "*"
				ItemType = enumHarborType.FH_BUILD_PLANES
				Me.Text = "Airfield at " & SectorString(secx, secy)
				Text1(1).Visible = True
				Label1(1).Visible = True
				Text1(1).Text = CStr(rsSectors.Fields("mil").Value)
				Label1(1).Text = "mil"
				Command1.Enabled = True
				Option2(0).Checked = Not AirportAvail
				Option2(1).Checked = AirportAvail
				
			Case "f"
				ItemType = enumHarborType.FH_IMPROVE_LANDS
				Me.Text = "Fortress at " & SectorString(secx, secy)
				Text1(1).Visible = True
				Label1(1).Visible = True
				Text1(1).Text = CStr(rsSectors.Fields("gun").Value)
				Label1(1).Text = "guns"
				Command1.Enabled = False
				Option2(0).Checked = Not FortAvail
				Option2(1).Checked = FortAvail
				
			Case Else
				HandleError("Sector not h,!, f, or *")
				Exit Sub 'error, not a sector type
		End Select
		
		'Check if avail used is before or after the update
		iCurrentAvail = CShort(CStr(rsSectors.Fields("avail").Value))
		
		If Len(Trim(rsSectors.Fields("sdes").Value)) = 0 Then
			iRolloverAvail = rsSectors.Fields("avail").Value
			If iRolloverAvail > rsVersion.Fields("Rollover Avail").Value Then
				iRolloverAvail = rsVersion.Fields("Rollover Avail").Value
			End If
		Else
			iRolloverAvail = 0
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		v = Production(rsSectors)
		If v.prodValidFlag Then
			'UPGRADE_WARNING: Couldn't resolve default property of object v.ProdAmount. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			iUpdateAvail = CShort(CStr(v.ProdAmount))
		Else
			iUpdateAvail = iRolloverAvail
		End If
		
		If Option2(0).Checked Then
			If VB6.GetItemData(cmbUpdates, cmbUpdates.SelectedIndex) = 1 Then
				iMaxLeft = iCurrentAvail
				Text1(0).Text = CStr(iCurrentAvail)
			Else
				If v.prodValidFlag Then
					iMaxLeft = iUpdateAvail
					Text1(0).Text = CStr(iCurrentAvail + iRolloverAvail + ((iUpdateAvail - iRolloverAvail) * (VB6.GetItemData(cmbUpdates, cmbUpdates.SelectedIndex) - 1)))
				Else
					iMaxLeft = iCurrentAvail
					Text1(0).Text = CStr(iCurrentAvail)
				End If
			End If
		Else
			If v.prodValidFlag Then
				iMaxLeft = iUpdateAvail
				Text1(0).Text = CStr(iRolloverAvail + ((iUpdateAvail - iRolloverAvail) * VB6.GetItemData(cmbUpdates, cmbUpdates.SelectedIndex)))
			Else
				iMaxLeft = iRolloverAvail
				Text1(0).Text = CStr(iRolloverAvail)
			End If
		End If
		
		If frmOptions.bSPAtlantis Then
			Text1(2).Text = CStr(rsSectors.Fields("shell").Value)
		Else
			Text1(2).Text = CStr(rsSectors.Fields("lcm").Value)
		End If
		Text1(3).Text = CStr(rsSectors.Fields("hcm").Value)
	End Sub
	
	Private Sub HandleError(ByRef ErrMsg As String)
		lblError.Text = ErrMsg
		Text1(0).Text = vbNullString
		Text1(1).Text = vbNullString
		Text1(2).Text = vbNullString
		Text1(3).Text = vbNullString
		Label2(0).Text = vbNullString
		Label2(1).Text = vbNullString
		Label2(2).Text = vbNullString
		grid1.Clear()
	End Sub
	
	Private Function WorkForce(ByRef strBuildType As String, ByRef strType As String, ByRef iCiv As Short, ByRef iMil As Short, ByRef iOnShip As Short) As Short
		Dim hldIndex As String
		
		WorkForce = 0
		hldIndex = ""
		
		On Error GoTo WorkForce_error
		Select Case strBuildType
			Case "s"
				If Not (rsBuild.BOF And rsBuild.EOF) Then
					rsBuild.Seek("=", strBuildType, strType)
					If Not rsBuild.NoMatch Then
						If rsBuild.Fields("stat6").Value > 0 Then
							WorkForce = iMil / 2
						Else
							WorkForce = iCiv / 2 + iMil / 5
						End If
					End If
				End If
			Case "p"
				If iOnShip <> -1 Then
					If Not (rsShip.BOF And rsShip.EOF) Then
						hldIndex = rsShip.Index
						rsShip.Index = "PrimaryKey"
						rsShip.Seek("=", iOnShip)
						If Not rsShip.NoMatch Then
							WorkForce = rsShip.Fields("mil").Value / 2
						End If
					End If
				End If
		End Select
		
		WorkForce = WorkForce * VB6.GetItemData(cmbUpdates, cmbUpdates.SelectedIndex)
		If hldIndex <> "" Then
			rsShip.Index = hldIndex
		End If
		Exit Function
		
WorkForce_error: 
		If hldIndex <> "" Then
			rsShip.Index = hldIndex
		End If
		WorkForce = 0
	End Function
End Class