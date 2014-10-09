Option Strict Off
Option Explicit On
Friend Class frmPromptTrans
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'xx0803 efj: removed dead variables
	'180903 rjk: Added Unit grid selection when activating this form or selecting fields
	'            and when disactivating return to standard Sector display.
	'            General reformatting. Added setting the initial field on start up.
	'250903 rjk: Added Quantity for nukes.
	'            Added Origin label.  Added timestamp ndump. Added path cost
	'            for nukes.
	'021003 rjk: Added check to see if there is also a nuke or plane in the sector before enabling second option.
	'051003 rjk: Removed timestamp ndump
	'061003 rjk: Changed Transport submit command to True so it appears in the F9 list.
	'141003 rjk: Added ndump timestamp (try #2)
	'171003 rjk: Added Strength fields to Sector database
	'181103 rjk: Moved DisplayPath to frmDrawMap
	'201103 rjk: Added option to control strength updates
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	'180206 rjk: Replace pdump, ndump and lost with GetPlanes, GetNukes and GetLost.
	'120305 rjk: Fixed transport costs for SP2005.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	'060506 rjk: Added supported for Nuke id changes in 4.3.3 servers.
	'140506 rjk: Fixed nuke transport to select nuke grid.
	'190606 rjk: Incorporate the sector mobility changes for 4.3.6 servers.
	'030706 rjk: Fixed the sector updating for nukes.  Update the grid when transporting.
	
	Public strCmd As String
	Public strSector As String
	
	'UPGRADE_WARNING: Event chkDisplayPath.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkDisplayPath_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkDisplayPath.CheckStateChanged
		ComputePathCost()
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		strCmd = vbNullString
		frmDrawMap.DisplayFirstSectorPanel()
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim X As Short
		Dim wp As Object
		Dim strFill1 As String
		Dim strFill2 As String
		Dim n As Short
		Dim CurrUnit As Object
		
		On Error Resume Next
		
		'Start with first list box results
		If Option1.Checked Then
			strCmd = "plane"
			strFill1 = " "
			strFill2 = " "
		Else
			strCmd = "nuke"
			If VersionCheck(4, 3, 3) <> WinAceRoutines.enumVersion.VER_LESS Then
				strFill1 = " "
				strFill2 = " "
			Else
				strFill1 = " " & strSector & " "
				strFill2 = " " & txtQuantity.Text & " "
			End If
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object ParseWaypoints(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object wp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		wp = ParseWaypoints((txtPath.Text))
		Dim strCmd2 As String
		If IsArray(wp) Then
			frmEmpCmd.SubmitEmpireCommand("bf1", True)
			For X = 1 To UBound(wp)
				'UPGRADE_WARNING: Couldn't resolve default property of object wp(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strCmd2 = "transport " & strCmd & strFill1 & txtUnit.Text & strFill2 + wp(X)
				If strCmd = "nuke" And VersionCheck(4, 3, 3) = WinAceRoutines.enumVersion.VER_LESS Then
					'UPGRADE_WARNING: Couldn't resolve default property of object wp(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strFill1 = " " + wp(X) + " "
				End If
				frmEmpCmd.SubmitEmpireCommand(strCmd2, True)
			Next X
			frmEmpCmd.SubmitEmpireCommand("bf2", True)
		Else
			strCmd = "transport " & strCmd & strFill1 & txtUnit.Text & strFill2 & txtPath.Text
			frmEmpCmd.SubmitEmpireCommand(strCmd, True) '100603 rjk: Changed to True so the command appears in the F9 list.
		End If
		
		'Check if planes moved have missions.  If so, remove them
		If Len(txtPath.Text) > 0 And Option1.Checked Then
			'UPGRADE_WARNING: Couldn't resolve default property of object GetUnitList(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CurrUnit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CurrUnit = GetUnitList((txtUnit.Text), "p")
			If IsArray(CurrUnit) Then
				n = LBound(CurrUnit) + 1
				Do While n <= UBound(CurrUnit)
					'UPGRADE_WARNING: Couldn't resolve default property of object CurrUnit(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					rsMissions.Seek("=", "p", CShort(CurrUnit(n)))
					If Not rsMissions.NoMatch Then
						rsMissions.Delete()
					End If
					n = n + 1
				Loop 
			End If
		End If
		
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		If Option1.Checked Then
			GetPlanes()
		End If
		If Option2.Checked Then
			GetNukes()
			GetLost()
		End If
		
		GetSectors()
		'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
		GetCurrentStrength(tsSectors)
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
		Me.Close()
	End Sub
	
	'UPGRADE_WARNING: Form event frmPromptTrans.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmPromptTrans_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		'100203 rjk: Added check to see if there is also a nuke or plane in the sector before enabling second option.
		Dim sx As Short
		Dim sy As Short
		
		lblOrigin.Text = strSector
		If Option1.Checked Then
			txtQuantity.Visible = False
			lblQuantity.Visible = False
			If ParseSectors(sx, sy, strSector) Then
				rsNuke.Seek("=", sx, sy)
				If rsNuke.NoMatch Then
					Option2.Enabled = False
				End If
			Else
				Option2.Enabled = False
			End If
		Else
			If VersionCheck(4, 3, 3) <> WinAceRoutines.enumVersion.VER_LESS Then
				txtQuantity.Visible = False
				txtQuantity.Text = "1"
				lblQuantity.Visible = False
			Else
				txtQuantity.Visible = True
				lblQuantity.Visible = True
			End If
			If ParseSectors(sx, sy, strSector) Then
				rsPlane.Seek("=", sx, sy)
				If rsPlane.NoMatch Then
					Option1.Enabled = False
				End If
			Else
				Option1.Enabled = False
			End If
		End If
	End Sub
	
	Private Sub frmPromptTrans_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		' Set parent for the toolbar to display on top of:
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
		txtPath.Text = strSector
	End Sub
	
	Private Sub frmPromptTrans_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
		
		If frmDrawMap.DisplayingPath Then
			frmDrawMap.DisplayingPath = False
			frmDrawMap.FinishPath()
		End If
		
	End Sub
	
	'UPGRADE_WARNING: Event Option1.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option1_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option1.CheckedChanged
		If eventSender.Checked Then
			ComputePathCost()
			If txtUnit.Visible Then
				txtUnit.Focus()
			End If
			txtQuantity.Visible = False
			txtQuantity.Text = "1"
			lblQuantity.Visible = False
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Option2.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option2_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option2.CheckedChanged
		If eventSender.Checked Then
			ComputePathCost()
			If txtUnit.Visible Then
				txtUnit.Focus()
			End If
			If VersionCheck(4, 3, 3) <> WinAceRoutines.enumVersion.VER_LESS Then
				txtQuantity.Visible = False
				txtQuantity.Text = "1"
				lblQuantity.Visible = False
			Else
				txtQuantity.Visible = True
				lblQuantity.Visible = True
			End If
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtPath.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtPath_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPath.TextChanged
		ComputePathCost()
	End Sub
	
	Private Sub txtPath_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPath.DoubleClick
		'Make sure starting sector is valid before prompting
		Dim sx As Short
		Dim sy As Short
		If Not ParseSectors(sx, sy, strSector) Then
			Exit Sub
		End If
		
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
		Load(frmPromptPath)
		frmPromptPath.lblSector.Text = strSector
		frmPromptPath.lblEndSector.Text = strSector
		frmPromptPath.Text = "Select Transport Route"
		frmPromptPath.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(frmDrawMap.Left) + VB6.PixelsToTwipsX(frmDrawMap.Width) - VB6.PixelsToTwipsX(frmPromptPath.Width))
		frmPromptPath.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(frmDrawMap.Top) + (VB6.PixelsToTwipsY(frmDrawMap.Height) - VB6.PixelsToTwipsY(frmPromptPath.Height)) \ 2)
		frmPromptPath.Show()
	End Sub
	
	Private Sub ComputePathCost()
		On Error Resume Next
		Dim pvar As Object
		Dim pcost As Single
		Dim pweight As Short
		Dim mused As Single
		Dim mleft As Single
		Dim tweight As Integer
		Dim oldx As Short
		Dim oldy As Short
		Dim sx As Short
		Dim sy As Short
		Dim n As Short
		Dim nx As Short
		Dim ny As Short
		Dim hldIndex As Object
		'UPGRADE_WARNING: Arrays in structure rs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim rs As DAO.Recordset
		Dim bpath As String
		Dim UnitList As Object
		Dim strNukeType As String
		Dim quantity As Short
		
		lblLeft.Text = vbNullString
		lblCost.Text = vbNullString
		lblBestPath.Text = vbNullString
		lblPathCost.Text = vbNullString
		
		If frmDrawMap.DisplayingPath Then
			frmDrawMap.DisplayingPath = False
			frmDrawMap.FinishPath()
		End If
		
		'planes only for now
		' 092503 rjk: Added code for Nukes
		'If Not Option1.Value Then
		'    Exit Sub
		'End If
		
		If Len(txtPath.Text) = 0 Then
			Exit Sub
		End If
		
		'Don't do costs for waypoints
		If InStr(txtPath.Text, ";") > 0 Then
			Exit Sub
		End If
		
		If Option1.Checked Then
			'UPGRADE_WARNING: Couldn't resolve default property of object GetUnitList(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object UnitList. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			UnitList = GetUnitList((txtUnit.Text), "p")
			If Not IsArray(UnitList) Then
				Exit Sub
			End If
			
			'set indexs
			rs = rsPlane
			'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			hldIndex = rs.Index
			rs.Index = "PrimaryKey"
			tweight = 0
			oldx = -1
			oldy = 0
			
			For n = LBound(UnitList) To UBound(UnitList)
				If Len(UnitList(n)) > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object UnitList(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					rs.Seek("=", CShort(UnitList(n)))
					If Not rs.NoMatch Then
						sx = rs.Fields("x").Value
						sy = rs.Fields("y").Value
						If (oldx <> sx) Or (oldy <> sy) Then
							'Compute the path cost
							If ParseSectors(nx, ny, (txtPath.Text)) Then
								'UPGRADE_WARNING: Couldn't resolve default property of object BestPath(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object pvar. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								pvar = BestPath(SectorString(sx, sy), (txtPath.Text), EmpCommon.enumMobType.MT_COMMODITY)
								'UPGRADE_WARNING: Couldn't resolve default property of object pvar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								bpath = pvar(1)
								'UPGRADE_WARNING: Couldn't resolve default property of object pvar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								pcost = pvar(2)
							Else
								pcost = PathCost(SectorString(sx, sy), (txtPath.Text), EmpCommon.enumMobType.MT_COMMODITY)
								bpath = txtPath.Text
							End If
							oldx = sx
							oldy = sy
						End If
						rsBuild.Seek("=", "p", rs.Fields("type"))
						If rsBuild.NoMatch Then
							'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							rs.Index = hldIndex
							Exit Sub
						End If
						If frmOptions.bSP2005 Then
							pweight = rsBuild.Fields("avail").Value - 20
						Else
							pweight = rsBuild.Fields("hcm").Value * 2 + rsBuild.Fields("lcm").Value
						End If
						tweight = tweight + pweight
					End If
				End If
			Next n
		Else 'Nukes
			If VersionCheck(4, 3, 6) <> WinAceRoutines.enumVersion.VER_LESS Then
				'UPGRADE_WARNING: Couldn't resolve default property of object GetUnitList(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object UnitList. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				UnitList = GetUnitList((txtUnit.Text), "n")
				If Not IsArray(UnitList) Then
					Exit Sub
				End If
				
				'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				hldIndex = rsNuke.Index
				rsNuke.Index = "id"
				For n = LBound(UnitList) To UBound(UnitList)
					If Len(UnitList(n)) > 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object UnitList(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						rsNuke.Seek("=", CShort(UnitList(n)))
						If Not rsNuke.NoMatch Then
							rsBuild.Seek("=", "n", rsNuke.Fields("type"))
							If Not rsBuild.NoMatch Then
								tweight = tweight + rsBuild.Fields("stat3").Value 'stat3 is lbs field
							End If
							If ParseSectors(nx, ny, (txtPath.Text)) Then
								'UPGRADE_WARNING: Couldn't resolve default property of object BestPath(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object pvar. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								pvar = BestPath(strSector, (txtPath.Text), EmpCommon.enumMobType.MT_COMMODITY)
								'UPGRADE_WARNING: Couldn't resolve default property of object pvar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								bpath = pvar(1)
								'UPGRADE_WARNING: Couldn't resolve default property of object pvar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								pcost = pvar(2)
							Else
								pcost = PathCost(strSector, (txtPath.Text), EmpCommon.enumMobType.MT_COMMODITY)
								bpath = txtPath.Text
							End If
						End If
					End If
				Next n
				'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				rsNuke.Index = hldIndex
			Else
				ParseSectors(sx, sy, strSector)
				If Len(txtUnit.Text) > 0 And Len(txtQuantity.Text) > 0 Then
					quantity = Val(txtQuantity.Text)
					
					rsBuild.Seek("=", "n", txtUnit.Text)
					If Not rsBuild.NoMatch Then
						tweight = rsBuild.Fields("stat3").Value * quantity 'stat3 is lbs field
					End If
					'Compute the path cost
					
					If ParseSectors(nx, ny, (txtPath.Text)) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object BestPath(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object pvar. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						pvar = BestPath(strSector, (txtPath.Text), EmpCommon.enumMobType.MT_COMMODITY)
						'UPGRADE_WARNING: Couldn't resolve default property of object pvar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						bpath = pvar(1)
						'UPGRADE_WARNING: Couldn't resolve default property of object pvar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						pcost = pvar(2)
					Else
						pcost = PathCost(strSector, (txtPath.Text), EmpCommon.enumMobType.MT_COMMODITY)
						bpath = txtPath.Text
					End If
				End If
			End If
		End If
		mused = tweight * pcost
		lblCost.Text = "mob. cost: " & CStr(CSng(CInt(CSng(mused) * 1000) / 1000))
		
		rsSectors.Seek("=", sx, sy)
		If Not rsSectors.NoMatch Then
			mleft = rsSectors.Fields("mob").Value - mused
			If mleft < 0 Then
				lblLeft.Text = "Insuff. mobility"
			Else
				lblLeft.Text = "mob. left: " & CStr(Int(mleft))
			End If
		End If
		
		lblBestPath.Text = "path: " & bpath
		lblPathCost.Text = "path cost: " & CStr(CSng(CInt(CSng(pcost) * 1000) / 1000))
		If chkDisplayPath.CheckState = System.Windows.Forms.CheckState.Checked Then
			frmDrawMap.DisplayPath(SectorString(sx, sy), bpath)
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rs.Index = hldIndex
	End Sub
	
	Private Sub txtPath_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPath.Enter
		frmDrawMap.DisplayFirstSectorPanel()
	End Sub
	
	Private Sub txtPath_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPath.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If KeyAscii = 32 And (LCase(Trim(Label2.Text)) = "transport") Then
			KeyAscii = 0
			'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
			Load(frmPromptWaypoints)
			frmPromptWaypoints.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(frmDrawMap.Left) + VB6.PixelsToTwipsX(frmDrawMap.Width) - VB6.PixelsToTwipsX(frmPromptWaypoints.Width))
			frmPromptWaypoints.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(frmDrawMap.Top) + (VB6.PixelsToTwipsY(frmDrawMap.Height) - VB6.PixelsToTwipsY(frmPromptWaypoints.Height)) \ 2)
			frmPromptWaypoints.Show()
			frmPromptWaypoints.txtOrigin.Focus()
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtUnit.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtUnit_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnit.TextChanged
		ComputePathCost()
	End Sub
	
	'UPGRADE_WARNING: Event txtQuantity.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtQuantity_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtQuantity.TextChanged
		ComputePathCost()
	End Sub
	
	Private Sub txtUnit_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnit.Enter
		If Option1.Checked Then
			frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udPLANE, frmDrawMap.enumUnitType.TYPE_ALL, False, False, False)
		Else
			frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udNUKE, frmDrawMap.enumUnitType.TYPE_ALL, False, False, False)
		End If
	End Sub
End Class