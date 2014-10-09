Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmPromptNav
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'051303 rjk: replaced DisplaySectorPanel 1 with DisplayFirstSectorPanel
	'                       replaced DisplaySectorPanel 3 with DisplayUnitBox
	'                       changes to accomodate how the tabs are controlled in
	'                       SectorPanel
	'081203 efj: removed dead variables
	'092703 rjk: general reformatting, removed hard coding for unit types
	'            set OK to be the default button
	'093003 rjk: added RouteInfo for indicating if multiple origins are present
	'093003 rjk: if multiple origins are present can not selected a target sector.
	'111803 rjk: Moved DisplayPath to frmDrawMap
	'120303 rjk: Switched to frmOptions and boolean options.
	'120703 rjk: Changed hldindex to hldIndex.
	'220404 rjk: Code cleanup.
	'140505 rjk: Fixed ship mobility cost to deal with the extra mobility costs with
	'            ships with SWEEP capabilities.  Removed correction factor from NavMarkers.
	'160705 rjk: Added ability to get OContent for Sea Sectors
	'180206 rjk: Replace sdump, ldump, and pdump with GetShips, GetLandUnits and GetPlanes.
	'190606 rjk: Incorporate the sector mobility changes for 4.3.6 servers.
	'160706 rjk: Added AutoView for navigate and march.
	'181006 rjk: Remove AutoView for march.
	'140708 rjk: Added GetLost to navigate and march.
	
	Public strSector As String
	Dim UnitList As Object
	
	'drk 7/13/02
	'UPGRADE_WARNING: Event chkRefreshBMapOnStop.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkRefreshBMapOnStop_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkRefreshBMapOnStop.CheckStateChanged
		frmNavCtrl.RefreshBmapOnStop = chkRefreshBMapOnStop.CheckState
	End Sub
	
	'UPGRADE_WARNING: Event chkDisplayPath.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkDisplayPath_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkDisplayPath.CheckStateChanged
		ComputePathCost()
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim strCmd As String
		Dim strCmd2 As String
		Dim wp As Object
		Dim X As Short
		Dim n As Short
		' Dim rs As Recordset    8/12/03 efj  removed
		Dim strType As String
		Dim CurrUnit As Object
		Dim strInterdict As String
		Dim strReserve As String
		Dim hldIndex As Object
		
		
		If Len(txtUnit.Text) = 0 Then
			Exit Sub
		End If
		
		frmNavCtrl.AutoLooking_when_notAutoStopping = False
		frmNavCtrl.AutoViewing_when_notAutoStopping = False
		
		'if the path is blank, ignore the autostop
		strCmd = LCase(Label2.Text) & " " & txtUnit.Text & " "
		strCmd2 = vbNullString
		
		'UPGRADE_WARNING: Couldn't resolve default property of object ParseWaypoints(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object wp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		wp = ParseWaypoints((txtPath.Text))
		If IsArray(wp) Then
			frmEmpCmd.SubmitEmpireCommand("bf1", True)
			For X = 1 To UBound(wp)
				If chkLookAlongTheWay.CheckState = System.Windows.Forms.CheckState.Checked Then
					If chkView.CheckState = System.Windows.Forms.CheckState.Checked And chkView.Visible Then
						frmEmpCmd.SubmitEmpireCommand("as1lv", False)
					Else
						frmEmpCmd.SubmitEmpireCommand("as1l", False)
					End If
				ElseIf chkView.CheckState = System.Windows.Forms.CheckState.Checked And chkView.Visible Then 
					frmEmpCmd.SubmitEmpireCommand("as1v", False)
				Else
					frmEmpCmd.SubmitEmpireCommand("as1 ", False)
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object wp(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strCmd2 = strCmd & " " + wp(X)
				frmEmpCmd.SubmitEmpireCommand(strCmd2, True)
			Next X
			frmEmpCmd.SubmitEmpireCommand("bf2", True)
		Else
			'Submit the autostop
			If chkStopAfterMove.CheckState = System.Windows.Forms.CheckState.Checked And Len(txtPath.Text) > 0 Then
				frmNavCtrl.AutoLooking_when_notAutoStopping = False
				frmNavCtrl.AutoViewing_when_notAutoStopping = False
				If chkLookAlongTheWay.CheckState = System.Windows.Forms.CheckState.Checked Then
					If chkView.CheckState = System.Windows.Forms.CheckState.Checked And chkView.Visible Then
						frmEmpCmd.SubmitEmpireCommand("as1lv", False)
					Else
						frmEmpCmd.SubmitEmpireCommand("as1l", False)
					End If
				ElseIf chkView.CheckState = System.Windows.Forms.CheckState.Checked And chkView.Visible Then 
					frmEmpCmd.SubmitEmpireCommand("as1v", False)
				Else
					frmEmpCmd.SubmitEmpireCommand("as1 ", False)
				End If
			ElseIf chkLookAlongTheWay.CheckState = System.Windows.Forms.CheckState.Checked Or (chkView.CheckState = System.Windows.Forms.CheckState.Checked And chkView.Visible) Then 
				If chkLookAlongTheWay.CheckState Then
					'no autostop, but we do want to look along the way
					frmNavCtrl.AutoLooking_when_notAutoStopping = True
				End If
				If chkView.CheckState = System.Windows.Forms.CheckState.Checked And chkView.Visible Then
					frmNavCtrl.AutoViewing_when_notAutoStopping = True
				End If
			End If
			
			strCmd = strCmd & txtPath.Text
			frmEmpCmd.SubmitEmpireCommand(strCmd, True)
		End If
		
		'Check if units moved have missions.  If so, remove them
		If LCase(Label2.Text) = "march" Then
			strType = "l"
		Else
			strType = "s"
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object GetUnitList(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CurrUnit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CurrUnit = GetUnitList((txtUnit.Text), strType)
		If IsArray(CurrUnit) Then
			
			n = LBound(CurrUnit) + 1
			Do While n <= UBound(CurrUnit)
				'UPGRADE_WARNING: Couldn't resolve default property of object CurrUnit(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				rsMissions.Seek("=", strType, CShort(CurrUnit(n)))
				If Not rsMissions.NoMatch Then
					rsMissions.Delete()
				End If
				n = n + 1
			Loop 
		End If
		
		'set retreat
		frmDrawMap.DoRetreat = False
		If chkStopAfterMove.CheckState = System.Windows.Forms.CheckState.Checked And Len(txtPath.Text) > 0 Then
			If chkRetreat.CheckState = System.Windows.Forms.CheckState.Checked And Len(txtRetreatPath.Text) > 0 Then
				strCmd = "retreat " & txtUnit.Text & " " & txtRetreatPath.Text & " " & VB.Left(Combo1.Text, 1)
				If LCase(Label2.Text) = "march" Then
					strCmd = "l" & strCmd
				End If
				frmEmpCmd.SubmitEmpireCommand(strCmd, True)
			End If
		ElseIf chkRetreat.CheckState <> System.Windows.Forms.CheckState.Unchecked Then 
			'Set up retreat string for NavCtrl prompt
			frmDrawMap.DoRetreat = True
			frmDrawMap.LastRetreatString = txtRetreatPath.Text
			frmDrawMap.LastRetreatUnits = txtUnit.Text
			If LCase(Label2.Text) = "march" Then
				frmDrawMap.LastRetreatType = "l"
			Else
				frmDrawMap.LastRetreatType = vbNullString
			End If
		End If
		
		'Set post move missions if requested
		If chkMission.CheckState = System.Windows.Forms.CheckState.Checked Then
			If Val(txtRadius.Text) <= 0 Then
				txtRadius.Text = vbNullString
			End If
			If strType = "s" Then
				'ships are alway interdiction
				strCmd = "mission ship " & txtUnit.Text & " i . " & txtRadius.Text
				frmEmpCmd.SubmitEmpireCommand(strCmd, True)
				
				'land can be interdiction or reserve - figure out which fits
			ElseIf IsArray(CurrUnit) Then 
				n = LBound(CurrUnit) + 1
				strInterdict = vbNullString
				strReserve = vbNullString
				'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				hldIndex = rsLand.Index
				rsLand.Index = "PrimaryKey"
				Do While n <= UBound(CurrUnit)
					'UPGRADE_WARNING: Couldn't resolve default property of object CurrUnit(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					rsLand.Seek("=", CShort(CurrUnit(n)))
					If Not rsLand.NoMatch Then
						If rsLand.Fields("frg").Value > 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object CurrUnit(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strInterdict = strInterdict & CStr(CurrUnit(n)) & "/"
						ElseIf rsLand.Fields("def").Value >= 1 Then 
							'UPGRADE_WARNING: Couldn't resolve default property of object CurrUnit(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							strReserve = strReserve & CStr(CurrUnit(n)) & "/"
						End If
					End If
					n = n + 1
				Loop 
				'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				hldIndex = rsLand.Index
				If Len(strInterdict) > 0 Then
					strCmd = "mission land " & VB.Left(strInterdict, Len(strInterdict) - 1) & " i . " & txtRadius.Text
					frmEmpCmd.SubmitEmpireCommand(strCmd, True)
				End If
				If Len(strReserve) > 0 Then
					strCmd = "mission land " & VB.Left(strReserve, Len(strReserve) - 1) & " r . " & txtRadius.Text
					frmEmpCmd.SubmitEmpireCommand(strCmd, True)
				End If
			End If
		End If
		
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		If LCase(Label2.Text) = "navigate" Then
			DoNavDumps()
			GetOContent()
		ElseIf LCase(Label2.Text) = "march" Then 
			GetLandUnits()
		Else '093003 rjk: This case should not happen as only nav and march use this form.
			GetPlanes()
		End If
		GetLost()
		
		
cmdOK_exit: 
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
		If chkStopAfterMove.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			SaveSetting(APPNAME, "frmPromptNav", "StopAfterMove", "true")
		Else
			SaveSetting(APPNAME, "frmPromptNav", "StopAfterMove", "false")
		End If
		If chkLookAlongTheWay.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			SaveSetting(APPNAME, "frmPromptNav", "LookAlongTheWay", "true")
		Else
			SaveSetting(APPNAME, "frmPromptNav", "LookAlongTheWay", "false")
		End If
		If chkRefreshBMapOnStop.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			SaveSetting(APPNAME, "frmPromptNav", "RefreshBMapOnStop", "true")
		Else
			SaveSetting(APPNAME, "frmPromptNav", "RefreshBMapOnStop", "false")
		End If
		If chkDisplayPath.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			SaveSetting(APPNAME, "frmPromptNav", "DisplayPath", "true")
		Else
			SaveSetting(APPNAME, "frmPromptNav", "DisplayPath", "false")
		End If
		If LCase(Label2.Text) = "navigate" Then
			If chkView.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
				SaveSetting(APPNAME, "frmPromptNav", "View", "true")
			Else
				SaveSetting(APPNAME, "frmPromptNav", "View", "false")
			End If
		End If
		Me.Close()
	End Sub
	
	'UPGRADE_WARNING: Form event frmPromptNav.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmPromptNav_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		If LCase(Label2.Text) = "navigate" Then
			chkView.Visible = True
		Else
			chkView.Visible = False
		End If
	End Sub
	
	Private Sub frmPromptNav_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Make Stay always on top
		'Dim success As Long  removed 8/12/03 efj
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
		ComputePathCost()
		LoadRetreatBox(Combo1)
		Combo1.SelectedIndex = 0
		If frmOptions.bDefaultRetreat Then
			chkRetreat.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkRetreat.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		
		'drk: persist these settings through the registry. I'd really prefer a
		'different method, but this works for now
		If GetSetting(APPNAME, "frmPromptNav", "StopAfterMove", "false") = "false" Then
			chkStopAfterMove.CheckState = System.Windows.Forms.CheckState.Unchecked
		Else
			chkStopAfterMove.CheckState = System.Windows.Forms.CheckState.Checked
		End If
		If GetSetting(APPNAME, "frmPromptNav", "LookAlongTheWay", "false") = "false" Then
			chkLookAlongTheWay.CheckState = System.Windows.Forms.CheckState.Unchecked
		Else
			chkLookAlongTheWay.CheckState = System.Windows.Forms.CheckState.Checked
		End If
		If GetSetting(APPNAME, "frmPromptNav", "RefreshBMapOnStop", "false") = "false" Then
			chkRefreshBMapOnStop.CheckState = System.Windows.Forms.CheckState.Unchecked
		Else
			chkRefreshBMapOnStop.CheckState = System.Windows.Forms.CheckState.Checked
		End If
		If GetSetting(APPNAME, "frmPromptNav", "DisplayPath", "false") = "false" Then
			chkDisplayPath.CheckState = System.Windows.Forms.CheckState.Unchecked
		Else
			chkDisplayPath.CheckState = System.Windows.Forms.CheckState.Checked
		End If
		If GetSetting(APPNAME, "frmPromptNav", "View", "false") = "false" Then
			chkView.CheckState = System.Windows.Forms.CheckState.Unchecked
		Else
			chkView.CheckState = System.Windows.Forms.CheckState.Checked
		End If
	End Sub
	
	Private Sub frmPromptNav_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
		If frmDrawMap.DisplayingPath Then
			frmDrawMap.DisplayingPath = False
			frmDrawMap.FinishPath()
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
		frmPromptPath.Text = "Select Route"
		frmPromptPath.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(frmDrawMap.Left) + VB6.PixelsToTwipsX(frmDrawMap.Width) - VB6.PixelsToTwipsX(frmPromptPath.Width))
		frmPromptPath.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(frmDrawMap.Top) + (VB6.PixelsToTwipsY(frmDrawMap.Height) - VB6.PixelsToTwipsY(frmPromptPath.Height)) \ 2)
		frmPromptPath.Show()
	End Sub
	
	Private Sub txtPath_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPath.Enter
		frmDrawMap.DisplayFirstSectorPanel() 'rjk 5/13/03: replaced DisplaySectorPanel 1 with DisplayFirstSectorPanel
		'             changes to accomodate how the tabs are controlled in
		'             SectorPanel
	End Sub
	
	Private Sub txtPath_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPath.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		If KeyAscii = 32 Then
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
	
	Private Sub DoNavDumps()
		Dim sx As Short
		Dim sy As Short
		Dim nID As Short
		Dim hldIndex1 As Object
		Dim hldIndex2 As Object
		Dim hldIndex3 As Object
		Dim LandFound As Boolean
		Dim PlaneFound As Boolean
		Dim oldx As Short
		Dim oldy As Short
		Dim n As Short
		
		GetShips()
		
		If Not IsArray(UnitList) Then
			Exit Sub
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hldIndex1 = rsShip.Index
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hldIndex2 = rsPlane.Index
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hldIndex3 = rsLand.Index
		rsShip.Index = "PrimaryKey"
		rsPlane.Index = "location"
		rsLand.Index = "location"
		
		'if there are planes and units in the save starting sector as
		'the ships, then go ahead and do the ldump and pdump as well
		n = 1
		While n <= UBound(UnitList) And Not (LandFound And PlaneFound)
			LandFound = False
			PlaneFound = False
			oldx = -1
			oldy = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object UnitList(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			nID = CShort(UnitList(n))
			rsShip.Seek("=", nID)
			If Not rsShip.NoMatch Then
				sx = rsShip.Fields("x").Value
				sy = rsShip.Fields("y").Value
				If (oldx <> sx) Or (oldy <> sy) Then
					'check for land units in the sector
					If Not LandFound Then
						rsLand.Seek("=", sx, sy)
						LandFound = Not rsLand.NoMatch
					End If
					'check for planes in the sector
					If Not PlaneFound Then
						rsPlane.Seek("=", sx, sy)
						PlaneFound = Not rsPlane.NoMatch
					End If
					oldx = sx
					oldy = sy
				End If
			End If
			n = n + 1
		End While
		
		If PlaneFound Then
			GetPlanes()
		End If
		
		If LandFound Then
			GetLandUnits()
		End If
		
		'restore indexes
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsShip.Index = hldIndex1
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsPlane.Index = hldIndex2
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsLand.Index = hldIndex3
	End Sub
	
	Private Sub ComputePathCost()
		On Error Resume Next
		Dim pvar As Object
		Dim pcost As Single
		Dim mcost As Single
		Dim mused As Single
		Dim mleft As Single
		Dim minmob As Single
		Dim oldx As Short
		Dim oldy As Short
		Dim sx As Short
		Dim sy As Short
		Dim n As Short
		' Dim m As Integer    8/12/03 efj  removed
		Dim nx As Short
		Dim ny As Short
		' Dim pb As Integer    8/12/03 efj  removed
		Dim hldIndex As Object
		'UPGRADE_WARNING: Arrays in structure rs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim rs As DAO.Recordset
		Dim Fleet As String
		Dim bpath As String
		Dim strRetreat As String
		Dim MobType As Short
		
		lblLeft.Text = vbNullString
		lblCost.Text = vbNullString
		lblBestPath.Text = vbNullString
		lblPathCost.Text = vbNullString
		'093003 rjk: added RouteInfo for indicating if multiple origins are present
		lblRouteInfo.Text = vbNullString
		
		If LCase(Label2.Text) <> "march" And LCase(Label2.Text) <> "navigate" Then
			Exit Sub
		End If
		
		If frmDrawMap.DisplayingPath Then
			frmDrawMap.DisplayingPath = False
			frmDrawMap.FinishPath()
		End If
		
		If Not IsArray(UnitList) Then
			Exit Sub
		End If
		
		If Len(txtPath.Text) = 0 Then
			Exit Sub
		End If
		
		'set indexs
		If LCase(Label2.Text) = "march" Then
			rs = rsLand
			Fleet = "army"
			MobType = EmpCommon.enumMobType.MT_LAND
		Else
			rs = rsShip
			Fleet = "flt"
			MobType = EmpCommon.enumMobType.MT_SHIP
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hldIndex = rs.Index
		rs.Index = "PrimaryKey"
		minmob = 9999
		oldx = -1
		oldy = 0
		
		'093003 rjk: if multiple origins are present can not selected a target sector.
		cmdOK.Enabled = True
		
		For n = LBound(UnitList) + 1 To UBound(UnitList)
			'UPGRADE_WARNING: Couldn't resolve default property of object UnitList(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rs.Seek("=", CShort(UnitList(n)))
			If Not rs.NoMatch Then
				If MobType = EmpCommon.enumMobType.MT_LAND Then
					If frmDrawMap.UnitCharacteristicCheck(frmDrawMap.enumUnitType.TYPE_L_TRAIN, rs.Fields("type").Value) Then
						MobType = EmpCommon.enumMobType.MT_RAIL
					End If
				ElseIf MobType = EmpCommon.enumMobType.MT_SHIP Then 
					If frmDrawMap.UnitCharacteristicCheck(frmDrawMap.enumUnitType.TYPE_S_CANAL, rs.Fields("type").Value) Then
						MobType = EmpCommon.enumMobType.MT_SMALL_SHIP
					End If
				End If
				sx = rs.Fields("x").Value
				sy = rs.Fields("y").Value
				'093003 rjk: added RouteInfo for indicating if multiple origins are present
				If (oldx = -1) And (oldy = 0) Then 'Not first location.
					'093003 rjk: Used first unit for costs and paths
					'Compute the path cost
					oldx = sx
					oldy = sy
					If ParseSectors(nx, ny, (txtPath.Text)) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object BestPath(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object pvar. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						pvar = BestPath(SectorString(sx, sy), (txtPath.Text), MobType)
						'UPGRADE_WARNING: Couldn't resolve default property of object pvar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						bpath = pvar(1)
						'UPGRADE_WARNING: Couldn't resolve default property of object pvar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						pcost = pvar(2)
					Else
						pcost = PathCost(SectorString(sx, sy), (txtPath.Text), MobType)
						bpath = txtPath.Text
					End If
				ElseIf (sx <> oldx) Or (sy <> oldy) Then 
					If ParseSectors(nx, ny, (txtPath.Text)) Then
						lblRouteInfo.ForeColor = System.Drawing.Color.Red
						lblRouteInfo.Text = "Error: Multiple Origins"
						'093003 rjk: if multiple origins are present can not selected a target sector.
						cmdOK.Enabled = False
					Else
						lblRouteInfo.ForeColor = System.Drawing.Color.Black
						lblRouteInfo.Text = "Warning: Multiple Origins"
					End If
					'Else -same location do nothing
				End If
				
				'Fill in retreat path
				txtRetreatPath.Text = vbNullString
				strRetreat = ReversePath(bpath)
				If Len(strRetreat) > 5 Then
					strRetreat = VB.Left(strRetreat, 5)
				End If
				txtRetreatPath.Text = strRetreat
				
				If Fleet = "flt" Then
					mcost = ShipMoveCost(rs.Fields("spd").Value, rs.Fields("eff").Value, rs.Fields("tech").Value, rs.Fields("type").Value)
					mused = mcost * pcost
				Else
					mcost = ShipMoveCost(rs.Fields("spd").Value, 100, rs.Fields("tech").Value, rs.Fields("type").Value)
					
					'only supply units and train suffer mob penalties for low eff
					'092403 rjk: Switched to UnitCharacteristicCheck to remove hard coding
					'If rs.Fields("eff") < 100 And (rs.Fields("type") = "sup" Or rs.Fields("type") = "tra") Then
					If rs.Fields("eff").Value < 100 And frmDrawMap.UnitCharacteristicCheck(frmDrawMap.enumUnitType.TYPE_L_SUPPLY, rs.Fields("type").Value) Then
						mcost = (mcost * 100) / rs.Fields("eff").Value
					End If
					mused = mcost * pcost * 5
				End If
				
				mleft = rs.Fields("mob").Value - mused
				If mleft < minmob Then
					minmob = mleft
					lblCost.Text = "mob. cost: " & CStr(CSng(CInt(CSng(mused) * 1000) / 1000))
					lblLeft.Text = "mob. left: " & CStr(Int(mleft))
					lblBestPath.Text = "path: " & bpath
					lblPathCost.Text = "path cost: " & CStr(CSng(CInt(CSng(pcost) * 1000) / 1000))
					If chkDisplayPath.CheckState = System.Windows.Forms.CheckState.Checked Then
						frmDrawMap.DisplayPath(SectorString(sx, sy), bpath)
					End If
				End If
			End If
		Next n
		'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rs.Index = hldIndex
	End Sub
	
	'UPGRADE_WARNING: Event txtUnit.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtUnit_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnit.TextChanged
		If LCase(Label2.Text) = "march" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object GetUnitList(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object UnitList. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			UnitList = GetUnitList((txtUnit.Text), "l")
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object GetUnitList(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object UnitList. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			UnitList = GetUnitList((txtUnit.Text), "s")
		End If
		
		ComputePathCost()
	End Sub
	
	Private Sub txtUnit_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnit.Enter
		frmDrawMap.DisplayUnitBox() 'rjk 5/13/03: replace DisplaySectorPanel 3 with DisplayUnit
		'             changes to accomodate how the tabs are controlled in
		'             SectorPanel
	End Sub
End Class