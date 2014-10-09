Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmPromptBomb
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'08xx03 efj: removed dead variables
	'051303 rjk: replaced DisplaySectorPanel 1 with DisplayFirstSectorPanel
	'            changes to accomodate how the tabs are controlled in SectorPanel
	'092803 rjk: switched to SetUnitType and DisplayFirstSectorPanel.
	'            general reformatting.
	'101703 rjk: Added Strength fields to Sector database
	'112003 rjk: Added option to control strength updates
	'060604 rjk: Added Display Path for Bombing
	'130604 rjk: Added the ability to display path from sector to sector
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	'080805 rjk: Fixed the Commodities box to start with military.
	'180206 rjk: Replace ldump, sdump, lost with GetLandUnits, GetShips and GetLost.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	'170506 rjk: Added infra for SP: Atlantis
	
	Public ShowRange As Boolean
	
	'UPGRADE_WARNING: Event chkDisplayPath.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkDisplayPath_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkDisplayPath.CheckStateChanged
		SetMobilityCost()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	'UPGRADE_WARNING: Event Combo1.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event Combo1.Change was upgraded to Combo1.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub Combo1_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo1.TextChanged
		Check1.CheckState = System.Windows.Forms.CheckState.Checked
	End Sub
	
	'UPGRADE_WARNING: Event Combo1.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Combo1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo1.SelectedIndexChanged
		Check1.CheckState = System.Windows.Forms.CheckState.Checked
	End Sub
	
	'UPGRADE_WARNING: Event Combo2.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event Combo2.Change was upgraded to Combo2.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub Combo2_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo2.TextChanged
		Select Case Combo2.SelectedIndex
			Case 0
				Combo3.Visible = True
				txtTarget.Visible = False
			Case 1, 2, 3
				txtTarget.Visible = True
				Combo3.Visible = False
			Case Else
				txtTarget.Visible = False
				Combo3.Visible = False
		End Select
		
	End Sub
	
	'UPGRADE_WARNING: Event Combo2.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Combo2_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo2.SelectedIndexChanged
		Select Case Combo2.SelectedIndex
			Case 0
				Combo3.Visible = True
				txtTarget.Visible = False
			Case 1, 2, 3
				txtTarget.Visible = True
				Combo3.Visible = False
			Case Else
				txtTarget.Visible = False
				Combo3.Visible = False
		End Select
		
	End Sub
	
	'UPGRADE_WARNING: Event Option1.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option1_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option1.CheckedChanged
		If eventSender.Checked Then
			Combo2.Visible = False
			Combo3.Visible = False
			txtTarget.Visible = False
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Option2.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option2_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option2.CheckedChanged
		If eventSender.Checked Then
			Combo2.Visible = True
			Combo2.SelectedIndex = 0
			Combo3.Visible = True
		End If
	End Sub
	
	Private Sub txtEscort_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEscort.Enter
		If Len(txtPath.Text) > 0 And ShowRange Then
			frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udPLANE, frmDrawMap.enumUnitType.TYPE_PG_ESCORTS, True, True, True, txtPath.Text)
		Else
			frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udPLANE, frmDrawMap.enumUnitType.TYPE_PG_ESCORTS, True, True, False)
		End If
	End Sub
	
	Private Sub txtOrigin_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOrigin.Enter
		frmDrawMap.DisplayFirstSectorPanel() 'rjk 5/13/03: replaced DisplaySectorPanel 1 with DisplayFirstSectorPanel
		'             changes to accomodate how the tabs are controlled in
		'             SectorPanel
	End Sub
	
	'UPGRADE_WARNING: Event txtPath.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtPath_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPath.TextChanged
		'If ShowRange Then
		'    If LastLoad = 1 Then
		'        LastLoad = 0
		'        txtUnit_GotFocus
		'    ElseIf LastLoad = 2 Then
		'        LastLoad = 0
		'        txtEscort_GotFocus
		'    End If
		'End If
		
		If Len(txtUnit.Text) = 0 Or Len(txtPath.Text) = 0 Then
			cmdOK.Enabled = False
		Else
			cmdOK.Enabled = True
		End If
		SetMobilityCost()
	End Sub
	
	Private Sub txtPath_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPath.DoubleClick
		
		'Make sure starting sector is valid before prompting
		Dim sx As Short
		Dim sy As Short
		If Not ParseSectors(sx, sy, (txtOrigin.Text)) Then
			Exit Sub
		End If
		
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
		Load(frmPromptPath)
		frmPromptPath.lblSector.Text = txtOrigin.Text
		frmPromptPath.lblEndSector.Text = txtOrigin.Text
		frmPromptPath.Text = "Select Bomb Route"
		frmPromptPath.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(frmDrawMap.Left) + VB6.PixelsToTwipsX(frmDrawMap.Width) - VB6.PixelsToTwipsX(frmPromptPath.Width))
		frmPromptPath.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(frmDrawMap.Top) + (VB6.PixelsToTwipsY(frmDrawMap.Height) - VB6.PixelsToTwipsY(frmPromptPath.Height)) \ 2)
		frmPromptPath.Show()
		
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		frmDrawMap.DisplayFirstSectorPanel()
		Me.Close()
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim strCmd As String
		Dim strCmd2 As String
		Dim strItem As String
		Dim strAP As String
		Dim num As Short
		Dim X As Short
		Dim n As Short
		Dim hldIndex As String
		Dim CurrUnit As Object
		
		'Must have a destination, or will bomb yourself
		If Len(txtPath.Text) = 0 Then
			MsgBox("Must have a destination/path, or you will bomb yourself !")
			Exit Sub
		End If
		
		strItem = LCase(Combo3.Text)
		If VB.Left(strItem, 1) = "u" Then
			strItem = "un"
		End If
		
		strCmd = LCase(Trim(Label2.Text)) & " "
		
		strCmd = strCmd & txtUnit.Text & " "
		
		If Len(txtEscort.Text) > 0 Then
			strCmd = strCmd & txtEscort.Text
		Else
			strCmd = strCmd & "."
		End If
		
		If Option1.Checked Then
			strCmd = strCmd & " " & Option1.Text
		Else
			strCmd = strCmd & " " & Option2.Text
		End If
		
		'Verify assembly point
		Dim sx As Short
		Dim sy As Short
		
		On Error GoTo AP_Error
		strAP = txtOrigin.Text
		Dim ap As Object
		If Not ParseSectors(sx, sy, (txtOrigin.Text)) Then
			'find the starting position of the first plane in the bombing run
			'UPGRADE_WARNING: Couldn't resolve default property of object ComputeAirports(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ap. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ap = ComputeAirports((txtUnit.Text))
			If IsArray(ap) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object ap(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If ParseSectors(sx, sy, CStr(ap(1))) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object ap(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strAP = CStr(ap(1))
				Else
					GoTo AP_Error
				End If
			Else
AP_Error: 
				MsgBox("Cannot compute assembly point.  Manually enter and resubmit")
				Exit Sub
			End If
		End If
		On Error GoTo 0
		
		'Build bomb command
		strCmd = strCmd & " " & strAP & " " & txtPath.Text
		
		'Execute fire command
		If Check1.CheckState = System.Windows.Forms.CheckState.Checked Then
			num = Val(Combo1.Text)
		Else
			num = 1
		End If
		
		If Option2.Checked Then 'pinbombing
			strCmd2 = VB.Left(Combo2.Text, 3)
			If Combo2.SelectedIndex = 0 Then 'commodities
				strCmd2 = "com" & VB.Left(Combo3.Text, 3)
			ElseIf Combo2.SelectedIndex < 4 Then 
				strCmd2 = strCmd2 & txtTarget.Text
			End If
		End If
		
		frmEmpCmd.SubmitEmpireCommand("ba1 " & strCmd2, False)
		For X = 1 To num
			frmEmpCmd.SubmitEmpireCommand(strCmd, True)
		Next 
		frmEmpCmd.SubmitEmpireCommand("ba2", False)
		
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		
		'Dump the airports from which the planes are flying
		GetSectors()
		'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
		GetCurrentStrength(tsSectors)
		
		'Check to see if any planes were on ships.  If so, dump them to
		'UPGRADE_WARNING: Couldn't resolve default property of object GetUnitList(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CurrUnit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CurrUnit = GetUnitList((txtUnit.Text), "p")
		If IsArray(CurrUnit) Then
			hldIndex = rsPlane.Index
			rsPlane.Index = "PrimaryKey"
			n = LBound(CurrUnit) + 1
			Do While n <= UBound(CurrUnit)
				'UPGRADE_WARNING: Couldn't resolve default property of object CurrUnit(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				rsPlane.Seek("=", CShort(CurrUnit(n)))
				If Not rsPlane.NoMatch Then
					If rsPlane.Fields("ship").Value > -1 Then
						frmEmpCmd.CarrierDumpRequested = True
						GetShips()
						Exit Do
					End If
				End If
				n = n + 1
			Loop 
			rsPlane.Index = hldIndex
			
			'Check to see if any of the planes were missioned
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
		
		'Check if escorts moved have missions.  If so, remove them
		If Len(txtEscort.Text) > 0 Then
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
		
		'now dump the planes themselves
		GetPlanes()
		GetLost()
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
		If chkDisplayPath.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			SaveSetting(APPNAME, "frmPromptBomb", "Display Path", CStr(True))
		Else
			SaveSetting(APPNAME, "frmPromptBomb", "Display Path", CStr(False))
		End If
		
		frmDrawMap.DisplayFirstSectorPanel()
		Me.Close()
	End Sub
	
	
	Private Sub frmPromptBomb_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		' Set parent for the toolbar to display on top of:
		' Dim success As Long    8/12/03 efj  removed
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
		
		txtTarget.SetBounds(Combo3.Left, Combo3.Top, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		LoadCommodityBox(Combo3)
		Combo3.SelectedIndex = 1 'set to military
		Combo2.SelectedIndex = 0
		
		If CBool(GetSetting(APPNAME, "frmPromptBomb", "Display Path", CStr(False))) Then
			chkDisplayPath.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			chkDisplayPath.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		
		If frmOptions.bSPAtlantis Then
			Combo2.Items.Add("infra")
		End If
	End Sub
	
	Private Sub frmPromptBomb_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
		If frmDrawMap.DisplayingPath Then
			frmDrawMap.DisplayingPath = False
			frmDrawMap.FinishPath()
		End If
	End Sub
	
	
	Private Sub txtPath_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPath.Enter
		If Not ShowRange Then
			frmDrawMap.DisplayFirstSectorPanel() 'rjk 5/13/03: replaced DisplaySectorPanel 1 with DisplayFirstSectorPanel
			'             changes to accomodate how the tabs are controlled in
			'             SectorPanel
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtUnit.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtUnit_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnit.TextChanged
		If Len(txtUnit.Text) = 0 Or Len(txtPath.Text) = 0 Then
			cmdOK.Enabled = False
		Else
			cmdOK.Enabled = True
		End If
		SetMobilityCost()
	End Sub
	
	Private Sub txtUnit_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnit.Enter
		If Len(txtPath.Text) > 0 And ShowRange Then
			frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udPLANE, frmDrawMap.enumUnitType.TYPE_PC_BOMB, True, True, True, txtPath.Text)
		Else
			frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udPLANE, frmDrawMap.enumUnitType.TYPE_PC_BOMB, True, True, False)
		End If
	End Sub
	
	Private Sub SetMobilityCost()
		
		Dim sx As Short
		Dim sy As Short
		Dim sx2 As Short
		Dim sy2 As Short
		Dim mstr As String
		Dim rstr As String
		Dim dstr As String
		Dim mcost As Single
		Dim hindex As Object
		Dim distance As Short
		Dim n As Short
		Dim bpath As String
		Dim pvar As Object
		
		On Error GoTo SetMobilityCostEnd
		
		mstr = "Mobility Cost: "
		dstr = "Round Trip Distance: "
		rstr = "Plane Range: "
		'UPGRADE_WARNING: Couldn't resolve default property of object hindex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hindex = rsPlane.Index
		rsPlane.Index = "PrimaryKey"
		
		If frmDrawMap.DisplayingPath Then
			frmDrawMap.DisplayingPath = False
			frmDrawMap.FinishPath()
		End If
		
		'must have an assembly point
		If Not ParseSectors(sx2, sy2, (txtOrigin.Text)) Then
			GoTo SetMobilityCostEnd
		End If
		
		'must have a bomber and targets
		If Len(txtUnit.Text) = 0 Then
			GoTo SetMobilityCostEnd
		End If
		
		'compute distance
		If Len(txtPath.Text) > 0 Then
			If ParseSectors(sx, sy, (txtPath.Text)) Then
				distance = SectorDistance(sx, sy, sx2, sy2)
				If chkDisplayPath.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
					'UPGRADE_WARNING: Couldn't resolve default property of object BestPath(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object pvar. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					pvar = BestPath(SectorString(sx2, sy2), (txtPath.Text), EmpCommon.enumMobType.MT_PLANE)
					'UPGRADE_WARNING: Couldn't resolve default property of object pvar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					bpath = pvar(1)
				End If
			Else
				If chkDisplayPath.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
					bpath = txtPath.Text
				End If
				distance = Len(txtPath.Text) - 1
			End If
		Else
			distance = 0
		End If
		dstr = dstr & CStr(distance * 2)
		
		'get unit list and compute mobility
		Dim us As Object
		Dim c As Short
		c = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object ParseUnitString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object us. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		us = ParseUnitString(txtUnit.Text)
		If IsArray(us) Then
			For n = LBound(us) To UBound(us)
				'UPGRADE_WARNING: Couldn't resolve default property of object us(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				rsPlane.Seek("=", CStr(us(n)))
				If Not rsPlane.NoMatch Then
					c = c + 1
					If c < 5 Then
						mcost = PlaneFlyCost(rsPlane.Fields("range").Value, rsPlane.Fields("eff").Value, 2 * distance)
						mstr = mstr & CStr(CShort(mcost)) & "/"
						rstr = rstr & CStr(CShort(rsPlane.Fields("range").Value)) & "/"
					ElseIf c = 5 Then 
						mstr = mstr & "... "
						rstr = rstr & "... "
					End If
				End If
			Next n
		End If
		
		If chkDisplayPath.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			frmDrawMap.DisplayPath(SectorString(sx2, sy2), bpath)
		End If
		
		
SetMobilityCostEnd: 
		Label8.Text = VB.Left(mstr, Len(mstr) - 1)
		Label9.Text = VB.Left(rstr, Len(rstr) - 1)
		Label10.Text = dstr
		'UPGRADE_WARNING: Couldn't resolve default property of object hindex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsPlane.Index = hindex
	End Sub
End Class