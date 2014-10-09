Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmPromptRecon
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'051303 rjk: replaced DisplaySectorPanel 1 with DisplayFirstSectorPanel
	'            changes to accomodate how the tabs are controlled in
	'            SectorPanel
	'08xx03 efj: remvoed dead variables
	'091803 rjk: Added Unit grid selection when activating this form or selecting fields
	'            and when disactivating return to standard Sector display.
	'            General reformatting. Added setting the initial field on start up.
	'100203 rjk: Repeat is not appropriate for fly so it is removed from the form.
	'101403 rjk: Added MT_DROP to commands requesting bmap to catch mining from a plane.
	'101703 rjk: Added Strength fields to Sector database
	'111803 rjk: Added DisplayPath for display path, does not do sector to sector yet.
	'112003 rjk: Added option to control strength updates
	'            Added red highlighting when one or more planes exceed their range
	'130604 rjk: Added the ability to display path from sector to sector
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	'180206 rjk: Replace sdump, pdump and lost with GetShips, GetLandUnits and GetLost.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	'210506 rjk: Fixed combobox commodities for SP: Atlantis
	
	Dim MissionType As Short
	Dim PlaneType As frmDrawMap.enumUnitType
	Public ShowRange As Boolean
	
	Const MT_FLY As Short = 1
	Const MT_RECON As Short = 2
	Const MT_DROP As Short = 3
	Const MT_PARADROP As Short = 4
	
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
	
	'UPGRADE_ISSUE: Label event Label2.Change was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub Label2_Change()
		Select Case LCase(Trim(Label2.Text))
			Case "recon"
				MissionType = MT_RECON
				PlaneType = frmDrawMap.enumUnitType.TYPE_P_SPY
				cbCom.Visible = False
				Label7.Visible = False
			Case "sweep"
				MissionType = MT_RECON
				PlaneType = frmDrawMap.enumUnitType.TYPE_P_SWEEP
				cbCom.Visible = False
				Label7.Visible = False
			Case "drop"
				MissionType = MT_DROP
				PlaneType = frmDrawMap.enumUnitType.TYPE_PC_DROP
				cbCom.Visible = True
				Label7.Visible = True
				LoadCommodityBox(cbCom)
				cbCom.Items.Add("none")
				If frmOptions.bSPAtlantis Then
					cbCom.SelectedIndex = 1
				Else
					cbCom.SelectedIndex = 8 'set to military
				End If
			Case "paradrop"
				MissionType = MT_PARADROP
				PlaneType = frmDrawMap.enumUnitType.TYPE_P_PARA
				cbCom.Visible = False
				Label7.Visible = False
			Case "fly"
				MissionType = MT_FLY
				PlaneType = frmDrawMap.enumUnitType.TYPE_ALL
				cbCom.Visible = True
				Label7.Visible = True
				LoadCommodityBox(cbCom)
				cbCom.Items.Add("none")
				'UPGRADE_ISSUE: ComboBox property cbCom.NewIndex was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="F649E068-7137-45E5-AC20-4D80A3CC70AC"'
				cbCom.SelectedIndex = cbCom.NewIndex 'set to nothing
				Check1.Visible = False '100203 rjk: Repeat is not appropriate for fly
				Combo1.Visible = False '100203 rjk: Repeat is not appropriate for fly
				Label1.Visible = False '100203 rjk: Repeat is not appropriate for fly
		End Select
	End Sub
	
	Private Sub txtEscort_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEscort.Enter
		If Len(txtPath.Text) > 0 And ShowRange Then
			frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udPLANE, frmDrawMap.enumUnitType.TYPE_PG_ESCORTS, True, True, True, txtPath.Text)
		Else
			frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udPLANE, frmDrawMap.enumUnitType.TYPE_PG_ESCORTS, True, True, False)
		End If
	End Sub
	
	Private Sub txtOrigin_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOrigin.Enter
		frmDrawMap.DisplayFirstSectorPanel() 'rjk 5/13/03: replace DisplaySectorPanel 1 with DisplayFirstSectorPanel
		'             changes to accomodate how the tabs are controlled in
		'             SectorPanel
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
		frmPromptPath.Text = "Select " & Label2.Text & " Route"
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
		' Dim strCmd2 As String    8/12/03 efj  removed
		Dim num As Short
		Dim strAP As String
		Dim X As Short
		Dim n As Short
		Dim hldIndex As String
		Dim CurrUnit As Object
		
		strCmd = LCase(Trim(Label2.Text)) & " "
		
		strCmd = strCmd & txtUnit.Text & " "
		
		If Len(txtEscort.Text) > 0 Then
			strCmd = strCmd & txtEscort.Text
		Else
			strCmd = strCmd & "."
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
		
		strCmd = strCmd & " " & strAP & " " & txtPath.Text
		
		Dim strItem As String
		If cbCom.Visible Then
			strItem = cbCom.Text
			If VB.Left(strItem, 1) = "u" Then
				strItem = "un"
			End If
			If VB.Left(strItem, 4) = "none" Then
				strItem = "."
			End If
			strCmd = strCmd & " " & strItem
		End If
		
		'Execute command
		If Check1.CheckState = System.Windows.Forms.CheckState.Checked Then
			num = Val(Combo1.Text)
		Else
			num = 1
		End If
		
		If Not cbCom.Visible Then
			If num > 1 Then
				frmEmpCmd.SubmitEmpireCommand("bf1", False)
				For X = 1 To num
					frmEmpCmd.SubmitEmpireCommand(strCmd, True)
				Next 
				frmEmpCmd.SubmitEmpireCommand("bf2", False)
			Else
				frmEmpCmd.SubmitEmpireCommand(strCmd, True)
			End If
		Else
			frmEmpCmd.SubmitEmpireCommand("ba1 " & strItem, False)
			For X = 1 To num
				frmEmpCmd.SubmitEmpireCommand(strCmd, True)
			Next 
			frmEmpCmd.SubmitEmpireCommand("ba2", False)
		End If
		
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		GetSectors()
		'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
		GetCurrentStrength(tsSectors)
		GetPlanes()
		GetLost()
		
		'Check for reconissence
		If MissionType = MT_RECON Or MissionType = MT_DROP Then '101403 rjk: Added MT_DROP to catch mining from a plane.
			frmEmpCmd.SubmitEmpireCommand("bmap *", False)
		End If
		
		'Check to see if any planes were on ships.  If so, dump them too
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
		
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
		frmDrawMap.DisplayFirstSectorPanel()
		Me.Close()
	End Sub
	
	Private Sub frmPromptRecon_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		' Set parent for the toolbar to display on top of:
		' Dim success As Long    8/12/03 efj  removed
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
	End Sub
	
	Private Sub frmPromptRecon_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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
	
	Private Sub txtPath_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPath.Enter
		If Not ShowRange Then
			frmDrawMap.DisplayFirstSectorPanel() 'rjk 5/13/03: replace DisplaySectorPanel 1 with DisplayFirstSectorPanel
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
			frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udPLANE, PlaneType, True, True, True, txtPath.Text)
		Else
			frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udPLANE, PlaneType, True, True, False)
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
		
		On Error GoTo SetMobilityCostEnd
		
		If MissionType = MT_DROP Or MissionType = MT_PARADROP Then
			dstr = "Round Trip Distance: "
		Else
			dstr = "Trip Distance: "
		End If
		
		mstr = "Mobility Cost: "
		rstr = "Plane Range: "
		'UPGRADE_WARNING: Couldn't resolve default property of object hindex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hindex = rsPlane.Index
		rsPlane.Index = "PrimaryKey"
		
		'must have an assembly point
		If Not ParseSectors(sx2, sy2, (txtOrigin.Text)) Then
			GoTo SetMobilityCostEnd
		End If
		
		'must have a bomber and targets
		If Len(txtUnit.Text) = 0 Then
			GoTo SetMobilityCostEnd
		End If
		
		'display path 111803 rjk: Added
		If frmDrawMap.DisplayingPath Then
			frmDrawMap.DisplayingPath = False
			frmDrawMap.FinishPath()
		End If
		Dim bpath As String
		Dim pvar As Object
		If chkDisplayPath.CheckState = System.Windows.Forms.CheckState.Checked And Len(txtPath.Text) > 0 Then
			If Not ParseSectors(sx, sy, (txtPath.Text)) Then
				frmDrawMap.DisplayPath((txtOrigin.Text), (txtPath.Text))
			Else
				
				'UPGRADE_WARNING: Couldn't resolve default property of object BestPath(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object pvar. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				pvar = BestPath(SectorString(sx2, sy2), (txtPath.Text), EmpCommon.enumMobType.MT_PLANE)
				'UPGRADE_WARNING: Couldn't resolve default property of object pvar(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				bpath = pvar(1)
				frmDrawMap.DisplayPath((txtOrigin.Text), bpath)
			End If
		End If
		
		'compute distance
		If Len(txtPath.Text) > 0 Then
			If ParseSectors(sx, sy, (txtPath.Text)) Then
				distance = SectorDistance(sx, sy, sx2, sy2)
			Else
				distance = Len(txtPath.Text) - 1
			End If
		Else
			distance = 0
		End If
		
		If MissionType = MT_DROP Or MissionType = MT_PARADROP Then
			distance = distance * 2
		End If
		dstr = dstr & CStr(distance)
		
		'get unit list and compute mobility
		Dim us As Object
		Dim c As Short
		Dim bExceededRange As Boolean '112003 rjk: Added red highlighting when one or more planes exceed their range
		
		bExceededRange = False
		
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
						mcost = PlaneFlyCost(rsPlane.Fields("range").Value, rsPlane.Fields("eff").Value, distance)
						mstr = mstr & CStr(CShort(mcost)) & "/"
						rstr = rstr & CStr(CShort(rsPlane.Fields("range").Value)) & "/"
					ElseIf c = 5 Then 
						mstr = mstr & "... "
						rstr = rstr & "... "
					End If
					If distance > rsPlane.Fields("range").Value Then '112003 rjk: Added red highlighting when one or more planes exceed their range
						bExceededRange = True
					End If
				End If
			Next n
		End If
		
		
SetMobilityCostEnd: 
		Label8.Text = VB.Left(mstr, Len(mstr) - 1)
		If bExceededRange Then '112003 rjk: Added red highlighting when one or more planes exceed their range
			Label9.ForeColor = System.Drawing.Color.Red
		Else
			Label9.ForeColor = System.Drawing.Color.Black
		End If
		Label9.Text = VB.Left(rstr, Len(rstr) - 1)
		Label10.Text = dstr
		'UPGRADE_WARNING: Couldn't resolve default property of object hindex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsPlane.Index = hindex
	End Sub
End Class