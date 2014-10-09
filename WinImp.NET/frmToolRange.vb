Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmToolRange
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'100303 rjk: Protect against too tech level for the Vscroll bar.
	'100304 rjk: Protect vsTech Min/Max against too high tech level.
	'140505 rjk: Fixed ship mobility cost to deal with the extra mobility costs with
	'            ships with SWEEP capabilities.
	'160705 rjk: Fixed LoadComboBox to deal with empty build table.
	'010106 rjk: Added Airport Range.
	'210506 rjk: Fixed Nuke combobox, broken when airport was added.
	
	Dim bscroll As Boolean
	Dim ShipMobTurn As Short
	Dim ShipMobMax As Short
	
	'UPGRADE_WARNING: Event chkRoundTrip.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkRoundTrip_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkRoundTrip.CheckStateChanged
		CalculatePlaneRange((vsTech.Value))
	End Sub
	
	'UPGRADE_WARNING: Event cmbLand.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cmbLand.Change was upgraded to cmbLand.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cmbLand_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbLand.TextChanged
		CalculateLandRanges(CSng(vsTech.Value))
	End Sub
	
	'UPGRADE_WARNING: Event cmbLand.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbLand_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbLand.SelectedIndexChanged
		CalculateLandRanges(CSng(vsTech.Value))
	End Sub
	
	'UPGRADE_WARNING: Event cmbNuke.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cmbNuke.Change was upgraded to cmbNuke.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cmbNuke_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbNuke.TextChanged
		If TabStrip1.SelectedItem.Index = 4 Then
			CalculateNukeDamage()
		Else
			CalculatePlaneRange((vsTech.Value))
		End If
	End Sub
	
	'UPGRADE_WARNING: Event cmbNuke.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbNuke_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbNuke.SelectedIndexChanged
		If TabStrip1.SelectedItem.Index = 4 Then
			CalculateNukeDamage()
		Else
			CalculatePlaneRange((vsTech.Value))
		End If
	End Sub
	
	'UPGRADE_WARNING: Event cmbShip.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cmbShip.Change was upgraded to cmbShip.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cmbShip_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbShip.TextChanged
		CalculateShipRanges(CSng(vsTech.Value))
	End Sub
	
	'UPGRADE_WARNING: Event cmbShip.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbShip_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbShip.SelectedIndexChanged
		CalculateShipRanges(CSng(vsTech.Value))
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Me.Close()
	End Sub
	
	Private Sub frmToolRange_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		bscroll = False
		
		'100303 rjk: Protect against too tech level for the Vscroll bar.
		If TechLevel > vsTech.Minimum Then 'scrollbar is setup in reverse direction.
			vsTech.Value = vsTech.Minimum
		ElseIf TechLevel < (vsTech.Maximum - vsTech.LargeChange + 1) Then 
			vsTech.Value = (vsTech.Maximum - vsTech.LargeChange + 1)
		Else
			vsTech.Value = TechLevel
		End If
		
		TabStrip1_ClickEvent(TabStrip1, New System.EventArgs())
		CalculateSectorRanges(TechLevel)
		LoadComboBox((Me.cmbShip), "s")
		LoadComboBox((Me.cmbLand), "l")
		LoadComboBox((Me.cmbNuke), "n")
		
		If Not (rsVersion.BOF And rsVersion.EOF) Then
			ShipMobMax = rsVersion.Fields("Max mob - Ship").Value
			ShipMobTurn = rsVersion.Fields("Mob gain - Ship").Value
		End If
		
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
	End Sub
	
	
	Private Sub frmToolRange_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		If Len(NukeDamageType) > 0 Or Len(PlaneRangeType) > 0 Then
			NukeDamageType = vbNullString
			PlaneRangeType = vbNullString
			frmDrawMap.DrawHexes()
		End If
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	'UPGRADE_WARNING: Event optAirBurst.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optAirBurst_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optAirBurst.CheckedChanged
		If eventSender.Checked Then
			CalculateNukeDamage()
		End If
	End Sub
	
	'UPGRADE_WARNING: Event optGroundBurst.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optGroundBurst_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optGroundBurst.CheckedChanged
		If eventSender.Checked Then
			CalculateNukeDamage()
		End If
	End Sub
	
	Private Sub TabStrip1_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TabStrip1.ClickEvent
		frmSectors.Visible = False
		frmShip.Visible = False
		frmLand.Visible = False
		frmNuke.Visible = False
		txtTech.Visible = True
		lblTech.Visible = True
		vsTech.Visible = True
		Select Case (TabStrip1.SelectedItem.Index)
			Case 1
				frmSectors.Visible = True
			Case 2
				frmShip.Visible = True
			Case 3
				frmLand.Visible = True
			Case 4
				txtTech.Visible = False
				lblTech.Visible = False
				vsTech.Visible = False
				frmNuke.Visible = True
				optAirBurst.Visible = True
				optGroundBurst.Visible = True
				lblNukeType.Text = "Nuke Type"
				chkRoundTrip.Visible = False
				LoadComboBox(cmbNuke, "n")
			Case 5
				txtTech.Visible = True
				lblTech.Visible = True
				vsTech.Visible = True
				frmNuke.Visible = True
				optAirBurst.Visible = False
				optGroundBurst.Visible = False
				lblNukeType.Text = "Plane Type"
				chkRoundTrip.Visible = True
				LoadComboBox(cmbNuke, "p")
		End Select
		CalculateRanges(CSng(vsTech.Value))
	End Sub
	
	'UPGRADE_WARNING: Event txtOrigin.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtOrigin_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOrigin.TextChanged
		If TabStrip1.SelectedItem.Index = 4 Then
			CalculateNukeDamage()
		Else
			CalculatePlaneRange((vsTech.Value))
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtTech.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtTech_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTech.TextChanged
		If Not bscroll Then
			bscroll = True
			'100304 rjk: Protect vsTech Min/Max
			If Val(txtTech.Text) > vsTech.Minimum Then
				vsTech.Value = vsTech.Minimum
				txtTech.Text = CStr(vsTech.Value)
			ElseIf Val(txtTech.Text) < (vsTech.Maximum - vsTech.LargeChange + 1) Then 
				vsTech.Value = (vsTech.Maximum - vsTech.LargeChange + 1)
				txtTech.Text = CStr(vsTech.Value)
			Else
				vsTech.Value = Val(txtTech.Text)
			End If
			bscroll = False
		End If
		
	End Sub
	
	'UPGRADE_NOTE: vsTech.Change was changed from an event to a procedure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="4E2DC008-5EDA-4547-8317-C9316952674F"'
	'UPGRADE_WARNING: VScrollBar event vsTech.Change has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub vsTech_Change(ByVal newScrollValue As Integer)
		If Not bscroll Then
			bscroll = True
			txtTech.Text = CStr(newScrollValue)
			bscroll = False
		End If
		CalculateRanges(CSng(newScrollValue))
	End Sub
	
	Private Sub CalculateRanges(ByRef Tech As Single)
		If frmSectors.Visible Then
			CalculateSectorRanges(Tech)
		ElseIf frmShip.Visible Then 
			CalculateShipRanges(Tech)
		ElseIf frmLand.Visible Then 
			CalculateLandRanges(Tech)
		End If
		If frmNuke.Visible Then
			If TabStrip1.SelectedItem.Index = 4 Then
				CalculateNukeDamage()
			Else
				CalculatePlaneRange((vsTech.Value))
			End If
		Else
			If Len(NukeDamageType) > 0 Or Len(PlaneRangeType) > 0 Then
				NukeDamageType = vbNullString
				PlaneRangeType = vbNullString
				frmDrawMap.DrawHexes()
			End If
		End If
	End Sub
	
	Private Sub CalculateSectorRanges(ByRef Tech As Single)
		'Sector Values
		lblFort100.Text = VB6.Format(FortRange(Tech), "#0.00")
		lblFort59.Text = VB6.Format(FortRange(Tech, 59), "#0.00")
		lblRadar.Text = CStr(RadarRange(Tech))
		lblCoastwatch.Text = CStr(CoastwatchRange(Tech))
	End Sub
	
	Private Sub CalculateShipRanges(ByRef Tech As Single)
		Dim frnge As Short
		Dim speed As Short
		Dim movecost As Single
		If cmbShip.SelectedIndex < 0 Then Exit Sub
		
		If Tech >= VB6.GetItemData(cmbShip, cmbShip.SelectedIndex) Then
			Label4.Text = "Firing Range"
			frnge = ComputeRangeStat(Tech, "s", Trim(VB.Right(VB6.GetItemString(cmbShip, cmbShip.SelectedIndex), 5)))
			lblShipRange.Text = VB6.Format(UnitFireRange(Tech, frnge), "#0.00")
			speed = ComputeSpeedStat(Tech, "s", Trim(VB.Right(VB6.GetItemString(cmbShip, cmbShip.SelectedIndex), 5)))
			If speed <= 0 Then
				lblShipMoveCost.Text = vbNullString
				lblShipTurn.Text = vbNullString
				lblShipMax.Text = vbNullString
			Else
				movecost = ShipMoveCost(speed, 100, CShort(Tech), Trim(VB.Right(VB6.GetItemString(cmbShip, cmbShip.SelectedIndex), 5)))
				lblShipMoveCost.Text = VB6.Format(movecost, "#0.0")
				lblShipTurn.Text = VB6.Format(ShipMobTurn / movecost, "#0.0")
				lblShipMax.Text = VB6.Format(ShipMobMax / movecost, "#0.0")
			End If
		Else
			Label4.Text = "Build Tech"
			lblShipRange.Text = CStr(VB6.GetItemData(cmbShip, cmbShip.SelectedIndex))
			lblShipMoveCost.Text = vbNullString
			lblShipTurn.Text = vbNullString
			lblShipMax.Text = vbNullString
		End If
	End Sub
	Private Sub CalculateLandRanges(ByRef Tech As Single)
		Dim frnge As Short
		If cmbLand.SelectedIndex < 0 Then Exit Sub
		
		If Tech >= VB6.GetItemData(cmbLand, cmbLand.SelectedIndex) Then
			Label7.Text = "Firing Range"
			frnge = ComputeRangeStat(Tech, "l", Trim(VB.Right(VB6.GetItemString(cmbLand, cmbLand.SelectedIndex), 5)))
			lblLandRange.Text = VB6.Format(UnitFireRange(Tech, frnge), "#0.00")
		Else
			Label7.Text = "Build Tech"
			lblLandRange.Text = CStr(VB6.GetItemData(cmbLand, cmbLand.SelectedIndex))
		End If
	End Sub
	
	Private Sub CalculateNukeDamage()
		Dim sx As Short
		Dim sy As Short
		
		If ParseSectors(sx, sy, (txtOrigin.Text)) And Len(cmbNuke.Text) > 0 Then
			If NukeDamageType <> cmbNuke.Text Or NukeTargetSector <> txtOrigin.Text Or (optAirBurst.Checked <> System.Windows.Forms.CheckState.Unchecked) <> bNukeAirBurst Then
				NukeDamageType = cmbNuke.Text
				NukeTargetSector = txtOrigin.Text
				If optAirBurst.Checked <> System.Windows.Forms.CheckState.Unchecked Then
					bNukeAirBurst = True
				Else
					bNukeAirBurst = False
				End If
				frmDrawMap.DrawHexes()
			End If
		Else
			If Len(NukeDamageType) > 0 Then
				NukeDamageType = vbNullString
				frmDrawMap.DrawHexes()
			End If
		End If
	End Sub
	
	Private Sub CalculatePlaneRange(ByRef Tech As Short)
		Dim sx As Short
		Dim sy As Short
		
		If ParseSectors(sx, sy, (txtOrigin.Text)) And Len(cmbNuke.Text) > 0 Then
			If PlaneRangeType <> cmbNuke.Text Or PlaneTargetSector <> txtOrigin.Text Or PlaneRoundTrip <> (chkRoundTrip.CheckState <> System.Windows.Forms.CheckState.Unchecked) Or PlaneTech <> Tech Then
				PlaneRangeType = cmbNuke.Text
				PlaneTargetSector = txtOrigin.Text
				PlaneRoundTrip = chkRoundTrip.CheckState <> System.Windows.Forms.CheckState.Unchecked
				PlaneTech = Tech
				frmDrawMap.DrawHexes()
			End If
		Else
			If Len(PlaneRangeType) > 0 Then
				PlaneRangeType = vbNullString
				frmDrawMap.DrawHexes()
			End If
		End If
	End Sub
	
	'UPGRADE_NOTE: Class was upgraded to Class_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub LoadComboBox(ByRef cmb As System.Windows.Forms.Control, ByRef Class_Renamed As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object cmb.Clear. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cmb.Clear()
		Dim hldIndex As String
		
		If Class_Renamed = "n" Then
			hldIndex = rsBuild.Index
			rsBuild.Index = "Tech"
			If rsBuild.BOF And rsBuild.EOF Then
				Exit Sub
			End If
			rsBuild.MoveFirst()
		End If
		
		If Not (rsBuild.BOF And rsBuild.EOF) Then
			rsBuild.MoveFirst()
			While Not rsBuild.EOF
				If rsBuild.Fields("type").Value = Class_Renamed Then
					If Class_Renamed = "n" Then
						'UPGRADE_WARNING: Couldn't resolve default property of object cmb.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						cmb.AddItem(rsBuild.Fields("id").Value + " - " + rsBuild.Fields("desc").Value)
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object cmb.AddItem. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						cmb.AddItem(rsBuild.Fields("desc").Value + "                                      " + VB.Left(rsBuild.Fields("id").Value + "     ", 5))
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object cmb.NewIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object cmb.ItemData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					cmb.ItemData(cmb.NewIndex) = rsBuild.Fields("tech").Value
				End If
				rsBuild.MoveNext()
			End While
		End If
		If Class_Renamed = "n" Then
			rsBuild.Index = hldIndex
		End If
	End Sub
	Private Sub vsTech_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ScrollEventArgs) Handles vsTech.Scroll
		Select Case eventArgs.type
			Case System.Windows.Forms.ScrollEventType.EndScroll
				vsTech_Change(eventArgs.newValue)
		End Select
	End Sub
End Class