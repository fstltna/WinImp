Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmPromptLook
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081203 efj: removed dead variables
	'091803 rjk: Added Unit grid selection when activating this form
	'092803 rjk: Made Okay the default button, added timestamp to ndump.
	'            For marker only, only do the first ship and reset the unit field
	'            Switched grid selection to use SetUnitType and DisplayFirstSectorPanel.
	'            Switch OK to Show for marking, as OK does not leave form
	'            General reformatting.
	'100503 rjk: Removed timestamp ndump
	'101403 rjk: Adding timestamp ndump
	'            Added bmap to get X's on the map when doing lmine or mine
	'102503 rjk: Added database update to unsail to get the UnitGrid to update
	'270304 rjk: Switched over to DeleteAllRecords for clearing a table
	'290404 rjk: Added the ability to include Update efficicency in NavMarkers
	'100504 rjk: Added the ability to include Update mobility in NavMarkers
	'060604 rjk: Changed the harbour check for NavMarkers to use the Bmap instead.
	'120405 rjk: Added Update combo box for NavMarkers
	'200405 rjk: Added checkbox for NavMarkers to control whether to limit
	'            mobility to max. mobility or not.
	'180206 rjk: Replace ndump with GetNukes.
	'010108 rjk: Switch to new GetOrders function.
	
	Public strCmd As String
	
	'UPGRADE_WARNING: Event chkIncludeUpdateEff.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkIncludeUpdateEff_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkIncludeUpdateEff.CheckStateChanged
		If chkIncludeUpdateEff.CheckState <> System.Windows.Forms.CheckState.Unchecked Or chkIncludeUpdateMob.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			cmbUpdates.Enabled = True
		Else
			cmbUpdates.Enabled = False
		End If
		If chkIncludeUpdateMob.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			chkLimitMobility.Enabled = True
		Else
			chkLimitMobility.Enabled = False
		End If
		DrawMarkers()
	End Sub
	
	'UPGRADE_WARNING: Event chkIncludeUpdateMob.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkIncludeUpdateMob_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkIncludeUpdateMob.CheckStateChanged
		If chkIncludeUpdateEff.CheckState <> System.Windows.Forms.CheckState.Unchecked Or chkIncludeUpdateMob.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			cmbUpdates.Enabled = True
		Else
			cmbUpdates.Enabled = False
		End If
		If chkIncludeUpdateMob.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			chkLimitMobility.Enabled = True
		Else
			chkLimitMobility.Enabled = False
		End If
		DrawMarkers()
	End Sub
	
	'UPGRADE_WARNING: Event chkLimitMobility.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkLimitMobility_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkLimitMobility.CheckStateChanged
		DrawMarkers()
	End Sub
	
	'UPGRADE_WARNING: Event cmbUpdates.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbUpdates_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbUpdates.SelectedIndexChanged
		DrawMarkers()
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		If strCmd = "marker" Then
			frmDrawMap.NavMarkerShip = vbNullString
			frmDrawMap.DrawHexes()
		End If
		
		frmDrawMap.DisplayFirstSectorPanel()
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		If strCmd = "marker" Then
			DrawMarkers()
		Else
			frmEmpCmd.SubmitEmpireCommand(strCmd & " " & txtUnit.Text, True)
			If strCmd = "look" Or strCmd = "llook" Or strCmd = "sonar" Then
				'just force the map to redraw
				frmEmpCmd.SubmitEmpireCommand("db1", False)
				frmEmpCmd.SubmitEmpireCommand("db2", False)
			ElseIf strCmd = "disarm" Then 
				frmEmpCmd.SubmitEmpireCommand("db1", False)
				GetNukes()
				frmEmpCmd.SubmitEmpireCommand("db2", False)
			ElseIf strCmd = "mine" Or strCmd = "lmine" Then 
				frmEmpCmd.SubmitEmpireCommand("db1", False)
				frmEmpCmd.SubmitEmpireCommand("bmap *", False) '101403 rjk: Added bmap to get X's on the map
				frmEmpCmd.SubmitEmpireCommand("db2", False)
			ElseIf strCmd = "unsail" Then  '102503 rjk: Added unsail to get the UnitGrid to update
				GetOrders(True)
			End If
			
			frmDrawMap.DisplayFirstSectorPanel()
			Me.Close()
		End If
	End Sub
	
	Private Sub frmPromptLook_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		' Set parent for the toolbar to display on top of:
		' Dim success As Long  removed 8/12/03 efj
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
		If strCmd = "marker" Then '092803 rjk: Switch OK to Show for marking, as OK does not leave form
			cmdOK.Text = "Show"
			lblUpdates.Visible = True
			cmbUpdates.Visible = True
			cmbUpdates.SelectedIndex = 0
			If frmDrawMap.bNavMarkerShipIncludeUpdateMob Or frmDrawMap.bNavMarkerShipIncludeUpdateEff Then
				cmbUpdates.Enabled = True
				chkLimitMobility.Enabled = True
			Else
				cmbUpdates.Enabled = False
				chkLimitMobility.Enabled = False
			End If
			
			chkIncludeUpdateMob.Visible = True
			If frmDrawMap.bNavMarkerShipIncludeUpdateMob Then
				chkIncludeUpdateMob.CheckState = System.Windows.Forms.CheckState.Checked
				cmbUpdates.Enabled = True
			Else
				chkIncludeUpdateMob.CheckState = System.Windows.Forms.CheckState.Unchecked
			End If
			chkIncludeUpdateEff.Visible = True
			If frmDrawMap.bNavMarkerShipIncludeUpdateEff Then
				chkIncludeUpdateEff.CheckState = System.Windows.Forms.CheckState.Checked
			Else
				chkIncludeUpdateEff.CheckState = System.Windows.Forms.CheckState.Unchecked
			End If
			chkLimitMobility.Visible = True
			If frmDrawMap.bNavMarkerShipLimitMobility Then
				chkLimitMobility.CheckState = System.Windows.Forms.CheckState.Checked
			Else
				chkLimitMobility.CheckState = System.Windows.Forms.CheckState.Unchecked
			End If
		Else
			cmdOK.Text = "OK"
		End If
	End Sub
	
	Private Sub frmPromptLook_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	'UPGRADE_WARNING: Event txtUnit.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtUnit_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnit.TextChanged
		If strCmd = "marker" Then
			DrawMarkers()
		End If
	End Sub
	
	Private Sub txtUnit_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnit.Enter
		Select Case strCmd
			Case "look", "marker", "radar", "unsail"
				frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udSHIP, frmDrawMap.enumUnitType.TYPE_ALL, False, False, False)
			Case "mine"
				frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udSHIP, frmDrawMap.enumUnitType.TYPE_S_MINE, True, True, False)
			Case "sonar"
				frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udSHIP, frmDrawMap.enumUnitType.TYPE_S_SONAR, True, True, False)
			Case "llook"
				frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udLAND, frmDrawMap.enumUnitType.TYPE_ALL, False, True, False)
			Case "lradar"
				frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udLAND, frmDrawMap.enumUnitType.TYPE_L_RADAR, False, True, False)
			Case "lmine"
				frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udLAND, frmDrawMap.enumUnitType.TYPE_L_ENGINEER, True, True, False)
			Case "supply"
				frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udLAND, frmDrawMap.enumUnitType.TYPE_L_SUPPLY, True, True, False)
			Case "disarm"
				frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udPLANE, frmDrawMap.enumUnitType.TYPE_P_MISSILE, False, False, False)
		End Select
	End Sub
	
	Private Sub DrawMarkers()
		Dim bEff As Boolean
		Dim hldIndex As String
		
		If InStrRev(txtUnit.Text, "/") > 0 Then '290404 rjk: only mark the last unit
			txtUnit.Text = VB.Right(txtUnit.Text, Len(txtUnit.Text) - InStrRev(txtUnit.Text, "/")) 'reset to only the last unit
		End If
		frmDrawMap.NavMarkerShip = txtUnit.Text
		If chkIncludeUpdateMob.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			frmDrawMap.bNavMarkerShipIncludeUpdateMob = True
		Else
			frmDrawMap.bNavMarkerShipIncludeUpdateMob = False
		End If
		
		bEff = False
		hldIndex = rsShip.Index
		rsShip.Index = "PrimaryKey"
		If Len(txtUnit.Text) > 0 Then
			rsShip.Seek("=", Val(txtUnit.Text))
			If Not rsShip.NoMatch Then
				rsBmap.Seek("=", rsShip.Fields("x"), rsShip.Fields("y"))
				If Not rsBmap.NoMatch Then
					If rsBmap.Fields("des").Value = "h" Then
						bEff = True
					End If
				End If
			End If
		End If
		rsShip.Index = hldIndex
		If bEff Then
			chkIncludeUpdateEff.Enabled = True
		Else
			chkIncludeUpdateEff.CheckState = System.Windows.Forms.CheckState.Unchecked
			chkIncludeUpdateEff.Enabled = False
		End If
		If chkIncludeUpdateEff.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			frmDrawMap.bNavMarkerShipIncludeUpdateEff = True
		Else
			frmDrawMap.bNavMarkerShipIncludeUpdateEff = False
		End If
		
		frmDrawMap.iNavMarkerShipUpdates = VB6.GetItemData(cmbUpdates, cmbUpdates.SelectedIndex)
		If chkLimitMobility.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			frmDrawMap.bNavMarkerShipLimitMobility = True
		Else
			frmDrawMap.bNavMarkerShipLimitMobility = False
		End If
		frmDrawMap.DrawHexes()
	End Sub
End Class