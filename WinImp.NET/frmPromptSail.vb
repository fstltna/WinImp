Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmPromptSail
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'091803 rjk: Added Unit grid selection when activating this form or selecting fields
	'            and when disactivating return to standard Sector display.
	'            General reformatting. Added setting the initial field on start up.
	'092503 rjk: Added timestamp to ndump.
	'100503 rjk: Removed timestamp ndump.
	'101403 rjk: Adding timestamp ndump, add lost * to pick nukes being attached to missiles.
	'101603 rjk: Update the map with info satellite report
	'270304 rjk: Switched over to DeleteAllRecords for clearing a table
	'180206 rjk: Replace ldump, pdump, sdump, ndump, and lost with GetLandUnits, GetPlanes,
	'            GetShips, GetNukes and GetLost.
	'010108 rjk: Switch to new GetOrders function.
	
	Public strCmd As String
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		frmDrawMap.DisplayFirstSectorPanel()
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		strCmd = strCmd & " " & txtUnit.Text & " " & txtPath.Text
		If Combo1.Visible Then
			strCmd = strCmd & " " & VB.Left(Combo1.Text, 1)
		End If
		frmEmpCmd.SubmitEmpireCommand(strCmd, True)
		'102503 rjk: Added database updates
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		Select Case VB.Left(strCmd, 3)
			Case "arm"
				'101403 rjk: To pick up arm activity which remove nukes and attaches them to missiles.
				GetPlanes()
				GetLost()
				GetNukes()
				'Case "sat"
			Case "sai"
				GetOrders(True)
			Case "ret", "nam", "mqu"
				GetShips()
			Case "ran", "har"
				GetPlanes()
			Case "lra", "for", "mor", "lre"
				GetLandUnits()
		End Select
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
		frmDrawMap.DisplayFirstSectorPanel()
		Me.Close()
	End Sub
	
	Private Sub frmPromptSail_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		' Set parent for the toolbar to display on top of:
		' Dim success As Long    8/2003 efj  removed
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
	End Sub
	
	Private Sub frmPromptSail_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	Private Sub txtPath_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPath.DoubleClick
		On Error GoTo txtPath_DblClick_Exit
		
		'if this is a sail command - path is not valid
		If (Not (Label2.Text = "Sail" Or Label2.Text = "Retreat")) Or Len(txtUnit.Text) = 0 Then
			Exit Sub
		End If
		
		'Make sure starting unit strind is valid before prompting
		Dim units As Object
		Dim strOrigin As String
		Dim hldIndex As Object
		
		strOrigin = vbNullString
		'UPGRADE_WARNING: Couldn't resolve default property of object ParseUnitString(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object units. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		units = ParseUnitString(txtUnit.Text)
		If Not IsArray(units) Then
			Exit Sub
		End If
		
		'    CurrUnit = GetUnitList(txtUnit.Text, "p")
		'    hldIndex = rsPlane.Index
		'    rsPlane.Index = "PrimaryKey"
		'    n = LBound(CurrUnit) + 1
		
		If VB.Left(strCmd, 1) = "l" Then
			'find the first land in unit string - if not found, exit
			'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			hldIndex = rsLand.Index
			rsLand.Index = "PrimaryKey"
			rsLand.Seek("=", units(LBound(units) + 1))
			If Not rsLand.NoMatch Then
				strOrigin = SectorString(rsLand.Fields("x").Value, rsLand.Fields("y").Value)
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rsLand.Index = hldIndex
			
		Else
			'find the first ship in unit string - if not found, exit
			'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			hldIndex = rsShip.Index
			rsShip.Index = "PrimaryKey"
			rsShip.Seek("=", units(LBound(units) + 1))
			If Not rsShip.NoMatch Then
				strOrigin = SectorString(rsShip.Fields("x").Value, rsShip.Fields("y").Value)
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rsShip.Index = hldIndex
			
		End If
		
		If Not Len(strOrigin) > 0 Then
			Exit Sub
		End If
		
		'setup path prompt
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
		Load(frmPromptPath)
		frmPromptPath.lblSector.Text = strOrigin
		frmPromptPath.lblEndSector.Text = strOrigin
		frmPromptPath.Text = "Exploration Route"
		frmPromptPath.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(frmDrawMap.Left) + VB6.PixelsToTwipsX(frmDrawMap.Width) - VB6.PixelsToTwipsX(frmPromptPath.Width))
		frmPromptPath.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(frmDrawMap.Top) + (VB6.PixelsToTwipsY(frmDrawMap.Height) - VB6.PixelsToTwipsY(frmPromptPath.Height)) \ 2)
		frmPromptPath.Show()
		
txtPath_DblClick_Exit: 
	End Sub
	
	Private Sub txtUnit_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnit.Enter
		Select Case strCmd
			Case "sail", "mquota", "name", "retreat"
				frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udSHIP, frmDrawMap.enumUnitType.TYPE_ALL, False, False, False)
			Case "fortify", "morale", "lrange", "lretreat"
				frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udLAND, frmDrawMap.enumUnitType.TYPE_ALL, False, False, False)
			Case "range"
				frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udPLANE, frmDrawMap.enumUnitType.TYPE_ALL, False, False, False)
			Case "satellite"
				frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udPLANE, frmDrawMap.enumUnitType.TYPE_P_SATELLITE, True, True, False)
			Case "arm", "harden"
				frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udPLANE, frmDrawMap.enumUnitType.TYPE_P_MISSILE, False, False, False)
		End Select
	End Sub
End Class