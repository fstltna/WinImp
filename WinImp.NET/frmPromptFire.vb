Option Strict Off
Option Explicit On
Friend Class frmPromptFire
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081203 efj: removed dead variables
	'091803 rjk: Added Unit grid selection when activating this form
	'092803 rjk: General reformatting. Switched to SetUnitType and
	'            DisplayFirstSectorPanel
	'101703 rjk: Added Strength fields to Sector database
	'112003 rjk: Added option to control strength updates
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	'180206 rjk: Replace sdump and ldump with GetShips and GetLandUnits.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	
	'UPGRADE_WARNING: Event Check1.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Check1_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Check1.CheckStateChanged
		chkAbort.Enabled = (Check1.CheckState = System.Windows.Forms.CheckState.Checked)
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		frmDrawMap.DisplayFirstSectorPanel()
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim strCmd As String
		Dim num As Short
		Dim X As Short
		
		If Option1.Checked Then
			strCmd = " ship " & txtUnit.Text
		ElseIf Option2.Checked Then 
			strCmd = " land " & txtUnit.Text
		Else
			strCmd = " sect " & txtOrigin.Text
		End If
		
		strCmd = "fire " & strCmd & " " & txtTarget.Text
		
		'Execute fire command
		If Check1.CheckState = System.Windows.Forms.CheckState.Checked Then
			num = Val(Combo1.Text)
		Else
			num = 1
		End If
		
		If num > 1 Then
			'Check abort flag
			If chkAbort.CheckState = System.Windows.Forms.CheckState.Checked Then
				frmEmpCmd.SubmitEmpireCommand("bfa", False)
			Else
				frmEmpCmd.SubmitEmpireCommand("bf1", False)
			End If
			For X = 1 To num
				frmEmpCmd.SubmitEmpireCommand(strCmd, True)
			Next 
			frmEmpCmd.SubmitEmpireCommand("bf2", False)
		Else
			frmEmpCmd.SubmitEmpireCommand(strCmd, True)
		End If
		
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		If Option1.Checked Then
			GetShips()
		ElseIf Option2.Checked Then 
			GetLandUnits()
		Else
			GetSectors()
			'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
			GetCurrentStrength(tsSectors)
		End If
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
		frmDrawMap.DisplayFirstSectorPanel()
		Me.Close()
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
	
	Private Sub frmPromptFire_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Make Stay always on top
		' Dim success As Long  removed 8/12/03 efj
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
		
		'Move unit box over origin box
		txtUnit.SetBounds(txtOrigin.Left, txtOrigin.Top, txtOrigin.Width, txtOrigin.Height)
		
		chkAbort.Enabled = False
	End Sub
	
	Private Sub frmPromptFire_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	'UPGRADE_WARNING: Event Option1.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option1_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option1.CheckedChanged
		If eventSender.Checked Then
			txtUnit.Visible = True
			txtOrigin.Visible = False
			SetFiringRange()
			txtUnit.Focus()
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Option2.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option2_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option2.CheckedChanged
		If eventSender.Checked Then
			txtUnit.Visible = True
			txtOrigin.Visible = False
			SetFiringRange()
			txtUnit.Focus()
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Option3.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option3_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option3.CheckedChanged
		If eventSender.Checked Then
			txtUnit.Visible = False
			txtOrigin.Visible = True
			SetFiringRange()
			txtOrigin.Focus()
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtOrigin.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtOrigin_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOrigin.TextChanged
		SetFiringRange()
		SetTargetRange()
	End Sub
	
	Private Sub txtOrigin_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOrigin.DoubleClick
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
		Load(frmPromptSectors)
		frmPromptSectors.strSectors = txtOrigin.Text
		frmPromptSectors.SetControls()
		frmPromptSectors.Text = "Select Sectors"
		frmPromptSectors.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(frmDrawMap.Left) + VB6.PixelsToTwipsX(frmDrawMap.Width) - VB6.PixelsToTwipsX(frmPromptSectors.Width))
		frmPromptSectors.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(frmDrawMap.Top) + (VB6.PixelsToTwipsY(frmDrawMap.Height) - VB6.PixelsToTwipsY(frmPromptSectors.Height)) \ 2)
		frmPromptSectors.Show()
	End Sub
	
	Private Sub txtOrigin_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOrigin.Enter
		frmDrawMap.DisplayFirstSectorPanel()
	End Sub
	'UPGRADE_WARNING: Event txtTarget.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtTarget_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTarget.TextChanged
		SetTargetRange()
	End Sub
	
	'UPGRADE_WARNING: Event txtUnit.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtUnit_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnit.TextChanged
		SetFiringRange()
		SetTargetRange()
	End Sub
	
	Private Sub txtUnit_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnit.Enter
		If Option1.Checked Then
			frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udSHIP, frmDrawMap.enumUnitType.TYPE_SC_FIRE, False, True, False)
		End If
		If Option2.Checked Then
			frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udLAND, frmDrawMap.enumUnitType.TYPE_LC_FIRE, False, True, False)
		End If
	End Sub
	
	Private Sub SetFiringRange()
		Dim frange As Single
		Dim sx As Short
		Dim sy As Short
		Dim hldIndex As Object
		
		On Error GoTo SetFiringRangeError
		
		If Option1.Checked Then
			'Check for ship
			'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			hldIndex = rsShip.Index
			rsShip.Index = "PrimaryKey"
			rsShip.Seek("=", CShort(txtUnit.Text))
			If Not rsShip.NoMatch Then
				frange = UnitFireRange(rsShip.Fields("tech").Value, rsShip.Fields("rng").Value)
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rsShip.Index = hldIndex
		ElseIf Option2.Checked Then 
			'Check for land
			'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			hldIndex = rsLand.Index
			rsLand.Index = "PrimaryKey"
			rsLand.Seek("=", CShort(txtUnit.Text))
			If Not rsLand.NoMatch Then
				frange = UnitFireRange(rsLand.Fields("tech").Value, rsLand.Fields("frg").Value)
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rsLand.Index = hldIndex
		Else
			If ParseSectors(sx, sy, (txtOrigin.Text)) Then
				'Get Sector Information
				rsSectors.Seek("=", sx, sy)
				If Not rsSectors.NoMatch Then
					frange = FortRange(TechLevel, rsSectors.Fields("eff").Value)
				Else
					frange = FortRange(TechLevel)
				End If
			End If
		End If
		
SetFiringRangeError: 
		Label4.Text = "Firing Range: " & VB6.Format(frange, "#0.00")
	End Sub
	
	Private Sub SetTargetRange()
		' Dim frange As Single    8/12/03 efj  removed
		Dim sx1 As Short
		Dim sy1 As Short
		Dim sx2 As Short
		Dim sy2 As Short
		Dim SetLabel As Boolean
		Dim hldIndex As Object
		
		On Error GoTo SetTargetRangeError
		
		If Len(txtTarget.Text) = 0 Then
			GoTo SetTargetRangeError
		End If
		
		SetLabel = True
		If Option1.Checked Then
			'Check for ship
			'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			hldIndex = rsShip.Index
			rsShip.Index = "PrimaryKey"
			rsShip.Seek("=", CShort(txtUnit.Text))
			If Not rsShip.NoMatch Then
				sx1 = rsShip.Fields("x").Value
				sy1 = rsShip.Fields("y").Value
			Else
				SetLabel = False
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rsShip.Index = hldIndex
		ElseIf Option2.Checked Then 
			'Check for land
			'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			hldIndex = rsLand.Index
			rsLand.Index = "PrimaryKey"
			rsLand.Seek("=", CShort(txtUnit.Text))
			If Not rsLand.NoMatch Then
				sx1 = rsLand.Fields("x").Value
				sy1 = rsLand.Fields("y").Value
			Else
				SetLabel = False
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rsLand.Index = hldIndex
		Else
			If Not ParseSectors(sx1, sy1, (txtOrigin.Text)) Then
				SetLabel = False
			End If
		End If
		
		'get the target sector
		If Not ParseSectors(sx2, sy2, (txtTarget.Text)) Then
			'Check for known enemy unit
			'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			hldIndex = rsEnemyUnit.Index
			rsEnemyUnit.Index = "ByID"
			rsEnemyUnit.Seek("=", "s", CShort(txtTarget.Text))
			If rsEnemyUnit.NoMatch Then
				SetLabel = False
			Else
				sx2 = rsEnemyUnit.Fields("x").Value
				sy2 = rsEnemyUnit.Fields("y").Value
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object hldIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			rsEnemyUnit.Index = hldIndex
		End If
		
		If SetLabel Then
			Label5.Text = "Range to Target: " & CStr(SectorDistance(sx1, sy1, sx2, sy2))
		Else
			Label5.Text = "Range to Target: "
		End If
		Exit Sub
		
SetTargetRangeError: 
		Label5.Text = "Range to Target: "
	End Sub
End Class