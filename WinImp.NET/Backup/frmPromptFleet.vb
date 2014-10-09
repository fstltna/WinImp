Option Strict Off
Option Explicit On
Friend Class frmPromptFleet
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'120803 efj: removed dead variables
	'280903 rjk: Added check to ensure the Fleet group is selected before
	'            submitting command. Added timestamp to sdump.
	'            Expand the form to work for planes and land units as well.
	'            Added initial field selection. General reformatting.
	'            Added Unit grid selection (SetUnitType, DisplayFirstSectorPanel)
	'271003 rjk: Added cmbSub_Click to pick when the combo box selected from a click, needed to
	'            enabled the Okay button if the user picks an existing fleet.
	'110106 rjk: Fixed FleetAdd to not cancel the OK button whenever the map is touched.
	'180206 rjk: Replace sdump, ldump, pdump with GetShips, GetLandUnits and GetPlanes.
	
	'UPGRADE_WARNING: Event cmbSub.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cmbSub.Change was upgraded to cmbSub.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cmbSub_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbSub.TextChanged
		If cmbSub.Text = "" Then
			cmdOK.Enabled = False
		Else
			cmdOK.Enabled = True
		End If
	End Sub
	
	'UPGRADE_WARNING: Event cmbSub.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbSub_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbSub.SelectedIndexChanged
		If cmbSub.Text = "" Then
			cmdOK.Enabled = False
		Else
			cmdOK.Enabled = True
		End If
	End Sub
	
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim strCmd As String
		
		strCmd = LCase(Label2.Text) & " " & cmbSub.Text & " " & txtUnit.Text
		frmEmpCmd.SubmitEmpireCommand(strCmd, True)
		
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		If Label2.Text = "Fleet" Then '092503 rjk: Added dumps for appropriate fleet
			GetShips()
		ElseIf Label2.Text = "Wing" Then 
			GetPlanes()
		Else
			GetLandUnits()
		End If
		
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
		Me.Close()
	End Sub
	
	'UPGRADE_WARNING: Form event frmPromptFleet.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmPromptFleet_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		If Len(cmbSub.Text) > 0 And Len(txtUnit.Text) > 0 Then
			cmdOK.Enabled = True
		Else
			cmdOK.Enabled = False
		End If
	End Sub
	
	Private Sub frmPromptFleet_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		' Set parent for the toolbar to display on top of:
		' Dim success As Long  removed 8/12/03 efj
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
	End Sub
	
	Private Sub frmPromptFleet_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	Public Sub LoadCombo() ' 8/12/03 efj  removed, 09/11/03 rjk added back, used for FleetAdd from menu
		'UPGRADE_WARNING: Arrays in structure rs may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim rs As DAO.Recordset
		Dim strSubFleet As String
		Dim Fleet As String
		
		'Clear old records
		strSubFleet = vbNullString
		cmbSub.Items.Clear()
		
		'Check recordset
		If Label2.Text = "Fleet" Then
			rs = rsShip
		ElseIf Label2.Text = "Wing" Then 
			rs = rsPlane
		Else
			rs = rsLand
		End If
		
		'Run thru list
		rs.MoveFirst()
		While Not rs.EOF
			'If fleet has not been added
			Fleet = rs.Fields(4).Value
			If InStr(strSubFleet, Fleet) = 0 Then
				strSubFleet = strSubFleet & Fleet
				cmbSub.Items.Add(Fleet)
			End If
			rs.MoveNext()
		End While
	End Sub
	
	Private Sub txtUnit_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnit.Enter
		If Label2.Text = "Fleet" Then
			frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udSHIP, frmDrawMap.enumUnitType.TYPE_ALL, False, False, False)
		ElseIf Label2.Text = "Wing" Then 
			frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udPLANE, frmDrawMap.enumUnitType.TYPE_ALL, False, False, False)
		Else 'Army
			frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udLAND, frmDrawMap.enumUnitType.TYPE_ALL, False, False, False)
		End If
	End Sub
End Class