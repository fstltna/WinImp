Option Strict Off
Option Explicit On
Friend Class frmPromptWarload
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'091803 rjk: Added Unit grid selection when activating this form
	'            and when disactivating Return to standard Sector display.
	'            Also general reformatting
	'101703 rjk: Added Strength fields to Sector database
	'111403 rjk: Added so all the commands are combined on the command display
	'112003 rjk: Added option to control strength updates
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	'011506 rjk: Set the OK button to default.
	'180206 rjk: Replace ldump and sdump with GetLandUnits and GetShips.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	
	Public strCmd As String
	' Public strSector As String    8/2003 efj  removed
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		frmDrawMap.DisplayFirstSectorPanel()
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		'111403 rjk: Added so all the commands are combined on the command display
		frmEmpCmd.SubmitEmpireCommand("bf1", False)
		
		If strCmd = "swarload" Then
			frmEmpCmd.SubmitEmpireCommand("load mil " & txtUnit.Text & " 999", True)
			frmEmpCmd.SubmitEmpireCommand("load gun " & txtUnit.Text & " 999", True)
			frmEmpCmd.SubmitEmpireCommand("load shell " & txtUnit.Text & " 999", True)
			frmEmpCmd.SubmitEmpireCommand("load food " & txtUnit.Text & " 999", True)
			frmEmpCmd.SubmitEmpireCommand("load oil " & txtUnit.Text & " 999", True)
			frmEmpCmd.SubmitEmpireCommand("load pet " & txtUnit.Text & " 999", True)
		End If
		
		If strCmd = "lwarload" Then
			frmEmpCmd.SubmitEmpireCommand("lload mil " & txtUnit.Text & " 999", True)
			frmEmpCmd.SubmitEmpireCommand("lload gun " & txtUnit.Text & " 999", True)
			frmEmpCmd.SubmitEmpireCommand("lload shell " & txtUnit.Text & " 999", True)
			frmEmpCmd.SubmitEmpireCommand("lload food " & txtUnit.Text & " 999", True)
			frmEmpCmd.SubmitEmpireCommand("lload oil " & txtUnit.Text & " 999", True)
		End If
		'111403 rjk: Added so all the commands are combined on the command display
		frmEmpCmd.SubmitEmpireCommand("bf2", False)
		
		'refresh database
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		GetSectors()
		'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
		GetCurrentStrength(tsSectors)
		
		If strCmd = "swarload" Then
			GetShips()
		End If
		
		If strCmd = "lwarload" Then
			GetLandUnits()
		End If
		
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
		frmDrawMap.DisplayFirstSectorPanel()
		Me.Close()
	End Sub
	
	Private Sub frmPromptWarload_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		' Set parent for the toolbar to display on top of:
		' Dim success As Long    8/2003 efj  removed
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
	End Sub
	
	Private Sub frmPromptWarload_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	Private Sub txtUnit_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnit.Enter
		Select Case strCmd
			Case "swarload"
				frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udSHIP, frmDrawMap.enumUnitType.TYPE_ALL, False, False, False)
			Case "lwarload"
				frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udLAND, frmDrawMap.enumUnitType.TYPE_ALL, False, False, False)
		End Select
	End Sub
End Class