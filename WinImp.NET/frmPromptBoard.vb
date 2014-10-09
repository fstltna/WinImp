Option Strict Off
Option Explicit On
Friend Class frmPromptBoard
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081203 efj: removed dead variables
	'091803 rjk: Added Unit grid selection when activating this form
	'092803 rjk: Switched to SetUnitType and DisplayFirstSectorPanel
	'            General reformatting.
	'101703 rjk: Added Strength fields to Sector database
	'112003 rjk: Added option to control strength updates
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	'160705 rjk: Added ability to get OContent for Sea Sectors
	'180206 rjk: Replace ldump and sdump with GetLandUnits and GetShips.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		frmDrawMap.DisplayFirstSectorPanel()
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim strCmd As String
		
		strCmd = "board " & txtTarget.Text & " "
		If txtUnit.Visible Then
			strCmd = strCmd & txtUnit.Text
		Else
			strCmd = strCmd & txtOrigin.Text
		End If
		
		frmEmpCmd.SubmitEmpireCommand(strCmd, True)
		
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		GetShips()
		If Not txtUnit.Visible Then
			GetSectors()
			'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
			GetCurrentStrength(tsSectors)
			GetOContent()
			GetLandUnits()
		End If
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
		frmDrawMap.DisplayFirstSectorPanel()
		Me.Close()
	End Sub
	
	Private Sub frmPromptBoard_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Make Stay always on top
		' Dim success As Long  removed 8/12/03 efj
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
	End Sub
	
	Private Sub frmPromptBoard_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	Private Sub txtUnit_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnit.Enter
		frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udSHIP, frmDrawMap.enumUnitType.TYPE_ALL, True, True, False)
	End Sub
	
	Private Sub txtTarget_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTarget.Enter
		frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udENEMY, frmDrawMap.enumUnitType.TYPE_ALL, False, False, False)
	End Sub
	
	Private Sub txtOrigin_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOrigin.Enter
		frmDrawMap.DisplayFirstSectorPanel()
	End Sub
End Class