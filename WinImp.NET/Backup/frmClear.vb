Option Strict Off
Option Explicit On
Friend Class frmClear
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'112903 rjk: Created to give more control to the clear functions.
	'120303 rjk: Added the ability to select by Allied/Neutral/Enemy as well
	'            Added post clear command to refill the intelligence.
	'120303 rjk: Moved the screen update to after the post clear commands, (display refreshed intelligence)
	'300704 rjk: Added Days range to clear function.
	'021104 rjk: Changed spy/llook/look/coastwatch default to off.
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		If chkTelegrams.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			frmTelegram.ClearTelegrams()
		End If
		
		ClearEnemyInfo((chkSectors.CheckState <> System.Windows.Forms.CheckState.Unchecked), (chkShips.CheckState <> System.Windows.Forms.CheckState.Unchecked), (chkLands.CheckState <> System.Windows.Forms.CheckState.Unchecked), (chkPlanes.CheckState <> System.Windows.Forms.CheckState.Unchecked), (chkAirCombat.CheckState <> System.Windows.Forms.CheckState.Unchecked), (chkAllied.CheckState <> System.Windows.Forms.CheckState.Unchecked), (chkNeutral.CheckState <> System.Windows.Forms.CheckState.Unchecked), (chkEnemy.CheckState <> System.Windows.Forms.CheckState.Unchecked), VB6.GetItemData(cbDays, cbDays.SelectedIndex))
		
		If chkSpy.CheckState <> System.Windows.Forms.CheckState.Unchecked Or chkLook.CheckState <> System.Windows.Forms.CheckState.Unchecked Or chkLlook.CheckState <> System.Windows.Forms.CheckState.Unchecked Or chkCoastWatch.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			frmEmpCmd.SubmitEmpireCommand("bf1", False)
			If chkSpy.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
				frmEmpCmd.SubmitEmpireCommand("spy *", False)
			End If
			If chkLook.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
				frmEmpCmd.SubmitEmpireCommand("look *", False)
			End If
			If chkLlook.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
				frmEmpCmd.SubmitEmpireCommand("llook *", False)
			End If
			If chkCoastWatch.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
				frmEmpCmd.SubmitEmpireCommand("coastwatch *", False)
			End If
			frmEmpCmd.SubmitEmpireCommand("bf2", False)
		End If
		
		If (chkSectors.CheckState <> System.Windows.Forms.CheckState.Unchecked) Or (chkShips.CheckState <> System.Windows.Forms.CheckState.Unchecked) Or (chkLands.CheckState <> System.Windows.Forms.CheckState.Unchecked) Or (chkPlanes.CheckState <> System.Windows.Forms.CheckState.Unchecked) Or (chkAirCombat.CheckState <> System.Windows.Forms.CheckState.Unchecked) Then
			frmDrawMap.DrawHexes()
		End If
		Me.Close()
	End Sub
	
	Private Sub frmClear_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		cbDays.SelectedIndex = 0
	End Sub
End Class