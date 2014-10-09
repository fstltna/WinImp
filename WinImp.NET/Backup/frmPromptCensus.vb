Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmPromptCensus
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081203 efj: Added Option Explicit and removed dead variables
	'092703 rjk: General reformatting.  Added initial field selection.
	'            Repositioned fields on form so label2 would not override label1.
	'101703 rjk: Added Strength fields to Sector database
	'112003 rjk: Added option to control strength updates
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	'250606 rjk: Added start/stop unit support for 4.3.6 servers.
	'181006 rjk: Ensure the unit list only appears for start and stop.
	
	Private Sub FillForm()
		Dim strCmd As String
		
		strCmd = LCase(Label2.Text)
		
		On Error Resume Next
		If VersionCheck(4, 3, 6) = WinAceRoutines.enumVersion.VER_LESS Or (strCmd <> "start" And strCmd <> "stop") Then
			cmbType.Visible = False
			cmbType.Text = ""
			txtOrigin.Visible = True
			txtUnit.Visible = False
			txtOrigin.Focus()
			Exit Sub
		End If
		cmbType.Visible = True
		If cmbType.SelectedIndex = 0 Then
			txtOrigin.Visible = True
			txtUnit.Visible = False
			txtOrigin.Focus()
		Else
			txtOrigin.Visible = False
			txtUnit.Visible = True
			txtUnit.Focus()
		End If
		Select Case cmbType.SelectedIndex
			Case 0
				frmDrawMap.DisplayFirstSectorPanel()
			Case 1
				frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udSHIP, frmDrawMap.enumUnitType.TYPE_ALL, False, False, False)
			Case 2
				frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udLAND, frmDrawMap.enumUnitType.TYPE_ALL, False, False, False)
			Case 3
				frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udPLANE, frmDrawMap.enumUnitType.TYPE_ALL, False, False, False)
			Case 4
				frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udNUKE, frmDrawMap.enumUnitType.TYPE_ALL, False, False, False)
		End Select
	End Sub
	'UPGRADE_WARNING: Event cmbType.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cmbType.Change was upgraded to cmbType.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cmbType_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbType.TextChanged
		FillForm()
	End Sub
	
	'UPGRADE_WARNING: Event cmbType.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cmbType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbType.SelectedIndexChanged
		FillForm()
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim strCmd As String
		
		strCmd = LCase(Label2.Text)
		'database update
		If strCmd = "start" Or strCmd = "stop" Then
			If cmbType.SelectedIndex <= 0 Then
				frmEmpCmd.SubmitEmpireCommand(strCmd & " " & LCase(VB.Left(cmbType.Text, 4)) & " " & txtOrigin.Text, True)
			Else
				frmEmpCmd.SubmitEmpireCommand(strCmd & " " & LCase(VB.Left(cmbType.Text, 4)) & " " & txtUnit.Text, True)
			End If
			frmEmpCmd.SubmitEmpireCommand("db1", False)
			Select Case (cmbType.SelectedIndex)
				Case -1, 0
					GetSectors()
					'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
					GetCurrentStrength(tsSectors)
				Case 1
					GetShips()
				Case 2
					GetLandUnits()
				Case 3
					GetPlanes()
				Case 4
					GetNukes()
			End Select
			frmEmpCmd.SubmitEmpireCommand("db2", False)
		ElseIf strCmd = "Import Offset" Then 
		Else
			frmEmpCmd.SubmitEmpireCommand(strCmd & " " & txtOrigin.Text, True)
		End If
		
		Me.Close()
	End Sub
	
	'UPGRADE_WARNING: Form event frmPromptCensus.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmPromptCensus_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		If (Label2.Text = "Start" Or Label2.Text = "Stop") And VersionCheck(4, 3, 6) <> WinAceRoutines.enumVersion.VER_LESS Then
			FillForm()
		Else
			cmbType.Visible = False
			cmbType.Text = ""
			txtOrigin.Visible = True
			txtUnit.Visible = False
			txtOrigin.Focus() '092203 rjk: added to ensure the sector map will update the txtOrigin field
		End If
	End Sub
	
	Private Sub frmPromptCensus_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Make Stay always on top
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
	End Sub
	
	Private Sub frmPromptCensus_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
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
End Class