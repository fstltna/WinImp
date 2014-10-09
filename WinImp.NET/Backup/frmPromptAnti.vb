Option Strict Off
Option Explicit On
Friend Class frmPromptAnti
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081203 efj: Added Option Explicit and removed dead variables
	'092803 rjk: General reformatting.
	'101703 rjk: Added Strength fields to Sector database
	'111403 rjk: Added so update commands are not shown in the command window
	'112003 rjk: Added option to control strength updates
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	'180206 rjk: Replace lost command with GetLost.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim strCmd As String
		
		strCmd = "anti " & txtOrigin.Text
		frmEmpCmd.SubmitEmpireCommand(strCmd, True)
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False) '111403 rjk: Added so update commands are not shown in the command window
		GetSectors()
		'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
		GetCurrentStrength(tsSectors)
		GetLost()
		frmEmpCmd.SubmitEmpireCommand("db2", False) '111403 rjk: Added so update commands are not shown in the command window
		Me.Close()
	End Sub
	
	Private Sub frmPromptAnti_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Make Stay always on top
		' Dim sucess As Long  removed 8/12/03 efj
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
	End Sub
	
	Private Sub frmPromptAnti_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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