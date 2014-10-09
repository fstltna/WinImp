Option Strict Off
Option Explicit On
Friend Class frmPromptRadar
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'092703 rjk: Add initial field selection. General reformatting.
	'101703 rjk: Added Strength fields to Sector database
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim strCmd As String
		
		strCmd = LCase(Label2.Text) & " " & txtOrigin.Text
		frmEmpCmd.SubmitEmpireCommand(strCmd, True)
		
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		'102703 rjk: Remove the dump and strength from the spy command
		If Trim(LCase(Label2.Text)) <> "spy" Then
			frmEmpCmd.SubmitEmpireCommand("bmap *", False)
		End If
		
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		Me.Close()
	End Sub
	
	'UPGRADE_WARNING: Form event frmPromptRadar.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmPromptRadar_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		txtOrigin.Focus()
	End Sub
	
	Private Sub frmPromptRadar_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Make Stay always on top
		' Dim success As Long    8/2003 efj  removed
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
		
	End Sub
	
	Private Sub frmPromptRadar_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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