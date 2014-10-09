Option Strict Off
Option Explicit On
Friend Class frmPromptMap
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081203 efj: removed dead variables
	'092803 rjk: General Reformatting. Made Okay the default
	'            Added Grid Unit selection on start up and field
	'            selection.  Added reset to normal sector display upon
	'            exit.  Added initial field selection.
	'            Move field selection logic to form from frmDrawMap
	'            Remove spaces between options so they were work
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim strCmd As String
		Dim n As Short
		strCmd = LCase(Trim(Label2.Text)) & " "
		If txtOrigin.Visible Then
			strCmd = strCmd & txtOrigin.Text & " " ' 092403 rjk: Removed the space between options as they are not accepted
		Else
			strCmd = strCmd & txtUnit.Text & " " ' 092403 rjk: Removed the space between options as they are not accepted
		End If
		
		'Check the check boxes for display options
		For n = 0 To 2
			If Check1(n).CheckState = System.Windows.Forms.CheckState.Checked Then
				strCmd = strCmd & Check1(n).Tag ' 092403 rjk: Removed the space between options as they are not accepted
			End If
		Next n
		
		frmEmpCmd.SubmitEmpireCommand(strCmd, True)
		
		Me.Close()
	End Sub
	
	'UPGRADE_WARNING: Form event frmPromptMap.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmPromptMap_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		Select Case LCase(Trim(Label2.Text))
			Case "smap", "lmap", "pmap", "sbmap", "lbmap", "pbmap"
				Me.txtUnit.Visible = True
				Me.txtUnit.Focus()
				Me.txtOrigin.Visible = False
			Case Else
				Me.txtUnit.Visible = False
				Me.txtOrigin.Visible = True
				Me.txtOrigin.Focus()
		End Select
	End Sub
	
	Private Sub frmPromptMap_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		' Set parent for the toolbar to display on top of:
		'Dim success As Long  removed 8/12/03 efj
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
	End Sub
	
	Private Sub frmPromptMap_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
		frmDrawMap.DisplayFirstSectorPanel()
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
	
	Private Sub txtUnit_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnit.Enter
		Select Case LCase(Trim(Label2.Text))
			Case "smap", "sbmap"
				frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udSHIP, frmDrawMap.enumUnitType.TYPE_ALL, False, False, False)
			Case "lmap", "lbmap"
				frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udLAND, frmDrawMap.enumUnitType.TYPE_ALL, False, False, False)
			Case "pmap", "pbmap"
				frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udPLANE, frmDrawMap.enumUnitType.TYPE_ALL, False, False, False)
		End Select
	End Sub
End Class