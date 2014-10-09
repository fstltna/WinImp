Option Strict Off
Option Explicit On
Friend Class frmPromptOrigin
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081203 efj: Added Option Explicit and removed dead variables
	'092703 rjk: Set initial field, general reformatting
	'200404 rjk: Added the ability to shift database origin without submitting
	'            a command to server.
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim strCmd As String
		Dim sx As Short
		Dim sy As Short
		
		If chkShiftOnly.CheckState <> System.Windows.Forms.CheckState.Unchecked Then
			If ParseSectors(sx, sy, (txtOrigin.Text)) Then
				ShiftOrigin(sx, sy)
				OriginX = OriginX - sx
				OriginY = OriginY - sy
				frmDrawMap.SelX = frmDrawMap.SelX - sx
				frmDrawMap.SelY = frmDrawMap.SelY - sy
				frmDrawMap.FillSectorBox((frmDrawMap.SelX), (frmDrawMap.SelY))
				'OriginY must be an even number
				If (CShort(OriginY \ 2)) * 2 <> OriginY Then
					OriginY = OriginY + 1
					OriginX = OriginX + 1
				End If
				frmDrawMap.DrawHexes()
			End If
		Else
			strCmd = LCase(Label2.Text) & " " & txtOrigin.Text
			frmEmpCmd.SubmitEmpireCommand(strCmd, True)
		End If
		Me.Close()
	End Sub
	
	'UPGRADE_WARNING: Form event frmPromptOrigin.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmPromptOrigin_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		If frmEmpireLogin.bOffline Then
			chkShiftOnly.CheckState = System.Windows.Forms.CheckState.Indeterminate
		End If
		txtOrigin.Focus()
	End Sub
	
	Private Sub frmPromptOrigin_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Make Stay always on top
		' Dim sucess As Long  removed 8/12/03 efj
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
	End Sub
	
	Private Sub frmPromptOrigin_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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