Option Strict Off
Option Explicit On
Friend Class frmPromptCom
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081103 efj: Added Option Explicit
	'101603 rjk: Added functionality for Buy, Sell, Reset and enhanced Market
	'101703 rjk: Added Strength fields to Sector database
	'101803 rjk: Added Route to using this form
	'112003 rjk: Added option to control strength updates
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	
	Public strSector As String
	Private bFirstLoad As Boolean
	
	'UPGRADE_WARNING: Event cbCom.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cbCom_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbCom.SelectedIndexChanged
		Dim strComm As String
		
		If Label2.Text = "Buy" Then
			strComm = cbCom.Text
			If strComm = "unc. workers" Then
				strComm = "uw"
			End If
			If bFirstLoad Then
				frmEmpCmd.SubmitEmpireCommand("market all", False)
			Else
				frmEmpCmd.SubmitEmpireCommand("market " & strComm, False)
			End If
		End If
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		If Label2.Text = "Buy" Or Label2.Text = "Reset" Then
			If frmReport.Visible Then
				frmReport.Close()
			End If
		End If
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim strComm As String
		Dim strCmd As String
		Dim sx As Short
		Dim sy As Short
		
		strComm = cbCom.Text
		If strComm = "unc. workers" Then
			strComm = "uw"
		End If
		
		Select Case Label2.Text
			Case "Buy"
				If frmReport.Visible Then
					frmReport.Close()
				End If
				strCmd = "buy " & strComm
				If txtLotNumber.Text <> "" Then
					strCmd = strCmd & " " & txtLotNumber.Text
					If txtPrice.Text <> "" Then
						strCmd = strCmd & " " & txtPrice.Text
						If ParseSectors(sx, sy, (txtOrigin.Text)) Then
							rsSectors.Seek("=", sx, sy)
							If rsSectors.NoMatch Then
								MsgBox("You do not own this sector")
								Exit Sub
							End If
							If rsSectors.Fields("des").Value <> "h" And rsSectors.Fields("des").Value <> "w" Then
								MsgBox("This is not a valid sector must be either harbor or warehouse")
								Exit Sub
							End If
							strCmd = strCmd & " " & txtOrigin.Text
						End If
					End If
				End If
			Case "Market"
				strCmd = "market " & strComm
				
			Case "Reset"
				If frmReport.Visible Then
					frmReport.Close()
				End If
				strCmd = "reset " & strComm
				If txtLotNumber.Text <> "" Then
					strCmd = strCmd & " " & txtLotNumber.Text
					If txtPrice.Text <> "" Then
						strCmd = strCmd & " " & txtPrice.Text
					End If
				End If
			Case "Sell"
				strCmd = "sell " & strComm
				If txtOrigin.Text <> "" Then
					'Manual indicates multiple sectors are supported but does not seem to work
					'so currently will ensure we are starting with a valid harbor or warehouse
					If ParseSectors(sx, sy, (txtOrigin.Text)) Then
						rsSectors.Seek("=", sx, sy)
						If rsSectors.NoMatch Then
							MsgBox("You do not own this sector")
							Exit Sub
						End If
						If rsSectors.Fields("des").Value <> "h" And rsSectors.Fields("des").Value <> "w" Then
							MsgBox("This is not a valid sector must be either harbor or warehouse")
							Exit Sub
						End If
						strCmd = strCmd & " " & txtOrigin.Text
						If txtLotNumber.Text <> "" Then
							strCmd = strCmd & " " & txtLotNumber.Text
							If txtPrice.Text <> "" Then
								strCmd = strCmd & " " & txtPrice.Text
							End If
						End If
					End If
				End If
			Case "Route"
				strCmd = "route " & strComm & " " & txtOrigin.Text
		End Select
		frmEmpCmd.SubmitEmpireCommand(strCmd, True)
		
		'database update
		If Label2.Text <> "Market" Then
			frmEmpCmd.SubmitEmpireCommand("db1", False)
			GetSectors()
			'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
			GetCurrentStrength(tsSectors)
			frmEmpCmd.SubmitEmpireCommand("db2", False)
		End If
		Me.Close()
	End Sub
	
	'UPGRADE_WARNING: Form event frmPromptCom.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmPromptCom_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		If bFirstLoad Then
			txtOrigin.Text = strSector
			
			Select Case Label2.Text
				Case "Buy"
				Case "Market"
					lblLotNumber.Visible = False
					txtLotNumber.Visible = False
					lblPrice.Visible = False
					txtPrice.Visible = False
					lblOrigin.Visible = False
					txtOrigin.Visible = False
					cbCom.Items.Insert(0, "all") 'note: "All" does not work
					cbCom.Items.Insert(1, "*")
				Case "Reset"
					frmEmpCmd.SubmitEmpireCommand("market all", False)
					lblOrigin.Visible = False
					txtOrigin.Visible = False
					lblCommodity.Visible = False
					cbCom.Visible = False
				Case "Sell"
					lblLotNumber.Text = "Quantity"
					lblOrigin.Text = "from"
				Case "Route"
					LoadCommodityBox(cbCom) 'load back missing items.
					lblPrice.Visible = False
					txtPrice.Visible = False
					lblOrigin.Visible = False
					txtOrigin.Visible = False
			End Select
			If cbCom.Visible Then
				cbCom.SelectedIndex = 0
				cbCom.Focus()
			Else
				txtLotNumber.Focus()
			End If
			bFirstLoad = False
		End If
	End Sub
	
	Private Sub frmPromptCom_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim Index As Short
		
		bFirstLoad = True
		
		'Fill list box with commodities
		LoadCommodityBox(cbCom)
		
		Index = 0
		While Index < cbCom.Items.Count - 1
			If VB6.GetItemString(cbCom, Index) = "military" Or VB6.GetItemString(cbCom, Index) = "civilians" Then
				cbCom.Items.RemoveAt(Index)
			End If
			Index = Index + 1
		End While
		
		Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(frmDrawMap.Left) + (VB6.PixelsToTwipsX(frmDrawMap.Width) - VB6.PixelsToTwipsX(Width)) \ 2)
		Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(frmDrawMap.Top) + VB6.PixelsToTwipsY(frmDrawMap.Height) - VB6.PixelsToTwipsY(Height))
		
		'Make Stay always on top
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
	End Sub
	
	Private Sub frmPromptCom_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	'Manual indicates multiple sectors are supported but does not seem to work
	'Private Sub txtOrigin_DblClick()
	'If Label2.Caption = "Sell" Then
	'    Load frmPromptSectors
	'    frmPromptSectors.strSectors = txtOrigin.Text
	'    frmPromptSectors.SetControls
	'    frmPromptSectors.Caption = "Select Sectors"
	'    frmPromptSectors.Left = frmDrawMap.Left + frmDrawMap.Width - frmPromptSectors.Width
	'    frmPromptSectors.top = frmDrawMap.top + (frmDrawMap.Height - frmPromptSectors.Height) \ 2
	'    frmPromptSectors.Show vbModeless
	'End If
	'End Sub
	'UPGRADE_WARNING: Event txtOrigin.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtOrigin_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOrigin.TextChanged
		If Label2.Text = "Route" Then
			'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
			Load(frmPromptSectors)
			frmPromptSectors.strSectors = txtOrigin.Text
			frmPromptSectors.SetControls()
			frmPromptSectors.Text = "Select Sectors"
			frmPromptSectors.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(frmDrawMap.Left) + VB6.PixelsToTwipsX(frmDrawMap.Width) - VB6.PixelsToTwipsX(frmPromptSectors.Width))
			frmPromptSectors.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(frmDrawMap.Top) + (VB6.PixelsToTwipsY(frmDrawMap.Height) - VB6.PixelsToTwipsY(frmPromptSectors.Height)) \ 2)
			frmPromptSectors.Show()
		End If
	End Sub
End Class