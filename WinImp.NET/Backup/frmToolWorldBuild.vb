Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmToolWorldBuild
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'031003 rjk: In non-deity mode wilderness is not the list so left button is blank
	'            but designates as headquarters, changed so blank does not designate.
	'            needs more work.
	'171003 rjk: Added Strength fields to Sector database
	'201103 rjk: Added option to control strength updates
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	
	Dim strCodes As String
	Dim strNewDes(1) As String
	
	'UPGRADE_WARNING: Event cbType.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cbType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbType.SelectedIndexChanged
		Dim Index As Short = cbType.GetIndex(eventSender)
		
		If cbType(Index).SelectedIndex < 0 Then
			Exit Sub
		End If
		strNewDes(Index) = Mid(strCodes, VB6.GetItemData(cbType(Index), cbType(Index).SelectedIndex), 1)
		
	End Sub
	
	'UPGRADE_WARNING: Event chkDes.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkDes_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkDes.CheckStateChanged
		If chkDes.CheckState = System.Windows.Forms.CheckState.Checked Then
			frmDrawMap.WorldBuilding = True
		Else
			frmDrawMap.WorldBuilding = False
		End If
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		GetSectors()
		'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
		GetCurrentStrength(tsSectors)
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		Me.Close()
	End Sub
	
	Private Sub cmdCopy_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCopy.Click
		
		Dim sx As Short
		Dim sy As Short
		Dim sx1 As Short
		Dim sy1 As Short
		Dim sx2 As Short
		Dim sy2 As Short
		On Error Resume Next
		
		If Not ParseSectors(sx, sy, (txtPath.Text)) Then
			MsgBox("Destination Invalid", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly, "Copy Error")
			Exit Sub
		End If
		
		If Not ParseSectorRange(sx1, sx2, sy1, sy2, txtOrigin.Text) Then
			MsgBox("Origin Invalid", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly, "Copy Error")
			Exit Sub
		End If
		CopySectorInfo(sx1, sx2, sy1, sy2, sx, sy, (cbCopy.SelectedIndex))
	End Sub
	
	Private Sub cmdSetResource_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSetResource.Click
		SetResources()
		
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		GetSectors()
		'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
		GetCurrentStrength(tsSectors)
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
	End Sub
	
	Private Sub frmToolWorldBuild_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		strCodes = LoadTypebox(cbType(0))
		strCodes = LoadTypebox(cbType(1))
		cbType(0).SelectedIndex = InStr(strCodes, "-") - 1
		If cbType(0).SelectedIndex = -1 Then '100303 rjk: in non-deity mode wilderness is not the list so left button is blank
			strNewDes(0) = ""
		End If
		cbType(1).SelectedIndex = cbType(1).Items.Count - 1
		cbCopy.SelectedIndex = 0
		
		'Make Stay always on top
		' Dim success As Long    8/2003 efj  removed
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
		
	End Sub
	
	Private Sub frmToolWorldBuild_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		frmDrawMap.WorldBuilding = False
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
	
	Public Sub DesignateHex(ByRef sectx As Short, ByRef secty As Short, ByRef Button As Short) '8/2003 efj  removed, 09/11/03 rjk added back used when selecting sector from map
		Dim strCmdDes As String
		Dim strDes As String
		
		strCmdDes = "designate " & SectorString(sectx, secty) & " "
		If Button = VB6.MouseButtonConstants.LeftButton Then
			strDes = strNewDes(0)
		Else
			strDes = strNewDes(1)
		End If
		
		If strDes = "" Then '100303 rjk: If Left or Right button is blank then do not designate
			Exit Sub
		End If
		
		strCmdDes = strCmdDes & strDes
		
		rsSectors.Seek("=", sectx, secty)
		If Not rsSectors.NoMatch Then
			rsSectors.Edit()
			rsSectors.Fields("des").Value = strDes
			rsSectors.Update()
		End If
		
		'now do the designation
		frmEmpCmd.SubmitEmpireCommand(strCmdDes, True)
		
		If Check1.CheckState = System.Windows.Forms.CheckState.Checked Then
			txtDest.Text = SectorString(sectx, secty)
			SetResources()
		End If
		
	End Sub
	
	
	Private Sub SetResources()
		Dim n As Short
		Dim strCmd As String
		
		For n = 0 To 4
			If Len(txtThresh(n).Text) > 0 Then
				strCmd = Trim("setresource " & LCase(lblThresh(n).Text) & " " & txtDest.Text & " " & txtThresh(n).Text)
				frmEmpCmd.SubmitEmpireCommand(strCmd, True)
			End If
		Next n
		
	End Sub
End Class