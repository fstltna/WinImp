Option Strict Off
Option Explicit On
Friend Class frmPromptNation
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'120803 efj: removed dead variables
	'270903 rjk: general reformatting
	'011103 rjk: Form cleanup
	'251103 rjk: Added relations update after doing a declare command.
	'180206 rjk: Replace relation with GetCountryList.
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim strCmd As String
		' Dim strCmd2 As String    8/2003 efj  removed
		
		strCmd = Trim(LCase(Label2.Text))
		If strCmd = "relations" Or strCmd = "accept" Then
			frmEmpCmd.SubmitEmpireCommand(strCmd & " " & cbNations.Text, True)
		ElseIf strCmd = "declare" Then 
			frmEmpCmd.SubmitEmpireCommand(strCmd & " " & cbRelations.Text & " " & cbNations.Text, True)
		ElseIf strCmd = "sharebmap" Then 
			frmEmpCmd.SubmitEmpireCommand(strCmd & " " & cbNations.Text & " " & txtOrigin.Text & " " & txtNum.Text, True)
		Else
			Exit Sub
		End If
		If strCmd = "declare" Then '112503 rjk: Added requesting a relations update to get the new declare status.
			frmEmpCmd.SubmitEmpireCommand("db1", False)
			GetCountryList()
			frmEmpCmd.SubmitEmpireCommand("db2", False)
		End If
		Me.Close()
	End Sub
	
	Private Sub frmPromptNation_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Make Stay always on top
		' Dim success As Long  removed 8/12/03 efj
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
		
		'Fill list box with nation names
		Nations.FillListBox(cbNations)
		
		cbRelations.Items.Clear()
		cbRelations.Items.Add("alliance")
		cbRelations.Items.Add("friendly")
		cbRelations.Items.Add("neutrality")
		cbRelations.Items.Add("hostility")
		cbRelations.Items.Add("war")
		cbRelations.SelectedIndex = 3 'default to hostile
	End Sub
	
	Private Sub frmPromptNation_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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