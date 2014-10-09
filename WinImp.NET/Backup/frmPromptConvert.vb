Option Strict Off
Option Explicit On
Friend Class frmPromptConvert
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081203 efj: added Option Explcit and removed dead variables
	'092803 rjk: Fill in txtNum based on the origin and command.
	'            Added initial field selection. General reformatting.
	'093003 rjk: Blank the txtNum field if it is not a valid origin.
	'            Only move focus to txtNum field if origin sector is valid.
	'            Update the mil reserve level after enlist and demobilize commands.
	'093003 rjk: removed the LCase for command determination and adjust the corresponding text.
	'101703 rjk: Added Strength fields to Sector database
	'112003 rjk: Added option to control strength updates
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	'180206 rjk: Replace nation with GetNations.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim strCmd As String
		
		If Label2.Text = "Realm" Then '093003 rjk: removed the LCase and adjust the text.
			strCmd = LCase(Label2.Text) & " " & txtNum.Text & " " & txtOrigin.Text
			frmEmpCmd.SubmitEmpireCommand(strCmd, True)
		Else
			strCmd = LCase(Label2.Text) & " " & txtOrigin.Text & " " & txtNum.Text
			'Handle reserve check box for demobilization
			If chkReserve.Visible Then
				If chkReserve.CheckState = System.Windows.Forms.CheckState.Checked Then
					strCmd = strCmd & " y"
				Else
					strCmd = strCmd & " n"
				End If
			End If
			
			frmEmpCmd.SubmitEmpireCommand(strCmd, True)
			'database update
			frmEmpCmd.SubmitEmpireCommand("db1", False)
			'093003 rjk: Update the mil reserve level after enlist or demobilize commands
			If Label2.Text = "Enlist" Or Label2.Text = "Demobilize" Then
				GetNation()
			End If
			GetSectors()
			'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
			GetCurrentStrength(tsSectors)
			frmEmpCmd.SubmitEmpireCommand("db2", False)
		End If
		Me.Close()
	End Sub
	
	'UPGRADE_WARNING: Form event frmPromptConvert.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmPromptConvert_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		txtNum.Focus()
	End Sub
	
	Private Sub frmPromptConvert_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Make Stay always on top
		' Dim sucess As Long  removed 8/12/03 efj
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
	End Sub
	
	Private Sub frmPromptConvert_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	'UPGRADE_WARNING: Event txtOrigin.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtOrigin_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOrigin.TextChanged
		Dim secx As Short
		Dim secy As Short
		
		If ParseSectors(secx, secy, (txtOrigin.Text)) Then
			rsSectors.Seek("=", secx, secy)
			If Not rsSectors.NoMatch Then
				Select Case Label2.Text
					Case "Grind"
						txtNum.Text = CStr(rsSectors.Fields("bar").Value)
					Case "Realm", "Enlist"
						txtNum.Text = vbNullString
					Case "Convert"
						txtNum.Text = CStr(rsSectors.Fields("civ").Value)
					Case "Demobilize"
						txtNum.Text = CStr(rsSectors.Fields("mil").Value)
				End Select
				If txtNum.Visible Then '093003 rjk: only move focus if sector is valid
					txtNum.Focus()
				End If
			Else
				txtNum.Text = vbNullString
			End If
		Else '093003 rjk: Blank the field if it is not a valid origin.
			txtNum.Text = vbNullString
		End If
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