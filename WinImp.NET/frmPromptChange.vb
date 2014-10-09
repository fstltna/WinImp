Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmPromptChange
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081203 efj: Added Option Explicit and removed dead variables
	'092803 rjk: general reformatting.
	'100703 rjk: Nation list now in relations only so do not need "report *"
	'100703 rjk: Does not accept a string longer than 19 characters for country name.
	'100703 rjk: remove any space in the front and rear of the country name to make matching easier.
	'180206 rjk: Replace relation with GetCountryList.
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim strCmd As String
		
		strCmd = "change "
		If Option1.Checked Then
			txtNew.Text = Trim(txtNew.Text) '100703 rjk: remove any space in the front and rear to make matching easier.
			If Len(txtNew.Text) > 19 Then '100703 rjk: Does not accept a string longer than 19 characters.
				txtNew.Text = VB.Left(txtNew.Text, 19)
			End If
			frmEmpCmd.CountryName = txtNew.Text '100703 rjk: Updated the name so relation grid will work
			strCmd = strCmd & "country " & Chr(34) & txtNew.Text & Chr(34)
			frmEmpCmd.SubmitEmpireCommand(strCmd, True)
			frmEmpCmd.SubmitEmpireCommand("db1", False)
			GetCountryList()
			frmEmpCmd.SubmitEmpireCommand("db2", False)
		Else
			strCmd = strCmd & "representative " & Chr(34) & txtNew.Text & Chr(34)
			frmEmpCmd.SubmitEmpireCommand(strCmd, True)
		End If
		'database update
		Me.Close()
	End Sub
	
	Private Sub frmPromptChange_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Make Stay always on top
		' Dim sucess As Long  removed 8/12/03 efj
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, Flags)
	End Sub
	
	Private Sub frmPromptChange_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
End Class