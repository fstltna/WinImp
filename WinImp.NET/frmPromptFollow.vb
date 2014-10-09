Option Strict Off
Option Explicit On
Friend Class frmPromptFollow
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081203 efj: unused file, commented everything out
	'092303 rjk: Added back and added menu click event to call the form
	'            added unit grid selection
	'102203 rjk: Added Unload Me' to follow form to make consist with other forms when presenting
	'            the Okay button.
	
	Public Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		frmDrawMap.DisplayPromptHelp((Label2.Text))
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim strCmd As String
		Dim num As Short
		Dim x As Short
		
		strCmd = "follow " & txtUnit.Text & " " & txtUnit2.Text
		frmEmpCmd.SubmitEmpireCommand(strCmd, True)
		Me.Close()
	End Sub
	
	Private Sub frmPromptFollow_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Make Stay always on top
		' Dim success As Long  removed 8/12/03 efj
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
		txtUnit.Visible = True
		frmDrawMap.SetUnitDisplay(frmDrawMap.enumUnitDisplay.udSHIP, frmDrawMap.enumUnitType.TYPE_ALL, False, False, False)
	End Sub
	
	Private Sub frmPromptFollow_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
		frmDrawMap.DisplayFirstSectorPanel()
	End Sub
End Class