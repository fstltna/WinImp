Option Strict Off
Option Explicit On
Friend Class frmPromptOffer
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081203 efj: removed dead variables
	'092403 rjk: Switched to label1 to left justified and moved txtNum 500 to make room for number from Consider
	'092703 rjk: General reformatting
	'100903 rjk: Moved the field logic to the form.  Modified the form to 'reject accept' and 'reject reject' commands
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Public Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim strCmd As String
		' Dim strCmd2 As String    8/2003 efj  removed
		
		'Build Loan offer
		strCmd = Trim(LCase(Label2.Text))
		If strCmd = "offer" Then
			If Option1.Checked Then
				strCmd = strCmd & " loan " & cbNations.Text
				If txtAmount.Visible = True Then
					If Len(txtAmount.Text) > 0 Then
						strCmd = strCmd & " " & txtAmount.Text
						If Len(txtDur.Text) > 0 Then
							strCmd = strCmd & " " & txtDur.Text
							If Len(txtRate.Text) > 0 Then
								strCmd = strCmd & " " & txtRate.Text
							End If
						End If
					End If
				Else
					strCmd = strCmd & " " & txtNum.Text
				End If
			Else
				strCmd = strCmd & " treaty " & cbNations.Text
			End If
		ElseIf strCmd = "consider" Then 
			If Option1.Checked Then
				strCmd = strCmd & " loan " & txtNum.Text
			Else
				strCmd = strCmd & " treaty " & txtNum.Text
			End If
		Else '100903 rjk: Added 'reject accept' and 'reject reject' commands
			strCmd = "reject " & strCmd
			If Option1.Checked Then
				strCmd = strCmd & " loan"
			ElseIf Option2.Checked Then 
				strCmd = strCmd & " treaties"
			ElseIf Option3.Checked Then 
				strCmd = strCmd & " mail"
			ElseIf Option4.Checked Then 
				strCmd = strCmd & " announcements"
			End If
			If Len(cbNations.Text) > 0 Then
				strCmd = strCmd & " " & cbNations.Text
			End If
		End If
		
		frmEmpCmd.SubmitEmpireCommand(strCmd, True)
		
		Me.Close()
	End Sub
	
	'UPGRADE_WARNING: Form event frmPromptOffer.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmPromptOffer_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		Select Case Trim(LCase(Label2.Text))
			Case "offer"
			Case "consider"
				txtNum.Visible = True
				txtNum.Text = vbNullString
				'092403 rjk: Added 500 to make room for number
				txtNum.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(cbNations.Left) + 500), cbNations.Top, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
				cbNations.Visible = False
				Label1.Text = "number"
				LoanPromptsVisible(False)
			Case "reject", "accept" '100903 rjk: Added additional fields for reject commands
				LoanPromptsVisible(False)
				Option3.Visible = True
				Option4.Visible = True
		End Select
		
		If Trim(LCase(Label2.Text)) <> "consider" Then
			'Fill list box with nation names
			Nations.FillListBox(cbNations)
			cbNations.SelectedIndex = 0
		End If
	End Sub
	
	Private Sub frmPromptOffer_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		On Error Resume Next 'In case nation list is empty
		'Make Stay always on top
		' Dim success As Long  removed 8/12/03 efj
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
		
		'Put form in proper place
		Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(frmDrawMap.Left) + (VB6.PixelsToTwipsX(frmDrawMap.Width) - VB6.PixelsToTwipsX(Width)) \ 2)
		Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(frmDrawMap.Top) + VB6.PixelsToTwipsY(frmDrawMap.Height) - VB6.PixelsToTwipsY(Height))
	End Sub
	
	Private Sub frmPromptOffer_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	'UPGRADE_WARNING: Event Option1.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option1_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option1.CheckedChanged
		If eventSender.Checked Then
			
			If LCase(Trim(Label2.Text)) = "offer" Then
				LoanPromptsVisible(True)
			End If
			
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Option2.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option2_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option2.CheckedChanged
		If eventSender.Checked Then
			LoanPromptsVisible(False)
		End If
	End Sub
	
	'100903 rjk: Added to manage the LoanPrompts fields on the screen.
	Private Sub LoanPromptsVisible(ByRef bVisible As Boolean)
		Label3.Visible = bVisible
		Label4.Visible = bVisible
		Label5.Visible = bVisible
		txtAmount.Visible = bVisible
		txtRate.Visible = bVisible
		txtDur.Visible = bVisible
	End Sub
End Class