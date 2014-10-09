Option Strict Off
Option Explicit On
Friend Class frmServerQuery
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081203 efj: removed dead variables and procedures
	'093003 rjk: general reformatting.
	
	
	Dim TextMaxWidth As Short
	Dim TextLines As Short
	
	Private Sub cmdAbort_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAbort.Click
		txtRespond.Text = "ctld"
		Me.Hide()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		If Len(txtRespond.Text) = 0 Then
			Exit Sub
		End If
		
		Me.Hide()
	End Sub
	
	'UPGRADE_WARNING: Form event frmServerQuery.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmServerQuery_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		Dim nLen As Integer
		
		If TextMaxWidth > 0 Then
			If TextMaxWidth > ((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) * 4) / 5) Then
				TextMaxWidth = (VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) * 4) / 5
			End If
			
			Me.Width = VB6.TwipsToPixelsX(TextMaxWidth * 1.15)
			Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) \ 2)
		End If
		
		If TextLines > 1 Then
			'UPGRADE_ISSUE: Form method frmServerQuery.TextHeight was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			nLen = Me.TextHeight("XX") * (TextLines + 2)
			If VB6.PixelsToTwipsY(lstMsgs.Height) > ((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) * 3) / 4) Then
				lstMsgs.Height = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) * 3) / 4)
			Else
				lstMsgs.Height = VB6.TwipsToPixelsY(nLen)
			End If
			
			Me.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(lstMsgs.Height) + (3 * VB6.PixelsToTwipsY(cmdOK.Height)))
			Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) \ 2)
		End If
		
		Me.Activate()
		txtRespond.Focus()
		lstMsgs.TopIndex = lstMsgs.Items.Count - 1
	End Sub
	
	'Private Sub Form_KeyPress(KeyAscii As Integer)   removed 08/12/03  efj
	'
	'If KeyAscii = 13 Then
	'    cmdOK.Value = True
	'End If
	'
	'End Sub
	
	'UPGRADE_WARNING: Event frmServerQuery.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmServerQuery_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		' Dim x As Integer  removed 8/12/03 efj
		Dim n As Short
		
		'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
		If WindowState <> vbNormal Then
			Exit Sub
		End If
		
		' x = Me.TextHeight("XX")  removed 8/12/03 efj
		n = lstMsgs.Items.Count ' changed from m to n  08/11/03 efj
		If n > 10 Then
			n = 10
		End If
		
		If VB6.PixelsToTwipsX(Me.Width) < 3 * VB6.PixelsToTwipsX(cmdOK.Width) Then
			Me.Width = VB6.TwipsToPixelsX(3 * VB6.PixelsToTwipsX(cmdOK.Width))
		End If
		
		If VB6.PixelsToTwipsY(Me.Height) < 5 * VB6.PixelsToTwipsY(cmdOK.Height) Then
			Me.Height = VB6.TwipsToPixelsY(5 * VB6.PixelsToTwipsY(cmdOK.Height))
		End If
		
		txtRespond.Width = lstMsgs.Width
		
		
		Me.Label1.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.lstMsgs.Top) + VB6.PixelsToTwipsY(lstMsgs.Height) + VB6.PixelsToTwipsY(Label1.Height))
		
		Me.txtRespond.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Label1.Top) + 2 * VB6.PixelsToTwipsY(Label1.Height))
		
		cmdOK.SetBounds(VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(Me.Width) - 2.5 * VB6.PixelsToTwipsX(cmdOK.Width)) \ 2), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(txtRespond.Top) + VB6.PixelsToTwipsY(txtRespond.Height) + VB6.PixelsToTwipsY(cmdOK.Height) * 0.5), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		cmdAbort.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(cmdOK.Left) + 1.5 * VB6.PixelsToTwipsX(cmdOK.Width)), cmdOK.Top, 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		
		Me.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(cmdOK.Top) + 3 * VB6.PixelsToTwipsY(cmdOK.Height))
		
		If VB6.PixelsToTwipsY(Me.Height) < 5 * VB6.PixelsToTwipsY(cmdOK.Height) Then
			Me.Height = VB6.TwipsToPixelsY(5 * VB6.PixelsToTwipsY(cmdOK.Height))
		End If
		
		Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) \ 2)
		Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) \ 2)
	End Sub
End Class