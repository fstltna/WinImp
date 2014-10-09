Option Strict Off
Option Explicit On
Friend Class frmPromptIsland
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'081203 efj: removed dead variables
	'092803 rjk: general reformatting.
	'100603 rjk: Changed OK/Cancel labels for Set Owner to be less confusing.
	'100603 rjk: Added initial field selection.
	
	Private Done As Short
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		' Dim strCmd As String       removed 3/2003 efj
		Dim x1 As Short
		Dim X2 As Short
		Dim y1 As Short
		Dim Y2 As Short
		
		If Not ParseSectors(x1, y1, (txtOrigin.Text)) Then
			MsgBox("Must fill in 'from' sector",  , "Error")
			txtOrigin.Focus()
			Exit Sub
		End If
		
		If Not ParseSectors(X2, Y2, (txtDest.Text)) Then
			If Not Label2.Text = "Set Owner" Then
				MsgBox("Must fill in 'to' sector",  , "Error")
				txtDest.Focus()
				Exit Sub
			Else
				X2 = x1
				Y2 = y1
			End If
		End If
		
		
		If Label2.Text = "Island Develop Tool" Then
			IslandDevelop(x1, y1, X2, Y2)
		ElseIf Label2.Text = "Island Setup Tool" Then 
			BlitzSetup(False, x1, y1, X2, Y2)
		ElseIf Label2.Text = "Set Owner" Then 
			SetOwner(x1, y1, X2, Y2, VB6.GetItemData(cbNations, cbNations.SelectedIndex))
		End If
		
		If Not Label2.Text = "Set Owner" Then
			Me.Close()
		End If
	End Sub
	
	'UPGRADE_WARNING: Form event frmPromptIsland.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmPromptIsland_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		On Error Resume Next
		
		If Done = 0 Then
			If Label2.Text = "Set Owner" Then '100603 rjk: Changed labels for Set Owner to be less confusing.
				cmdOK.Text = "Set Owner"
				cmdCancel.Text = "Exit"
				cbNations.Focus()
			Else
				If Len(txtOrigin.Text) = 0 Or Len(txtDest.Text) > 0 Then
					txtOrigin.Focus()
				Else
					txtDest.Focus()
				End If
			End If
			Done = 1
		End If
		
	End Sub
	
	Private Sub frmPromptIsland_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Make Stay always on top
		' Dim success As Long       removed 3/2003 efj
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
		
		Done = 0 '100603 rjk: Ensure the form runs activated logic each startup.
		
		'Fill list box with nation names
		Nations.FillListBox(cbNations)
		cbNations.SelectedIndex = 0
	End Sub
	
	Private Sub frmPromptIsland_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		'UPGRADE_NOTE: Object frmDrawMap.PromptForm may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmDrawMap.PromptForm = Nothing
		frmDrawMap.PromptUp = False
	End Sub
	
	'UPGRADE_WARNING: Event txtOrigin.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtOrigin_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOrigin.TextChanged
		Dim sx As Short
		Dim sy As Short
		
		On Error Resume Next
		'If first box has been "clicked", more cursor to second box
		If ParseSectors(sx, sy, (txtOrigin.Text)) Then
			txtDest.Focus()
		End If
	End Sub
End Class