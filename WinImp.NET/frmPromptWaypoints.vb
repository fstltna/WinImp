Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmPromptWaypoints
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'08xx03 efj: removed dead variables
	'092703 rjk: general reformatting
	
	Public HoldSource As System.Windows.Forms.Form
	' Public strSectors As String    8/2003 efj  removed
	
	Private Receiver As System.Windows.Forms.Control
	' Private Done As Integer    8/2003 efj  removed
	
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		
		Dim n As Short
		Dim x As Short
		Dim strSector As String
		
		'Load all waypoints into a string
		For x = 0 To lstPoints.Items.Count - 1
			n = InStr(VB6.GetItemString(lstPoints, x), " - ")
			If n > 0 Then
				strSector = strSector & VB.Left(VB6.GetItemString(lstPoints, x), n) & ";"
			End If
		Next x
		
		'Put the string in the Source
		HoldSource.Enabled = True
		Receiver.Text = strSector
		Receiver.Focus()
		Me.Close()
	End Sub
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		lstPoints.Items.Clear()
	End Sub
	
	'UPGRADE_WARNING: Form event frmPromptWaypoints.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmPromptWaypoints_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		On Error Resume Next
		HoldSource.Enabled = False
	End Sub
	
	Private Sub frmPromptWaypoints_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'Make Stay always on top
		' Dim success As Long    8/2003 efj  removed
		Call SetWindowPos(Me.Handle.ToInt32, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
		
		'Set to recieve focus for clicks on DrawMap
		HoldSource = frmDrawMap.PromptForm
		frmDrawMap.PromptForm = Me
		Receiver = HoldSource.ActiveControl
		'Hide the txtOrigin box
		'txtOrigin.Move lstPoints.Left, lstPoints.top
	End Sub
	
	Private Sub frmPromptWaypoints_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		frmDrawMap.PromptForm = HoldSource
		HoldSource.Enabled = True
		'UPGRADE_NOTE: Object HoldSource may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		HoldSource = Nothing
	End Sub
	
	'UPGRADE_WARNING: Event txtOrigin.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtOrigin_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOrigin.TextChanged
		Dim sx As Short
		Dim sy As Short
		Dim strText As String
		
		On Error Resume Next
		If ParseSectors(sx, sy, (txtOrigin.Text)) Then
			'Get Sector Information
			rsSectors.Seek("=", sx, sy)
			If Not rsSectors.NoMatch Then
				strText = txtOrigin.Text & " - " & CStr(rsSectors.Fields("eff").Value) & "% "
				'UPGRADE_WARNING: Couldn't resolve default property of object colDes.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				strText = strText + colDes.Item(rsSectors.Fields("des").Value)
			Else
				rsBmap.Seek("=", sx, sy)
				If Not rsBmap.NoMatch Then
					strText = txtOrigin.Text & " - "
					'UPGRADE_WARNING: Couldn't resolve default property of object colDes.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					strText = strText + colDes.Item(rsBmap.Fields("des").Value)
				Else
					strText = txtOrigin.Text & " - unknown"
				End If
			End If
			
			lstPoints.Items.Add(strText)
		End If
	End Sub
End Class