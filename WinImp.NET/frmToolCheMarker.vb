Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmToolCheMarker
	Inherits System.Windows.Forms.Form
	
	'070604 rjk: Added Plague Marking
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Dim strSource As String
		Dim strTemp As String
		Dim n As Integer
		
		'loop through the telegram and mark che events
		strSource = Text1.Text
		Do 
			n = InStr(strSource, vbLf)
			If n = 0 Then n = Len(strSource) + 1
			strTemp = VB.Left(strSource, n - 1)
			strSource = Mid(strSource, n + 1)
			EventChe(strTemp)
		Loop Until Len(strSource) = 0
		
		Me.Close()
	End Sub
	
	Private Sub cmdPlague_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPlague.Click
		Dim strSource As String
		Dim strTemp As String
		Dim n As Integer
		
		'loop through the telegram and mark che events
		strSource = Text1.Text
		Do 
			n = InStr(strSource, vbLf)
			If n = 0 Then n = Len(strSource) + 1
			strTemp = VB.Left(strSource, n - 1)
			strSource = Mid(strSource, n + 1)
			EventPlague(strTemp)
		Loop Until Len(strSource) = 0
		
		Me.Close()
	End Sub
	Private Sub frmToolCheMarker_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		LoadList()
		If List1.Items.Count > 0 Then
			List1.SelectedIndex = 0
		End If
	End Sub
	
	Private Sub LoadList()
		
		Dim hldindex As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object hldindex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		hldindex = rsTeleHead.Index
		rsTeleHead.Index = "Received"
		If Not (rsTeleHead.BOF And rsTeleHead.EOF) Then
			rsTeleHead.MoveLast()
			While Not rsTeleHead.BOF
				If rsTeleHead.Fields("Class").Value = TELEGRAM_CLASS_PRODUCTION Then
					List1.Items.Add(New VB6.ListBoxItem(rsTeleHead.Fields("Subject").Value, rsTeleHead.Fields("ID").Value))
				End If
				rsTeleHead.MovePrevious()
			End While
			rsTeleHead.MoveLast()
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object hldindex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsTeleHead.Index = hldindex
		
	End Sub
	
	'UPGRADE_WARNING: Event List1.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub List1_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles List1.SelectedIndexChanged
		
		'Load Text box with Telegram
		Dim teleID As Short
		teleID = VB6.GetItemData(List1, List1.SelectedIndex)
		rsTeleBody.Index = "PrimaryKey"
		rsTeleBody.Seek("=", teleID, 1)
		Text1.Text = vbNullString
		'Load telegram body into text box
		If Not rsTeleBody.NoMatch Then
			Do While (Not rsTeleBody.EOF)
				If rsTeleBody.Fields("ID").Value <> teleID Then
					Exit Do
				End If
				Text1.Text = Text1.Text + rsTeleBody.Fields("Body").Value + vbLf
				rsTeleBody.MoveNext()
			Loop 
		End If
		
	End Sub
End Class