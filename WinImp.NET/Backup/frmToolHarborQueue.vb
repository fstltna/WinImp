Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmToolHarborQueue
	Inherits System.Windows.Forms.Form
	
	'Change Log:
	'171003 rjk: Added Strength fields to Sector database
	'161103 rjk: Changed the name TxtSector to strSector, code cleanup.
	'201103 rjk: Added option to control strength updates
	'180304 rjk: if the build list is empty, do not attempt a build, saving generating a server problem.
	'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
	'270904 rjk: Fixed the crash if the build list is empty.
	'261105 rjk: Do not compute avail use the value from the server.
	'180206 rjk: Replace sdump, ldump, and pdump with GetShips, GetLandUnits and GetPlanes.
	'210306 rjk: Switched SendFullDumpCommand to GetSectors
	
	Dim BuildType As String
	Public strSector As String
	
	Private Sub cmdAdd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAdd.Click
		Dim X As Short
		Dim i As Short
		Static c As Short
		'check number if input
		If Len(txtNum.Text) > 0 Then
			X = Val(txtNum.Text)
			If X < 1 Then
				MsgBox("Number String not valid", MsgBoxStyle.Exclamation + MsgBoxStyle.OKOnly, "Entry Error")
				txtNum.Focus()
				Exit Sub
			End If
		Else
			X = 1
		End If
		
		For i = 1 To X
			c = c + 1
			List1.Items.Add(VB.Left(txtType.Text, 5) & "#" & CStr(c) & " (" & Trim(Mid(txtType.Text, 6)) & ")")
		Next i
		ComputeCosts()
		Label1.Visible = (List1.Items.Count > 0)
	End Sub
	
	Private Sub cmdBuild_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBuild.Click
		
		'Scan Build Queue
		Dim BuildCount As Short
		Dim BuildID As String
		Dim strID As String
		Dim i As Short
		For i = 1 To List1.Items.Count
			strID = VB.Left(VB6.GetItemString(List1, i - 1), InStr(VB6.GetItemString(List1, i - 1), "#") - 1)
			If strID = BuildID Then
				BuildCount = BuildCount + 1
			ElseIf Len(BuildID) = 0 Then 
				BuildID = strID
				BuildCount = 1
			Else
				BuildUnit(BuildID, BuildCount)
				BuildID = strID
				BuildCount = 1
			End If
		Next i
		
		'build the last one
		BuildUnit(BuildID, BuildCount)
		
		'database update
		frmEmpCmd.SubmitEmpireCommand("db1", False)
		If BuildType = "s" Then
			GetShips()
		ElseIf BuildType = "l" Then 
			GetLandUnits()
		Else
			GetPlanes()
		End If
		
		GetSectors()
		'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
		GetCurrentStrength(tsSectors)
		frmEmpCmd.SubmitEmpireCommand("db2", False)
		
		List1.Items.Clear()
		Me.Hide()
		
	End Sub
	Private Sub BuildUnit(ByRef BuildID As String, ByRef BuildCount As Short)
		Dim strCmd As String
		
		If BuildCount > 0 Then '180304 rjk: ensure the list is not empty, otherwise it generate a server problem.
			strCmd = "build " & BuildType & " " & strSector & " " & BuildID & " " & CStr(BuildCount)
			frmEmpCmd.SubmitEmpireCommand(strCmd, True)
		End If
	End Sub
	
	Private Sub cmdClear_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClear.Click
		List1.Items.Clear()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		Me.Hide()
	End Sub
	
	Public Sub LoadTypebox(ByRef UnitClass As String)
		
		txtType.Items.Clear()
		txtType.Text = vbNullString
		
		If Not (rsBuild.BOF And rsBuild.EOF) Then
			rsBuild.MoveFirst()
			While Not rsBuild.EOF
				If rsBuild.Fields("type").Value = UnitClass Then
					If TechLevel < 0 Or rsBuild.Fields("tech").Value <= TechLevel Then
						txtType.Items.Add(VB.Left(rsBuild.Fields("id").Value + "     ", 5) + rsBuild.Fields("desc").Value)
					End If
				End If
				rsBuild.MoveNext()
			End While
		End If
		
		If txtType.Items.Count > 0 Then
			txtType.SelectedIndex = 0
		End If
		BuildType = UnitClass
		ComputeCosts()
	End Sub
	
	
	'UPGRADE_WARNING: Form event frmToolHarborQueue.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmToolHarborQueue_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		Label1.Visible = (List1.Items.Count > 0)
	End Sub
	
	Private Sub List1_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles List1.DoubleClick
		List1.Items.RemoveAt(List1.SelectedIndex)
		ComputeCosts()
		Label1.Visible = (List1.Items.Count > 0)
	End Sub
	
	
	Public Sub ComputeCosts()
		
		Dim i As Short
		Dim Cost As Integer
		Dim lcm As Integer
		Dim hcm As Integer
		Dim avail As Integer
		Dim Mil As Integer
		Dim gun As Integer
		
		For i = 0 To List1.Items.Count - 1
			
			'get the build requirements
			rsBuild.Seek("=", BuildType, Trim(VB.Left(VB6.GetItemString(List1, i), 5)))
			If Not rsBuild.NoMatch Then
				'load costs
				Cost = Cost + rsBuild.Fields("cost").Value
				lcm = lcm + rsBuild.Fields("lcm").Value
				hcm = hcm + rsBuild.Fields("hcm").Value
				avail = avail + rsBuild.Fields("avail").Value
				Select Case BuildType
					Case "p"
						Mil = Mil + rsBuild.Fields("mil").Value
					Case "l"
						gun = gun + rsBuild.Fields("gun").Value
				End Select
			End If
		Next i
		
		If List1.Items.Count > 0 Then 'make sure something was built
			'Fill cost string
			Frame2.Text = CStr(List1.Items.Count) & " Units in Queue"
			Label2(1).Text = "Cost: $" & CStr(Cost)
			Label2(0).Text = vbNullString
			'Fill resource expended string
			If lcm > 0 Then
				Label2(0).Text = CStr(lcm) & " lcms, "
			End If
			If hcm > 0 Then
				Label2(0).Text = Label2(0).Text & CStr(hcm) & " hcms, "
			End If
			If avail > 0 Then
				Label2(0).Text = Label2(0).Text & CStr(avail) & " avail, "
			End If
			Select Case BuildType
				Case "p"
					If avail > 0 Then
						Label2(0).Text = Label2(0).Text & CStr(Mil) & " mil, "
					End If
				Case "l"
					If gun > 0 Then
						Label2(0).Text = Label2(0).Text & CStr(gun) & " guns, "
					End If
			End Select
			
			If Len(Label2(0).Text) > 0 Then
				Label2(0).Text = VB.Left(Label2(0).Text, Len(Label2(0).Text) - 2)
			End If
		Else
			Frame2.Text = "No Units in Queue"
			Label2(0).Text = vbNullString
			Label2(1).Text = vbNullString
		End If
		
	End Sub
End Class