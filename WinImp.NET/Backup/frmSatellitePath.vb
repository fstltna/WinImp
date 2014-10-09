Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmSatellitePath
	Inherits System.Windows.Forms.Form
	
	'Change Log
	'190604 rjk: Created
	'010704 rjk: Fixed a satellite path calculation. Added the ability to plot unlaunched satellites.
	'            Still a problem with limits of the sin function, rounding errors.
	
	
	'UPGRADE_WARNING: Event cbSatellites.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	'UPGRADE_WARNING: ComboBox event cbSatellites.Change was upgraded to cbSatellites.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
	Private Sub cbSatellites_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbSatellites.TextChanged
		txtOrbitalStartLocation.Text = vbNullString
		ComputeSatellitePath()
	End Sub
	
	'UPGRADE_WARNING: Event cbSatellites.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cbSatellites_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cbSatellites.SelectedIndexChanged
		txtOrbitalStartLocation.Text = vbNullString
		ComputeSatellitePath()
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	'    newtheta = pp->pln_theta + .05;
	'
	'    if (newtheta >= 1.0) {
	'       newtheta -= 1.0;
	'    }
	'
	'    x1 = (coord)(2 * pp->pln_theta * WORLD_X);
	'    x1 = xnorm(x1);
	'    y1 = (coord)(sin(6 * PI * pp->pln_theta) * (WORLD_Y / 4));
	'    x2 = (coord)(2 * newtheta * WORLD_X);
	'    x2 = xnorm(x2);
	'    y2 = (coord)(sin(6 * PI * newtheta) * (WORLD_Y / 4));
	'    dx = x1 - pp->pln_x;
	'    dy = y1 - pp->pln_y;
	'    x2 -= dx;
	'    y2 -= dy;
	'
	'    if ((x2 + y2) & 1) {
	'       x2++;
	'    }
	'
	'    pp->pln_x = xnorm(x2);
	'    pp->pln_y = ynorm(y2);
	'    pp->pln_theta = newtheta;
	'    getsect(pp->pln_x, pp->pln_y, &sect);
	'    if (sect.sct_own)
	'       if (pp->pln_own != sect.sct_own)
	'           wu(0, sect.sct_own, "%s satellite spotted over %s\n",
	'              cname(pp->pln_own), xyas(pp->pln_x, pp->pln_y,sect.sct_own));
	'    return;
	Private Sub ComputeSatellitePath()
		Dim theta As Single
		Dim X As Short
		Dim Y As Short
		Dim WORLD_X As Short
		Dim WORLD_Y As Short
		Dim sx As Short
		Dim sy As Short
		Dim SatX As Short
		Dim SatY As Short
		Const PI As Double = 3.14159265358979
		Const MAX_SATELLITE_UPDATES As Short = 50
		Dim strSat As String
		Dim hldIndex As String
		Dim strPaths(MAX_SATELLITE_UPDATES) As String
		Dim strPath As String
		Dim iIndex As Short
		Dim StartX As Short
		Dim StartY As Short
		Dim PreviousX As Short
		Dim PreviousY As Short
		Dim DeltaX As Short
		Dim DeltaY As Short
		Dim iStartIndex As Short
		
		WORLD_X = rsVersion.Fields("World X").Value
		WORLD_Y = rsVersion.Fields("World Y").Value
		
		lstPaths.Items.Clear()
		
		If Not ParseSectors(StartX, StartY, (txtOrbitalStartLocation.Text)) Then
			lstPaths.Items.Add("Invalid start location")
			Exit Sub
		End If
		strSat = StringBetween((cbSatellites.Text), "#", " ")
		hldIndex = rsPlane.Index
		rsPlane.Index = "PrimaryKey"
		rsPlane.Seek("=", strSat)
		hldIndex = rsPlane.Index
		
		iStartIndex = -1
		
		If Not rsPlane.NoMatch Then
			SatX = rsPlane.Fields("X").Value
			SatY = rsPlane.Fields("Y").Value
			PreviousX = 0
			PreviousY = 0
			sx = StartX
			sy = StartY
			theta = 0#
			For iIndex = 0 To MAX_SATELLITE_UPDATES
				X = Fix(2 * theta * WORLD_X)
				Y = Fix(System.Math.Sin(6 * PI * theta) * WORLD_Y / 4)
				DeltaX = X - PreviousX
				DeltaY = Y - PreviousY
				If ((sx + sy + DeltaX + DeltaY) Mod 2) <> 0 Then
					DeltaX = DeltaX + 1
				End If
				PreviousX = X
				PreviousY = Y
				OffsetSectorLocation(sx, sy, DeltaX, DeltaY)
				If SatX = sx And SatY = sy And rsPlane.Fields("laun").Value = "Y" Then
					iStartIndex = iIndex
				End If
				strPaths(iIndex) = VB6.Format(CStr(sx), "\ 000;-000") & "," & VB6.Format(CStr(sy), "\ 000;-000")
				theta = theta + 0.05
			Next iIndex
			If iStartIndex = -1 Then
				If rsPlane.Fields("laun").Value = "Y" Then
					lstPaths.Items.Add("Could match current location")
					lstPaths.Items.Add("with path from start origin")
					Exit Sub
				Else
					iStartIndex = 0
				End If
			End If
			For iIndex = iStartIndex To MAX_SATELLITE_UPDATES
				If iIndex = iStartIndex Then
					strPath = "Current      " & strPaths(iIndex)
				Else
					strPath = VB6.Format(iIndex - iStartIndex, "\U\p\d\a\t\e\:00\ ") & strPaths(iIndex)
				End If
				lstPaths.Items.Add(strPath)
			Next iIndex
		End If
	End Sub
	
	'UPGRADE_WARNING: Form event frmSatellitePath.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmSatellitePath_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		If cbSatellites.Items.Count > 0 Then
			cbSatellites.SelectedIndex = 0
			Me.Visible = True
		Else
			Me.Visible = False
			MsgBox("No Satellites Found")
			Me.Close()
		End If
	End Sub
	
	Private Sub frmSatellitePath_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		cbSatellites.Items.Clear()
		If Not (rsPlane.EOF And rsPlane.BOF) Then
			rsPlane.MoveFirst()
			Do 
				If frmDrawMap.UnitCharacteristicCheck(frmDrawMap.enumUnitType.TYPE_P_SATELLITE, rsPlane.Fields("type").Value) And Not frmDrawMap.UnitCharacteristicCheck(frmDrawMap.enumUnitType.TYPE_P_MISSILE, rsPlane.Fields("type").Value) Then
					cbSatellites.Items.Add(rsPlane.Fields("type").Value + " #" + CStr(rsPlane.Fields("id").Value) + " " + CStr(rsPlane.Fields("x").Value) + "," + CStr(rsPlane.Fields("y").Value))
				End If
				rsPlane.MoveNext()
			Loop Until rsPlane.EOF
		End If
	End Sub
	
	'UPGRADE_WARNING: Event frmSatellitePath.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmSatellitePath_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		If Me.WindowState = System.Windows.Forms.FormWindowState.Minimized Then
			Exit Sub
		End If
		
		If VB6.PixelsToTwipsX(Me.Width) <> 2790 Then
			Me.Width = VB6.TwipsToPixelsX(2790)
		End If
		If (VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(txtOrbitalStartLocation.Height) - VB6.PixelsToTwipsY(cbSatellites.Height) - 1850) < 0 Then
			Me.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(txtOrbitalStartLocation.Height) + VB6.PixelsToTwipsY(cbSatellites.Height) + 1850)
		End If
		lstPaths.SetBounds(lstPaths.Left, lstPaths.Top, lstPaths.Width, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - VB6.PixelsToTwipsY(txtOrbitalStartLocation.Height) - VB6.PixelsToTwipsY(cbSatellites.Height) - 1100))
		'rtbText.Move 120, 120, Me.Width - 360, Me.Height - txtMessage.Height - cbAllies.Height - 850
		'txtMessage.Move rtbText.Left, rtbText.top + rtbText.Height + 120, rtbText.Width, txtMessage.Height
		'txtUser.Move txtUser.Left, Me.Height - 850
		'cbAllies.Move cbAllies.Left, Me.Height - 850
		'chkBeep.Move chkBeep.Left, Me.Height - 850
		'chkTop.Move chkTop.Left, Me.Height - 850
		'optAll.Move optAll.Left, Me.Height - 850
		'optUser.Move optUser.Left, Me.Height - 850
	End Sub
	
	'UPGRADE_WARNING: Event lstPaths.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstPaths_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstPaths.SelectedIndexChanged
		Dim sx As Short
		Dim sy As Short
		
		If ParseSectors(sx, sy, VB.Right(lstPaths.Text, 9)) Then
			frmDrawMap.MoveMap(sx, sy)
		End If
	End Sub
	
	'UPGRADE_WARNING: Event txtOrbitalStartLocation.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtOrbitalStartLocation_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOrbitalStartLocation.TextChanged
		ComputeSatellitePath()
	End Sub
	
	Private Sub txtOrbitalStartLocation_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOrbitalStartLocation.Click
		ComputeSatellitePath()
	End Sub
End Class