VERSION 5.00
Begin VB.Form frmSatellitePath 
   Caption         =   "Satellite Path"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2670
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   2670
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox txtOrbitalStartLocation 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.ListBox lstPaths 
      Height          =   1980
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   2415
   End
   Begin VB.ComboBox cbSatellites 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Orbital Start Location"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "frmSatellitePath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log
'190604 rjk: Created
'010704 rjk: Fixed a satellite path calculation. Added the ability to plot unlaunched satellites.
'            Still a problem with limits of the sin function, rounding errors.


Private Sub cbSatellites_Change()
txtOrbitalStartLocation.Text = vbNullString
ComputeSatellitePath
End Sub

Private Sub cbSatellites_Click()
txtOrbitalStartLocation.Text = vbNullString
ComputeSatellitePath
End Sub

Private Sub cmdCancel_Click()
Unload Me
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
Dim X As Integer
Dim Y As Integer
Dim WORLD_X As Integer
Dim WORLD_Y As Integer
Dim sx As Integer
Dim sy As Integer
Dim SatX As Integer
Dim SatY As Integer
Const PI = 3.14159265358979
Const MAX_SATELLITE_UPDATES = 50
Dim strSat As String
Dim hldIndex As String
Dim strPaths(MAX_SATELLITE_UPDATES) As String
Dim strPath As String
Dim iIndex As Integer
Dim StartX As Integer
Dim StartY As Integer
Dim PreviousX As Integer
Dim PreviousY As Integer
Dim DeltaX As Integer
Dim DeltaY As Integer
Dim iStartIndex As Integer

WORLD_X = rsVersion.Fields("World X")
WORLD_Y = rsVersion.Fields("World Y")

lstPaths.Clear

If Not ParseSectors(StartX, StartY, txtOrbitalStartLocation.Text) Then
    lstPaths.AddItem "Invalid start location"
    Exit Sub
End If
strSat = StringBetween(cbSatellites.Text, "#", " ")
hldIndex = rsPlane.Index
rsPlane.Index = "PrimaryKey"
rsPlane.Seek "=", strSat
hldIndex = rsPlane.Index

iStartIndex = -1

If Not rsPlane.NoMatch Then
    SatX = rsPlane.Fields("X")
    SatY = rsPlane.Fields("Y")
    PreviousX = 0
    PreviousY = 0
    sx = StartX
    sy = StartY
    theta = 0#
    For iIndex = 0 To MAX_SATELLITE_UPDATES
        X = Fix(2 * theta * WORLD_X)
        Y = Fix(Sin(6 * PI * theta) * WORLD_Y / 4)
        DeltaX = X - PreviousX
        DeltaY = Y - PreviousY
        If ((sx + sy + DeltaX + DeltaY) Mod 2) <> 0 Then
            DeltaX = DeltaX + 1
        End If
        PreviousX = X
        PreviousY = Y
        OffsetSectorLocation sx, sy, DeltaX, DeltaY
        If SatX = sx And SatY = sy And rsPlane.Fields("laun") = "Y" Then
            iStartIndex = iIndex
        End If
        strPaths(iIndex) = Format$(CStr(sx), "\ 000;-000") + "," + Format$(CStr(sy), "\ 000;-000")
        theta = theta + 0.05
    Next iIndex
    If iStartIndex = -1 Then
        If rsPlane.Fields("laun") = "Y" Then
            lstPaths.AddItem "Could match current location"
            lstPaths.AddItem "with path from start origin"
            Exit Sub
        Else
            iStartIndex = 0
        End If
    End If
    For iIndex = iStartIndex To MAX_SATELLITE_UPDATES
        If iIndex = iStartIndex Then
            strPath = "Current      " + strPaths(iIndex)
        Else
            strPath = Format$(iIndex - iStartIndex, "\U\p\d\a\t\e\:00\ ") + strPaths(iIndex)
        End If
        lstPaths.AddItem strPath
    Next iIndex
End If
End Sub

Private Sub Form_Activate()
If cbSatellites.ListCount > 0 Then
    cbSatellites.ListIndex = 0
    Me.Visible = True
Else
    Me.Visible = False
    MsgBox "No Satellites Found"
    Unload Me
End If
End Sub

Private Sub Form_Load()
cbSatellites.Clear
If Not (rsPlane.EOF And rsPlane.BOF) Then
    rsPlane.MoveFirst
    Do
        If frmDrawMap.UnitCharacteristicCheck(TYPE_P_SATELLITE, rsPlane.Fields("type")) And _
            Not frmDrawMap.UnitCharacteristicCheck(TYPE_P_MISSILE, rsPlane.Fields("type")) Then
            cbSatellites.AddItem rsPlane.Fields("type") + " #" + CStr(rsPlane.Fields("id")) + " " + CStr(rsPlane.Fields("x")) + "," + CStr(rsPlane.Fields("y"))
        End If
        rsPlane.MoveNext
    Loop Until rsPlane.EOF
End If
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then
    Exit Sub
End If

If Me.Width <> 2790 Then
    Me.Width = 2790
End If
If (Me.Height - txtOrbitalStartLocation.Height - cbSatellites.Height - 1850) < 0 Then
    Me.Height = txtOrbitalStartLocation.Height + cbSatellites.Height + 1850
End If
lstPaths.Move lstPaths.Left, lstPaths.top, lstPaths.Width, Me.Height - txtOrbitalStartLocation.Height - cbSatellites.Height - 1100
'rtbText.Move 120, 120, Me.Width - 360, Me.Height - txtMessage.Height - cbAllies.Height - 850
'txtMessage.Move rtbText.Left, rtbText.top + rtbText.Height + 120, rtbText.Width, txtMessage.Height
'txtUser.Move txtUser.Left, Me.Height - 850
'cbAllies.Move cbAllies.Left, Me.Height - 850
'chkBeep.Move chkBeep.Left, Me.Height - 850
'chkTop.Move chkTop.Left, Me.Height - 850
'optAll.Move optAll.Left, Me.Height - 850
'optUser.Move optUser.Left, Me.Height - 850
End Sub

Private Sub lstPaths_Click()
Dim sx As Integer
Dim sy As Integer

If ParseSectors(sx, sy, Right$(lstPaths.Text, 9)) Then
    frmDrawMap.MoveMap sx, sy
End If
End Sub

Private Sub txtOrbitalStartLocation_Change()
ComputeSatellitePath
End Sub

Private Sub txtOrbitalStartLocation_Click()
ComputeSatellitePath
End Sub
