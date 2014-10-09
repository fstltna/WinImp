VERSION 5.00
Begin VB.Form frmPromptCopyGameDB 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2148
   ClientLeft      =   48
   ClientTop       =   48
   ClientWidth     =   8040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2148
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cbGameName 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2775
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   6960
      TabIndex        =   18
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Copy"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6960
      TabIndex        =   17
      Top             =   600
      Width           =   855
   End
   Begin VB.Frame frameType 
      Caption         =   "Items"
      Height          =   1935
      Left            =   3120
      TabIndex        =   5
      Top             =   120
      Width           =   3615
      Begin VB.CheckBox chkItems 
         Caption         =   "Nation"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   11
         Top             =   1540
         Width           =   1695
      End
      Begin VB.CheckBox chkItems 
         Caption         =   "Enemy Sectors"
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox chkItems 
         Caption         =   "Enemy Ships"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   13
         Top             =   500
         Width           =   1695
      End
      Begin VB.CheckBox chkItems 
         Caption         =   "Enemy Planes"
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   14
         Top             =   760
         Width           =   1695
      End
      Begin VB.CheckBox chkItems 
         Caption         =   "Enemy Land Units"
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   15
         Top             =   1020
         Width           =   1695
      End
      Begin VB.CheckBox chkItems 
         Caption         =   "Enemy Nukes"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   1800
         TabIndex        =   16
         Top             =   1280
         Width           =   1695
      End
      Begin VB.CheckBox chkItems 
         Caption         =   "Owner Sectors"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkItems 
         Caption         =   "Owner Ships"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   500
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkItems 
         Caption         =   "Owner Planes"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   8
         Top             =   760
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkItems 
         Caption         =   "Owner Land Units"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   9
         Top             =   1020
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkItems 
         Caption         =   "Owner Nukes"
         Enabled         =   0   'False
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   10
         Top             =   1280
         Width           =   1695
      End
   End
   Begin VB.Frame frameOffset 
      Caption         =   "Offset"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   2775
      Begin VB.TextBox txtOffsetX 
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtOffsetY 
         Height          =   285
         Left            =   2160
         TabIndex        =   4
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblOffsetX 
         Alignment       =   1  'Right Justify
         Caption         =   "X Offset"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblOffsetY 
         Alignment       =   1  'Right Justify
         Caption         =   "Y Offset"
         Height          =   255
         Left            =   1440
         TabIndex        =   20
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Copy Game Database"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.6
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Game Name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmPromptCopyGameDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
'112004 rjk: Fixed LandUnits
'030405 rjk: Fixed RestoreDummyProductionSector, had references to rsSectors.
'160705 rjk: Added ability to get OContent for Sea Sectors
'180206 rjk: Replace ndump, ldump, sdump, pdump with GetNukes, GetLandUnits, GetShips and GetPlanes.
'210306 rjk: Switched SendFullDumpCommand to GetSectors
'090607 rjk: Added Land Unit counts and Plane counts for Server 4.3.6.
'120108 rjk: Removed unused local variable strCmd from CopyGameInfo.
'            Removed unused local variables iSecX and iSecY from CreatePlanes, CreateLandUnits
'            and CreateShips.


Private Sub cbGameName_Change()
If Len(Trim$(cbGameName.Text)) = 0 Then
    cmdOK.Enabled = False
Else
    cmdOK.Enabled = True
End If
End Sub

Private Sub cbGameName_Click()
If Len(Trim$(cbGameName.Text)) = 0 Then
    cmdOK.Enabled = False
Else
    cmdOK.Enabled = True
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
CopyGameInfo
Unload Me
End Sub

Private Sub CopyGameInfo()
Dim iOffsetX As Integer
Dim iOffsetY As Integer
Dim dbEmpireCopy As Database
Dim DatabaseName As String

iOffsetX = Val(txtOffsetX)
iOffsetY = Val(txtOffsetY)

DatabaseName = Trim$(cbGameName.Text)

'Allow the user to specifiy the path
If InStr(DatabaseName, "\") = 0 Then
    DatabaseName = App.path + "\Games\" + DatabaseName + ".mdb"
End If

Set dbEmpireCopy = OpenDatabase(DatabaseName, False, False)

'Copy sectors and bmap
If chkItems(5).Value <> vbUnchecked Then
    CopySectors dbEmpireCopy, iOffsetX, iOffsetY
End If

If chkItems(10).Value <> vbUnchecked Then
    UpdateNation dbEmpireCopy, iOffsetX, iOffsetY
End If

'Copy Ships
If chkItems(6).Value <> vbUnchecked Then
    CreateHarbor
    CreateShips dbEmpireCopy, iOffsetX, iOffsetY
    CleanupShips
End If

'Copy Plane
If chkItems(7).Value <> vbUnchecked Then
    CreateAirport
    CreatePlanes dbEmpireCopy, iOffsetX, iOffsetY
    CleanupPlanes
End If

'Copy Land Units
If chkItems(8).Value <> vbUnchecked Then
    CreateHeadquarters
    CreateLandUnits dbEmpireCopy, iOffsetX, iOffsetY
    CleanupLandUnits
End If
    
If chkItems(6).Value <> vbUnchecked Or _
   chkItems(7).Value <> vbUnchecked Or _
   chkItems(8).Value <> vbUnchecked Or _
   chkItems(9).Value <> vbUnchecked Then
    RestoreDummyProductionSector dbEmpireCopy, iOffsetX, iOffsetY
End If
    
dbEmpireCopy.Close

'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
GetSectors
GetCurrentStrength tsSectors
GetOContent

If chkItems(6).Value <> vbUnchecked Then
    GetShips
End If
If chkItems(7).Value <> vbUnchecked Then
    GetLandUnits
End If
If chkItems(8).Value <> vbUnchecked Then
    GetPlanes
End If
If chkItems(9).Value <> vbUnchecked Then
    GetNukes
End If
frmEmpCmd.SubmitEmpireCommand "db2", False
End Sub

Private Sub Form_Activate()
Dim FileName As String   ' Walking filename variable.
Dim strDirectoryName As String

strDirectoryName = App.path + "\Games\*.mdb"

' Search through this directory and sum file sizes.
FileName = Dir(strDirectoryName, vbNormal Or vbArchive)
While Len(FileName) <> 0
   cbGameName.AddItem Left$(FileName, Len(FileName) - 4)
   FileName = Dir()  ' Get next file.
Wend
End Sub

Private Sub CopySectors(dbEmpireCopy As Database, iOffsetX As Integer, iOffsetY As Integer)
Dim rsB As Recordset
Dim rs As Recordset
Dim iSecX As Integer
Dim iSecY As Integer
Dim strCmd As String
Dim strDes As String
Dim strCom As String
Dim n As Integer

Set rs = dbEmpireCopy.OpenRecordset("Sectors")
rs.Index = "PrimaryKey"
If Not (rs.BOF And rs.EOF) Then
    rs.MoveFirst
End If
    
While Not rs.EOF
    CopySectorAndOwner rs, iOffsetX, iOffsetY
    rs.MoveNext
Wend

Set rsB = dbEmpireCopy.OpenRecordset("Bmap")
rsB.Index = "PrimaryKey"
If Not (rsB.BOF And rsB.EOF) Then
    rsB.MoveFirst
End If
    
While Not rsB.EOF
    iSecX = rsB.Fields("x")
    iSecY = rsB.Fields("y")
    rs.Seek "=", iSecX, iSecY
    If rs.NoMatch Then
        OffsetSectorLocation iSecX, iSecY, iOffsetX, iOffsetY
        strDes = rsB.Fields("des")
        strCmd = "des " + SectorString(iSecX, iSecY) + " " + strDes
        frmEmpCmd.SubmitEmpireCommand strCmd, False
    End If
    rsB.MoveNext
Wend

'wipe out commodities from sea sectors
strCom = "mil  gun  shellfood civ  uw   iron lcm  hcm  oil  dust pet  rad  bar  "
For n = 0 To 13
    strCmd = Trim$(Mid$(strCom, 5 * n + 1, 5))
    strCmd = "give " + strCmd + " * ?des=.&" + strCmd + ">0 -9999"
    frmEmpCmd.SubmitEmpireCommand strCmd, False
Next n

rs.Close
rsB.Close
End Sub

Private Sub CreateHarbor()
'Setup a dummy harbor
CreateDummyProductionSector "h"
'scuttle any ships that already exist
ScuttleUnits "ship"
End Sub

Private Sub CreateHeadquarters()
'Setup a dummy headquarters
CreateDummyProductionSector "!"
'scuttle any land units that already exist
ScuttleUnits "land"
End Sub

    
Private Sub CreateAirport()
'Setup a dummy airport
CreateDummyProductionSector "*"
'scuttle any planes that already exist
ScuttleUnits "plane"
End Sub

Private Sub CreateShips(dbEmpireCopy As Database, iOffsetX As Integer, iOffsetY As Integer)
Dim rs As Recordset

Set rs = dbEmpireCopy.OpenRecordset("ship")
rs.Index = "PrimaryKey"

If Not (rs.BOF And rs.EOF) Then
    If VersionCheck(4, 3, 0) <> VER_LESS Then
        ComputeUnitCountsForShips
    End If
    CreateUnits rs, "ship", "fb"
    AdjustShips rs, iOffsetX, iOffsetY
End If
rs.Close
End Sub

Private Sub CreateLandUnits(dbEmpireCopy As Database, iOffsetX As Integer, iOffsetY As Integer)
Dim rs As Recordset

Set rs = dbEmpireCopy.OpenRecordset("land")
rs.Index = "PrimaryKey"

If Not (rs.BOF And rs.EOF) Then
    If VersionCheck(4, 3, 0) <> VER_LESS Then
        ComputeUnitCountsForLandUnits
    End If
    CreateUnits rs, "land", "cav", "unit"
    AdjustLandUnits rs, iOffsetX, iOffsetY
End If
rs.Close
End Sub

Private Sub CreatePlanes(dbEmpireCopy As Database, iOffsetX As Integer, iOffsetY As Integer)
Dim rs As Recordset

Set rs = dbEmpireCopy.OpenRecordset("plane")
rs.Index = "PrimaryKey"

If Not (rs.BOF And rs.EOF) Then
    CreateUnits rs, "plane", "f1"
    AdjustPlanes rs, iOffsetX, iOffsetY
End If
End Sub

Private Sub AdjustShips(rs As Recordset, iOffsetX As Integer, iOffsetY As Integer)
Dim iSecX As Integer
Dim iSecY As Integer
Dim strCmd As String

rs.MoveFirst
While Not rs.EOF
    iSecX = rs.Fields("x")
    iSecY = rs.Fields("y")
    OffsetSectorLocation iSecX, iSecY, iOffsetX, iOffsetY
    strCmd = " L " + SectorString(iSecX, iSecY)
    If rs.Fields("tech") > 0 Then
        strCmd = strCmd + " T " + CStr(rs.Fields("tech"))
    End If
    If rs.Fields("mob") > 0 Then
        strCmd = strCmd + " M " + CStr(rs.Fields("mob"))
    End If
    If rs.Fields("land") > 0 Then
        strCmd = strCmd + " Y " + CStr(rs.Fields("land"))
    End If
    If rs.Fields("eff") > 0 Then
        strCmd = strCmd + " E " + CStr(rs.Fields("eff"))
    End If
    If rs.Fields("pln") > 0 Then
        strCmd = strCmd + " P " + CStr(rs.Fields("pln"))
    End If
    If rs.Fields("xl") > 0 Then
        strCmd = strCmd + " X " + CStr(rs.Fields("xl"))
    End If
    If rs.Fields("he") > 0 Then
        strCmd = strCmd + " H " + CStr(rs.Fields("he"))
    End If
    If rs.Fields("mil") > 0 Then
        strCmd = strCmd + " m " + CStr(rs.Fields("mil"))
    End If
    If rs.Fields("civ") > 0 Then
        strCmd = strCmd + " c " + CStr(rs.Fields("civ"))
    End If
    If rs.Fields("uw") > 0 Then
        strCmd = strCmd + " u " + CStr(rs.Fields("uw"))
    End If
    If rs.Fields("food") > 0 Then
        strCmd = strCmd + " f " + CStr(rs.Fields("food"))
    End If
    If rs.Fields("gun") > 0 Then
        strCmd = strCmd + " g " + CStr(rs.Fields("gun"))
    End If
    If rs.Fields("shell") > 0 Then
        strCmd = strCmd + " s " + CStr(rs.Fields("shell"))
    End If
    If rs.Fields("iron") > 0 Then
        strCmd = strCmd + " i " + CStr(rs.Fields("iron"))
    End If
    If rs.Fields("dust") > 0 Then
        strCmd = strCmd + " d " + CStr(rs.Fields("dust"))
    End If
    If rs.Fields("oil") > 0 Then
        strCmd = strCmd + " o " + CStr(rs.Fields("oil"))
    End If
    If rs.Fields("petrol") > 0 Then
        strCmd = strCmd + " p " + CStr(rs.Fields("petrol"))
    End If
    If rs.Fields("lcm") > 0 Then
        strCmd = strCmd + " l " + CStr(rs.Fields("lcm"))
    End If
    If rs.Fields("hcm") > 0 Then
        strCmd = strCmd + " h " + CStr(rs.Fields("hcm"))
    End If
    If rs.Fields("rad") > 0 Then
        strCmd = strCmd + " r " + CStr(rs.Fields("rad"))
    End If
    strCmd = "edit ship " + CStr(rs.Fields("id")) + strCmd
    frmEmpCmd.SubmitEmpireCommand strCmd, False
    rs.MoveNext
Wend
End Sub

Private Sub AdjustLandUnits(rs As Recordset, iOffsetX As Integer, iOffsetY As Integer)
Dim iSecX As Integer
Dim iSecY As Integer
Dim strCmd As String
    
rs.MoveFirst
    
While Not rs.EOF
    iSecX = rs.Fields("x")
    iSecY = rs.Fields("y")
    OffsetSectorLocation iSecX, iSecY, iOffsetX, iOffsetY
    strCmd = " L " + SectorString(iSecX, iSecY)
    strCmd = strCmd + " t " + CStr(rs.Fields("tech"))
    strCmd = strCmd + " M " + CStr(rs.Fields("mob"))
    strCmd = strCmd + " e " + CStr(rs.Fields("eff"))
    strCmd = strCmd + " F " + CStr(rs.Fields("fort"))
    strCmd = strCmd + " B " + CStr(rs.Fields("fuel"))
    strCmd = strCmd + " S " + CStr(rs.Fields("ship"))
    strCmd = strCmd + " Y " + CStr(rs.Fields("land"))
    strCmd = strCmd + " X " + CStr(rs.Fields("xl"))
    strCmd = strCmd + " m " + CStr(rs.Fields("mil"))
    strCmd = strCmd + " f " + CStr(rs.Fields("food"))
    strCmd = strCmd + " g " + CStr(rs.Fields("gun"))
    strCmd = strCmd + " s " + CStr(rs.Fields("shell"))
    strCmd = strCmd + " i " + CStr(rs.Fields("iron"))
    strCmd = strCmd + " d " + CStr(rs.Fields("dust"))
    strCmd = strCmd + " o " + CStr(rs.Fields("oil"))
    strCmd = strCmd + " p " + CStr(rs.Fields("petrol"))
    strCmd = strCmd + " l " + CStr(rs.Fields("lcm"))
    strCmd = strCmd + " h " + CStr(rs.Fields("hcm"))
    strCmd = strCmd + " r " + CStr(rs.Fields("rad"))
    strCmd = "edit unit " + CStr(rs.Fields("id")) + strCmd
    frmEmpCmd.SubmitEmpireCommand strCmd, False
    rs.MoveNext
Wend
End Sub

Private Sub AdjustPlanes(rs As Recordset, iOffsetX As Integer, iOffsetY As Integer)
Dim iSecX As Integer
Dim iSecY As Integer
Dim strCmd As String

rs.MoveFirst
While Not rs.EOF
    iSecX = rs.Fields("x")
    iSecY = rs.Fields("y")
    OffsetSectorLocation iSecX, iSecY, iOffsetX, iOffsetY
    strCmd = " l " + SectorString(iSecX, iSecY)
    strCmd = strCmd + " t " + CStr(rs.Fields("tech"))
    strCmd = strCmd + " m " + CStr(rs.Fields("mob"))
    strCmd = strCmd + " e " + CStr(rs.Fields("eff"))
    strCmd = strCmd + " a " + CStr(rs.Fields("att"))
    strCmd = strCmd + " d " + CStr(rs.Fields("def"))
    strCmd = strCmd + " r " + CStr(rs.Fields("range"))
    strCmd = strCmd + " s " + CStr(rs.Fields("ship"))
    strCmd = strCmd + " l " + CStr(rs.Fields("land"))
    strCmd = "edit plane " + CStr(rs.Fields("id")) + strCmd
    frmEmpCmd.SubmitEmpireCommand strCmd, False
    rs.MoveNext
Wend
End Sub

Private Sub CleanupShips()
'now scrap the filler Ships
frmEmpCmd.SubmitEmpireCommand "sc1", False
frmEmpCmd.SubmitEmpireCommand "scuttle ship * ?type=fb&eff=20&xloc=0&yloc=0&own=98", False
End Sub

Private Sub CleanupLandUnits()
'now scrap the filler Land Units
frmEmpCmd.SubmitEmpireCommand "sc1", False
frmEmpCmd.SubmitEmpireCommand "scuttle land * ?type=cav&eff=10&xloc=0&yloc=0&own=98", False
End Sub

Private Sub CleanupPlanes()
'now scrap the filler Planes
frmEmpCmd.SubmitEmpireCommand "sc1", False
frmEmpCmd.SubmitEmpireCommand "scuttle plane * ?type=f1&eff=10&xloc=0&yloc=0&own=98", False
End Sub

Private Sub CreateDummyProductionSector(strDes As String)
Dim strCmd As String

strCmd = "designate 0,0 " + strDes
frmEmpCmd.SubmitEmpireCommand strCmd, False
strCmd = "setsector eff 0,0 100"
frmEmpCmd.SubmitEmpireCommand strCmd, False
strCmd = "setsector work 0,0 100"
frmEmpCmd.SubmitEmpireCommand strCmd, False
strCmd = "setsector oldown 0,0 0"
frmEmpCmd.SubmitEmpireCommand strCmd, False
strCmd = "setsector own 0,0 0"
frmEmpCmd.SubmitEmpireCommand strCmd, False
End Sub

Private Sub RestoreDummyProductionSector(dbEmpireCopy As Database, iOffsetX As Integer, iOffsetY As Integer)
Dim rs As Recordset
Dim strCmd As String
Dim strCom As String
Dim strComm As String
Dim iIndex As Integer
Dim iSecDistX As Integer
Dim iSecDistY As Integer

Set rs = dbEmpireCopy.OpenRecordset("Sectors")
rs.Index = "PrimaryKey"
rs.Seek "=", -iOffsetX, -iOffsetY
If rs.NoMatch Then
    strCmd = "designate 0,0 ."
    frmEmpCmd.SubmitEmpireCommand strCmd, False
Else
    strCmd = "setsector oldown 0,0 " + CStr(rs.Fields("Country"))
    frmEmpCmd.SubmitEmpireCommand strCmd, False
    strCmd = "setsector own 0,0 " + CStr(rs.Fields("Country"))
    frmEmpCmd.SubmitEmpireCommand strCmd, False
    strCmd = "setsector mob 0,0 " + CStr(rs.Fields("mob"))
    frmEmpCmd.SubmitEmpireCommand strCmd, False
    strCmd = "setsector work 0,0 " + CStr(rs.Fields("work"))
    frmEmpCmd.SubmitEmpireCommand strCmd, False
    strCmd = "setsector avail 0,0 " + CStr(rs.Fields("avail"))
    frmEmpCmd.SubmitEmpireCommand strCmd, False
    frmEmpCmd.SubmitEmpireCommand "designate 0,0 " + rs.Fields("des"), False
    strCom = "mil  gun  shelllcm  hcm  "
    For iIndex = 0 To 4
        strComm = Trim$(Mid$(strCom, 5 * iIndex + 1, 5))
        strCmd = "give " + strComm + " 0,0 ?" + strComm + ">0 -9999"
        frmEmpCmd.SubmitEmpireCommand strCmd, True
        If rs.Fields(strComm) > 0 Then
            strCmd = "give " + strComm + " 0,0 " + CStr(rs.Fields(strComm))
            frmEmpCmd.SubmitEmpireCommand strCmd, True
        End If
    Next iIndex
    strCom = "eff  work avail"
    For iIndex = 0 To 2
        strComm = Trim$(Mid$(strCom, 5 * iIndex + 1, 5))
        strCmd = "setsector " + strComm + " 0,0 -9999"
        frmEmpCmd.SubmitEmpireCommand strCmd, False
        If rs.Fields(strComm) > 0 Then
            strCmd = "setsector " + strComm + " 0,0 " + CStr(rs.Fields(strComm))
            frmEmpCmd.SubmitEmpireCommand strCmd, False
        End If
    Next iIndex
    iSecDistX = rs.Fields("dist_x")
    iSecDistY = rs.Fields("dist_y")
    strCmd = "edit land 0,0 D " + SectorString(iSecDistX, iSecDistY)
    frmEmpCmd.SubmitEmpireCommand strCmd, False
    If rs.Fields("off") = 0 Then
        frmEmpCmd.SubmitEmpireCommand "start 0,0", False
    Else
        frmEmpCmd.SubmitEmpireCommand "stop 0,0", False
    End If
    SetThreshold "0,0", "c", rs
    SetThreshold "0,0", "m", rs
    SetThreshold "0,0", "u", rs
    SetThreshold "0,0", "s", rs
    SetThreshold "0,0", "g", rs
    SetThreshold "0,0", "l", rs
    SetThreshold "0,0", "h", rs
    SetThreshold "0,0", "o", rs
    SetThreshold "0,0", "p", rs
    SetThreshold "0,0", "f", rs
    SetThreshold "0,0", "d", rs
    SetThreshold "0,0", "b", rs
    SetThreshold "0,0", "i", rs
    SetThreshold "0,0", "r", rs
    strCmd = "terr 0,0 " + CStr(rs.Fields("terr")) + " 0"
    frmEmpCmd.SubmitEmpireCommand strCmd, False
    strCmd = "terr 0,0 " + CStr(rs.Fields("terr1")) + " 1"
    frmEmpCmd.SubmitEmpireCommand strCmd, False
    strCmd = "terr 0,0 " + CStr(rs.Fields("terr2")) + " 2"
    frmEmpCmd.SubmitEmpireCommand strCmd, False
    strCmd = "terr 0,0 " + CStr(rs.Fields("terr3")) + " 3"
    frmEmpCmd.SubmitEmpireCommand strCmd, False
End If
rs.Close
End Sub

Private Sub UpdateNation(dbEmpireCopy As Database, iOffsetX As Integer, iOffsetY As Integer)
Dim rs As Recordset
Dim strCmd As String
Dim CountryNumber As Integer
Dim iSecX As Integer
Dim iSecY As Integer

Set rs = dbEmpireCopy.OpenRecordset("Nations")
If Not (rs.BOF And rs.EOF) Then
    rs.MoveFirst
    CountryNumber = Val(rs.Fields("Number"))
    rs.Close
Else
    rs.Close
    Exit Sub
End If

Set rs = dbEmpireCopy.OpenRecordset("Nation")
If Not (rs.BOF And rs.EOF) Then
    rs.MoveFirst
    strCmd = "add " + CStr(CountryNumber) + " " + CStr(CountryNumber) + " " + CStr(CountryNumber) + " new i"
    frmEmpCmd.SubmitEmpireCommand strCmd, False
    strCmd = "edit country " + CStr(CountryNumber)
    strCmd = strCmd + " E " + CStr(rs.Fields("Education"))
    strCmd = strCmd + " T " + CStr(rs.Fields("TechLevel"))
    strCmd = strCmd + " H " + CStr(rs.Fields("Happiness"))
    strCmd = strCmd + " R " + CStr(rs.Fields("Research"))
    If ParseSectors(iSecX, iSecY, rs.Fields("CapSect")) Then
        OffsetSectorLocation iSecX, iSecY, iOffsetX, iOffsetY
        strCmd = strCmd + " c " + CStr(iSecX) + "," + CStr(iSecY)
    End If
        
    strCmd = strCmd + " m " + CStr(rs.Fields("MilReserves"))
    strCmd = strCmd + " o " + CStr(iOffsetX) + "," + CStr(iOffsetY)
    strCmd = strCmd + " s 5 "
    strCmd = strCmd + " M 25000 "
    strCmd = strCmd + " b 640 "
    frmEmpCmd.SubmitEmpireCommand strCmd, False
End If
rs.Close
End Sub

Private Sub CreateUnits(rs As Recordset, strUnit As String, strDefaultUnit As String, Optional strEdit As String)
Dim strCmd As String
Dim nUnit As Integer
Dim iDefaultMil As Integer
Dim iDefaultGuns As Integer
Dim iDefaultAvail As Integer
Dim iDefaultHcms As Integer
Dim iDefaultLcms As Integer

If Len(strEdit) = 0 Then
    strEdit = strUnit
End If

If strUnit = "unit" Then
    rsBuild.Seek "=", "l", strDefaultUnit
Else
    rsBuild.Seek "=", Left$(strUnit, 1), strDefaultUnit
End If
If rsBuild.NoMatch Then
    EmpireError "CreateUnits", "Can not find default unit characteristics", strDefaultUnit
    Exit Sub
End If
iDefaultMil = rsBuild.Fields("mil")
iDefaultGuns = rsBuild.Fields("gun")
iDefaultAvail = rsBuild.Fields("avail")
iDefaultHcms = rsBuild.Fields("hcm")
iDefaultLcms = rsBuild.Fields("lcm")

rs.MoveFirst

nUnit = 0
While Not rs.EOF
    'need a filler Unit
    If nUnit < rs.Fields("id") Then
        SupplySector iDefaultMil, iDefaultGuns, iDefaultAvail, iDefaultHcms, iDefaultLcms
        strCmd = "build " + strUnit + " 0,0 " + strDefaultUnit
        frmEmpCmd.SubmitEmpireCommand strCmd, False
        strCmd = "edit " + strEdit + " " + CStr(nUnit) + " O 98"
        frmEmpCmd.SubmitEmpireCommand strCmd, False
    'need to build the Unit
    ElseIf nUnit = rs.Fields("id") Then
        If strUnit = "unit" Then
            rsBuild.Seek "=", "l", rs.Fields("type")
        Else
            rsBuild.Seek "=", Left$(strUnit, 1), rs.Fields("type")
        End If
        If rsBuild.NoMatch Then
            EmpireError "CreateUnits", "Can not find unit characteristics", rs.Fields("type")
            Exit Sub
        End If
        SupplySector CInt(rsBuild.Fields("mil")), CInt(rsBuild.Fields("gun")), CInt(rsBuild.Fields("avail")), _
            CInt(rsBuild.Fields("hcm")), CInt(rsBuild.Fields("lcm"))
        strCmd = "build " + strUnit + " 0,0 " + rs.Fields("type") + " 1 " + CStr(rs.Fields("tech"))
        frmEmpCmd.SubmitEmpireCommand strCmd, False
        strCmd = "edit " + strEdit + " " + CStr(nUnit) + " O " + CStr(rs.Fields("country"))
        frmEmpCmd.SubmitEmpireCommand strCmd, False
        rs.MoveNext
    End If
    nUnit = nUnit + 1
Wend
End Sub

Private Sub SupplySector(iMil As Integer, iGuns As Integer, iAvail As Integer, iHcms As Integer, iLcms As Integer)
If iMil > 0 Then
    frmEmpCmd.SubmitEmpireCommand "give mil 0,0 " + CStr(iMil), False
End If
If iGuns > 0 Then
    frmEmpCmd.SubmitEmpireCommand "give gun 0,0 " + CStr(iGuns), False
End If
If iHcms > 0 Then
    frmEmpCmd.SubmitEmpireCommand "give hcm 0,0 " + CStr(iHcms), False
End If
If iLcms > 0 Then
    frmEmpCmd.SubmitEmpireCommand "give lcm 0,0 " + CStr(iLcms), False
End If
If iAvail > 0 Then
    frmEmpCmd.SubmitEmpireCommand "setsector avail 0,0 " + CStr(iAvail), False
End If
End Sub

Private Sub ScuttleUnits(strUnit As String)
Dim strCmd As String

strCmd = "sc1"
frmEmpCmd.SubmitEmpireCommand strCmd, False

strCmd = "scuttle " + strUnit + " *"
frmEmpCmd.SubmitEmpireCommand strCmd, False
End Sub

Private Sub CopySectorAndOwner(rs As Recordset, iOffsetX As Integer, iOffsetY As Integer)
Dim vhld() As Variant
Dim iSecX As Integer
Dim iSecY As Integer
Dim iSecDistX As Integer
Dim iSecDistY As Integer
Dim strCmd As String
Dim strSector As String
Dim n As Integer
Dim secowner As Integer
    
ReDim vhld(0 To rs.Fields.Count)
   
For n = 0 To rs.Fields.Count - 1
    vhld(rs.Fields(n).OrdinalPosition) = rs.Fields(n).Value
Next n
iSecX = rs.Fields("x")
iSecY = rs.Fields("y")
OffsetSectorLocation iSecX, iSecY, iOffsetX, iOffsetY
CopySector vhld, iSecX, iSecY, 3

strSector = SectorString(iSecX, iSecY)
secowner = rs.Fields("country")
strCmd = "setsector own " + strSector + " " + CStr(secowner)
frmEmpCmd.SubmitEmpireCommand strCmd, False
'    secowner = rs.Fields("oldown")
strCmd = "setsector oldown " + strSector + " " + CStr(secowner)
frmEmpCmd.SubmitEmpireCommand strCmd, False
iSecDistX = rs.Fields("dist_x")
iSecDistY = rs.Fields("dist_y")
OffsetSectorLocation iSecDistX, iSecDistY, iOffsetX, iOffsetY
strCmd = "edit land " + strSector + " D " + SectorString(iSecDistX, iSecDistY)
frmEmpCmd.SubmitEmpireCommand strCmd, False
If rs.Fields("off") = 0 Then
    frmEmpCmd.SubmitEmpireCommand "start " + strSector, False
Else
    frmEmpCmd.SubmitEmpireCommand "stop " + strSector, False
End If
SetThreshold strSector, "c", rs
SetThreshold strSector, "m", rs
SetThreshold strSector, "u", rs
SetThreshold strSector, "s", rs
SetThreshold strSector, "g", rs
SetThreshold strSector, "l", rs
SetThreshold strSector, "h", rs
SetThreshold strSector, "o", rs
SetThreshold strSector, "p", rs
SetThreshold strSector, "f", rs
SetThreshold strSector, "d", rs
SetThreshold strSector, "b", rs
SetThreshold strSector, "i", rs
SetThreshold strSector, "r", rs
strCmd = "terr " + strSector + " " + CStr(rs.Fields("terr")) + " 0"
frmEmpCmd.SubmitEmpireCommand strCmd, False
strCmd = "terr " + strSector + " " + CStr(rs.Fields("terr1")) + " 1"
frmEmpCmd.SubmitEmpireCommand strCmd, False
strCmd = "terr " + strSector + " " + CStr(rs.Fields("terr2")) + " 2"
frmEmpCmd.SubmitEmpireCommand strCmd, False
strCmd = "terr " + strSector + " " + CStr(rs.Fields("terr3")) + " 3"
frmEmpCmd.SubmitEmpireCommand strCmd, False
End Sub

Private Sub SetThreshold(strSector As String, strComm As String, rs As Recordset)
Dim strCmd As String

strCmd = "threshold " + strComm + " " + strSector + " " + CStr(rs.Fields(strComm + "_dist"))
frmEmpCmd.SubmitEmpireCommand strCmd, False
If rs.Fields(strComm + "_del") = "." Then
    strCmd = "deliver " + strComm + " " + strSector + " " + CStr(rs.Fields(strComm + "_cut")) + " h"
Else
    strCmd = "deliver " + strComm + " " + strSector + " " + CStr(rs.Fields(strComm + "_cut")) + " " + CStr(rs.Fields(strComm + "_del"))
End If
frmEmpCmd.SubmitEmpireCommand strCmd, False
End Sub
