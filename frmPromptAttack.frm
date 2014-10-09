VERSION 5.00
Begin VB.Form frmPromptAttack 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2355
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5130
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Caption         =   "Assaulting Units"
      Height          =   2295
      Left            =   3720
      TabIndex        =   33
      Top             =   0
      Width           =   1335
      Begin VB.ListBox lstUnits 
         Height          =   1425
         Left            =   120
         MultiSelect     =   1  'Simple
         TabIndex        =   34
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Move in"
      Height          =   2295
      Left            =   3720
      TabIndex        =   13
      Top             =   0
      Width           =   1335
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Caption         =   "Units"
         Height          =   975
         Left            =   120
         TabIndex        =   27
         Top             =   1200
         Width           =   1095
         Begin VB.TextBox txtUnit2 
            Height          =   285
            Left            =   0
            TabIndex        =   30
            Top             =   600
            Width           =   1095
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Y"
            Height          =   255
            Left            =   0
            TabIndex        =   29
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Command3 
            Caption         =   "N"
            Height          =   255
            Left            =   600
            TabIndex        =   28
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "Units"
            Height          =   255
            Left            =   360
            TabIndex        =   31
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.OptionButton Option8 
         Caption         =   "AS"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   720
         TabIndex        =   20
         Top             =   720
         Width           =   495
      End
      Begin VB.OptionButton Option7 
         Caption         =   "All"
         Height          =   255
         Left            =   720
         TabIndex        =   19
         Top             =   480
         Width           =   495
      End
      Begin VB.OptionButton Option5 
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   495
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Prompt"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   1200
         Y1              =   1080
         Y2              =   1080
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Attack Force"
      Height          =   2295
      Left            =   2280
      TabIndex        =   11
      Top             =   0
      Width           =   1335
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   720
         TabIndex        =   12
         Top             =   720
         Width           =   495
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Units"
         Height          =   975
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   1095
         Begin VB.CommandButton Command2 
            Caption         =   "N"
            Height          =   255
            Left            =   600
            TabIndex        =   26
            Top             =   240
            Width           =   495
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Y"
            Height          =   255
            Left            =   0
            TabIndex        =   25
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtUnit 
            Height          =   285
            Left            =   0
            TabIndex        =   24
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Units"
            Height          =   255
            Left            =   360
            TabIndex        =   23
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.OptionButton Option3 
         Caption         =   "AS"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         Caption         =   "All"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Prompt"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   1200
         Y1              =   1080
         Y2              =   1080
      End
   End
   Begin VB.Frame frmMain 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   2055
      Begin VB.ComboBox cbShips 
         Height          =   315
         Left            =   960
         TabIndex        =   32
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtOrigin 
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Must Fill"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "from ship"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Support"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      Begin VB.CheckBox Check1 
         Caption         =   "planes"
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   4
         Top             =   480
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "artillery"
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "ships"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "forts"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmPromptAttack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'081203 efj: Removed Dead Variables
'092403 rjk: Switched to UnitCharacteristicCheck to remove hard coding.
'            General reformatting. Clear to ship combo to prevent invalid
'            ships appearing in the list after the target was moved.
'            Switched to SetUnitType and DisplayFirstSectorPanel
'101703 rjk: Added Strength fields to Sector database
'112003 rjk: Added option to control strength updates
'120703 rjk: Changed hldindex to hldIndex.
'270104 rjk: St@r W@rs, must be on top of the sector to assault.
'270304 rjk: Retro, assault does not required land or semi-land capability
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
'180206 rjk: Replace ldump and sdump with GetLandUnits and GetShips.
'210306 rjk: Switched SendFullDumpCommand to GetSectors
'120108 rjk: Removed unused local variables iCountry and Des from txtOriginChange.
'260408 rjk: Added assault in the same sector for Heavy Metal II
'140708 rjk: Added GetLost to attack and assault.
'160209 rjk: Allow assault of the sector when in it (new in server 4.3.19).
'100410 rjk: Fixed Attack/Assault support to use y/n instead 1/0.

Dim Support As Boolean
Dim Assault As Boolean

Private Sub cbShips_Change()
LoadUnitBox
End Sub

Private Sub cbShips_Click()
LoadUnitBox
End Sub

Private Sub cmdCancel_Click()
frmDrawMap.DisplayFirstSectorPanel
Unload Me
End Sub

Public Sub cmdOK_Click()
Dim strCmd As String
Dim strCmd2 As String
Dim strSupport As String
' Dim strTemp As String
Dim X As Integer
' Dim sx As Integer
' Dim sy As Integer

strSupport = vbNullString

For X = 0 To 3
    If Check1(X).Value = vbChecked Then
        strSupport = strSupport + " y"
    Else
        strSupport = strSupport + " n"
    End If
Next X

If Label2.Caption = "Attack" Then
    strCmd = "attack " + txtOrigin.Text
Else
    strCmd = StringBetween(cbShips.Text, "#", "(")
    If Len(strCmd) = 0 Then
        strCmd = cbShips.Text
    End If
    strCmd = "assault " + txtOrigin.Text + " " + strCmd
End If

If strSupport <> " y y y y" Then
    strCmd = strCmd + strSupport
End If

'build header
strCmd2 = "at1 "
If Not Option1.Value Then
    If Option2.Value Then
        strCmd2 = strCmd2 + "/mil=29999;"
    ElseIf Option3.Value And Len(Text1.Text) > 0 Then
        strCmd2 = strCmd2 + "/mil=" + Trim$(Text1.Text) + ";"
    End If
    strCmd2 = strCmd2 + "/sec=" + txtOrigin.Text + ";"
End If

If Not Option4.Value Then
    If Option5.Value Then
        strCmd2 = strCmd2 + "/occ=0;"
    ElseIf Option7.Value Then
        strCmd2 = strCmd2 + "/occ=29999;"
    ElseIf Option8.Value And Len(Text2.Text) > 0 Then
        strCmd2 = strCmd2 + "/occ=" + Trim$(Text2.Text) + ";"
    End If
End If

'Attacking Units
If Len(txtUnit.Text) > 0 Then
    strCmd2 = strCmd2 + "/unt=" + txtUnit.Text + ";"
End If

'Occupying Units
If Len(txtUnit2.Text) > 0 Then
    strCmd2 = strCmd2 + "/ocu=" + txtUnit2.Text + ";"
End If

If Len(strCmd2) > 4 Then
    frmEmpCmd.SubmitEmpireCommand strCmd2, False
    frmEmpCmd.SubmitEmpireCommand strCmd, True
    frmEmpCmd.SubmitEmpireCommand "at2", False
Else
    frmEmpCmd.SubmitEmpireCommand strCmd, True
End If

'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
GetSectors
'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
GetCurrentStrength tsSectors
GetLandUnits
If Label2.Caption <> "Attack" Then
    GetShips
End If
GetLost
frmEmpCmd.SubmitEmpireCommand "db2", False

frmDrawMap.DisplayFirstSectorPanel
Unload Me
End Sub


Private Sub Command1_Click()
txtUnit.Text = "Y"
End Sub

Private Sub Command2_Click()
txtUnit.Text = "N"
End Sub

Private Sub Command3_Click()
txtUnit2.Text = "N"
End Sub

Private Sub Command4_Click()
txtUnit2.Text = "Y"
End Sub

Private Sub Form_Load()
    'Make Stay always on top
    ' Dim sucess As Long  removed 8/12/03 efj
    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
    Support = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub


Private Sub Frame1_Click()
Dim X As Integer

'Toggle support
Support = Not Support

If Support Then
    For X = 0 To 3
        Check1(X).Value = vbChecked
    Next X
Else
    For X = 0 To 3
        Check1(X).Value = 0
    Next X
End If

End Sub


Private Sub Label2_Change()
If Label2.Caption = "Attack" Then
    Frame3.Enabled = True
    Frame3.Visible = True
    Frame6.Visible = False
    Assault = False
Else
    Frame6.Visible = True
    Frame3.Visible = False
    Me.cbShips.Visible = True
    Me.Label1.Visible = True
    Assault = True
End If
End Sub

Private Sub lstUnits_Click()
Dim n As Integer
Dim str As String
str = vbNullString
For n = 0 To lstUnits.ListCount - 1
    If lstUnits.Selected(n) Then
        str = str + StringBetween(lstUnits.List(n), "#", "(") + "/"
    End If
Next n
If Len(str) > 0 Then
    str = Left$(str, Len(str) - 1)
End If
txtUnit.Text = str
End Sub

Private Sub Option3_Click()
If Option3.Value Then
    Text1.SetFocus
End If
End Sub

Private Sub Option8_Click()
If Option8.Value Then
    Text2.SetFocus
End If
End Sub

Private Sub Text1_Change()
If Me.Visible Then
    If Len(Text1.Text) > 0 Then
        Option3.Value = True
    Else
        Option1.Value = True
    End If
End If
End Sub

Private Sub Text2_Change()
If Len(Text1.Text) > 0 Then
    Option8.Value = True
Else
    Option4.Value = True
End If
End Sub

Private Sub txtOrigin_Change()
Dim iEnemyMil As Integer

iEnemyMil = EnemyMil(txtOrigin)

If Assault Then
    LoadCombo iEnemyMil
    If cbShips.ListCount > 0 Then
        cbShips.ListIndex = 0
    End If
Else
    frmDrawMap.SetUnitDisplay udLAND, TYPE_ALL, True, True, True, txtOrigin.Text
End If

If iEnemyMil > 0 Then
    Text1 = CStr(iEnemyMil * 1.5)
End If
End Sub

Private Sub LoadCombo(DefMil As Integer)

Dim sx As Integer
Dim sy As Integer
Dim ass_str As Long
Dim n As Integer
Dim i As Integer
Dim strAdd As String
Dim Done As Boolean
Dim hldIndex As Variant

If Not ParseSectors(sx, sy, txtOrigin.Text) Then
    Exit Sub
End If

'save and set land unit index
hldIndex = rsLand.Index
rsLand.Index = "location"

cbShips.Clear '092403 rjk: Remove any ship not valid anymore

If Not (rsShip.BOF And rsShip.EOF) Then
    rsShip.MoveFirst
End If
While Not rsShip.EOF
    
    'choose the ships that could assault the target sector
    '270104 rjk: St@r W@rs, must be on top of the sector to assault.
    If (SectorDistance(sx, sy, rsShip.Fields("x"), rsShip.Fields("y")) = 1 And Not frmOptions.bStarWars) Or _
       (sx = rsShip.Fields("x") And sy = rsShip.Fields("y") And frmOptions.bStarWars) Or _
       (sx = rsShip.Fields("x") And sy = rsShip.Fields("y") And frmOptions.bHeavyMetal) Or _
       (sx = rsShip.Fields("x") And sy = rsShip.Fields("y") And VersionCheck(4, 3, 19) <> VER_LESS) Then
        'compute the number of ship-board military who could assault
        ass_str = rsShip.Fields("mil") - 1

'092403 rjk: switched to UnitCharacteristicCheck to remove hard coding
'        Select Case rsShip.Fields("type")
'            Case "ls"
'                ass_str = ass_str
'            Case "tt", "frg"
'                ass_str = Int(ass_str / 4#)
'            Case Else
'                ass_str = Int(ass_str / 10#)
'        End Select
        If DefMil > 0 Then
            If Not frmDrawMap.UnitCharacteristicCheck(TYPE_S_LAND, rsShip.Fields("type")) Then
                If frmDrawMap.UnitCharacteristicCheck(TYPE_S_SEMI_LAND, rsShip.Fields("type")) Then
                    ass_str = Int(ass_str / 4#)
                Else
                    ass_str = Int(ass_str / 10#)
                End If
            End If
        End If

        'check for lands on the ship
        If Not (rsLand.EOF And rsLand.BOF) Then
            rsLand.MoveFirst
            rsLand.Seek "=", rsShip.Fields("x"), rsShip.Fields("y")
            Done = rsLand.NoMatch
        Else
            Done = True
        End If
        
        While Not rsLand.EOF And Not Done
            'make sure they are still in the hex
            If rsShip.Fields("x") = rsLand.Fields("x") And rsShip.Fields("y") = rsLand.Fields("y") Then
                'make sure they are on the ship
                If rsLand.Fields("ship") = rsShip("id") Then
                    'make sure they can assault
                    '092403 rjk: switched to UnitCharacteristicCheck to remove hard coding
                    'If InStr(" linf inf eng mar meng ", " " + rsLand.Fields("type") + " ") > 0 Then
                    If frmDrawMap.UnitCharacteristicCheck(TYPE_L_ASSAULT, rsLand.Fields("type")) Then
                        ass_str = ass_str + CInt(rsLand.Fields("mil") * (rsLand.Fields("eff") / 100#) * rsLand.Fields("att") * 0.5)
                    End If
                End If
                rsLand.MoveNext
            Else
                Done = True
            End If
        Wend
        
        If ass_str > 0 Then
            
            'do a simple insertion sort into the list.
            n = 0
            i = cbShips.ListCount
            While n <= cbShips.ListCount - 1
                If cbShips.ItemData(n) > ass_str Then
                    n = n + 1
                Else
                    i = n
                    n = cbShips.ListCount
                End If
            Wend
            'build the display sting
            strAdd = rsShip.Fields("type") + "#" + CStr(rsShip.Fields("id")) + " (" + CStr(ass_str) + ")"
            If i = cbShips.ListCount Then
                cbShips.AddItem strAdd
            Else
                cbShips.AddItem strAdd, i
            End If
            cbShips.ItemData(cbShips.NewIndex) = ass_str
        End If
    End If
    rsShip.MoveNext
Wend

'restore land index
rsLand.Index = hldIndex

End Sub

Private Sub LoadUnitBox()
Dim strShip As String
Dim hldIndex As Variant
Dim hldIndex2 As Variant
Dim Done As Boolean
Dim ass_str As Integer
Dim strEntry As String

If Len(cbShips.Text) = 0 Then
    Exit Sub
End If

'save and set land unit index
hldIndex = rsLand.Index
rsLand.Index = "location"
hldIndex2 = rsShip.Index
rsShip.Index = "PrimaryKey"

On Error GoTo cbShips_Change_exit


'get ship number
strShip = StringBetween(cbShips.Text, "#", "(")
If Len(strShip) = 0 Then
    strShip = cbShips.Text
End If


'Load unit box
lstUnits.Clear
rsShip.Seek "=", CInt(strShip)
If Not rsShip.NoMatch Then
    'check for lands on the ship
    rsLand.MoveFirst
    rsLand.Seek "=", rsShip.Fields("x"), rsShip.Fields("y")
    Done = rsLand.NoMatch
    
    While Not rsLand.EOF And Not Done
        'make sure they are still in the hex
        If rsShip.Fields("x") = rsLand.Fields("x") And rsShip.Fields("y") = rsLand.Fields("y") Then
            'make sure they are on the ship
            If rsLand.Fields("ship") = rsShip.Fields("id") Then
                'make sure they can assault
                '092403 rjk: switched to UnitCharacteristicCheck to remove hard coding
                'If InStr(" linf inf eng mar meng ", " " + rsLand.Fields("type") + " ") > 0 Then
                If frmDrawMap.UnitCharacteristicCheck(TYPE_L_ASSAULT, rsLand.Fields("type")) Then
                    ass_str = CInt(rsLand.Fields("mil") * (rsLand.Fields("eff") / 100#) * rsLand.Fields("att") * 0.5)
                    strEntry = rsLand.Fields("type") + "#" + CStr(rsLand.Fields("id")) + " (" + CStr(ass_str) + ")"
                    lstUnits.AddItem strEntry
                End If
            End If
            rsLand.MoveNext
        Else
            Done = True
        End If
    Wend

End If

cbShips_Change_exit:

'reset index
rsLand.Index = hldIndex
rsShip.Index = hldIndex2

End Sub
