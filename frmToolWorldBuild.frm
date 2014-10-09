VERSION 5.00
Begin VB.Form frmToolWorldBuild 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2955
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8010
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   8010
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Copy/Paste"
      Height          =   2535
      Left            =   6000
      TabIndex        =   21
      Top             =   240
      Width           =   1815
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         Height          =   375
         Left            =   360
         TabIndex        =   27
         Top             =   2040
         Width           =   975
      End
      Begin VB.ComboBox cbCopy 
         Height          =   315
         ItemData        =   "frmToolWorldBuild.frx":0000
         Left            =   120
         List            =   "frmToolWorldBuild.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtOrigin 
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Destination Sector"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Source Sectors"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CheckBox chkDes 
      Caption         =   "Click to Designate"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   720
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Set Resources"
      Height          =   2535
      Left            =   3600
      TabIndex        =   8
      Top             =   240
      Width           =   2295
      Begin VB.TextBox txtDest 
         Height          =   285
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmdSetResource 
         Caption         =   "Set "
         Height          =   375
         Left            =   1200
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtThresh 
         Height          =   285
         Index           =   4
         Left            =   1680
         TabIndex        =   13
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtThresh 
         Height          =   285
         Index           =   0
         Left            =   600
         TabIndex        =   12
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtThresh 
         Height          =   285
         Index           =   1
         Left            =   600
         TabIndex        =   11
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtThresh 
         Height          =   285
         Index           =   2
         Left            =   600
         TabIndex        =   10
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtThresh 
         Height          =   285
         Index           =   3
         Left            =   1680
         TabIndex        =   9
         Top             =   1320
         Width           =   375
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   2160
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label6 
         Caption         =   "Sector"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblThresh 
         Caption         =   "Uran"
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   18
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblThresh 
         Caption         =   "Iron"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblThresh 
         Caption         =   "Gold"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lblThresh 
         Caption         =   "Fert"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblThresh 
         Caption         =   "Oil"
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   14
         Top             =   1320
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Mouse Buttons"
      Height          =   1695
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   3255
      Begin VB.CheckBox Check1 
         Caption         =   "Set Resources when you designate"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.ComboBox cbType 
         Height          =   315
         Index           =   1
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   840
         Width           =   2055
      End
      Begin VB.ComboBox cbType 
         Height          =   315
         Index           =   0
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Right"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Left "
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Done"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "World Builder"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmToolWorldBuild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'031003 rjk: In non-deity mode wilderness is not the list so left button is blank
'            but designates as headquarters, changed so blank does not designate.
'            needs more work.
'171003 rjk: Added Strength fields to Sector database
'201103 rjk: Added option to control strength updates
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
'210306 rjk: Switched SendFullDumpCommand to GetSectors

Dim strCodes As String
Dim strNewDes(0 To 1) As String

Private Sub cbType_Click(Index As Integer)

If cbType(Index).ListIndex < 0 Then
    Exit Sub
End If
strNewDes(Index) = Mid$(strCodes, cbType(Index).ItemData(cbType(Index).ListIndex), 1)

End Sub

Private Sub chkDes_Click()
If chkDes = vbChecked Then
    frmDrawMap.WorldBuilding = True
Else
    frmDrawMap.WorldBuilding = False
End If
End Sub

Private Sub cmdCancel_Click()
'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
GetSectors
'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
GetCurrentStrength tsSectors
frmEmpCmd.SubmitEmpireCommand "db2", False
Unload Me
End Sub

Private Sub cmdCopy_Click()

Dim sx As Integer
Dim sy As Integer
Dim sx1 As Integer
Dim sy1 As Integer
Dim sx2 As Integer
Dim sy2 As Integer
On Error Resume Next

If Not ParseSectors(sx, sy, txtPath.Text) Then
    MsgBox "Destination Invalid", vbExclamation + vbOKOnly, "Copy Error"
    Exit Sub
End If

If Not ParseSectorRange(sx1, sx2, sy1, sy2, txtOrigin.Text) Then
    MsgBox "Origin Invalid", vbExclamation + vbOKOnly, "Copy Error"
    Exit Sub
End If
CopySectorInfo sx1, sx2, sy1, sy2, sx, sy, cbCopy.ListIndex
End Sub

Private Sub cmdSetResource_Click()
SetResources

'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
GetSectors
'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
GetCurrentStrength tsSectors
frmEmpCmd.SubmitEmpireCommand "db2", False

End Sub

Private Sub Form_Load()
strCodes = LoadTypebox(cbType(0))
strCodes = LoadTypebox(cbType(1))
cbType(0).ListIndex = InStr(strCodes, "-") - 1
If cbType(0).ListIndex = -1 Then '100303 rjk: in non-deity mode wilderness is not the list so left button is blank
    strNewDes(0) = ""
End If
cbType(1).ListIndex = cbType(1).ListCount - 1
cbCopy.ListIndex = 0

'Make Stay always on top
' Dim success As Long    8/2003 efj  removed
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmDrawMap.WorldBuilding = False
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub


Private Sub txtOrigin_DblClick()
Load frmPromptSectors
frmPromptSectors.strSectors = txtOrigin.Text
frmPromptSectors.SetControls
frmPromptSectors.Caption = "Select Sectors"
frmPromptSectors.Left = frmDrawMap.Left + frmDrawMap.Width - frmPromptSectors.Width
frmPromptSectors.top = frmDrawMap.top + (frmDrawMap.Height - frmPromptSectors.Height) \ 2
frmPromptSectors.Show vbModeless
End Sub

Public Sub DesignateHex(sectx As Integer, secty As Integer, Button As Integer) '8/2003 efj  removed, 09/11/03 rjk added back used when selecting sector from map
Dim strCmdDes As String
Dim strDes As String

strCmdDes = "designate " + SectorString(sectx, secty) + " "
If Button = vbLeftButton Then
    strDes = strNewDes(0)
Else
    strDes = strNewDes(1)
End If

If strDes = "" Then '100303 rjk: If Left or Right button is blank then do not designate
    Exit Sub
End If

strCmdDes = strCmdDes + strDes

rsSectors.Seek "=", sectx, secty
If Not rsSectors.NoMatch Then
    rsSectors.Edit
    rsSectors.Fields("des") = strDes
    rsSectors.Update
End If

'now do the designation
frmEmpCmd.SubmitEmpireCommand strCmdDes, True

If Check1.Value = vbChecked Then
    txtDest.Text = SectorString(sectx, secty)
    SetResources
End If

End Sub


Private Sub SetResources()
Dim n As Integer
Dim strCmd As String

For n = 0 To 4
    If Len(txtThresh(n).Text) > 0 Then
        strCmd = Trim$("setresource " + LCase$(lblThresh(n)) + " " + txtDest.Text + " " + txtThresh(n).Text)
        frmEmpCmd.SubmitEmpireCommand strCmd, True
    End If
Next n

End Sub

