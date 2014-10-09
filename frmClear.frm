VERSION 5.00
Begin VB.Form frmClear 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clear"
   ClientHeight    =   1545
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cbDays 
      Height          =   315
      ItemData        =   "frmClear.frx":0000
      Left            =   6360
      List            =   "frmClear.frx":003E
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame fraPostOptions 
      Caption         =   "Post Clear Commands"
      Height          =   1335
      Left            =   4320
      TabIndex        =   13
      Top             =   120
      Width           =   1815
      Begin VB.CheckBox chkCoastWatch 
         Caption         =   "coastwatch"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1215
      End
      Begin VB.CheckBox chkSpy 
         Caption         =   "spy *"
         Height          =   255
         Left            =   960
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox chkLook 
         Caption         =   "look *"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox chkLlook 
         Caption         =   "llook *"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Unit Options"
      Height          =   1335
      Left            =   2880
      TabIndex        =   9
      Top             =   120
      Width           =   1215
      Begin VB.CheckBox chkEnemy 
         Caption         =   "Enemy"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkNeutral 
         Caption         =   "Neutral"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkAllied 
         Caption         =   "Allied"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
   End
   Begin VB.CheckBox chkTelegrams 
      Caption         =   "Telegrams"
      Height          =   255
      Left            =   6360
      TabIndex        =   8
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame frmEnemy 
      Caption         =   "Enemy Intelligence"
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2535
      Begin VB.CheckBox chkAirCombat 
         Caption         =   "Air Combat"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkPlanes 
         Caption         =   "Planes"
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         Top             =   960
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkShips 
         Caption         =   "Ships"
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkLands 
         Caption         =   "Land Units"
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   600
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkSectors 
         Caption         =   "Sectors"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Value           =   1  'Checked
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Abort"
      Default         =   -1  'True
      Height          =   375
      Left            =   7680
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Clear"
      Height          =   375
      Left            =   7680
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblDays 
      Caption         =   "Clear records older then days"
      Height          =   375
      Left            =   6360
      TabIndex        =   18
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmClear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'112903 rjk: Created to give more control to the clear functions.
'120303 rjk: Added the ability to select by Allied/Neutral/Enemy as well
'            Added post clear command to refill the intelligence.
'120303 rjk: Moved the screen update to after the post clear commands, (display refreshed intelligence)
'300704 rjk: Added Days range to clear function.
'021104 rjk: Changed spy/llook/look/coastwatch default to off.

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If chkTelegrams <> vbUnchecked Then
    frmTelegram.ClearTelegrams
End If

ClearEnemyInfo (chkSectors <> vbUnchecked), (chkShips <> vbUnchecked), (chkLands <> vbUnchecked), _
(chkPlanes <> vbUnchecked), (chkAirCombat <> vbUnchecked), _
(chkAllied <> vbUnchecked), (chkNeutral <> vbUnchecked), (chkEnemy <> vbUnchecked), cbDays.ItemData(cbDays.ListIndex)

If chkSpy <> vbUnchecked Or chkLook <> vbUnchecked Or chkLlook <> vbUnchecked Or chkCoastWatch <> vbUnchecked Then
    frmEmpCmd.SubmitEmpireCommand "bf1", False
    If chkSpy <> vbUnchecked Then
        frmEmpCmd.SubmitEmpireCommand "spy *", False
    End If
    If chkLook <> vbUnchecked Then
        frmEmpCmd.SubmitEmpireCommand "look *", False
    End If
    If chkLlook <> vbUnchecked Then
        frmEmpCmd.SubmitEmpireCommand "llook *", False
    End If
    If chkCoastWatch <> vbUnchecked Then
        frmEmpCmd.SubmitEmpireCommand "coastwatch *", False
    End If
    frmEmpCmd.SubmitEmpireCommand "bf2", False
End If

If (chkSectors <> vbUnchecked) Or (chkShips <> vbUnchecked) Or (chkLands <> vbUnchecked) Or _
   (chkPlanes <> vbUnchecked) Or (chkAirCombat <> vbUnchecked) Then
    frmDrawMap.DrawHexes
End If
Unload Me
End Sub

Private Sub Form_Load()
cbDays.ListIndex = 0
End Sub
