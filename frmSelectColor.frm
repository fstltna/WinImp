VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmSelectColor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Empire Color"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset Colors"
      Height          =   495
      Left            =   1560
      TabIndex        =   47
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Apply"
      Height          =   495
      Left            =   1560
      TabIndex        =   46
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3120
      TabIndex        =   45
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   495
      Left            =   120
      TabIndex        =   44
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "To change colors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   41
      Top             =   4080
      Width           =   4335
      Begin VB.Label Label2 
         Caption         =   "Left click square to change background color"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label3 
         Caption         =   "Right click square to change text color"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   600
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Player Colors"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   39
         Left            =   3600
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   40
         Top             =   2280
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   38
         Left            =   3120
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   39
         Top             =   2280
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   37
         Left            =   2640
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   38
         Top             =   2280
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   36
         Left            =   2160
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   37
         Top             =   2280
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   35
         Left            =   1680
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   36
         Top             =   2280
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   34
         Left            =   1200
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   35
         Top             =   2280
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   33
         Left            =   720
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   34
         Top             =   2280
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   32
         Left            =   240
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   33
         Top             =   2280
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   31
         Left            =   3600
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   32
         Top             =   1800
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   30
         Left            =   3120
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   31
         Top             =   1800
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   29
         Left            =   2640
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   30
         Top             =   1800
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   28
         Left            =   2160
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   29
         Top             =   1800
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   27
         Left            =   1680
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   28
         Top             =   1800
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   26
         Left            =   1200
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   27
         Top             =   1800
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   25
         Left            =   720
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   26
         Top             =   1800
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   24
         Left            =   240
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   25
         Top             =   1800
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   23
         Left            =   3600
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   24
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   22
         Left            =   3120
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   23
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   21
         Left            =   2640
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   22
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   20
         Left            =   2160
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   21
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         FillStyle       =   0  'Solid
         FontTransparent =   0   'False
         Height          =   495
         Index           =   0
         Left            =   240
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   20
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   1
         Left            =   720
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   19
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   2
         Left            =   1200
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   18
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   3
         Left            =   1680
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   17
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   4
         Left            =   2160
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   16
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   5
         Left            =   2640
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   15
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   6
         Left            =   3120
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   14
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   7
         Left            =   3600
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         FontTransparent =   0   'False
         Height          =   495
         Index           =   8
         Left            =   240
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   12
         Top             =   840
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   9
         Left            =   720
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   11
         Top             =   840
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   10
         Left            =   1200
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   10
         Top             =   840
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   11
         Left            =   1680
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   9
         Top             =   840
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   12
         Left            =   2160
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   8
         Top             =   840
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   13
         Left            =   2640
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   7
         Top             =   840
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   14
         Left            =   3120
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   6
         Top             =   840
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   15
         Left            =   3600
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   5
         Top             =   840
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   16
         Left            =   240
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   4
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   17
         Left            =   720
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   3
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   18
         Left            =   1200
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   2
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         Height          =   495
         Index           =   19
         Left            =   1680
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   1
         Top             =   1320
         Width           =   495
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   4200
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
End
Attribute VB_Name = "frmSelectColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim HldC(1 To NUMBER_OF_PLAYER_COLORS) As Long
Dim HldT(1 To NUMBER_OF_PLAYER_COLORS) As Long

Private Sub cmdCancel_Click()
Dim n As Integer

'Restore Colors
For n = 1 To NUMBER_OF_PLAYER_COLORS
    PlayerColors(n) = HldC(n)
    PlayerText(n) = HldT(n)
Next n

Unload Me

End Sub

Private Sub cmdOK_Click()
Dim n As Integer

For n = 1 To NUMBER_OF_PLAYER_COLORS
    PlayerColors(n) = Picture3(n - 1).BackColor
    PlayerText(n) = Picture3(n - 1).ForeColor
Next n
frmDrawMap.DrawHexes
Unload Me
End Sub

Private Sub cmdReset_Click()
SetDefaultPlayerColors
FillColors
End Sub

Private Sub cmdTest_Click()

Dim n As Integer
For n = 1 To NUMBER_OF_PLAYER_COLORS
    PlayerColors(n) = Picture3(n - 1).BackColor
    PlayerText(n) = Picture3(n - 1).ForeColor
Next n
frmDrawMap.DrawHexes

End Sub

Private Sub Form_Load()
'Save Colors
Dim n As Integer
For n = 1 To NUMBER_OF_PLAYER_COLORS
    HldC(n) = PlayerColors(n)
    HldT(n) = PlayerText(n)
Next n

FillColors
'Center form on screen
Left = (Screen.Width - Width) \ 2
top = (Screen.Height - Height) \ 2

FillColors
End Sub

Private Sub Picture3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    cd1.Color = PlayerColors(Index + 1)
    cd1.ShowColor
    PlayerColors(Index + 1) = cd1.Color
Else
    cd1.Color = PlayerText(Index + 1)
    cd1.ShowColor
    PlayerText(Index + 1) = cd1.Color
End If
FillColors
End Sub

Private Sub FillColors()
Dim n As Integer
Dim ch As String

For n = 1 To NUMBER_OF_PLAYER_COLORS
    Picture3(n - 1).BackColor = PlayerColors(n)
Next n

For n = 1 To NUMBER_OF_PLAYER_COLORS
    Picture3(n - 1).AutoRedraw = True
    Picture3(n - 1).ForeColor = PlayerText(n)
    SetDrawingFont Picture3(n - 1)
    Picture3(n - 1).FontSize = 10
    ch = CStr(n)
    Picture3(n - 1).CurrentX = (Picture3(n - 1).ScaleWidth - Picture3(n - 1).TextWidth(ch)) / 2
    Picture3(n - 1).CurrentY = (Picture3(n - 1).ScaleHeight - Picture3(n - 1).TextHeight(ch)) / 2
    Picture3(n - 1).Print ch
Next n
End Sub

