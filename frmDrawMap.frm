VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDrawMap 
   Caption         =   "Map"
   ClientHeight    =   7065
   ClientLeft      =   165
   ClientTop       =   1035
   ClientWidth     =   10695
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDrawMap.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer PlayersTimer 
      Left            =   600
      Top             =   1320
   End
   Begin MSComctlLib.Toolbar tbMain 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   179
      Top             =   0
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   794
      ButtonWidth     =   656
      ButtonHeight    =   635
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ImageTB2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   28
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SetColors"
            Object.ToolTipText     =   "Set Colors"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "AutoUpdate"
            Object.ToolTipText     =   "Auto Update is on"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "AutoRead"
            Object.ToolTipText     =   "AutoRead is on"
            ImageIndex      =   24
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "UpdateSector"
            Object.ToolTipText     =   "Update Sector Info Now"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "UpdateUnits"
            Object.ToolTipText     =   "Update Unit Info Now"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "UpdateBmap"
            Object.ToolTipText     =   "Update Bmap Now"
            ImageIndex      =   27
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Budget"
            Object.ToolTipText     =   "Run Budget Report"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Coastwatch"
            Object.ToolTipText     =   "Run Coastwatch Report"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Nation"
            Object.ToolTipText     =   "Run Nation Report"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Newspaper"
            Object.ToolTipText     =   "Get Today's Paper"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PowerNew"
            Object.ToolTipText     =   "Get a New Power Report"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Report"
            Object.ToolTipText     =   "Get the Tech Report"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Starvation"
            Object.ToolTipText     =   "Check Starvation"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ReportShow"
            Object.ToolTipText     =   "Show Report"
            ImageIndex      =   26
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Version"
            Object.ToolTipText     =   "Show the Version Report"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Wire"
            Object.ToolTipText     =   "Read today's wires"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Declare"
            Object.ToolTipText     =   "Declare Relations"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Relations"
            Object.ToolTipText     =   "Check Relations"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Script"
            Object.ToolTipText     =   "Script Tool"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Build"
            Object.ToolTipText     =   "Build Projection Tool"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Famine"
            Object.ToolTipText     =   "Famine Relief Tool"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NationLevels"
            Object.ToolTipText     =   "Nation Levels Tool"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Chi"
            Object.ToolTipText     =   "Che/Plague Marker Tool"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture4 
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   1995
      TabIndex        =   4
      Top             =   3480
      Width           =   2055
      Begin RichTextLib.RichTextBox rtbTelegram 
         Height          =   735
         Left            =   0
         TabIndex        =   8
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1296
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmDrawMap.frx":030A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   5775
      Left            =   2520
      ScaleHeight     =   5715
      ScaleWidth      =   8955
      TabIndex        =   3
      Top             =   480
      Width           =   9015
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Index           =   2
         Left            =   240
         TabIndex        =   13
         Top             =   1200
         Width           =   3975
         Begin VB.Label lblDes 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   182
            Top             =   0
            Width           =   3255
         End
         Begin VB.Label Label1 
            Caption         =   "Food"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   76
            Left            =   480
            TabIndex        =   75
            Top             =   2280
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Shell"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   77
            Left            =   480
            TabIndex        =   74
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Guns"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   78
            Left            =   480
            TabIndex        =   73
            Top             =   2760
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   103
            Left            =   0
            LinkTimeout     =   61
            TabIndex        =   72
            Top             =   2280
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   104
            Left            =   0
            LinkTimeout     =   62
            TabIndex        =   71
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   105
            Left            =   0
            TabIndex        =   70
            Top             =   2760
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Bars"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   82
            Left            =   1440
            TabIndex        =   69
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Oil"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   83
            Left            =   1440
            TabIndex        =   68
            Top             =   2280
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Lcms"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   84
            Left            =   1440
            TabIndex        =   67
            Top             =   2520
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   109
            Left            =   960
            TabIndex        =   66
            Top             =   2040
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   110
            Left            =   960
            TabIndex        =   65
            Top             =   2280
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   111
            Left            =   960
            TabIndex        =   64
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   113
            Left            =   960
            TabIndex        =   63
            Top             =   3000
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   112
            Left            =   960
            TabIndex        =   62
            Top             =   2760
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Rads"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   90
            Left            =   1440
            TabIndex        =   61
            Top             =   3000
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Hcms"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   91
            Left            =   1440
            TabIndex        =   60
            Top             =   2760
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   108
            Left            =   960
            TabIndex        =   59
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   107
            Left            =   960
            TabIndex        =   58
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   106
            Left            =   0
            TabIndex        =   57
            Top             =   3000
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Dust"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   95
            Left            =   1440
            TabIndex        =   56
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Iron"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   96
            Left            =   1440
            TabIndex        =   55
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Pet"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   97
            Left            =   480
            TabIndex        =   54
            Top             =   3000
            Width           =   375
         End
         Begin VB.Label lblDes 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   53
            Top             =   300
            Width           =   3255
         End
         Begin VB.Label Label1 
            Caption         =   "Uw"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   94
            Left            =   480
            TabIndex        =   52
            Top             =   2040
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   102
            Left            =   0
            TabIndex        =   51
            Top             =   2040
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   101
            Left            =   0
            TabIndex        =   50
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   100
            Left            =   0
            LinkTimeout     =   62
            TabIndex        =   49
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Mil"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   80
            Left            =   480
            TabIndex        =   48
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Civ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   99
            Left            =   480
            TabIndex        =   47
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "Thresholds"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   46
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Distribution Sect:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   45
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lblDes 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   44
            Top             =   900
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Food"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   88
            Left            =   2520
            TabIndex        =   43
            Top             =   2280
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Shell"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   87
            Left            =   2520
            TabIndex        =   42
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Guns"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   86
            Left            =   2520
            TabIndex        =   41
            Top             =   2760
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999:h"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   123
            Left            =   2000
            LinkTimeout     =   61
            TabIndex        =   40
            Top             =   2280
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999:h"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   124
            Left            =   2000
            LinkTimeout     =   62
            TabIndex        =   39
            Top             =   2520
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999:h"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   125
            Left            =   2000
            TabIndex        =   38
            Top             =   2760
            Width           =   480
         End
         Begin VB.Label Label1 
            Caption         =   "Bars"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   75
            Left            =   3480
            TabIndex        =   37
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Oil"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   74
            Left            =   3480
            TabIndex        =   36
            Top             =   2280
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Lcms"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   73
            Left            =   3480
            TabIndex        =   35
            Top             =   2520
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999:h"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   129
            Left            =   2960
            TabIndex        =   34
            Top             =   2040
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999:h"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   130
            Left            =   2960
            TabIndex        =   33
            Top             =   2280
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999:h"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   131
            Left            =   2960
            TabIndex        =   32
            Top             =   2520
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999:h"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   133
            Left            =   2960
            TabIndex        =   31
            Top             =   3000
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999:h"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   132
            Left            =   2960
            TabIndex        =   30
            Top             =   2760
            Width           =   480
         End
         Begin VB.Label Label1 
            Caption         =   "Rads"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   56
            Left            =   3480
            TabIndex        =   29
            Top             =   3000
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Hcms"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   52
            Left            =   3480
            TabIndex        =   28
            Top             =   2760
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999:h"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   128
            Left            =   2960
            TabIndex        =   27
            Top             =   1800
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999:h"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   127
            Left            =   2960
            TabIndex        =   26
            Top             =   1560
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999:h"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   126
            Left            =   2000
            TabIndex        =   25
            Top             =   3000
            Width           =   480
         End
         Begin VB.Label Label1 
            Caption         =   "Dust"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   45
            Left            =   3480
            TabIndex        =   24
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Iron"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   44
            Left            =   3480
            TabIndex        =   23
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Pet"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   43
            Left            =   2520
            TabIndex        =   22
            Top             =   3000
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Uw"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   42
            Left            =   2520
            TabIndex        =   21
            Top             =   2040
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999:h"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   122
            Left            =   2000
            TabIndex        =   20
            Top             =   2040
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999:h"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   121
            Left            =   2000
            TabIndex        =   19
            Top             =   1800
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999:h"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   120
            Left            =   2000
            LinkTimeout     =   62
            TabIndex        =   18
            Top             =   1560
            Width           =   480
         End
         Begin VB.Label Label1 
            Caption         =   "Mil"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   36
            Left            =   2520
            TabIndex        =   17
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Civ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   33
            Left            =   2520
            TabIndex        =   16
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Cutoff/Direction"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   2040
            TabIndex        =   15
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Delivery"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   1920
            TabIndex        =   14
            Top             =   600
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Index           =   3
         Left            =   720
         TabIndex        =   171
         Top             =   960
         Visible         =   0   'False
         Width           =   4095
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   276
            Left            =   0
            TabIndex        =   172
            Top             =   0
            Width           =   8400
            _ExtentX        =   14817
            _ExtentY        =   476
            ButtonWidth     =   572
            ButtonHeight    =   402
            AllowCustomize  =   0   'False
            Appearance      =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   9
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Ship"
                  Object.ToolTipText     =   "Ships"
                  ImageIndex      =   4
                  Style           =   2
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Land"
                  Object.ToolTipText     =   "Lands"
                  ImageIndex      =   2
                  Style           =   2
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Plane"
                  Object.ToolTipText     =   "Planes"
                  ImageIndex      =   3
                  Style           =   2
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Nuke"
                  Object.ToolTipText     =   "Nukes"
                  ImageIndex      =   10
                  Style           =   2
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Enemy"
                  Object.ToolTipText     =   "Enemy Units"
                  ImageIndex      =   1
                  Style           =   2
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "List"
                  Object.ToolTipText     =   "Summary List"
                  ImageIndex      =   9
                  Style           =   2
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Mob"
                  Object.ToolTipText     =   "Only Show if Mobility >0"
                  ImageIndex      =   7
                  Style           =   1
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Eff"
                  Object.ToolTipText     =   "Only Show if at least Min. Eff."
                  ImageIndex      =   5
                  Style           =   1
               EndProperty
               BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Delete"
                  Object.ToolTipText     =   "Delete Unit from Database"
                  Object.Tag             =   "Delete"
                  ImageIndex      =   6
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   120
            Top             =   120
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   20
            ImageHeight     =   13
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   10
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDrawMap.frx":0381
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDrawMap.frx":08D7
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDrawMap.frx":0E2D
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDrawMap.frx":1383
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDrawMap.frx":18D9
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDrawMap.frx":1BF3
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDrawMap.frx":1D05
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDrawMap.frx":2397
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDrawMap.frx":26B1
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmDrawMap.frx":29CB
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin RichTextLib.RichTextBox rtbUnitList 
            Height          =   375
            Left            =   3960
            TabIndex        =   180
            Top             =   2040
            Visible         =   0   'False
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393217
            ScrollBars      =   3
            TextRTF         =   $"frmDrawMap.frx":2CE5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid grid1 
            Height          =   2535
            Left            =   0
            TabIndex        =   174
            Top             =   720
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   4471
            _Version        =   393216
            GridLines       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   1200
            TabIndex        =   181
            Top             =   360
            Width           =   2895
            Begin VB.ComboBox cmbUnitFilter 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               ItemData        =   "frmDrawMap.frx":2D65
               Left            =   0
               List            =   "frmDrawMap.frx":2D67
               Style           =   2  'Dropdown List
               TabIndex        =   188
               Top             =   0
               Width           =   2895
            End
         End
         Begin VB.ComboBox cmbSub 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "frmDrawMap.frx":2D69
            Left            =   0
            List            =   "frmDrawMap.frx":2D6B
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   173
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Index           =   4
         Left            =   120
         TabIndex        =   175
         Top             =   1000
         Width           =   4095
         Begin VB.Label SeaFrameLabel 
            Caption         =   "Sea"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   187
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   146
            Left            =   720
            TabIndex        =   186
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Fert:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   145
            Left            =   360
            TabIndex        =   185
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   143
            Left            =   720
            TabIndex        =   184
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Oil:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   144
            Left            =   360
            TabIndex        =   183
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Loading..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   176
            Top             =   1320
            Width           =   2535
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Index           =   1
         Left            =   3600
         TabIndex        =   76
         Top             =   360
         Width           =   4455
         Begin VB.Line Line1 
            Index           =   5
            X1              =   2280
            X2              =   4440
            Y1              =   3240
            Y2              =   3240
         End
         Begin VB.Label lblPlayers 
            Caption         =   "12 45 78 90 23 56 89 12 45 78 01 34 67 90 23 56 78"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2400
            TabIndex        =   207
            Top             =   3240
            Width           =   1815
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            Caption         =   "Ter3:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   163
            Left            =   3720
            TabIndex        =   206
            Top             =   2640
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   164
            Left            =   4200
            TabIndex        =   205
            Top             =   2640
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Ter2:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   161
            Left            =   3720
            TabIndex        =   204
            Top             =   2400
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   162
            Left            =   4200
            TabIndex        =   203
            Top             =   2400
            Width           =   255
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   160
            Left            =   4200
            TabIndex        =   118
            Top             =   2160
            Width           =   255
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   25
            Left            =   1200
            TabIndex        =   201
            Top             =   2520
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Terr:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   28
            Left            =   840
            TabIndex        =   202
            Top             =   2520
            Width           =   375
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   3600
            X2              =   4440
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "2.25"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   154
            Left            =   4080
            TabIndex        =   195
            Top             =   840
            Width           =   375
         End
         Begin VB.Line Line2 
            Index           =   1
            X1              =   3600
            X2              =   3600
            Y1              =   600
            Y2              =   3240
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   158
            Left            =   4080
            TabIndex        =   199
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   152
            Left            =   4080
            TabIndex        =   193
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   156
            Left            =   4080
            TabIndex        =   197
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   150
            Left            =   4080
            TabIndex        =   191
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   148
            Left            =   4080
            TabIndex        =   189
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "TD:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   157
            Left            =   3720
            TabIndex        =   200
            ToolTipText     =   "Total defense - (sector defense) + (reacting units)"
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "SD:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   155
            Left            =   3720
            TabIndex        =   198
            ToolTipText     =   "Sector defense - (mil + units) * mult"
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "SM:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   153
            Left            =   3720
            TabIndex        =   196
            ToolTipText     =   "The defensive multiplier of the sector including the land mine bonus"
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Lm:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   151
            Left            =   3720
            TabIndex        =   194
            ToolTipText     =   "Land Mines"
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "RU:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   149
            Left            =   3720
            TabIndex        =   192
            ToolTipText     =   "The total strengths of all supplied mobile reacting units in range"
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Un:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   147
            Left            =   3720
            TabIndex        =   190
            ToolTipText     =   "The total defensive strength of units in the sector"
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   58
            Left            =   1920
            TabIndex        =   79
            ToolTipText     =   "Food Required"
            Top             =   1560
            Width           =   255
         End
         Begin VB.Label lblWarn 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   110
            Top             =   3000
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Civ:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   143
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Mil:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   142
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Uw:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   141
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   140
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   480
            TabIndex        =   139
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   480
            TabIndex        =   138
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Mob:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   1080
            TabIndex        =   137
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Work:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   1080
            TabIndex        =   136
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Avail:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   1080
            TabIndex        =   135
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   1560
            TabIndex        =   134
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   1560
            TabIndex        =   133
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   1560
            TabIndex        =   132
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label lblDes 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   131
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   12
            Left            =   1200
            TabIndex        =   130
            Top             =   2040
            Width           =   255
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   13
            Left            =   1200
            TabIndex        =   129
            Top             =   1800
            Width           =   255
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   14
            Left            =   1200
            TabIndex        =   128
            Top             =   1560
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Def:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   15
            Left            =   840
            TabIndex        =   127
            Top             =   2040
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Rail:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   16
            Left            =   840
            TabIndex        =   126
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Rd:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   17
            Left            =   840
            TabIndex        =   125
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   18
            Left            =   480
            TabIndex        =   124
            Top             =   2040
            Width           =   255
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   19
            Left            =   480
            TabIndex        =   123
            Top             =   1800
            Width           =   255
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   20
            Left            =   480
            TabIndex        =   122
            Top             =   1560
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Fert:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   21
            Left            =   120
            TabIndex        =   121
            Top             =   2040
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Gld:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   22
            Left            =   120
            TabIndex        =   120
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Min:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   23
            Left            =   120
            TabIndex        =   119
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   26
            Left            =   1200
            TabIndex        =   117
            Top             =   2280
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Ter1:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   159
            Left            =   3720
            TabIndex        =   116
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Fall:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   29
            Left            =   840
            TabIndex        =   115
            Top             =   2280
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   31
            Left            =   480
            TabIndex        =   114
            Top             =   2520
            Width           =   255
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   32
            Left            =   480
            TabIndex        =   113
            Top             =   2280
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Urn:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   34
            Left            =   120
            TabIndex        =   112
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Oil:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   35
            Left            =   120
            TabIndex        =   111
            Top             =   2280
            Width           =   375
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   0
            X2              =   2280
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Line Line1 
            Index           =   3
            X1              =   0
            X2              =   2280
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Label lblNDes 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   109
            Top             =   0
            Width           =   3255
         End
         Begin VB.Label Label1 
            Caption         =   "Food"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   55
            Left            =   2880
            TabIndex        =   108
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Shell"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   54
            Left            =   2880
            TabIndex        =   107
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Guns"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   53
            Left            =   2880
            TabIndex        =   106
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   60
            Left            =   2400
            LinkTimeout     =   61
            TabIndex        =   105
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   61
            Left            =   2400
            LinkTimeout     =   62
            TabIndex        =   104
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   62
            Left            =   2400
            TabIndex        =   103
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Bars"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   49
            Left            =   2880
            TabIndex        =   102
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Oil"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   48
            Left            =   2880
            TabIndex        =   101
            Top             =   2280
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Lcms"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   47
            Left            =   2880
            TabIndex        =   100
            Top             =   2520
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   66
            Left            =   2400
            TabIndex        =   99
            Top             =   2040
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   67
            Left            =   2400
            TabIndex        =   98
            Top             =   2280
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   68
            Left            =   2400
            TabIndex        =   97
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   70
            Left            =   2400
            TabIndex        =   96
            Top             =   3000
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   69
            Left            =   2400
            TabIndex        =   95
            Top             =   2760
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Rads"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   39
            Left            =   2880
            TabIndex        =   94
            Top             =   3000
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Hcms"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   38
            Left            =   2880
            TabIndex        =   93
            Top             =   2760
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   65
            Left            =   2400
            TabIndex        =   92
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   64
            Left            =   2400
            TabIndex        =   91
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   63
            Left            =   2400
            TabIndex        =   90
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Dust"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   30
            Left            =   2880
            TabIndex        =   89
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Iron"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   27
            Left            =   2880
            TabIndex        =   88
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Pet"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   24
            Left            =   2880
            TabIndex        =   87
            Top             =   1320
            Width           =   735
         End
         Begin VB.Line Line2 
            Index           =   0
            X1              =   2280
            X2              =   2280
            Y1              =   600
            Y2              =   3840
         End
         Begin VB.Label Label1 
            Caption         =   "Civ:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   37
            Left            =   1560
            TabIndex        =   86
            ToolTipText     =   $"frmDrawMap.frx":2D6D
            Top             =   2280
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "NE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   40
            Left            =   1560
            TabIndex        =   85
            ToolTipText     =   "New Efficiency"
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   41
            Left            =   1800
            TabIndex        =   84
            ToolTipText     =   $"frmDrawMap.frx":2DFD
            Top             =   2280
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "100"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   46
            Left            =   1920
            TabIndex        =   83
            ToolTipText     =   "New Efficiency"
            Top             =   2520
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "FR:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   50
            Left            =   1560
            TabIndex        =   82
            ToolTipText     =   "Food Required"
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Pd:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   51
            Left            =   1560
            TabIndex        =   81
            ToolTipText     =   "Production"
            Top             =   1800
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Mx:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   57
            Left            =   1560
            TabIndex        =   80
            ToolTipText     =   "Max Production"
            Top             =   2040
            Width           =   255
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   85
            Left            =   1800
            TabIndex        =   78
            ToolTipText     =   "Production"
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   142
            Left            =   1800
            TabIndex        =   77
            ToolTipText     =   "Max Production"
            Top             =   2040
            Width           =   375
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Index           =   0
         Left            =   2160
         TabIndex        =   144
         Top             =   120
         Width           =   3015
         Begin VB.Label Label8 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   512
            Left            =   0
            TabIndex        =   145
            Top             =   2880
            Width           =   1815
         End
         Begin VB.Label Label7 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   170
            Top             =   120
            Width           =   2895
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   0
            X2              =   1920
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Line Line1 
            Index           =   4
            X1              =   0
            X2              =   1920
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Label Label1 
            Caption         =   "Iron:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   59
            Left            =   0
            TabIndex        =   169
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Food:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   71
            Left            =   0
            TabIndex        =   168
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Fd:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   72
            Left            =   0
            TabIndex        =   167
            Top             =   2040
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   79
            Left            =   480
            TabIndex        =   166
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   81
            Left            =   480
            TabIndex        =   165
            Top             =   2040
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Road:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   89
            Left            =   1080
            TabIndex        =   164
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Rail:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   92
            Left            =   1080
            TabIndex        =   163
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Def:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   93
            Left            =   1080
            TabIndex        =   162
            Top             =   2280
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   98
            Left            =   1560
            TabIndex        =   161
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   114
            Left            =   1560
            TabIndex        =   160
            Top             =   2040
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   115
            Left            =   1560
            TabIndex        =   159
            Top             =   2280
            Width           =   375
         End
         Begin VB.Label lblDes 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   158
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   116
            Left            =   1560
            TabIndex        =   157
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   117
            Left            =   1560
            TabIndex        =   156
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   118
            Left            =   1560
            TabIndex        =   155
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Pet:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   119
            Left            =   1080
            TabIndex        =   154
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Shell:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   134
            Left            =   1080
            TabIndex        =   153
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Gun:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   135
            Left            =   1080
            TabIndex        =   152
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   136
            Left            =   480
            TabIndex        =   151
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   137
            Left            =   480
            TabIndex        =   150
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "9999"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   138
            Left            =   480
            TabIndex        =   149
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Bar:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   139
            Left            =   0
            TabIndex        =   148
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Mil:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   140
            Left            =   0
            TabIndex        =   147
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Civ:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   141
            Left            =   0
            TabIndex        =   146
            Top             =   840
            Width           =   375
         End
      End
      Begin VB.CommandButton cmdMaxUnit 
         Caption         =   "p"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   178
         Top             =   0
         Width           =   255
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   3855
         Left            =   0
         TabIndex        =   10
         Top             =   360
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   6800
         TabWidthStyle   =   2
         TabFixedWidth   =   1764
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Census"
               Object.ToolTipText     =   "Sector Information"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Dist"
               Object.ToolTipText     =   "Distribution and Threshold Information"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Units"
               Object.ToolTipText     =   "Ship/Land/Plane Information"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblSect 
         Alignment       =   2  'Center
         Caption         =   "Sector 15,-5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1515
      ScaleWidth      =   2355
      TabIndex        =   2
      Top             =   4800
      Width           =   2415
      Begin VB.TextBox txtEntry 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         TabIndex        =   9
         Top             =   1080
         Width           =   3135
      End
      Begin VB.ListBox lstCmdResult 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   0
         MultiSelect     =   2  'Extended
         TabIndex        =   7
         Top             =   0
         Width           =   2295
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   480
      ScaleHeight     =   1515
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   1800
      Width           =   2055
      Begin VB.CommandButton cmdMaxMap 
         Caption         =   "p"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   177
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox picMap 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   0
         ScaleHeight     =   975
         ScaleWidth      =   1335
         TabIndex        =   12
         Top             =   0
         Width           =   1335
      End
      Begin VB.HScrollBar hsMap 
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   1320
         Width           =   1935
      End
      Begin VB.VScrollBar vsMap 
         Height          =   1215
         Left            =   1680
         TabIndex        =   5
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Timer Msg_Timer 
      Left            =   0
      Top             =   2280
   End
   Begin MSComctlLib.StatusBar sb1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6810
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer UpdateTimer 
      Left            =   0
      Top             =   1800
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   0
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageTB2 
      Left            =   0
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   27
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawMap.frx":2E8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawMap.frx":30BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawMap.frx":3351
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawMap.frx":3A63
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawMap.frx":3CF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawMap.frx":4007
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawMap.frx":4859
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawMap.frx":4AEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawMap.frx":4D7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawMap.frx":502F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawMap.frx":52B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawMap.frx":547B
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawMap.frx":5B25
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawMap.frx":5DB7
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawMap.frx":60C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawMap.frx":639B
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawMap.frx":661D
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawMap.frx":6957
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawMap.frx":6BE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawMap.frx":6E4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawMap.frx":70CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawMap.frx":734F
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawMap.frx":7BA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawMap.frx":83F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawMap.frx":8C45
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawMap.frx":9497
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrawMap.frx":A089
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   1
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuFileOption 
         Caption         =   "Import &Telegrams"
         Index           =   1
      End
      Begin VB.Menu mnuFileOption 
         Caption         =   "Import &Intelligence"
         Index           =   2
      End
      Begin VB.Menu mnuFileOption 
         Caption         =   "Import Sea Information"
         Index           =   3
      End
      Begin VB.Menu mnuFileOption 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuFileOption 
         Caption         =   "&Export All Telegrams"
         Index           =   5
      End
      Begin VB.Menu mnuFileOption 
         Caption         =   "Export &Recent Telegrams"
         Index           =   6
      End
      Begin VB.Menu mnuFileOption 
         Caption         =   "Ex&port Intelligence"
         Index           =   7
      End
      Begin VB.Menu mnuFileOption 
         Caption         =   "Export &Nation"
         Index           =   8
      End
      Begin VB.Menu mnuFileOption 
         Caption         =   "Export Sea Information"
         Index           =   9
      End
      Begin VB.Menu mnuFileOption 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuFileOption 
         Caption         =   "Copy Game Database"
         Index           =   11
      End
      Begin VB.Menu mnuFileOption 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuFileOption 
         Caption         =   "&Clear "
         Index           =   13
      End
      Begin VB.Menu mnuFileOption 
         Caption         =   "Print"
         Index           =   14
      End
      Begin VB.Menu mnuFileOption 
         Caption         =   "E&xit"
         Index           =   15
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewOptions 
         Caption         =   "Select Color Scheme"
         Index           =   1
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "Set Player Colors"
         Index           =   2
         Shortcut        =   +{F4}
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "Unit Viewing Options"
         Index           =   3
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "Show Ship Wakes"
         Index           =   4
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "Clear &Event Markers"
         Index           =   5
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "Zoom/UnZoom"
         Index           =   7
         Begin VB.Menu mnuZoomOptions 
            Caption         =   "UnZoom (25%)"
            Index           =   1
         End
         Begin VB.Menu mnuZoomOptions 
            Caption         =   "UnZoom (50%)"
            Index           =   2
         End
         Begin VB.Menu mnuZoomOptions 
            Caption         =   "UnZoom (80%)"
            Index           =   3
            Shortcut        =   ^{F11}
         End
         Begin VB.Menu mnuZoomOptions 
            Caption         =   "Zoom (125%)"
            Index           =   4
            Shortcut        =   ^{F12}
         End
         Begin VB.Menu mnuZoomOptions 
            Caption         =   "Zoom (175%)"
            Index           =   5
         End
         Begin VB.Menu mnuZoomOptions 
            Caption         =   "Zoom (250%)"
            Index           =   6
         End
         Begin VB.Menu mnuZoomOptions 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnuZoomOptions 
            Caption         =   "Set Magnifcation Level"
            Index           =   8
         End
         Begin VB.Menu mnuZoomOptions 
            Caption         =   "Return to Default"
            Index           =   9
         End
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Center"
         Index           =   9
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Go to...."
         Index           =   10
         Begin VB.Menu mnuGotoOptions 
            Caption         =   "&Last Event"
            Index           =   1
            Shortcut        =   ^Z
         End
         Begin VB.Menu mnuGotoOptions 
            Caption         =   "Jump Point 1"
            Index           =   2
            Shortcut        =   ^{F1}
         End
         Begin VB.Menu mnuGotoOptions 
            Caption         =   "Jump Point 2"
            Index           =   3
            Shortcut        =   ^{F2}
         End
         Begin VB.Menu mnuGotoOptions 
            Caption         =   "Jump Point 3"
            Index           =   4
            Shortcut        =   ^{F3}
         End
         Begin VB.Menu mnuGotoOptions 
            Caption         =   "Jump Point 4"
            Index           =   5
            Shortcut        =   ^{F4}
         End
         Begin VB.Menu mnuGotoOptions 
            Caption         =   "Jump Point 5"
            Index           =   6
            Shortcut        =   ^{F5}
         End
         Begin VB.Menu mnuGotoOptions 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnuGotoOptions 
            Caption         =   "Set Jump Point...."
            Index           =   8
            Begin VB.Menu mnuSetJumpPoint 
               Caption         =   "Set Jump Point 1"
               Index           =   1
            End
            Begin VB.Menu mnuSetJumpPoint 
               Caption         =   "Set Jump Point 2"
               Index           =   2
            End
            Begin VB.Menu mnuSetJumpPoint 
               Caption         =   "Set Jump Point 3"
               Index           =   3
            End
            Begin VB.Menu mnuSetJumpPoint 
               Caption         =   "Set Jump Point 4"
               Index           =   4
            End
            Begin VB.Menu mnuSetJumpPoint 
               Caption         =   "Set Jump Point 5"
               Index           =   5
            End
         End
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Redraw"
         Index           =   11
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "Options"
         Index           =   13
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "Refresh Database"
         Index           =   18
         Begin VB.Menu mnuRefresh 
            Caption         =   "Update Se&ctors"
            Index           =   1
         End
         Begin VB.Menu mnuRefresh 
            Caption         =   "Update &Bmap"
            Index           =   2
         End
         Begin VB.Menu mnuRefresh 
            Caption         =   "Update Units"
            Index           =   3
         End
         Begin VB.Menu mnuRefresh 
            Caption         =   "Update &Player List"
            Index           =   4
         End
         Begin VB.Menu mnuRefresh 
            Caption         =   "Update All"
            Index           =   5
         End
         Begin VB.Menu mnuRefresh 
            Caption         =   "Update Changed Info"
            Index           =   6
         End
         Begin VB.Menu mnuRefresh 
            Caption         =   "-"
            Index           =   7
         End
         Begin VB.Menu mnuRefresh 
            Caption         =   "Update Orders"
            Index           =   8
         End
         Begin VB.Menu mnuRefresh 
            Caption         =   "Update Missions"
            Index           =   9
         End
      End
   End
   Begin VB.Menu mnuUnits 
      Caption         =   "&Units"
      Begin VB.Menu mnuUnitsMarkNukes 
         Caption         =   "Mark Nukes"
      End
      Begin VB.Menu mnuUnitLand 
         Caption         =   "Land Units"
      End
      Begin VB.Menu mnuLandStat 
         Caption         =   "Land Stats"
      End
      Begin VB.Menu mnuLandCargo 
         Caption         =   "Land Cargo"
      End
      Begin VB.Menu mnuSep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUnitShip 
         Caption         =   "Ships"
      End
      Begin VB.Menu mnuShipStat 
         Caption         =   "Ship Stats"
      End
      Begin VB.Menu mnuShipCargo 
         Caption         =   "Ship Cargo"
      End
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUnitPlanes 
         Caption         =   "Planes"
      End
      Begin VB.Menu mnuPlaneStats 
         Caption         =   "Plane Stats"
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReportNuke 
         Caption         =   "Nukes"
      End
      Begin VB.Menu mnuLandsep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUnitShow 
         Caption         =   "Show Build Info"
      End
      Begin VB.Menu mnuUnitBuild 
         Caption         =   "Build Units"
      End
      Begin VB.Menu mnuUnitCommand 
         Caption         =   "&Unit Commands"
         Begin VB.Menu mnuUnitAttack 
            Caption         =   "Attack"
            Begin VB.Menu mnuUnitAttAttack 
               Caption         =   "&Probe"
               Index           =   0
            End
            Begin VB.Menu mnuUnitAttAttack 
               Caption         =   "&Attack"
               Index           =   1
               Shortcut        =   ^A
            End
            Begin VB.Menu mnuUnitAttAttack 
               Caption         =   "A&ssault"
               Index           =   2
               Shortcut        =   ^L
            End
            Begin VB.Menu mnuUnitAttAttack 
               Caption         =   "&Bomb"
               Index           =   3
            End
            Begin VB.Menu mnuUnitAttAttack 
               Caption         =   "&Paradrop"
               Index           =   4
            End
            Begin VB.Menu mnuUnitAttAttack 
               Caption         =   "&Missile Strike"
               Index           =   5
            End
            Begin VB.Menu mnuUnitAttAttack 
               Caption         =   "Center Map"
               Index           =   6
            End
         End
         Begin VB.Menu mnuSectCmd 
            Caption         =   "Se&ctor "
            Begin VB.Menu mnuSectCmdMove 
               Caption         =   "Move"
               Shortcut        =   ^M
            End
            Begin VB.Menu mnuSectCmdMoveAll 
               Caption         =   "MoveAll"
            End
            Begin VB.Menu mnuSectCmdDes 
               Caption         =   "Designate"
               Shortcut        =   ^D
            End
            Begin VB.Menu mnuSectCmdDist 
               Caption         =   "Distribute"
            End
            Begin VB.Menu mnuSectCmdThr 
               Caption         =   "Threshhold"
            End
            Begin VB.Menu mnuSectCmdThrAll 
               Caption         =   "Threshhold All"
               Shortcut        =   ^T
            End
            Begin VB.Menu mnuSectCmdFire 
               Caption         =   "Fire"
               Shortcut        =   ^F
            End
            Begin VB.Menu mnuSectCmdSpy 
               Caption         =   "Spy"
            End
            Begin VB.Menu mnuSectCmdBuild 
               Caption         =   "Build"
               Shortcut        =   ^B
            End
            Begin VB.Menu mnuSectCmdExplore 
               Caption         =   "Explore"
               Shortcut        =   ^E
            End
            Begin VB.Menu mnuSep5 
               Caption         =   "Delivery"
               Begin VB.Menu mnuSectCmdDel 
                  Caption         =   "Deliver"
               End
               Begin VB.Menu mnuSectCmdWipe 
                  Caption         =   "Wipe"
               End
            End
            Begin VB.Menu mnuSep2 
               Caption         =   "Production"
               Begin VB.Menu mnuSectCmdEstimate 
                  Caption         =   "Estimate Tool"
               End
               Begin VB.Menu mnuSectCmdProduction 
                  Caption         =   "Production"
               End
               Begin VB.Menu mnuSectCmdImprove 
                  Caption         =   "Improve"
                  Shortcut        =   ^I
               End
               Begin VB.Menu mnuSectCmdNeweff 
                  Caption         =   "New efficency"
               End
               Begin VB.Menu mnuSectCmdRadar 
                  Caption         =   "Radar"
               End
               Begin VB.Menu mnuSectCmdStart 
                  Caption         =   "Start"
               End
               Begin VB.Menu mnuSectCmdStop 
                  Caption         =   "Stop"
               End
            End
            Begin VB.Menu mnuSep3 
               Caption         =   "Combat"
               Begin VB.Menu mnuSectCmdAnti 
                  Caption         =   "Anti"
               End
               Begin VB.Menu mnuSectCmdAssault 
                  Caption         =   "Assault"
               End
               Begin VB.Menu mnuSectCmdBoard 
                  Caption         =   "Board"
               End
               Begin VB.Menu mnuSectCmdConvert 
                  Caption         =   "Convert"
               End
               Begin VB.Menu mnuSectCmdDemob 
                  Caption         =   "Demobilize"
               End
               Begin VB.Menu mnuSectCmdEnlist 
                  Caption         =   "Enlist"
               End
               Begin VB.Menu mnuSectCmdShoot 
                  Caption         =   "Shoot"
               End
               Begin VB.Menu mnuSectCmdStrength 
                  Caption         =   "Strength"
               End
            End
            Begin VB.Menu mnuSep4 
               Caption         =   "Misc"
               Begin VB.Menu mnuSectCmdCap 
                  Caption         =   "Capital"
               End
               Begin VB.Menu mnuSectCmdCenter 
                  Caption         =   "Center Map"
               End
               Begin VB.Menu mnuSectCmdGrind 
                  Caption         =   "Grind"
               End
               Begin VB.Menu mnuSectCmdOrigin 
                  Caption         =   "Origin"
               End
               Begin VB.Menu mnuSectCmdTerr 
                  Caption         =   "Terr"
               End
            End
         End
         Begin VB.Menu mnuDeity 
            Caption         =   "&Deity"
            Begin VB.Menu mnuDietyOption 
               Caption         =   "Give(take) "
               Index           =   1
            End
            Begin VB.Menu mnuDietyOption 
               Caption         =   "Set Sector"
               Index           =   2
            End
            Begin VB.Menu mnuDietyOption 
               Caption         =   "Set Resource"
               Index           =   3
            End
            Begin VB.Menu mnuDietyOption 
               Caption         =   "Peek/Hidden"
               Index           =   4
            End
         End
         Begin VB.Menu mnuShip 
            Caption         =   "&Ship"
            Begin VB.Menu mnuShipNav 
               Caption         =   "&Navigate"
            End
            Begin VB.Menu mnuShipNavMark 
               Caption         =   "Nav. &Markers"
            End
            Begin VB.Menu mnuShipLoad 
               Caption         =   "&Load"
            End
            Begin VB.Menu mnuShipWarload 
               Caption         =   "&Warload"
            End
            Begin VB.Menu mnuShipUnload 
               Caption         =   "&Unload"
            End
            Begin VB.Menu mnuShipLook 
               Caption         =   "&Look"
            End
            Begin VB.Menu mnuShipFire 
               Caption         =   "&Fire"
            End
            Begin VB.Menu mnuShipSonar 
               Caption         =   "&Sonar"
            End
            Begin VB.Menu mnuShipMission 
               Caption         =   "Mission"
            End
            Begin VB.Menu mnuSep13 
               Caption         =   "Combat"
               Begin VB.Menu mnuShipTorp 
                  Caption         =   "&Torpedo"
               End
               Begin VB.Menu mnuShipAssult 
                  Caption         =   "&Assault"
               End
               Begin VB.Menu mnuShipBoard 
                  Caption         =   "&Board"
               End
               Begin VB.Menu mnuShipMine 
                  Caption         =   "Mine"
               End
            End
            Begin VB.Menu MnuSep15 
               Caption         =   "Orders"
               Begin VB.Menu mnuShipOrder 
                  Caption         =   "&Order"
               End
               Begin VB.Menu mnuShipRetreat 
                  Caption         =   "Retreat"
               End
               Begin VB.Menu mnuShipMquota 
                  Caption         =   "Mquota"
               End
               Begin VB.Menu mnuShipSail 
                  Caption         =   "&Sail"
               End
               Begin VB.Menu mnuShipUnsail 
                  Caption         =   "Unsail"
               End
               Begin VB.Menu mnuShipFollow 
                  Caption         =   "Follow"
               End
            End
            Begin VB.Menu MnuSep14 
               Caption         =   "Misc"
               Begin VB.Menu mnuShipMisc 
                  Caption         =   "Center Map"
                  Index           =   0
               End
               Begin VB.Menu mnuShipMisc 
                  Caption         =   "Fleetadd"
                  Index           =   1
               End
               Begin VB.Menu mnuShipMisc 
                  Caption         =   "Fuel"
                  Index           =   2
               End
               Begin VB.Menu mnuShipMisc 
                  Caption         =   "Name"
                  Index           =   3
               End
               Begin VB.Menu mnuShipMisc 
                  Caption         =   "&Radar"
                  Index           =   4
               End
               Begin VB.Menu mnuShipMisc 
                  Caption         =   "Scrap"
                  Index           =   5
               End
               Begin VB.Menu mnuShipMisc 
                  Caption         =   "Scuttle"
                  Index           =   6
               End
               Begin VB.Menu mnuShipMisc 
                  Caption         =   "Start"
                  Index           =   7
               End
               Begin VB.Menu mnuShipMisc 
                  Caption         =   "Stop"
                  Index           =   8
               End
               Begin VB.Menu mnuShipMisc 
                  Caption         =   "&Tend"
                  Index           =   9
               End
               Begin VB.Menu mnuShipMisc 
                  Caption         =   "Upgrade"
                  Index           =   10
               End
            End
         End
         Begin VB.Menu mnuLand 
            Caption         =   "&Land"
            Begin VB.Menu mnulandMarch 
               Caption         =   "&March"
            End
            Begin VB.Menu mnuLandLook 
               Caption         =   "&Look"
            End
            Begin VB.Menu mnuLandLoad 
               Caption         =   "&Load"
            End
            Begin VB.Menu mnuLandWarload 
               Caption         =   "&Warload"
            End
            Begin VB.Menu mnuLandUnload 
               Caption         =   "&Unload"
            End
            Begin VB.Menu mnuLandMission 
               Caption         =   "Mission"
            End
            Begin VB.Menu mnuLandFire 
               Caption         =   "&Fire"
            End
            Begin VB.Menu mnuSepL2 
               Caption         =   "Misc"
               Begin VB.Menu mnuLandMisc 
                  Caption         =   "&Army"
                  Index           =   0
               End
               Begin VB.Menu mnuLandMisc 
                  Caption         =   "Center Map"
                  Index           =   1
               End
               Begin VB.Menu mnuLandMisc 
                  Caption         =   "Fortify"
                  Index           =   2
               End
               Begin VB.Menu mnuLandMisc 
                  Caption         =   "Mine"
                  Index           =   3
               End
               Begin VB.Menu mnuLandMisc 
                  Caption         =   "&Range"
                  Index           =   4
               End
               Begin VB.Menu mnuLandMisc 
                  Caption         =   "Retreat"
                  Index           =   5
               End
            End
            Begin VB.Menu mnuSepL1 
               Caption         =   "Other"
               Begin VB.Menu mnuLandFuel 
                  Caption         =   "Fuel"
               End
               Begin VB.Menu mnuLandRadar 
                  Caption         =   "Radar"
               End
               Begin VB.Menu mnuLandTend 
                  Caption         =   "&Tend"
               End
               Begin VB.Menu mnuLandScrap 
                  Caption         =   "Scrap"
               End
               Begin VB.Menu mnuLandScuttle 
                  Caption         =   "Scuttle"
               End
               Begin VB.Menu mnuLandStart 
                  Caption         =   "Start"
                  Index           =   0
               End
               Begin VB.Menu mnuLandStart 
                  Caption         =   "Stop"
                  Index           =   1
               End
               Begin VB.Menu mnuLandSupply 
                  Caption         =   "&Supply"
               End
               Begin VB.Menu mnuLandUpgrade 
                  Caption         =   "Upgrade"
               End
               Begin VB.Menu mnuSepL3 
                  Caption         =   "-"
               End
               Begin VB.Menu mnuLandMorale 
                  Caption         =   "Morale"
               End
            End
         End
         Begin VB.Menu mnuPlane 
            Caption         =   "&Plane"
            Begin VB.Menu mnuPlaneFly 
               Caption         =   "&Fly"
            End
            Begin VB.Menu mnuPlaneTrans 
               Caption         =   "&Transport"
            End
            Begin VB.Menu mnuPlaneBomb 
               Caption         =   "&Bomb"
            End
            Begin VB.Menu mnuPlanePara 
               Caption         =   "&Paradrop"
            End
            Begin VB.Menu mnuPlaneMission 
               Caption         =   "&Mission"
            End
            Begin VB.Menu mnuSepPlane1 
               Caption         =   "&Combat"
               Begin VB.Menu mnuPlaneRecon 
                  Caption         =   "&Recon"
               End
               Begin VB.Menu mnuPlaneSweep 
                  Caption         =   "&Sweep"
               End
               Begin VB.Menu mnuPlaneDrop 
                  Caption         =   "&Drop"
               End
               Begin VB.Menu mnuPlaneRange 
                  Caption         =   "&Range"
               End
            End
            Begin VB.Menu mnuSepPlane3 
               Caption         =   "&Missile"
               Begin VB.Menu mnuPlaneMissile 
                  Caption         =   "&Arm"
                  Index           =   0
               End
               Begin VB.Menu mnuPlaneMissile 
                  Caption         =   "&Disarm"
                  Index           =   1
               End
               Begin VB.Menu mnuPlaneMissile 
                  Caption         =   "&Harden"
                  Index           =   2
               End
               Begin VB.Menu mnuPlaneMissile 
                  Caption         =   "&Launch"
                  Index           =   3
               End
            End
            Begin VB.Menu mnuSepPlane2 
               Caption         =   "&Other"
               Begin VB.Menu mnuPlaneOther 
                  Caption         =   "Center Map"
                  Index           =   0
               End
               Begin VB.Menu mnuPlaneOther 
                  Caption         =   "Sa&tellite"
                  Index           =   1
               End
               Begin VB.Menu mnuPlaneOther 
                  Caption         =   "&Scrap"
                  Index           =   2
               End
               Begin VB.Menu mnuPlaneOther 
                  Caption         =   "S&cuttle"
                  Index           =   3
               End
               Begin VB.Menu mnuPlaneOther 
                  Caption         =   "Start"
                  Index           =   4
               End
               Begin VB.Menu mnuPlaneOther 
                  Caption         =   "Stop"
                  Index           =   5
               End
               Begin VB.Menu mnuPlaneOther 
                  Caption         =   "&Upgrade"
                  Index           =   6
               End
               Begin VB.Menu mnuPlaneOther 
                  Caption         =   "&Wing"
                  Index           =   7
               End
            End
         End
         Begin VB.Menu mnuNuke 
            Caption         =   "&Nuke"
            Begin VB.Menu mnuNukeList 
               Caption         =   "List Nukes"
            End
            Begin VB.Menu mnuNukeStart 
               Caption         =   "Start"
               Index           =   0
            End
            Begin VB.Menu mnuNukeStart 
               Caption         =   "Stop"
               Index           =   1
            End
            Begin VB.Menu mnuNukeTrans 
               Caption         =   "Transport"
            End
         End
      End
   End
   Begin VB.Menu mnuReportS 
      Caption         =   "&Reports"
      Begin VB.Menu mnuReportBudget 
         Caption         =   "&Budget"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuReportCensus 
         Caption         =   "&Census"
      End
      Begin VB.Menu mnuReportCoastwatch 
         Caption         =   "Coastwatch"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuReportCommodity 
         Caption         =   "Commodity"
         Index           =   1
      End
      Begin VB.Menu mnuReportCommodity 
         Caption         =   "Commodity Total"
         Index           =   2
      End
      Begin VB.Menu mnuReportCountry 
         Caption         =   "Countries"
      End
      Begin VB.Menu mnuReportCutoff 
         Caption         =   "Cuttoff"
      End
      Begin VB.Menu mnuReportLevel 
         Caption         =   "&Level"
      End
      Begin VB.Menu mnuReportMaps 
         Caption         =   "&Maps..."
         Begin VB.Menu mnuReportMap 
            Caption         =   "&Map "
            Index           =   0
         End
         Begin VB.Menu mnuReportMap 
            Caption         =   "&Nmap"
            Index           =   1
         End
         Begin VB.Menu mnuReportMap 
            Caption         =   "&Smap"
            Index           =   2
         End
         Begin VB.Menu mnuReportMap 
            Caption         =   "&Lmap"
            Index           =   3
         End
         Begin VB.Menu mnuReportMap 
            Caption         =   "&Pmap"
            Index           =   4
         End
         Begin VB.Menu mnuReportMap 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu mnuReportMap 
            Caption         =   "&Bmap"
            Index           =   6
         End
         Begin VB.Menu mnuReportMap 
            Caption         =   "Sbmap"
            Index           =   7
         End
         Begin VB.Menu mnuReportMap 
            Caption         =   "LBmap"
            Index           =   8
         End
         Begin VB.Menu mnuReportMap 
            Caption         =   "Pbmap"
            Index           =   9
         End
         Begin VB.Menu mnuReportMap 
            Caption         =   "-"
            Index           =   10
         End
         Begin VB.Menu mnuReportRoute 
            Caption         =   "Route"
         End
      End
      Begin VB.Menu mnuReportMotd 
         Caption         =   "Message of the Day"
      End
      Begin VB.Menu mnuReportNation 
         Caption         =   "&Nation"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuReportNewsPaper 
         Caption         =   "Newspaper"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuReportPlayers 
         Caption         =   "Players"
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuReportPower 
         Caption         =   "Power"
      End
      Begin VB.Menu mnuReportPowerNew 
         Caption         =   "&Power New"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuReportProduction 
         Caption         =   "Production"
      End
      Begin VB.Menu mnuReportRead 
         Caption         =   "Read (Quick)"
      End
      Begin VB.Menu mnuReportReport 
         Caption         =   "&Report (Tech/Edu/Hap)"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuReportResource 
         Caption         =   "Resource"
      End
      Begin VB.Menu mnuReportShow 
         Caption         =   "&Show"
      End
      Begin VB.Menu mnuReportSky 
         Caption         =   "Skywatch"
      End
      Begin VB.Menu mnuReportSpy 
         Caption         =   "Sp&y"
      End
      Begin VB.Menu mnuReportStarve 
         Caption         =   "Starvation"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuReportStart 
         Caption         =   "Start"
         Index           =   1
      End
      Begin VB.Menu mnuReportStart 
         Caption         =   "Stop"
         Index           =   2
      End
      Begin VB.Menu mnuReportUpdate 
         Caption         =   "&Update"
      End
      Begin VB.Menu mnuReportVersion 
         Caption         =   "&Version"
      End
      Begin VB.Menu mnuReportWire 
         Caption         =   "&Wire"
      End
   End
   Begin VB.Menu mnuDiplomacy 
      Caption         =   "&Diplomacy"
      Begin VB.Menu mnuDiploAccept 
         Caption         =   "Acc&ept"
      End
      Begin VB.Menu mnuDiploAcceptanceReport 
         Caption         =   "&Acceptance Report"
      End
      Begin VB.Menu mnuDiploCollect 
         Caption         =   "&Collect"
      End
      Begin VB.Menu mnuDiploConsider 
         Caption         =   "Co&nsider"
      End
      Begin VB.Menu mnuDiploDeclare 
         Caption         =   "&Declare"
      End
      Begin VB.Menu mnuReportFinancial 
         Caption         =   "&Financial (World Loans)"
      End
      Begin VB.Menu mnuDiploOffer 
         Caption         =   "&Offer"
      End
      Begin VB.Menu mnuReportsLedger 
         Caption         =   "&Ledger (Loans)"
      End
      Begin VB.Menu mnuDiploRelations 
         Caption         =   "&Relations"
      End
      Begin VB.Menu mnuDiploRelationsgrid 
         Caption         =   "Relations Grid"
      End
      Begin VB.Menu mnuDiploReject 
         Caption         =   "Re&ject"
      End
      Begin VB.Menu mnuDiploRepay 
         Caption         =   "Re&pay"
      End
      Begin VB.Menu mnuDiploSharebmap 
         Caption         =   "&Share bmap"
      End
      Begin VB.Menu mnuDiploShark 
         Caption         =   "Shar&k"
      End
      Begin VB.Menu mnuDiploTreaty 
         Caption         =   "&Treaty"
      End
   End
   Begin VB.Menu mnuMarket 
      Caption         =   "&Market"
      Begin VB.Menu mnuMarketBuy 
         Caption         =   "&Buy"
      End
      Begin VB.Menu mnuMarketMarket 
         Caption         =   "&Market report"
      End
      Begin VB.Menu mnuMarketReset 
         Caption         =   "&Reset"
      End
      Begin VB.Menu mnuMarketSell 
         Caption         =   "&Sell"
      End
      Begin VB.Menu mnuMarketSet 
         Caption         =   "S&et"
      End
      Begin VB.Menu mnuMarketTrade 
         Caption         =   "&Trade"
      End
   End
   Begin VB.Menu mnuGen 
      Caption         =   "&General "
      Begin VB.Menu mnuGenCmd 
         Caption         =   "&Break Sanctuary"
         Index           =   0
      End
      Begin VB.Menu mnuGenCmd 
         Caption         =   "&Change Country Name / Representative (Password)"
         Index           =   1
      End
      Begin VB.Menu mnuGenCmd 
         Caption         =   "Set &Realm"
         Index           =   2
      End
      Begin VB.Menu mnuGenCmd 
         Caption         =   "Set &Territory"
         Index           =   3
      End
      Begin VB.Menu mnuGenCmd 
         Caption         =   "Set &Owner"
         Index           =   4
      End
   End
   Begin VB.Menu mnuAttack 
      Caption         =   "&Attack "
      Begin VB.Menu mnuAttAttack 
         Caption         =   "&Probe"
         Index           =   0
      End
      Begin VB.Menu mnuAttAttack 
         Caption         =   "&Attack"
         Index           =   1
      End
      Begin VB.Menu mnuAttAttack 
         Caption         =   "A&ssault"
         Index           =   2
      End
      Begin VB.Menu mnuAttAttack 
         Caption         =   "&Bomb"
         Index           =   3
      End
      Begin VB.Menu mnuAttAttack 
         Caption         =   "&Paradrop"
         Index           =   4
      End
      Begin VB.Menu mnuAttAttack 
         Caption         =   "&Missile Strike"
         Index           =   5
      End
      Begin VB.Menu mnuAttAttack 
         Caption         =   "&Center Map"
         Index           =   6
      End
      Begin VB.Menu mnuClearSectorIntelligence 
         Caption         =   "Clear &Intelligence"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsOption 
         Caption         =   "&Script Files.."
         Index           =   1
      End
      Begin VB.Menu mnuToolsOption 
         Caption         =   "Build &Projection"
         Index           =   2
      End
      Begin VB.Menu mnuToolsOption 
         Caption         =   "&Famine Relief"
         Index           =   3
      End
      Begin VB.Menu mnuToolsOption 
         Caption         =   "Nation Levels"
         Index           =   4
      End
      Begin VB.Menu mnuToolsOption 
         Caption         =   "&Che/Plague Marker"
         Index           =   5
      End
      Begin VB.Menu mnuToolsOption 
         Caption         =   "Production Summary"
         Index           =   6
      End
      Begin VB.Menu mnuToolsOption 
         Caption         =   "Air Combat Report"
         Index           =   7
      End
      Begin VB.Menu mnuToolsOption 
         Caption         =   "Range Table"
         Index           =   8
      End
      Begin VB.Menu mnuToolsOption 
         Caption         =   "Fortify "
         Index           =   9
      End
      Begin VB.Menu mnuToolsOption 
         Caption         =   "Satellite Path"
         Index           =   10
      End
      Begin VB.Menu mnuToolsOption 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnuToolsOption 
         Caption         =   "&Blitz Setup"
         Index           =   12
      End
      Begin VB.Menu mnuToolsOption 
         Caption         =   "&Island Setup"
         Index           =   13
      End
      Begin VB.Menu mnuToolsOption 
         Caption         =   "Island &Develop"
         Index           =   14
      End
      Begin VB.Menu mnuToolsOption 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnuToolsOption 
         Caption         =   "Import &Foreign Bmaps"
         Index           =   16
      End
      Begin VB.Menu mnuToolsOption 
         Caption         =   "World Builder"
         Index           =   17
      End
   End
   Begin VB.Menu mnuHelpMenu 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Command list"
         Index           =   1
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Contents"
         Index           =   2
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Index"
         Index           =   3
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Info"
         Index           =   4
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Quick Start Guide"
         Index           =   5
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "User Guide"
         Index           =   6
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&About"
         Index           =   7
      End
   End
   Begin VB.Menu mnuSea 
      Caption         =   "Sea"
      Visible         =   0   'False
      Begin VB.Menu mnuSeaCenter 
         Caption         =   "&Center Map"
      End
      Begin VB.Menu mnuClearSeaIntelligence 
         Caption         =   "Clear &Intelligence"
      End
   End
   Begin VB.Menu mnuTelegram 
      Caption         =   "Telegram"
      Visible         =   0   'False
      Begin VB.Menu mnuTelegramCenter 
         Caption         =   "Center Map"
      End
      Begin VB.Menu mnuCopyReportWindow 
         Caption         =   "Copy to Report Window"
      End
   End
End
Attribute VB_Name = "frmDrawMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'050602 drk: enable or disable menu items based on sector context
'120503 rjk: added new designatioin info. to distribution tab,
'            increased the designation label on the distribution tab to show all of the name,
'            added a check to prevent runtime error when doing ctrl-T at sea or unknown sector
'
'130503 rjk: revamped the tab functions for Sector Panel.  Only displays vaild tabs.
'130503 drk,rjk: Added information for sea
'140503 rjk: If the Sector Panel can not show the current tab then go back to the first tab.
'110803 rjk: Added Airport Code to FillGrid and SortGrid so the SortGrid can sort the Range
'            column when formated with an Airport Code
'130803 efj: removed dead variables and procedures
'140903 rjk: Added Center Map menu item to attack, sector(under misc) and sea right clicks,
'            Removed Sect menu item from Sector->Misc menu, Fixed Warload from code cleanup,
'            Added more checks to Sector menus
'150903 rjk: Fixed Fuel from code cleanup, added some units types for LOTR II
'160903 rjk: Added Sea title for Sea Frame, reposition the rest of the text.
'            Fixed runtime error deleting Owners with "???" for enemy units.
'170903 rjk: Changed the Unit Filters to be a combo box instead of buttons.
'210903 rjk: Removed Path and Route menu items from Sector->Deliver menu
'230903 rjk: Removed Plane Cargo menu item
'280903 rjk: General reformatting.  Moved most of form logic to individual forms.
'            Switched type to an enum.  Added SetUnitDisplay to control the unit grid.
'            Made Display info private.  Removed FillSubSetGrid and automated.
'            Added enum for type of unit grid display.
'            Switched to SelX and  SelY for top level menu.
'            Fixed to show Unit grid shows when InRange selected.
'            Added ndump to Unit database refreshes.
'            Added follow menu click to call frmPromptFollow.
'            Fixed so Cancel from CenterMap does not cause an error.
'            Set up automatic load for Plane and Land similar to Ship for single
'            unit entry.  Position the fields for Consider cmd to appear correctly.
'            Removed hard coding for units and switched to UnitCharacteristicCheck.
'            Fixed frmPromptFleet to work for Planes and Land units.
'            Fixed map(s) command (removed & from command string).
'            Combined substype and fleet code into single argument for AddSubBox.
'            Moved enableOrDisableMenu call for top level menu.
'            Change "Fleet" to contain the Fleet name as well, allows separation and
'            identification for Fleet A from Wing A.
'            Created SetIndexSubCombo so the code could be called from more then
'            one location. Added a ClearSubCombo as well.
'            Removed ChangeUnitDisplay and combined with Toolbar1_ButtonClick.
'            Eliminate ShowPlaneInRange, ShowMissileInRange and ShowAttackInRange,
'            functionality now in SetUnitRange.  Added Type for "Train" units.
'011003 rjk: Added map, bmap, report(tech), relations, nation, orders, and missions to update all option
'            Clear Markers in when switching territories.
'021003 rjk: Changed to Anti menu item to not have to be occuppied but must have mil. and mob.
'            In the deliver menu item, the default des. for selecting shells was an "s" and now is an "i".
'            Moved the field logic for frmPromptMove to form.
'031003 rjk: For ltend switched to UnitPrompt ship from land so initial field shows ships.
'061003 rjk: Removed query on rsEnemyUnit, does not work this.  Need to investigate more later.
'            Also removed rs.Index = "Primary Key" as it caused an error on Unit tables.  Need to investigate more later.
'            Updated udENEMY so it produced output in FillUnitList.
'051003 rjk: Removed timestamp ndump
'061003 rjk: Added nuke information to bottom status line.
'071003 rjk: Nation list now in relations only so do not need "report *"
'081003 rjk: Added CStr for nation number in FillGrid when doing AddSubBox
'091003 rjk: Split Reject menu into two menus, modify frmPromptOffer to work for the new Reject/Accept menu items
'            Rename Accept to AcceptReport so as to not be confused with Accept from the Reject command
'091003 rjk: Moved the field logic to the form for frmPromptOffer
'141003 rjk: Moved "show sect s" here from FormLoad frmDrawMap to RefreshDatabase to be inside the database refresh batch
'141003 rjk: Moved BuildType to be private variable of FillGrid
'            Change mnuPlaneScutle_Click() to mnuPlaneScuttle_Click(), fixed spelling
'            Added timestamp for lost when doing update all changes
'151003 rjk: Switch Oil and Fert fields so they display correctly on the Sea sector panel
'161003 rjk: Removed dump request from spy as none of our sectors should change.
'161003 rjk: Added functionality for Buy, Sell, Reset and enhanced Market
'171003 rjk: Removed the show sect s from the timer as this information is static
'            and is checked upon startup, added to Update All
'171003 rjk: Added Strength fields to Sector database
'171003 rjk: Added Export Nation to the file menu
'181003 rjk: Fixed that if enemy unit is at x=-1 or y=-1, it does not display ???
'            Added Route menu to Maps, it called frmPromptCom
'181003 rjk: Added multiple selection to MouseDown, builds sector list with '\' separating sectors.
'181003 rjk: Modified Distribute to support multiple sectors select.
'191003 rjk: Modified Threshold/Move/Deliver to support multiple sector selection.
'231003 rjk: Added three territory fields to the sector display, moved the strength fields to the right.
'251003 rjk: Added Shift F9 to forward in the command list
'            No Repeat Panel exists, so some code cleanup was done in sb1_PanelClick
'261003 rjk: Added Spy to report to allow section of multiple sectors.
'271003 rjk: Switched Spy report to use frmPromptRadar instead of frmPromptCensus.
'301003 rjk: Added Keys for moving about on the map (only active from the map)
'            ExportIntelligence has its own form for selecting parameters.
'            For active menu determination, switched to SelX and SelY from MxPos and MyPos.
'311003 rjk: For frmPromptMove - allow Destination selection while not in the destination field
'031103 rjk: Added relation command to Form load for when the AutoUpdate is off to build the basic relation info.
'061103 rjk: Modified the Key pattern for moving around the map using the Keyboard
'131103 rjk: Added World wraparound checks to AddtoPath (frmPromptPath) to fixed problem doing a recon over a boundary.
'161103 rjk: Added txtSector to sector selection using the mouse for frmPromptSector
'181103 rjk: Created one copy of the DisplayPath code and moved here.
'191103 rjk: Changed MoveAll to accomodate the multiple sector selection, added Commodity Total Report
'201103 rjk: Added a new option for controlling strength updates
'            Suppress Sector menu when on sea sector
'211103 rjk: Resolved nav. sector visibility run-time error by moving the mnuSectCmd and mnuUnits menu determination
'            to enableOrDisableMenus
'241103 rjk: Code Cleanup, changed muuUnitsMarkNukes to mnuUnitsMarkNukes
'            If the strength information is unknown display "???" (UpdateSectors, deletes the records, an suppress option on)
'251103 rjk: Fixed cmbSub to show all when All is selected.
'            Moved mnuAttCenter to mnuAttAttack to save a control. Moved the General into an array to save controls
'            Moved the ShipMisc/PlaneOthers/PlaneMissile/LandMisc commands into an array to save controls.
'            Added the ability to center the map to the first unit in the list.
'271103 rjk: Changed to show Attack and Unit menus together on enemy sectors.
'281103 rjk: Moved the pictures to EmpUnitView and added initialize for unitView object and removed ViewUnits
'021203 rjk: Changed from mnuViewOptions to mnuOptions to fixed status bar option.
'            Added Start and Stop to Reports Menu to allow multiple sector section.
'            Moved options to frmOptions and switched to boolean options.
'031203 rjk: display enemy mines in red "?".  Added F3 shortcut to UnitView menu.
'071203 rjk: Changed hldindex to hldIndex.
'121203 rjk: Switched to Items and Item classes and objects.
'121203 rjk: Added closing bracket to display of newdes in Deity mode.
'211203 rjk: Separate Unit number as separate column in the unit grid.
'291203 rjk: Added Local Help
'260104 rjk: Modify production double click to product info (St@r W@rs, y changes at the same time)
'040204 rjk: Fixed Help menu.
'050204 rjk: Fixed Unit List with 'All', and unit filter, moved boxes to fix overlap.
'070204 rjk: Update NavMarkers to use Collection to do an exhaustive path search
'260204 rjk: Added show sect c/b to RefreshAll ensure the database is fill, as it starts empty
'070304 rjk: Switch the Enemy timestamp to UTC
'080304 rjk: Fixed the grid selection to pass unit id to form when the row is not actually selected
'110304 rjk: Switched to use IsNull check for EnemySector designation, display Unknown designation for null entry
'            if you can not find the information in the rsBmap on the sector panel. Show ??? for unknown efficiency.
'150304 rjk: Added Estimate Tool from Jim's original code
'270304 rjk: Switched over to DeleteAllRecords for clearing a table
'010404 rjk: Added actual Firing Range column to Unit grid.
'030404 rjk: Added for Copy Game Database form, only visible for Deity
'160404 rjk: Fixed the sort column for id to be numeric instead of text.
'200404 rjk: Fixed the redrawing of EventMarkers.
'250404 rjk: Fixed run-time error drawing NavMarkers, color variable should have been a long.
'250404 rjk: Added the ability to include Update mobility in NavMarkers
'250404 rjk: Added the remaining fields from the rangetool to calculated fields in the grid.
'250404 rjk: Fixed the column width for enemy timestamp.
'250404 rjk: Added the ability to save Telegram/Server window output to log or file.
'250404 rjk: Added the ability to right click on the Telegram/Server window to center map if
'            cursor is on a sector address.
'260404 rjk: Fixed run-time crash if deleting unit enemy by selecting the whole grid or including the total line.
'290404 rjk: Added the ability to include Update efficicency in NavMarkers
'110604 rjk: Unload the chat/flash windows when closing.
'110604 rjk: Updated the menu for plague marking.
'200604 rjk: Updated Tooltip for Che Marker.  Added Satellite Path Calculation and Nuke Damage Report
'010704 rjk: Filled the default origin for Nuke Damage (Range Tool).
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
'230704 rjk: Added Radius for delivery for StarWars.
'010704 rjk: Added "!" as an end delimiter for sector check for Telegram Window.
'250804 rjk: Enlist does not require civilians in 2K4.
'290804 rjk: Fixed read (quick) to work for deity mode.
'250904 rjk: Added Simple Territory Form.
'270904 rjk: Added the saving of display mode (DD_CURRENT, DD_NEW and DD_BOTH).
'151104 rjk: Added the ability to start with your Capital instead of 0,0
'011204 rjk: Added new PrintMap function.
'190105 rjk: Set UTF8 on the server based on the option on startup.  Fixed spelling.
'            Postpone the update commands by 10 seconds to prevent the command from being aborted.
'220305 rjk: Added SubmitTelegramRead to single function for requesting telegrams.
'220305 rjk: Added SendTelegramRead to single function for requesting telegrams.
'            Added version check to deal syntax change in 4.2.21.
'260305 rjk: Fixed NavMarkers to deal with the extra mobility costs with ships with
'            SWEEP capabilities.
'060405 rjk: Changed the cmbSub store Nations instead of Fleets for deity mode.
'            Fixed the unit filter for enemy to work with country #s greater then 10.
'100405 rjk: UTF8 encapsulated in server version check for 4.2.21 or newer.
'120405 rjk: Added Update combo box for NavMarkers
'200405 rjk: Added checkbox for NavMarkers to control whether to limit
'            mobility to max. mobility or not.
'240405 rjk: Added Three custom script buttons.
'140505 rjk: Fixed ship mobility cost to deal with the extra mobility costs with
'            ships with SWEEP capabilities.  Removed correction factor from NavMarkers.
'140505 rjk: Added @sector parsing in the telegram frame in main form.
'260505 rjk: Switched UTF8 from an registry option to login option.
'290505 rjk: Added ability show Commodity Ratio and an option to control it.
'160705 rjk: Added GetOilContent, added Clear Intelligence for sector.
'230705 rjk: Moved Exit to the bottom of the menu.
'080805 rjk: Added 'n' to the build menu check.
'010106 rjk: Change Hidden command to Peek to reflect the server change in 4.2.22.
'280106 rjk: Added mission type and radius to plane grid.
'180206 rjk: Replace nation with GetNation.  Replace sdump with GetShips.
'            Replace ldump with GetLandUnits.  Replace ndump with GetNukes.
'            Replace pdump with GetPlanes.  Replace lost with GetLost.
'            Replace relation with GetCountryList.
'            Use xdump to replace mission for 4.3.0.  Use xdump to replace orders for 4.3.0.
'120306 rjk: Add a call to Compute Unit Counts for ships for 4.3.0.
'210306 rjk: Switched SendFullDumpCommand to GetSectors
'210306 rjk: Moved the RequestMetaTables to DrawMap form loading.
'210306 rjk: Added the peeper account to the visitor account.
'230406 rjk: Added the Unit Counts for Land Units for 4.3.X servers.
'060506 rjk: Added Nuke grid.  Remove Nuke warning label.
'140506 rjk: Fixed mobility and effeciency buttons from grid changes.
'            Fixed the distribution panel disappearing from grid changes.
'200506 rjk: Enable menus for radar and improve all the time for SP: Atlantis.
'220506 rjk: Incorporate 4.3.4 server changes for xdump nat and coun.
'            Also SP: Atlantis will have some the xdump nat and coun changes.
'250506 rjk: Added Four SP: Atlantis infrastructures to the Sector Panel,
'            use runway -> min, radar -> gld, fort -> fert nad navigate -> oil.
'270506 rjk: Increased the size of civilians to 5 digits and exceed civilians to
'            4 digits.
'            Do not reset the database you on startup for SP: Atlantis.
'300506 rjk: Added shortcuts for Jump points.
'100606 rjk: Added double click to centre map for the unit grid.
'            Fixed a problem where the Unit Text List in Unit grid.
'            Not able to requery against a table, only works for the queries.
'110606 rjk: Enhanced script buttons to support jump points.
'            Fixed radar and fire menus to check sector fields for SP:Atlantis.
'            Change toolbar button for bmap not to erase all the map information.
'            for SP:Atlantis.
'180606 rjk: Added support for canal flag to ship characteristics.
'250606 rjk: Added start/stop unit support for 4.3.6 servers.
'            Set the id to red if stopped.  Added menu for start/stop for units.
'030706 rjk: Update the grid when transporting.
'230806 rjk: Added Import/Export Sea Information
'181006 rjk: Added Import Offset for Intelligence
'311206 rjk: Added an option to show short names in the Unit grid.
'090607 rjk: Display whether production is material limited or now.
'090907 rjk: Added code for using xdump game and xdump update to get
'            next update for server 4.3.10 and newer.  Use new functions
'            GetUpdate, UpdateCommands() and StartUpdateTimer().
'221007 rjk: Use for xdump ver for 4.3.10 servers and newer for version button
'311007 rjk: Mark sectors that are now fully ours in yellow
'031107 rjk: Filter the load/unload/lload/lunload commands to relevant units only.
'251207 rjk: Changed to cargo to start cargo and added end cargo to mission string.
'010108 rjk: Switch to GetOrders function.
'100408 rjk: Fixed Version to ensure it shows in a menu when selected from the menu
'            or toolbar (broken 221007).
'190708 rjk: Fixed Escape Key to sent to 'aborted' instead of 'ctld', fixes Escape
'            cause disconnect from the session.
'250309 rjk: Changed to use RefreshAll for UpdateAll.
'110410 rjk: Probe changed to n n n n instead of 0 0 0 0 (attack using y/n instead of 1/0)

Public ShutDown As Boolean
Public PositionHelp As Boolean

Public PromptForm As Form  'Hold form currently on prompt
Public PromptUp As Boolean
Public DrawingPath As Boolean
Public DisplayingPath As Boolean
Public MarkingTerritory As Boolean
Public WorldBuilding As Boolean
Public MapValid As Boolean
Public NavMarkerShip As String
Public bNavMarkerShipIncludeUpdateMob As Boolean
Public bNavMarkerShipIncludeUpdateEff As Boolean
Public iNavMarkerShipUpdates As Integer
Public bNavMarkerShipLimitMobility As Boolean

'Track which control activated the pop-up menu
Dim PopUpMenuSource As Integer
Const DMAP_PUMS_MAP = 0
Const DMAP_PUMS_GRID = 1

' Selected Sectors
Public SelX As Integer
Public SelY As Integer
Private SelSectType As Integer

Const SEL_SECT_OWN = 1
Const SEL_SECT_ENEMY = 2
Const SEL_SECT_UNKNOWN = 3
Const SEL_SECT_SEA = 4 'Added rjk 05/13/03: Sea Sector that is not a bridge span or bridge tower

' Const MAPBORDER = 10   removed efj 8/2003

'This static var contains the last retreat string so that it can
'be pulled up in the Nav Control
Public DoRetreat As Boolean
Public LastRetreatString As String
Public LastRetreatUnits As String
Public LastRetreatType As String

'This holds the cursor position on the command line
'and if it currently has the focus
Dim CommandCursorPos As Integer
Dim CommandCursorFocus As Boolean

'This holds the pending telegrams
Public MsgQ As New Collection

'These hold the timer variables
Public SecondsToUpdate As Long
Public TimerAtUpdate As Long

Dim MxPos As Integer
Dim MyPos As Integer
Public Magnification As Single
Dim strLandid As String
Dim strShipid As String
Dim strPlaneid As String
Dim strNukeid As String
Dim strEnemyid As String
Dim ViewShipWake As Boolean
Dim GridSelection As Boolean
Dim strTextBox As String


Dim bShips As Boolean
Dim bLands As Boolean
Dim bPlanes As Boolean
Dim benemys As Boolean
Dim bNukes As Boolean

'Private mintCurFrame As Integer ' Current Frame visible rjk 05/13/03: removed,
' DisplaySelectPanel determines which frame should be visible.

' Dim Grid() As Integer   removed efj 8/2003
' Dim Gridxy() As Integer   removed efj 8/2003
Dim WorkingPath() As Integer

'unit box vars
' 0901903 rjk: must match to button index in Toolbar1,
'              also FillUnitList depends on the numbers and order
Public Enum enumUnitDisplay
    udUNKNOWN = 0
    udSHIP = 1
    udLAND = 2
    udPLANE = 3
    udNUKE = 4
    udENEMY = 5
    udList = 6
End Enum
Private Enum enumSubset
    ssALL
    ssSECT
    ssFLEET
    ssPLANE_RANGE
    ssMISSILE_RANGE
    ssATTACK_RANGE
End Enum
Public Enum enumUnitType
    TYPE_START 'must be the first type
    TYPE_ALL
    'qualifiers
    TYPE_LAND_UNLOADED
    TYPE_LAND_LOADED
    TYPE_SHIP_UNLOADED
    TYPE_SHIP_LOADED
    'ship capabilities/abilities
    TYPE_S_FISH
    TYPE_S_TORP
    TYPE_S_DCHRG
    TYPE_S_PLANE
    TYPE_S_MISS
    TYPE_S_OIL
    TYPE_S_OILER
    TYPE_S_SONAR
    TYPE_S_MINE
    TYPE_S_SWEEP
    TYPE_S_SUB
    TYPE_S_SPY
    TYPE_S_LAND
    TYPE_S_SEMI_LAND
    TYPE_S_SUB_TORP
    TYPE_S_TRADE
    TYPE_S_CANAL
    TYPE_S_ANTI_MISSILE
    'ship commands
    TYPE_SC_FIRE
    'plane capabilities/abilities
    TYPE_P_TACTICAL
    TYPE_P_BOMBER
    TYPE_P_INTERCEPT
    TYPE_P_CARGO
    TYPE_P_VTOL
    TYPE_P_MISSILE
    TYPE_P_LIGHT
    TYPE_P_SPY
    TYPE_P_IMAGE
    TYPE_P_SATELLITE
    TYPE_P_STEALTH
    TYPE_P_SDI
    TYPE_P_HALF_STEALTH
    TYPE_P_X_LIGHT
    TYPE_P_CHOPPER
    TYPE_P_ANTISUB
    TYPE_P_PARA
    TYPE_P_ESCORT
    TYPE_P_MINE
    TYPE_P_SWEEP
    TYPE_P_MARINE
    'plane groups and commands
    TYPE_PG_ESCORTS
    'plane commands
    TYPE_PC_BOMB
    TYPE_PC_LAUNCH
    TYPE_PC_DROP
    'land units capabilities/abilities
    TYPE_L_XLIGHT
    TYPE_L_ENGINEER
    TYPE_L_SUPPLY
    TYPE_L_SECURITY
    TYPE_L_LIGHT
    TYPE_L_MARINE
    TYPE_L_RECON
    TYPE_L_RADAR
    TYPE_L_ASSAULT
    TYPE_L_TRAIN
    'land unit commands
    TYPE_LC_FIRE
    TYPE_END 'Must be the last type
End Enum
Private DisplaySelect As enumUnitDisplay
Private LastUnitDisplay As enumUnitDisplay
Private DisplaySubset As enumSubset '092503 rjk: Switched to enum from Integer
'Public secx As Integer
'Public secy As Integer
Public Fleet As String
'Public BuildType As String 101403 rjk: Moved to be private variable of FillGrid
Public SubType As enumUnitType '092503 rjk: Switched to enum from Integer

'Public FillSubSetFlag As Boolean 092503 rjk: Removed, logic in FillGrid.
Private strSubSect As String 'used to keep a list of Sectors in subCombo
Private strSubFleet As String 'used to keep a list of Fleets in subCombo
'Private strSel As String moved to MouseUp function, only place used

'Sort Variables
Dim SortCol As Integer
Dim SortDecending As Boolean

'Selection Variables
Private NeedMob As Boolean
Private Needeff As Boolean

Private Const SPLITTER_HEIGHT = 40

' The percentage occupied by the PictureBox.
Private PercentageY As Single
Private PercentageX1 As Single
Private PercentageX2 As Single

'True when we are dragging the splitter.
Private Dragging As Boolean

'Which Splitter are we draging
Private DragY As Boolean
Private DragX1 As Boolean

'Map Maximized Vars
Private MapMax As Boolean
Private MMPercentageY As Single
Private MMPercentageX1 As Single
Private MMPercentageX2 As Single

'Unit Maximized Vars
Private UnitMax As Boolean

'Telegram Center Map
Dim iTelegramSelX As Integer
Dim iTelegramSelY As Integer

Private strType(TYPE_START To TYPE_END) As String
Private strTypeTitle(TYPE_START To TYPE_END) As String

Private Declare Function HtmlHelp Lib "HHCtrl.ocx" Alias "HtmlHelpA" _
    (ByVal hWndCaller As Long, _
     ByVal pszFile As String, _
     ByVal uCommand As Long, _
     dwData As Any) As Long

Const HH_DISPLAY_TOPIC As Long = 0
Const HH_DISPLAY_TOC As Long = &H1
Const HH_DISPLAY_INDEX As Long = &H2
Const HH_HELP_CONTEXT As Long = &HF

Private Sub cmbUnitFilter_Click()
SubType = cmbUnitFilter.ItemData(cmbUnitFilter.ListIndex)
FillGrid
End Sub

Private Sub cmdMaxMap_Click()
'Handle the Max Maximize option
If MapMax Then
    MapMax = False
    PercentageY = MMPercentageY
    PercentageX1 = MMPercentageX1
    PercentageX2 = MMPercentageX2
    tbMain.Visible = frmOptions.bToolbar
    sb1.Visible = frmOptions.bStatusBar
Else
    MapMax = True
    MMPercentageY = PercentageY
    MMPercentageX1 = PercentageX1
    MMPercentageX2 = PercentageX2
    PercentageY = 1
    PercentageX1 = 1
    tbMain.Visible = False
    sb1.Visible = False
End If
ResizePanels
ArrangeControls

End Sub
Private Sub cmdMaxUnit_Click()
'Handle the Max Maximize option
If UnitMax Then
    UnitMax = False
    PercentageY = MMPercentageY
    PercentageX1 = MMPercentageX1
    PercentageX2 = MMPercentageX2
    tbMain.Visible = frmOptions.bToolbar
    sb1.Visible = frmOptions.bStatusBar
Else
    UnitMax = True
    MMPercentageY = PercentageY
    MMPercentageX1 = PercentageX1
    MMPercentageX2 = PercentageX2
    PercentageY = 0
    PercentageX2 = 1
    DisplayUnitBox 'rjk 05/13/03: which to DisplayUnitBox from DisplaySectorPanel 3
    tbMain.Visible = False
    sb1.Visible = False
End If
ResizePanels
ArrangeControls

End Sub

Private Sub Form_Activate()
'030404 rjk: Added for Copy Game Database form, only visible for Deity
mnuFileOption(11).Visible = bDeity
mnuFileOption(12).Visible = bDeity
End Sub

Private Sub Form_Load()
Dim iCapX As Integer
Dim iCapY As Integer
Dim m As Integer

frmOptions.LoadOptions '120203 rjk: Moved to frmOptions

MapValid = True
MarkingTerritory = False
DrawingPath = False
PositionHelp = False
NavMarkerShip = vbNullString
SecondsToUpdate = 0
TimerAtUpdate = 0
' LowResGraphics = False    8/2003 efj  removed
picMap.ScaleMode = vbPixels

'Refresh Database from the server data
If frmEmpCmd.ConnectedtoHost Then
    RequestMetaTables
    
    'Set the map to be invalid
    MapValid = False
    
    'Skip database refill if you are visitor or guest
    If Not (LCase$(frmEmpCmd.CountryName) = "visitor" Or LCase$(frmEmpCmd.CountryName) = "guest" Or _
        LCase$(frmEmpCmd.CountryName) = "peeper") Then
        MousePointer = vbHourglass
        'set AutoRead
        UpdateAutoRead '120203 rjk: Switched to frmOptions and boolean options
            
        UpdateAutoUpdate '120203 rjk: Switched to frmOptions and boolean options
        StartUpdateTimer
        'fill the database
        If frmEmpCmd.bAutoUpdate Then
            frmEmpCmd.ForceUpdates = True
            'clear database
            If Not frmOptions.bSPAtlantis Then
                frmEmpCmd.SubmitEmpireCommand "dbr", False
            End If
            
            ' thomas lullier - fix show sect s here
            ' used to calculate the max safe civ at connection
            ' when autoupdate is on
            '101403 rjk: Moved to RefreshDatabase
            'frmEmpCmd.SubmitEmpireCommand "bf1", False 'drk 2/26/03: keep the show sect s window from popping up...
            'frmEmpCmd.SubmitEmpireCommand "show sect s", False
            'frmEmpCmd.SubmitEmpireCommand "bf2", False
            
            RefreshDataBase ' Fill it back up
        Else
            'submit the db2 to get the initial map update
            frmEmpCmd.SubmitEmpireCommand "db1", False
            GetCountryList
            frmEmpCmd.SubmitEmpireCommand "db2", False
        End If
    Else
        Label4.Caption = "Visitor"
    End If
End If
        
'get layout info from registry '060304 rjk: Added Val to deal with regional settings.
Magnification = Val(GetSetting(APPNAME, "Layout", "Magnification", 1))
PercentageY = Val(GetSetting(APPNAME, "Layout", "Percentage Y", 0))
PercentageX1 = Val(GetSetting(APPNAME, "Layout", "Percentage X1", 0))
PercentageX2 = Val(GetSetting(APPNAME, "Layout", "Percentage X2", 0))

'Check for error - can happen when country setting changes.
'Make sure values are reasonable
If Magnification > 15 Then
    Magnification = 1
ElseIf Magnification < 0.3 Then
    Magnification = 1
End If

'Same here, make sure settings are in range
If PercentageY < 0.05 Or PercentageY > 1 Then
    PercentageY = 0.65
End If
If PercentageX1 < 0.05 Or PercentageX1 > 1 Then
    PercentageX1 = 0.75
End If
If PercentageX2 < 0.05 Or PercentageX2 > 1 Then
    PercentageX2 = 0.55
End If
    
'Now load as normal
SetHexSideLength picMap, (Screen.Width / Screen.TwipsPerPixelX) / (30 * 0.866 * 2) * Magnification
DrawingPath = False

'MsgBox "Loading Pictures"
frmUnitView.unitView.Load '112803 rjk: Moved the pictures to EmpUnitView, and initialize the unitView object and removed ViewUnits
LoadPictures '112803 rjk: Moved the pictures to EmpUnitView

'MsgBox "Loading Descriptions"
LoadSectorDesc

'Set Tab box
'mintCurFrame = 1 rjk 05/13/03: not used more, done in DisplaySectorPanel
Frame1(0).Visible = False
Frame1(1).Visible = False
Frame1(1).Height = TabStrip1.Height
Frame1(2).Visible = False
Frame1(3).Visible = False 'rjk 05/13/03: ensure Frame1(3) is initialized as well
Frame1(4).Visible = False
'rjk 05/13/03: Initialize Sector Panel to location 0,0
If frmOptions.bCapital And ParseSectors(iCapX, iCapY, CapSect) Then
    FillSectorBox iCapX, iCapY
Else
    FillSectorBox 0, 0
End If
DisplayFirstSectorPanel

'Set Origins
OriginX = -6
OriginY = -6

'Set map size from version info - may need to change later
If rsVersion.BOF And rsVersion.EOF Then
    'If there is no version record, use 96x48 as default
    MapSizeX = 48
    MapSizeY = 48
Else
    rsVersion.MoveFirst
    MapSizeX = rsVersion.Fields("World X")
    MapSizeY = rsVersion.Fields("World Y")
    If MapSizeX = 0 Then MapSizeX = 48
    If MapSizeY = 0 Then MapSizeY = 48
End If

'Set ScrollBar Properties
Dim bhld As Boolean
bhld = MapValid
MapValid = False
vsMap.Min = -1.15 * (MapSizeY / 2)
vsMap.SmallChange = 2
vsMap.Value = OriginY

hsMap.Min = -1.15 * (MapSizeX / 2)
hsMap.SmallChange = 2
hsMap.Value = OriginX
MapValid = bhld

'Set Window Caption
'092603 rjk: Added Offline mode
If frmEmpireLogin.bOffline Then
    Caption = App.title + " - " + GameName + " (Offline)"
Else
    Caption = App.title + " - " + GameName
End If

'Max to fill the full screen
Me.WindowState = vbMaximized

'Set status bar panels
sb1.Panels.Item(1).key = "Timer"
sb1.Panels.Item(1).Text = "Unknown"
sb1.Panels.Item(1).ToolTipText = "Time Remaining until Update"
sb1.Panels.Item(1).MinWidth = 1

sb1.Panels.Add 2, "BTU"
sb1.Panels.Item(2).Text = " BTU: 640 "
sb1.Panels.Item(2).ToolTipText = "Bureaucratic Time Units Remaining"
sb1.Panels.Item(2).MinWidth = 1

sb1.Panels.Add 3, "Lights"
Set sb1.Panels.Item(3).Picture = picGreenLight
sb1.Panels.Item(3).MinWidth = 1
sb1.Panels.Item(3).ToolTipText = "Empire Server Available"

sb1.Panels.Add 4, "Mail"
sb1.Panels.Item(4).Text = vbNullString
sb1.Panels.Item(4).MinWidth = 600
sb1.Panels.Item(4).ToolTipText = "Telegram Indicator"

sb1.Panels.Add 5, "Mail Count"
sb1.Panels.Item(5).MinWidth = 300
sb1.Panels.Item(5).ToolTipText = "Number of pending telegrams"

sb1.Panels.Add 6, "Anno"
sb1.Panels.Item(6).Text = vbNullString
sb1.Panels.Item(6).ToolTipText = "Announcement Indicator"
sb1.Panels.Item(6).MinWidth = 600

For m = 1 To 6
    sb1.Panels.Item(m).AutoSize = sbrContents
    sb1.Panels.Item(m).Alignment = sbrCenter
Next

sb1.Panels.Add 7, "Hex"
sb1.Panels.Item(7).AutoSize = sbrSpring
'set rtb font
Set rtbTelegram.Font = Me.Font
'120203 rjk: Moved loading the options to frmOptions

LastShip = -1

'FillSubSetFlag = True 092503 rjk: Eliminated
'092103 rjk: Fill the strType dynamically
BuildUnitFilters
End Sub

Public Sub SetMapSize(sx As Integer, sy As Integer)
'error check
If sx <= 0 Or sy <= 0 Then
    Exit Sub
End If

MapSizeX = sx
MapSizeY = sy
vsMap.Min = -1.15 * CInt(MapSizeY / 2)
hsMap.Min = -1.15 * CInt(MapSizeX / 2)
DrawHexes
End Sub

'Resize the Panels of on the screen
Private Sub ResizePanels()
Dim hgt11 As Single
Dim hgt12 As Single
Dim hgt21 As Single
Dim hgt22 As Single
Dim wdt1 As Single
Dim wdt2 As Single
Dim ohgt As Single
Dim ftop As Single

'take into account toolbar and status bar height
ohgt = 0
ftop = 0
If tbMain.Visible Then
    ohgt = tbMain.Height
    ftop = tbMain.Height
End If

If sb1.Visible Then
    ohgt = ohgt + sb1.Height
End If


' Don't bother if we're iconized.
If WindowState = vbMinimized Then Exit Sub

wdt1 = (ScaleWidth - SPLITTER_HEIGHT) * PercentageY
wdt2 = (ScaleWidth - SPLITTER_HEIGHT) - wdt1
hgt11 = (ScaleHeight - ohgt - SPLITTER_HEIGHT) * PercentageX1
hgt12 = (ScaleHeight - ohgt - SPLITTER_HEIGHT) - hgt11
hgt21 = (ScaleHeight - ohgt - SPLITTER_HEIGHT) * PercentageX2
hgt22 = (ScaleHeight - ohgt - SPLITTER_HEIGHT) - hgt21

Picture1.Move 0, ftop, wdt1, hgt11
Picture2.Move 0, ftop + hgt11 + SPLITTER_HEIGHT, wdt1, hgt12
Picture3.Move wdt1 + SPLITTER_HEIGHT, ftop, wdt2, hgt21
Picture4.Move wdt1 + SPLITTER_HEIGHT, ftop + hgt21 + SPLITTER_HEIGHT, wdt2, hgt22
End Sub
' Arrange the controls on the form.
Private Sub ArrangeControls()

On Error Resume Next

'Picture Box1 - scroll bars and map
    vsMap.Left = Picture1.ScaleWidth - vsMap.Width
    vsMap.top = 0
    
    hsMap.Left = 0
    hsMap.Width = vsMap.Left
    hsMap.top = Picture1.ScaleHeight - hsMap.Height
    vsMap.Height = hsMap.top
    
    'Set up map picture box
    picMap.Move 0, 0, vsMap.Left, hsMap.top
    
    'Set the map Max Button
    cmdMaxMap.Move vsMap.Left - cmdMaxMap.Width, 0

'Picture Box2 - Command Display and Text Entry
    txtEntry.Move vsMap.Width, Picture2.ScaleHeight - txtEntry.Height, Picture2.ScaleWidth
    lstCmdResult.Visible = (txtEntry.Height < Picture2.ScaleHeight)
    If lstCmdResult.Visible Then
        lstCmdResult.Move vsMap.Width, 0, hsMap.Width, Picture2.ScaleHeight - txtEntry.Height
    End If
    
'Picture Box 3 - Sector Marker
    lblSect.Move 0, 0, Picture3.ScaleWidth - cmdMaxUnit.Width, lblSect.Height
    TabStrip1.Move 0, lblSect.Height, Picture3.Width, Picture3.ScaleHeight - lblSect.Height
    
    'Put the Frames on the tabstrip
    Dim X As Integer
    With TabStrip1
        For X = 0 To 3
            Frame1(X).Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
        Next X
        grid1.Width = .ClientWidth
        grid1.Height = .ClientHeight - grid1.top
        rtbUnitList.Move grid1.Left, grid1.top, grid1.Width, grid1.Height
    End With
    cmdMaxUnit.Move Picture3.Width - cmdMaxUnit.Width, 0

'Picture Box 4 - Telegram Box
    'Update buttons

    'Set telegram text box
    rtbTelegram.Move 0, 0, Picture4.ScaleWidth, Picture4.ScaleHeight

    DrawHexes
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim wdt1 As Single

    wdt1 = (ScaleWidth - SPLITTER_HEIGHT) * PercentageY
    Dragging = True
    DragY = (X >= wdt1 And X <= wdt1 + SPLITTER_HEIGHT)
    DragX1 = (X < wdt1)
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
Dim wdt1 As Single

If MapMax Then Exit Sub

If Not Dragging Then
    wdt1 = (ScaleWidth - SPLITTER_HEIGHT) * PercentageY
    If X >= wdt1 And X <= wdt1 + SPLITTER_HEIGHT Then
        Me.MousePointer = vbSizeWE
    Else
        Me.MousePointer = vbSizeNS
    End If
Else
    If DragY Then
        PercentageY = X / ScaleWidth
        If PercentageY < 0 Then PercentageY = 0
        If PercentageY > 1 Then PercentageY = 1
    ElseIf DragX1 Then
        PercentageX1 = Y / ScaleHeight
        If PercentageX1 < 0 Then PercentageX1 = 0
        If PercentageX1 > 1 Then PercentageX1 = 1
    Else
        PercentageX2 = Y / ScaleHeight
        If PercentageX2 < 0 Then PercentageX2 = 0
        If PercentageX2 > 1 Then PercentageX2 = 1
    End If
    
    If PercentageY > 0.98 Then PercentageY = 0.98
    If PercentageY < 0.02 Then PercentageY = 0.02
    If PercentageX1 > 0.98 Then PercentageX1 = 0.98
    If PercentageX1 < 0.02 Then PercentageX1 = 0.02
    If PercentageX2 > 0.98 Then PercentageX2 = 0.98
    If PercentageX2 < 0.02 Then PercentageX2 = 0.02

    ResizePanels
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dragging = False
    ArrangeControls
End Sub

Private Sub Form_Resize()

' Dim x As Integer   removed efj 8/2003

'Don't try to redraw when form is minimized
If WindowState = vbMinimized Then
    If PromptUp Then
        PromptForm.WindowState = vbMinimized
    End If
    frmEmpCmd.WindowState = vbMinimized
    frmTelegram.WindowState = vbMinimized
    Exit Sub
End If

'Check to make sure its at least the minimum width
If Me.Width < (TabStrip1.Width + vsMap.Width * 5) Then
    Me.Width = TabStrip1.Width + vsMap.Width * 5
End If

'Check make sure its at least the minimum height
If Me.Height < TabStrip1.Height * 1.5 Then
    Me.Height = TabStrip1.Height * 1.5
End If

'Check make sure its at least the minimum height
If Me.Height < lstCmdResult.Height + txtEntry.Height * 6 Then
    Me.Height = lstCmdResult.Height + txtEntry.Height * 6
End If

sb1.Panels.Item("Hex").Text = "Sect:"

'Re-arrange the controls
ResizePanels
ArrangeControls

'if a prompt is up, reposition the prompt
If PromptUp Then
    PromptForm.WindowState = vbNormal
    'Put form in proper place
    PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
    PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height
End If

' Titleheight = Me.Height - Me.ScaleHeight   removed efj 8/2003

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim n As Integer

'save colors to the registry
For n = 1 To NUMBER_OF_BASIC_COLORS
    SaveSetting APPNAME, "Colors", "BasicColor" + CStr(n), BasicColors(n)
    SaveSetting APPNAME, "Colors", "BasicText" + CStr(n), BasicText(n)
Next n
For n = 1 To NUMBER_OF_PLAYER_COLORS
    SaveSetting APPNAME, "Colors", "PlayerColors" + CStr(n), PlayerColors(n)
    SaveSetting APPNAME, "Colors", "PlayerText" + CStr(n), PlayerText(n)
Next n

'save layout settings
If MapMax Or UnitMax Then
    PercentageY = MMPercentageY
    PercentageX1 = MMPercentageX1
    PercentageX2 = MMPercentageX2
End If

SaveSetting APPNAME, "Layout", "Magnification", Magnification
SaveSetting APPNAME, "Layout", "Percentage Y", PercentageY
SaveSetting APPNAME, "Layout", "Percentage X1", PercentageX1
SaveSetting APPNAME, "Layout", "Percentage X2", PercentageX2

frmOptions.SaveOptions '120203 rjk: Moved saving options to form

'Set Nations = Nothing
ShutDown = True
Unload frmEmpCmd
Unload frmEmpireLogin
Unload frmTelegram
If frmEmpCmd.bEmpireHub Then
    Unload frmChat
End If
frmChat.ControlFlashForm

If PromptUp Then
    Unload PromptForm
End If

Unload frmScript

'Display forms that are not unloaded
'Dim f As Object
'For Each f In Forms
'    MsgBox f.Name
'Next

End Sub

Public Sub DisplayUnitBox()
'rjk 05/13/03: Set the Sector Panel Tab Display to display the Unit Tab
'
'DisplaySectorPanel 3 rjk 05/13/03: switched over to tab in DisplaySectorPanel

DisplaySectorPanel '092803 rjk: Moved early to ensure unit is present for InRange option
If TabStrip1.Tabs(TabStrip1.Tabs.Count).key = "Unit" Then
    Set TabStrip1.SelectedItem = TabStrip1.Tabs("Unit")
    SelectTabContent
End If

End Sub

Public Sub DisplayFirstSectorPanel()
'Added rjk 05/13/03: Set the Sector Panel Tab Display to display the First Tab, normally the
'Census Tab display
'
Set TabStrip1.SelectedItem = TabStrip1.Tabs(1)
DisplaySectorPanel

End Sub

Private Sub DisplaySectorPanel()
' Changed 05/13/03: this function to determines which Tab to display and ensures the proper frame is displayed
' Also removed the argument and made private, use the DisplayUnitBox or DisplayFirstSectorPanel functions instead
'
Dim CurrentTab As String

'rjk 05/14/03: Determine what tab the Sector Panel was showing later decision on what to Tab show
CurrentTab = TabStrip1.Tabs(TabStrip1.SelectedItem.Index).key

TabStrip1.Tabs.Clear 'Remove all Tabs
'Added the appropriate Tabs for the sector
If DisplaySelect <> udUNKNOWN Then
    TabStrip1.Tabs.Add 1, "Unit", "Units" 'Add first it will be at the end of the list
End If
If SelSectType = SEL_SECT_OWN Then
    TabStrip1.Tabs.Add 1, "Census", "Census"
    TabStrip1.Tabs.Add 2, "Dist", "Dist."
ElseIf SelSectType = SEL_SECT_ENEMY Then
    TabStrip1.Tabs.Add 1, "Enemy", "Census"
ElseIf SelSectType = SEL_SECT_SEA Then
    TabStrip1.Tabs.Add 1, "Sea", "Census"
Else
    If DisplaySelect = udUNKNOWN Then
        TabStrip1.Tabs.Add 1, "Unknown", vbNullString
    End If
End If

'rjk 05/14/03: If the Sector Panel can not show the current tab then go back to the first tab.
Select Case CurrentTab
    Case "Unit"
        If DisplaySelect <> udUNKNOWN Then
            Set TabStrip1.SelectedItem = TabStrip1.Tabs(TabStrip1.Tabs.Count)
        Else
            Set TabStrip1.SelectedItem = TabStrip1.Tabs(1) 'always select the first item if no unit tab
        End If
    Case "Dist"
        If TabStrip1.Tabs.Count >= 2 Then
            If TabStrip1.Tabs(2).key <> "Dist" Then
                Set TabStrip1.SelectedItem = TabStrip1.Tabs(1) 'always select the first item if no dist tab
            End If
        Else
            Set TabStrip1.SelectedItem = TabStrip1.Tabs(1) 'always select the first item if no dist tab
        End If
    Case Else
        If TabStrip1.SelectedItem.Index > TabStrip1.Tabs.Count Then
            Set TabStrip1.SelectedItem = TabStrip1.Tabs(1) 'always select the first item if not enough tabs
        End If
End Select

SelectTabContent 'Match the frame to the Selected Tab
End Sub

Private Sub SelectTabContent()
'Added rjk 05/13/03: Display the appropriate frame for the tab selected.
'drk 05/16/03: make everything hidden by default and only show what we need.  much fewer lines of code.
    
Frame1(0).Visible = False
Frame1(1).Visible = False
Frame1(2).Visible = False
Frame1(3).Visible = False
Frame1(4).Visible = False

Select Case TabStrip1.SelectedItem.key
    Case "Census"
            Frame1(1).Visible = True
    Case "Dist"
            Frame1(2).Visible = True
    Case "Unit"
            Frame1(3).Visible = True
    Case "Sea"
            Frame1(4).Visible = True
    Case "Enemy"
            Frame1(0).Visible = True
    Case "Unknown"
End Select
End Sub

Public Function IsCensusTabActive() As Boolean
If TabStrip1.SelectedItem.key = "Census" Then
    IsCensusTabActive = True
Else
    IsCensusTabActive = False
End If
End Function

Public Function IsAnyTabActive() As Boolean
If TabStrip1.SelectedItem.key = "Unknown" Then
    IsAnyTabActive = False
Else
    IsAnyTabActive = True
End If
End Function

Private Sub Frame1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
Me.MousePointer = vbDefault
End Sub

Private Sub grid1_DblClick()
Dim iRow As Integer
Dim iCol As Integer
Dim ix As Integer
Dim iy As Integer

iRow = grid1.MouseRow
iCol = grid1.MouseCol
If (iRow = 1 Or iCol = 0) And ((iRow > 0) And (iRow < grid1.Rows - 1)) Then
    Select Case DisplaySelect
    Case udSHIP
        ix = grid1.TextMatrix(iRow, 8)
        iy = grid1.TextMatrix(iRow, 9)
    Case udLAND
        ix = grid1.TextMatrix(iRow, 9)
        iy = grid1.TextMatrix(iRow, 10)
    Case udPLANE
        ix = grid1.TextMatrix(iRow, 7)
        iy = grid1.TextMatrix(iRow, 8)
    Case udNUKE, udENEMY
        ix = grid1.TextMatrix(iRow, 2)
        iy = grid1.TextMatrix(iRow, 3)
    Case Else
        Exit Sub
    End Select
    MoveMap ix, iy
End If
End Sub

Private Sub grid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim mx As Integer
Dim my As Integer
Dim strTemp As String
Dim StartX As Integer
Dim Finishx As Integer
Dim x1 As Integer

GridSelection = True
'Unit Pop up window
If Button = vbRightButton Then
    If Not PromptUp Then        'Don't prompt when already prompted
        GridSelection = True
        StartX = grid1.row
        Finishx = grid1.RowSel
        
        If Finishx < StartX Then
            x1 = StartX
            StartX = Finishx
            Finishx = x1
        End If

        'go thru selected rows, and add units to select String
        For x1 = StartX To Finishx
            '080304 rjk: Added row check, ensure the total line is not checked,
            '            also allows the selection of a ship actually selecting the line
            If x1 < grid1.Rows - 1 Then
                strTemp = strTemp + Trim$(grid1.TextMatrix(x1, 1)) + "/"
            End If
        Next x1
        If Len(strTemp) > 0 Then
            strTemp = Left$(strTemp, Len(strTemp) - 1)
        Else
            mx = grid1.MouseRow
            my = grid1.MouseCol
            If (my = 1 Or my = 0) And ((mx > 0) And (mx < grid1.Rows - 1)) Then
                strTemp = Trim$(grid1.TextMatrix(mx, 1))
            End If
        End If
        PopUpMenuSource = DMAP_PUMS_GRID
        'Switched to DisplaySelect from BuildType, BuildType going private/local
        Select Case DisplaySelect
        Case udSHIP
            strShipid = strTemp
            PopupMenu mnuShip
        Case udPLANE
            strPlaneid = strTemp
            PopupMenu mnuPlane
        Case udLAND
            strLandid = strTemp
            PopupMenu mnuLand
        Case udNUKE
            strNukeid = strTemp
            PopupMenu mnuNuke
        End Select
    End If
End If
End Sub

Private Sub grid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim mx As Integer
Dim my As Integer
Dim id As Integer
Dim strTemp As String

Me.MousePointer = vbDefault

mx = grid1.MouseRow
my = grid1.MouseCol
strTemp = vbNullString

If (my = 1 Or my = 0) And ((mx > 0) And (mx < grid1.Rows - 1)) Then
    id = CInt(Trim$(grid1.TextMatrix(mx, 1)))
    '101403 rjk: Switched to DisplaySelect from BuildType, BuildType going local/private
    'strTemp = grid1.TextMatrix(mx, my) + "  " + BuildMissionString(BuildType, id)
    Select Case DisplaySelect
    Case udSHIP
        strTemp = grid1.TextMatrix(mx, 1) + "  " + BuildMissionString("s", id)
    Case udPLANE
        strTemp = grid1.TextMatrix(mx, 1) + "  " + BuildMissionString("p", id)
    Case udLAND
        strTemp = grid1.TextMatrix(mx, 1) + "  " + BuildMissionString("l", id)
    Case udNUKE
        strTemp = grid1.TextMatrix(mx, 1) + "  " + BuildMissionString("n", id)
    End Select
End If
sb1.Panels.Item("Hex").Text = strTemp

End Sub

Private Sub lblSect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = vbDefault
End Sub

Private Sub lstCmdResult_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = vbDefault
End Sub

Private Sub mnuAttAttack_Click(Index As Integer)
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Select Case Index
Case 0 'Probe
    frmEmpCmd.SubmitEmpireCommand "at1 /mil=0; /probe=y; /unt=N;", False
    ' 092303 rjk: Switched to SelX and SelY for top level menu access
    'frmEmpCmd.SubmitEmpireCommand "attack " + SectorString(MxPos, MyPos) + " 0 0 0 0", True
    frmEmpCmd.SubmitEmpireCommand "attack " + SectorString(SelX, SelY) + " n n n n", True
    frmEmpCmd.SubmitEmpireCommand "at2", False
        
Case 1 'Attack
    'Setup Designation dialog and execute
    frmPromptAttack.Label2 = "Attack"
    ' 092303 rjk: Switched to SelX and SelY for top level menu access
    'frmPromptAttack.txtOrigin = SectorString(MxPos, MyPos)
    frmPromptAttack.txtOrigin = SectorString(SelX, SelY)
    'Put form in proper place
    frmPromptAttack.Left = frmDrawMap.Left + (frmDrawMap.Width - frmPromptAttack.Width) \ 2
    frmPromptAttack.top = frmDrawMap.top + frmDrawMap.Height - frmPromptAttack.Height
    Set PromptForm = frmPromptAttack
    PromptUp = True
    'Dialog box will take it from here
    frmPromptAttack.Show vbModeless

Case 2 'Assault
    'Setup Designation dialog and execute
    frmPromptAttack.Label2 = "Assault"
    ' 092303 rjk: Switched to SelX and SelY for top level menu access
    'frmPromptAttack.txtOrigin = SectorString(MxPos, MyPos)
    frmPromptAttack.txtOrigin = SectorString(SelX, SelY)
    'Put form in proper place
    frmPromptAttack.Left = frmDrawMap.Left + (frmDrawMap.Width - frmPromptAttack.Width) \ 2
    frmPromptAttack.top = frmDrawMap.top + frmDrawMap.Height - frmPromptAttack.Height
    Set PromptForm = frmPromptAttack
    PromptUp = True
    
    'Dialog box will take it from here
    frmPromptAttack.Show vbModeless

Case 3 'Bomb
    Set PromptForm = frmPromptBomb
    'Setup op sector hex
    ' 092303 rjk: Switched to SelX and SelY for top level menu access
    'PromptForm.txtPath.Text = SectorString(MxPos, MyPos)
    PromptForm.txtPath.Text = SectorString(SelX, SelY)
    PromptForm.ShowRange = True
    'Put form in proper place
    PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
    PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height
    PromptUp = True
    'Show Prompt
    PromptForm.Show vbModeless
    PromptForm.txtUnit.SetFocus

Case 4 'Paradrop
    Set PromptForm = frmPromptRecon
    'Setup op sector hex
    PromptForm.Label2 = "Paradrop"
    ' 092303 rjk: Switched to SelX and SelY for top level menu access
    'PromptForm.txtPath.Text = SectorString(MxPos, MyPos)
    PromptForm.txtPath.Text = SectorString(SelX, SelY)
    PromptForm.ShowRange = True
    'Put form in proper place
    PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
    PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height
    PromptUp = True
    'Show Prompt
    PromptForm.Show vbModeless
    PromptForm.txtUnit.SetFocus

Case 5 'Missile Strike
    Set PromptForm = frmPromptLaunch
    'Setup op sector hex
    PromptForm.strCmd = "launch"
    PromptForm.Label2 = "Launch"
    ' 092303 rjk: Switched to SelX and SelY for top level menu access
    'PromptForm.txtPath.Text = SectorString(MxPos, MyPos)
    PromptForm.txtPath.Text = SectorString(SelX, SelY)
    PromptForm.ShowRange = True
    'Put form in proper place
    PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
    PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height
    PromptUp = True
    'Show Prompt
    PromptForm.Show vbModeless
    PromptForm.txtUnit.SetFocus

Case 6 'Center Map 112503 rjk: Moved mnuAttCenter to here to save a control
    MoveMap SelX, SelY
End Select
End Sub

Private Sub mnuCopyReportWindow_Click()
frmReport.Caption = "Telegram/Server Output"
frmReport.AddReportLine rtbTelegram.Text
frmReport.Show
End Sub

Private Sub mnuHelp_Click(Index As Integer)
Dim strSubject As String

Select Case Index
Case 1 'Command List
    If frmOptions.bLocalHelp Then
        HtmlHelp hwnd, App.path + "/help/winace.chm", HH_DISPLAY_TOPIC, ByVal "all.html"
    Else
        frmEmpCmd.SubmitEmpireCommand "list", True
    End If
Case 2 'Contents
    HtmlHelp hwnd, App.path + "/help/winace.chm", HH_DISPLAY_TOC, ByVal 0
Case 3 'Index
    HtmlHelp hwnd, App.path + "/help/winace.chm", HH_DISPLAY_INDEX, ByVal ""
Case 4 'Info
    strSubject = InputBox("Enter Subject", "Empire Help")
    If frmOptions.bLocalHelp Then
        HtmlHelp hwnd, App.path + "/help/winace.chm", HH_DISPLAY_TOPIC, ByVal strSubject + ".html"
    Else
        frmEmpCmd.SubmitEmpireCommand "info " + strSubject, True
    End If
Case 5 'Quick Start Guide
    HtmlHelp hwnd, App.path + "/help/winace.chm", HH_DISPLAY_TOPIC, ByVal "Quick Start.mht"
Case 6 'User Guide
    HtmlHelp hwnd, App.path + "/help/winace.chm", HH_DISPLAY_TOPIC, ByVal "WinACE User's Guide.htm"
Case 7 'About
    frmAbout.Show vbModal
End Select
End Sub

Private Sub mnuLandStart_Click(Index As Integer)
Select Case Index
Case 0 'Start
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptCensus
    frmPromptCensus.Label2 = "Start"
    frmPromptCensus.cmbType.ListIndex = 2
    ShowUnitPrompt "land"
Case 1 'Stop
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptCensus
    frmPromptCensus.Label2 = "Stop"
    frmPromptCensus.cmbType.ListIndex = 2
    ShowUnitPrompt "land"
End Select
End Sub

Private Sub mnuNukeStart_Click(Index As Integer)
Select Case Index
Case 0 'Start
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptCensus
    frmPromptCensus.Label2 = "Start"
    frmPromptCensus.cmbType.ListIndex = 4
    ShowUnitPrompt "nuke"
Case 1 'Stop
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptCensus
    frmPromptCensus.Label2 = "Stop"
    frmPromptCensus.cmbType.ListIndex = 4
    ShowUnitPrompt "nuke"
End Select
End Sub

Private Sub mnuSectCmdEstimate_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If
Set PromptForm = frmToolProduction

'Setup Designation dialog and execute
frmToolProduction.txtOrigin = SectorString(SelX, SelY)
'Put form in proper place
PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height
PromptUp = True

'Dialog box will take it from here
PromptForm.Show vbModeless
PromptForm.txtOrigin.SetFocus
End Sub

Private Sub mnuTelegramCenter_Click()
MoveMap iTelegramSelX, iTelegramSelY
End Sub

Private Sub mnuUnitAttAttack_Click(Index As Integer)
mnuAttAttack_Click (Index)
End Sub

Private Sub mnuClearSeaIntelligence_Click()
mnuClearSectorIntelligence_Click

End Sub

Private Sub mnuClearSectorIntelligence_Click()
If rsEnemySect.BOF And rsEnemySect.EOF Then
    Exit Sub
End If

rsEnemySect.Seek "=", SelX, SelY
If Not rsEnemySect.NoMatch Then
    rsEnemySect.Delete
    DrawHexes
End If
End Sub

Private Sub mnuDietyOption_Click(Index As Integer)
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Select Case Index
    Case 1, 2, 3 'Give and set resource
    Set PromptForm = frmPromptLoad
    Select Case Index
    Case 1
        PromptForm.strCmd = "give"
    Case 2
        PromptForm.strCmd = "setsector"
    Case 3
        PromptForm.strCmd = "setresource"
    Case 1
    End Select
'    PromptForm.strSector = SectorString(MxPos, MyPos) 092103 rjk: Switched to SelX and SelY for top level menu
    PromptForm.strSector = SectorString(SelX, SelY)
    'Show Prompt
    ShowUnitPrompt "ship"
Case 4
    If VersionCheck(4, 2, 21) = VER_LESS Then
        frmPromptCensus.Label2.Caption = "Hidden"
    Else
        frmPromptCensus.Label2.Caption = "Peek"
    End If
    mnuReportCensus_Click
End Select
End Sub

Private Sub mnuDiploRelationsgrid_Click()
frmRelationsGrid.Show vbModeless, Me
frmRelationsGrid.go
End Sub

Private Sub mnuFileOption_Click(Index As Integer)
Select Case Index
Case 1 'Import Telegrams
    ImportTelegrams
Case 2 'Import Intelligence
    ImportIntelligenceOffset vbNullString
Case 3 'Import sea information
    ImportSeaInformation
Case 5 'Export All Telegrams
    ExportTelegrams True
Case 6 'Export Recent Telegrams
    ExportTelegrams False
Case 7 'Export Intelligence
    '103003 rjk: Added its own form for selecting parameters.
    Set PromptForm = frmPromptExportIntelligence
    PromptUp = True
    PromptForm.Show vbModeless
Case 8 'Export Nation
    ExportNation
Case 9 'Export Sea Information
    ExportSeaInformation
Case 11 'Copy Game Database
    frmPromptCopyGameDB.Show vbModal
Case 13 'Clear
    '112903 rjk: Created a separate form for clearing.
    frmClear.Show vbModal
Case 14 'Print
    frmPrintMap.Show
Case 15 'Exit
    Unload Me
End Select
    
End Sub

Private Sub mnuGenCmd_Click(Index As Integer)
Select Case Index
Case 0 'Break Sanctuary
    frmEmpCmd.SubmitEmpireCommand "break", True

Case 1 'Change
    'Put form in proper place
    frmPromptChange.Left = frmDrawMap.Left + (frmDrawMap.Width - frmPromptChange.Width) \ 2
    frmPromptChange.top = frmDrawMap.top + frmDrawMap.Height - frmPromptChange.Height
    Set PromptForm = frmPromptChange
    PromptUp = True
    'Dialog box will take it from here
    frmPromptChange.Show vbModeless

Case 2 'Realm
    frmPromptConvert.Label2.Caption = "Realm"
    frmPromptConvert.Label1.Caption = "Sectors:"
    mnuSectCmdConvert_Click

Case 3 'Territory
    'check for current prompt
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptTerritory
    PromptUp = True
    'Put form in proper place
    PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
    PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height
    'Dialog box will take it from here
    PromptForm.Show vbModeless

Case 4 'Set Owner
    'check for current prompt
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptIsland
    PromptUp = True
    'Put form in proper place
    PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
    PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height
    PromptForm.Label2 = "Set Owner"
    PromptForm.cbNations.Visible = True
    'Dialog box will take it from here
    PromptForm.Show vbModeless
End Select
End Sub

Private Sub mnuGotoOptions_Click(Index As Integer)
On Error Resume Next
Dim sx As Integer
Dim sy As Integer
        
If Not MapValid Then Exit Sub
        
Select Case Index
    Case 1 ' Last event
        sx = EventX
        sy = EventY
    Case 2, 3, 4, 5, 6 ' Jump Point 1-5
        If Not ParseSectors(sx, sy, rsNation.Fields("JumpPoint" + CStr(Index - 1))) Then
            Exit Sub
        End If
    Case Else
        Exit Sub
End Select

MoveMap sx, sy
End Sub

Private Sub mnuLandMisc_Click(Index As Integer)
Select Case Index
Case 0 'Army
    'check for current prompt
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptFleet
    PromptForm.Label2.Caption = "Army"
    If Len(strLandid) > 0 Then
        PromptForm.txtUnit.Text = strLandid
    End If
    PromptForm.LoadCombo
    'Show Prompt
    ShowUnitPrompt "land"
Case 1 'Center Map
    Dim sx As Integer, sy As Integer
    Dim slist As Variant
    Dim hldIndex As String
    
    If strLandid <> "" Then
        slist = ParseUnitString(strLandid)
        If IsArray(slist) Then
            hldIndex = rsLand.Index
            rsLand.Index = "PrimaryKey"
            rsLand.Seek "=", slist(1)
            If Not rsLand.NoMatch Then
                MoveMap rsLand.Fields("x"), rsLand.Fields("y")
            End If
            rsLand.Index = hldIndex
        End If
    End If
Case 2 'Fortify
    'check for current prompt
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptSail
    PromptUp = True
    PromptForm.strCmd = "fortify"
    PromptForm.Label2 = "Fortify"
    PromptForm.Label1 = "mobility"
    'Show Prompt
    ShowUnitPrompt "land"
Case 3 'Mine
    'check for current prompt
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptLook
    PromptForm.strCmd = "lmine"
    PromptForm.Label2 = "Land Mine"
    'Show Prompt
    ShowUnitPrompt "land"
Case 4 'Range
    'check for current prompt
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptSail
    PromptForm.strCmd = "lrange"
    PromptForm.Label2 = "Range"
    PromptForm.Label1 = "land radius"
    'Show Prompt
    ShowUnitPrompt "land"
Case 5 'Retreat
    'check for current prompt
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptSail
    PromptForm.strCmd = "lretreat"
    PromptForm.Label2 = "Retreat"
    PromptForm.Label1 = "path"
    PromptForm.Label3.Visible = True
    PromptForm.Combo1.Visible = True
    LoadRetreatBox PromptForm.Combo1
    PromptForm.Combo1.ListIndex = 0
    'Show Prompt
    ShowUnitPrompt "land"
End Select
End Sub

Private Sub mnuLandWarload_Click()
' added by thomas lullier
' load max mil/shell/gun to land units

'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptWarload
PromptForm.strCmd = "lwarload"
'PromptForm.strSector = SectorString(MxPos, MyPos) 09/14/03 rjk: strSector was removed 2.2.13, not used in function
'PromptForm.SetSectorLabels - thomas lullier - what is it used for ?
'Show Prompt
ShowUnitPrompt "land"

End Sub

Private Sub mnuMarketBuy_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptCom
PromptUp = True
PromptForm.Label2.Caption = "Buy"
PromptForm.strSector = SectorString(SelX, SelY)

'Dialog box will take it from here
PromptForm.Show vbModeless
End Sub

Private Sub mnuMarketMarket_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptCom
PromptUp = True
PromptForm.Label2.Caption = "Market"

'Dialog box will take it from here
PromptForm.Show vbModeless
End Sub

Private Sub mnuMarketReset_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptCom
PromptUp = True
PromptForm.Label2.Caption = "Reset"

'Dialog box will take it from here
PromptForm.Show vbModeless
End Sub

Private Sub mnuMarketSell_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptCom
PromptUp = True
PromptForm.Label2.Caption = "Sell"
PromptForm.strSector = SectorString(SelX, SelY)

'Dialog box will take it from here
PromptForm.Show vbModeless
End Sub

Private Sub mnuMarketSet_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptScrap
PromptUp = True
PromptForm.Label2.Caption = "Set"
PromptForm.strCmd = "set"
PromptForm.Option1 = True

'Dialog box will take it from here
ShowUnitPrompt "ship"
End Sub

Private Sub mnuNukeList_Click()
' 092303 rjk: Switched to SelX and SelY for top level menu access
'frmEmpCmd.SubmitEmpireCommand "nuke " + SectorString(MxPos, MyPos), True
frmEmpCmd.SubmitEmpireCommand "nuke " + SectorString(SelX, SelY), True
End Sub

Private Sub mnuNukeTrans_Click()
Set PromptForm = frmPromptTrans
PromptUp = True
    
'Setup op sector hex
' 092303 rjk: Switched to SelX and SelY for top level menu access
'PromptForm.strSector = SectorString(MxPos, MyPos)
PromptForm.strSector = SectorString(SelX, SelY)
frmPromptTrans.Option2 = True

'Put form in proper place
PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height
PromptUp = True

'Dialog box will take it from here
ShowUnitPrompt "nuke"
On Error Resume Next
PromptForm.txtUnit.SetFocus
End Sub

Private Sub mnuPlaneMissile_Click(Index As Integer)
Select Case Index
Case 0 'Arm
    'check for current prompt
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptSail
    PromptForm.strCmd = "arm"
    PromptForm.Label2 = "Arm"
    If VersionCheck(4, 3, 3) <> VER_LESS Then
        PromptForm.Label1 = "nuke id"
    Else
        PromptForm.Label1 = "nuke type"
    End If
    'Show Prompt
    ShowUnitPrompt "plane"
Case 1 'Disarm
    'check for current prompt
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptLook
    PromptForm.strCmd = "disarm"
    PromptForm.Label2 = "Disarm"
    'Show Prompt
    ShowUnitPrompt "plane"
Case 2 'Harden
    'check for current prompt
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptSail
    PromptForm.strCmd = "harden"
    PromptForm.Label2 = "Harden"
    PromptForm.Label1 = "amount"
    'Show Prompt
    ShowUnitPrompt "plane"
Case 3 'Launch
    'check for current prompt
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptLaunch
    PromptForm.strCmd = "launch"
    PromptForm.Label2 = "Lanuch"
    'Show Prompt
    ShowUnitPrompt "plane"
End Select
End Sub

Private Sub mnuPlaneOther_Click(Index As Integer)
Select Case Index
Case 0 'Center Map
    Dim sx As Integer, sy As Integer
    Dim slist As Variant
    Dim hldIndex As String
    
    If strPlaneid <> "" Then
        slist = ParseUnitString(strPlaneid)
        If IsArray(slist) Then
            hldIndex = rsPlane.Index
            rsPlane.Index = "PrimaryKey"
            rsPlane.Seek "=", slist(1)
            If Not rsPlane.NoMatch Then
                MoveMap rsPlane.Fields("x"), rsPlane.Fields("y")
            End If
            rsPlane.Index = hldIndex
        End If
    End If
Case 1 'Sat
    'check for current prompt
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptSail
    PromptForm.strCmd = "satellite"
    PromptForm.Label2 = "Satellite"
    PromptForm.Label1 = "report"
    'Show Prompt
    ShowUnitPrompt "plane"
Case 2 'Scrap
    'check for current prompt
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptScrap
    PromptForm.strCmd = "scrap"
    PromptForm.Label2 = "Scrap"
    PromptForm.Option3 = True
    'Show Prompt
    ShowUnitPrompt "plane"
Case 3 'Scuttle
    'check for current prompt
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptScrap
    PromptForm.strCmd = "scuttle"
    PromptForm.Label2 = "Scuttle"
    PromptForm.Option3 = True
    'Show Prompt
    ShowUnitPrompt "plane"
Case 4 'Start
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptCensus
    frmPromptCensus.Label2 = "Start"
    frmPromptCensus.cmbType.ListIndex = 3
    ShowUnitPrompt "plane"
Case 5 'Stop
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptCensus
    frmPromptCensus.Label2 = "Stop"
    frmPromptCensus.cmbType.ListIndex = 3
    ShowUnitPrompt "plane"
Case 6 'Upgrade
    'check for current prompt
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptScrap
    PromptForm.strCmd = "upgrade "
    PromptForm.Label2 = "Upgrade"
    PromptForm.Option3 = True
    'Show Prompt
    ShowUnitPrompt "plane"
Case 7 'Wing
    'check for current prompt
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptFleet
    PromptForm.Label2.Caption = "Wing"
    If Len(strPlaneid) > 0 Then
        PromptForm.txtUnit.Text = strPlaneid
    End If
    PromptForm.LoadCombo
    'Show Prompt
    ShowUnitPrompt "plane"
End Select
End Sub

Public Sub mnuRefresh_Click(Index As Integer)
Select Case Index
Case 1 'Update Sectors
    frmEmpCmd.ForceUpdates = True
    'Send Code to start database update
    frmEmpCmd.SubmitEmpireCommand "db1", False
    'Request an update of Sector information
    frmEmpCmd.SubmitEmpireCommand "cs1", False
    GetSectors True
    '101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
    GetCurrentStrength
    GetOContent '110605 rjk: Added ability to get OContent for Sea Sectors
    'send Code to end database update
    frmEmpCmd.SubmitEmpireCommand "db2", False
Case 2 'Update BMap
    frmEmpCmd.ForceUpdates = True
    If frmOptions.bSuppressBmapRefresh Then
        MsgBox ("Bmap updates are suppressed. Change option and then request again")
        Exit Sub
    End If
    'Clear Bmap file
    DeleteAllRecords rsBmap
    'Send Code to start database update
    frmEmpCmd.SubmitEmpireCommand "db1", False
    'Request an update of map information
    frmEmpCmd.SubmitEmpireCommand "map *", False
    frmEmpCmd.SubmitEmpireCommand "bmap *", False
    'send Code to end database update
    frmEmpCmd.SubmitEmpireCommand "db2", False
Case 3 'Update Units
    'check for current prompt
    If PromptUp Then
        Exit Sub
    End If
    'Set parms to make sure all dumps are done
    frmEmpCmd.ForceUpdates = True
    frmEmpCmd.HaveLands = True
    frmEmpCmd.HaveShips = True
    frmEmpCmd.HavePlanes = True
    frmEmpCmd.HaveNukes = True '092503 rjk: Added to Units option
    
    On Error Resume Next        'MoveFirst will give error on empty dataset
    'Send Code to start database update
    frmEmpCmd.SubmitEmpireCommand "db1", False
    GetLost True
    frmEmpCmd.SubmitEmpireCommand "cu1", False
    'Request an update of Plane information
    GetPlanes True
    'Request an update of land information
    GetLandUnits True
    'Request an update of ship information
    GetShips True
    'Request an update of nuke information 092503 rjk: Added to Units option
    GetNukes True
    'send Code to end database update
    frmEmpCmd.SubmitEmpireCommand "db2", False
Case 4 'Update Players List
    frmEmpCmd.SubmitEmpireCommand "db1", False
    'Request an update of player information
    If VersionCheck(4, 3, 4) = VER_LESS And Not frmOptions.bSPAtlantis Then
        GetNation
        GetCountryList
    Else
        GetCountryList
        GetNation
    End If
    'send Code to end database update
    frmEmpCmd.SubmitEmpireCommand "db2", False
Case 5 'Update All
    frmEmpCmd.SubmitEmpireCommand "db1", False
    frmEmpCmd.SubmitEmpireCommand "cs1", False
    frmEmpCmd.SubmitEmpireCommand "cu1", False
    GetUpdate False
    If VersionCheck(4, 3, 10) <> VER_LESS Then
        frmEmpCmd.SubmitEmpireCommand "xdump game *", False
    End If
    
    GetSectors True
    '101703 rjk: Added Strength fields to Sector database
    GetCurrentStrength '112003 rjk: Added option to control
    GetOContent '110605 rjk: Added ability to get OContent for Sea Sectors
    GetPlanes True
    GetLandUnits True
    GetShips True
    GetNukes True
    GetLost True
    If VersionCheck(4, 3, 4) = VER_LESS And Not frmOptions.bSPAtlantis Then
        GetNation
        GetCountryList
    Else
        GetCountryList
        GetNation
    End If
    frmToolShow.cmdRefreshAll_Click
    frmEmpCmd.SubmitEmpireCommand "map *", False '100103 rjk: Added to all options
    frmEmpCmd.SubmitEmpireCommand "bmap *", False '100103 rjk: Added to all options
    If VersionCheck(4, 3, 0) = VER_LESS Then
        GetOrders False 'inside if because GetShips has already been done
        DeleteAllRecords rsMissions '100103 rjk: Added to all options
        frmEmpCmd.SubmitEmpireCommand "mission land * q", False '100103 rjk: Added to all options
        frmEmpCmd.SubmitEmpireCommand "mission ship * q", False '100103 rjk: Added to all options
        frmEmpCmd.SubmitEmpireCommand "mission plane * q", False '100103 rjk: Added to all options
    End If
    frmEmpCmd.SubmitEmpireCommand "db2", False
Case 6 'Update Changed
    frmEmpCmd.SubmitEmpireCommand "db1", False
    GetSectors
    '101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
    GetCurrentStrength tsSectors
    GetOContent '110605 rjk: Added ability to get OContent for Sea Sectors
    GetPlanes
    GetLandUnits
    GetShips
    GetNukes
    GetLost
    frmEmpCmd.SubmitEmpireCommand "db2", False
Case 8 'Update Orders
    GetOrders True
Case 9 'Update Missions
    frmEmpCmd.SubmitEmpireCommand "db1", False
    If VersionCheck(4, 3, 0) = VER_LESS Then
        DeleteAllRecords rsMissions
        frmEmpCmd.SubmitEmpireCommand "mission land * q", False
        frmEmpCmd.SubmitEmpireCommand "mission ship * q", False
        frmEmpCmd.SubmitEmpireCommand "mission plane * q", False
        frmEmpCmd.SubmitEmpireCommand "db2", False
    Else
        GetLandUnits True
        GetShips True
        GetPlanes True
    End If
End Select
End Sub

Private Sub mnuReportRoute_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptCom
PromptUp = True
PromptForm.Label2.Caption = "Route"
PromptForm.strSector = SectorString(SelX, SelY)

'Dialog box will take it from here
PromptForm.Show vbModeless
End Sub

Private Sub mnuReportShow_Click()
frmToolShow.Show
End Sub

Private Sub mnuReportSpy_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptRadar
PromptUp = True

PromptForm.Label2.Caption = "Spy"
PromptForm.txtOrigin.Text = "*"
'Put form in proper place
PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height

'Dialog box will take it from here
PromptForm.Show vbModeless
End Sub

Private Sub mnuSeaCenter_Click()
'092303 rjk: Switched to SelX and SelY for top level menu access
'MoveMap MxPos, MyPos
MoveMap SelX, SelY
End Sub

Private Sub mnuSectCmdAssault_Click()
mnuAttAttack_Click 2
End Sub

Private Sub mnuSectCmdCenter_Click()
'092303 rjk: Switched to SelX and SelY for top level menu access
'MoveMap MxPos, MyPos
MoveMap SelX, SelY
End Sub

Private Sub mnuSectCmdMoveAll_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptLoad
PromptForm.strCmd = "move"
'PromptForm.strSector = SectorString(MxPos, MyPos) 092103 rjk: Switched to SelX and SelY for top level menu
PromptForm.strSector = SectorString(SelX, SelY)
'Show Prompt
'ShowUnitPrompt "ship"  'removed drk@unxsoft.com 2/26/02. replaced with the more sane code below...
PromptUp = True
PromptForm.Show vbModeless
PromptForm.txtMultDest.SetFocus '111903 rjk: Multiple Sector Selection
End Sub

Private Sub mnuSectCmdTerr_Click()
If PromptUp Then
    Exit Sub
End If
Set PromptForm = frmPromptSimpleTerritory
PromptUp = True

frmPromptSimpleTerritory.txtMultOrigin.Text = SectorString(SelX, SelY)

'Put form in proper place
PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height
'Dialog box will take it from here
PromptForm.Show vbModeless
End Sub

Private Sub mnuSetJumpPoint_Click(Index As Integer)
If rsNation.BOF And rsNation.EOF Then
    Exit Sub
End If

rsNation.MoveFirst
rsNation.Edit
rsNation.Fields("JumpPoint" + CStr(Index)).Value = _
    SectorString(SelX, SelY)
rsNation.Update
End Sub

Private Sub mnuShipFollow_Click() '092303 rjk: Added, frmPromptFollow exists but no call to
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptFollow
ShowUnitPrompt "ship"

End Sub

Private Sub mnuShipMisc_Click(Index As Integer)
Select Case Index
Case 0 'Center Map 112503 rjk Added the ability to center the map to the first ship in the list
    Dim sx As Integer, sy As Integer
    Dim slist As Variant
    Dim hldIndex As String
    
    If strShipid <> "" Then
        slist = ParseUnitString(strShipid)
        If IsArray(slist) Then
            hldIndex = rsShip.Index
            rsShip.Index = "PrimaryKey"
            rsShip.Seek "=", slist(1)
            If Not rsShip.NoMatch Then
                MoveMap rsShip.Fields("x"), rsShip.Fields("y")
            End If
            rsShip.Index = hldIndex
        End If
    End If
Case 1 'Fleet
    'check for current prompt
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptFleet
    PromptUp = True
    If Len(strShipid) > 0 Then
        PromptForm.txtUnit.Text = strShipid
    End If
    PromptForm.LoadCombo
    'Show Prompt
    ShowUnitPrompt "ship"
Case 2 'Fuel
    'check for current prompt
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptFuel
    PromptForm.Option1.Value = True
    'Show Prompt
    ShowUnitPrompt "ship"
Case 3 'Name
    'check for current prompt
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptSail
    PromptForm.strCmd = "name"
    PromptForm.Label2 = "Name"
    PromptForm.Label1 = "to"
    'Show Prompt
    ShowUnitPrompt "ship"
Case 4 'Radar
    'check for current prompt
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptLook
    PromptForm.strCmd = "radar"
    PromptForm.Label2 = "Radar"
    'Show Prompt
    ShowUnitPrompt "ship"
Case 5 'Scrap
    'check for current prompt
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptScrap
    PromptForm.strCmd = "scrap "
    PromptForm.Label2 = "Scrap"
    PromptForm.Option1 = True
    'Show Prompt
    ShowUnitPrompt "ship"
Case 6 'Scuttle
    'check for current prompt
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptScrap
    PromptForm.strCmd = "scuttle"
    PromptForm.Label2 = "Scuttle"
    PromptForm.Option1 = True
    'Show Prompt
    ShowUnitPrompt "ship"
Case 7 'Start
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptCensus
    frmPromptCensus.Label2 = "Start"
    frmPromptCensus.cmbType.ListIndex = 1
    ShowUnitPrompt "ship"
Case 8 'Stop
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptCensus
    frmPromptCensus.Label2 = "Stop"
    frmPromptCensus.cmbType.ListIndex = 1
    ShowUnitPrompt "ship"
Case 9 'Tend
    'check for current prompt
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptLoad
    PromptForm.strCmd = "tend"
    'Show Prompt
    ShowUnitPrompt "ship"
Case 10 'Upgrade
    'check for current prompt
    If PromptUp Then
        Exit Sub
    End If
    Set PromptForm = frmPromptScrap
    PromptForm.strCmd = "upgrade "
    PromptForm.Label2 = "Upgrade"
    PromptForm.Option1 = True
    'Show Prompt
    ShowUnitPrompt "ship"
End Select
End Sub

Private Sub mnuShipWarload_Click()
' added by thomas lullier
' load max mil/shell/gun to ship

'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptWarload
PromptForm.strCmd = "swarload"
'PromptForm.strSector = SectorString(MxPos, MyPos) 09/14/03 rjk: strSector was removed 2.2.13, not used in function
'PromptForm.SetSectorLabels - thomas lullier - what is it used for ?
'Show Prompt
ShowUnitPrompt "ship"

End Sub

Private Sub mnuToolsOption_Click(Index As Integer)
Dim strWarning As String

strWarning = "Warning: This utility is a tool designed to quickly set up a test country for" + vbLf _
        + "the Vampire Blitz.  It does not take into account GO_RENEW or BIG_CITIES." + vbLf _
        + "and it does not do the most efficent job.  It should be used with care, " + vbLf _
        + "and I recommend that it NEVER be used in a serious game. - Escher"

Select Case Index
    Case 1  'Script Files
        Load frmScript
        frmScript.Show vbModeless

    Case 2  'Build Projection
        'check for current prompt
        If PromptUp Then
            Exit Sub
        End If
        
        'Setup  dialog and execute
        Load frmToolHarbor
        frmToolHarbor.txtOrigin.Text = SectorString(SelX, SelY)
        Set PromptForm = frmToolHarbor
        
        'Put form in proper place
        frmToolHarbor.Left = frmDrawMap.Left + frmDrawMap.Width - frmToolHarbor.Width
        frmToolHarbor.top = frmDrawMap.top + (frmDrawMap.Height - frmToolHarbor.Height) \ 2
        PromptUp = True
        
        'Dialog box will take it from here
        frmToolHarbor.Show vbModeless

    Case 3
        'check for current prompt
        If PromptUp Then
            Exit Sub
        End If
        
        'Setup  dialog and execute
        Load frmToolFeeder
        frmToolFeeder.txtOrigin.Text = SectorString(SelX, SelY)
        Set PromptForm = frmToolFeeder
        
        'Put form in proper place
        PromptForm.Left = frmDrawMap.Left + frmDrawMap.Width - PromptForm.Width
        PromptForm.top = frmDrawMap.top + (frmDrawMap.Height - PromptForm.Height) \ 2
        PromptUp = True
        
        'Dialog box will take it from here
        frmToolFeeder.Show vbModeless

    Case 4
        frmToolNationLvl.Show vbModal
    Case 5
        frmToolCheMarker.Show vbModal
        DrawHexes
    Case 6 'production summary
        'check for current prompt
        If PromptUp Then
            Exit Sub
        End If
        
        Set PromptForm = frmPromptProdSummary
        PromptUp = True
        
        frmPromptProdSummary.txtOrigin.Text = SectorString(SelX, SelY)
                        
        PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
        PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height
       
        'Dialog box will take it from here
        PromptForm.Show vbModeless
    Case 7
        AirCombatReport
    Case 8
        Set PromptForm = frmToolRange
        frmToolRange.txtOrigin = SectorString(SelX, SelY)

        PromptUp = True
'        PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
 '       PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height
       
        'Dialog box will take it from here
        PromptForm.Show vbModeless
    Case 9
        'check for current prompt
        If PromptUp Then
            Exit Sub
        End If
        
        'Setup  dialog and execute
        Load frmToolFortify
        Set PromptForm = frmToolFeeder
        
        'Put form in proper place
        PromptForm.Left = frmDrawMap.Left + frmDrawMap.Width - PromptForm.Width
        PromptForm.top = frmDrawMap.top + (frmDrawMap.Height - PromptForm.Height) \ 2
        PromptUp = True
        
        'Dialog box will take it from here
        frmToolFortify.Show vbModeless
    Case 10
        frmSatellitePath.Show vbModeless
    Case 12
        Dim Reply As Integer
        'check for current prompt
        If PromptUp Then
            Exit Sub
        End If
    
        Reply = MsgBox("Run Blitz Setup ?" + vbLf + vbLf + strWarning, vbOKCancel + vbQuestion, "Tool Confirmation")
        If Reply = vbCancel Then
            Exit Sub
        End If
        
        BlitzSetup True, 0, 0, 0, 0
    Case 13
        Reply = MsgBox("Run Island Setup ?" + vbLf + vbLf + strWarning, vbOKCancel + vbQuestion, "Tool Confirmation")
        If Reply = vbCancel Then
            Exit Sub
        End If
        
        'Setup dialog and execute
        Set PromptForm = frmPromptIsland
        PromptUp = True
        
        'Put form in proper place
        PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
        PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height
        
        'Dialog box will take it from here
        PromptForm.Show vbModeless
    Case 14
        Reply = MsgBox("Run Island Develop Setup ?" + vbLf + vbLf + strWarning, vbOKCancel + vbQuestion, "Tool Confirmation")
        If Reply = vbCancel Then
            Exit Sub
        End If
        
        'check for current prompt
        If PromptUp Then
            Exit Sub
        End If
        
        Set PromptForm = frmPromptIsland
        PromptUp = True
        
        'Put form in proper place
        PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
        PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height
        PromptForm.Label2 = "Island Develop Tool"
        
        'Dialog box will take it from here
        PromptForm.Show vbModeless
    
    
    Case 16
        ImportBmap
        
    Case 17  'WorldBuilder
    
        'If bDeity Then
            'Setup dialog and execute
            Set PromptForm = frmToolWorldBuild
            PromptUp = True
            
            'Put form in proper place
            PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
            PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height
            
            'Dialog box will take it from here
            PromptForm.Show vbModeless
        'Else
            'MsgBox "World Builder is a Deity only Tool", vbOKOnly + vbExclamation, "Tool Confirmation"
        'End If
End Select
End Sub

Private Sub mnuUnitShow_Click()
frmToolShow.Show
End Sub

Private Sub mnuViewOptions_Click(Index As Integer)
Static SaveUnitView As Boolean

Select Case Index
Case 1
    If ColorScheme = COLORSCHEME_NORMAL Then
        Load frmColorScheme
        frmColorScheme.optScheme(2).Value = True
        frmColorScheme.Show vbModal
    Else
        ColorScheme = COLORSCHEME_NORMAL
        If frmOptions.dspDesignationSave = DD_BOTH_CURRENT Or _
           frmOptions.dspDesignationSave = DD_BOTH_NEW Then
            dspDesignation = DD_BOTH
            frmOptions.dspDesignationSave = DD_BOTH
        End If
        frmUnitView.unitView.Load
    End If
    DrawHexes
Case 2
    frmSelectColor.Show vbModal
    DrawHexes
Case 3
    frmUnitView.Show
Case 4
    If ViewShipWake Or LastShip >= 0 Then
        mnuViewOptions(4).Caption = "Show Ship Wakes"
        ViewShipWake = False
        LastShip = -1
    Else
        mnuViewOptions(4).Caption = "Hide Ship Wakes"
        ViewShipWake = True
        LastShip = -1
    End If
    DrawHexes
Case 5
    EventMarkers.Clear
    DrawHexes
Case 9
    Dim sx As Integer
    Dim sy As Integer
    Dim cstring As String
    On Error Resume Next
    
    If Not MapValid Then Exit Sub
    
    cstring = InputBox("Enter sector coordinates (x,y)", "Center Map", SectorString(SelX, SelY))
    
    '092403 rjk: Cancel should not return an error msgbox, this is also check blank entry
    If cstring = "" Then
        Exit Sub
    End If
    
    If ParseSectors(sx, sy, cstring) Then
        MoveMap sx, sy
    Else
        MsgBox "Sector Entry Invalid - Center Function Aborted", vbExclamation + vbOKOnly
    End If
Case 11 'Refresh screen
    DrawHexes
Case 13 'Options
    frmOptions.Show
End Select
End Sub

Private Sub mnuZoomOptions_Click(Index As Integer)
Dim strTemp As String
Select Case Index
    Case 1
        Magnification = Magnification * 0.25
    Case 2
        Magnification = Magnification * 0.5
    Case 3
        Magnification = Magnification * 0.8
    Case 4
        Magnification = Magnification * 1.25
    Case 5
        Magnification = Magnification * 1.75
    Case 6
        Magnification = Magnification * 2.5
    Case 8
        strTemp = InputBox("Current Magnification - " + Format$(Magnification, "#0.00"), "Set Magnification")
        If Len(strTemp) > 0 Then
            If Val(strTemp) > 0 Then
                Magnification = Val(strTemp)
            End If
        End If
    Case 9
        Magnification = GetSetting(APPNAME, "Layout", "Magnification", 1)
        
End Select

'Make sure values are reasonable
If Magnification > 6 Then
    Magnification = 6
ElseIf Magnification < 0.2 Then
    Magnification = 0.2
End If

'Erase the grid
SetHexSideLength picMap, (Screen.Width / Screen.TwipsPerPixelX) / (30 * 0.866 * 2) * Magnification
DrawHexes
End Sub

Private Sub mnuUnitsMarkNukes_Click()
EventMarkNukes
End Sub

'103003 rjk: Added Keys for moving about on the map
'110603 rjk: Modified the Key pattern
Private Sub picMap_KeyDown(KeyCode As Integer, Shift As Integer)
Dim bSectorFound As Boolean
Select Case KeyCode
Case vbKeyLeft
    bSectorFound = FillSectorBox(SelX - 2, SelY)
Case vbKeyRight
    bSectorFound = FillSectorBox(SelX + 2, SelY)
Case vbKeyUp
    If (SelX Mod 2) = 1 Then
        bSectorFound = FillSectorBox(SelX - 1, SelY - 1)
    Else
        bSectorFound = FillSectorBox(SelX + 1, SelY - 1)
    End If
Case vbKeyDown
    If (SelX Mod 2) = 1 Then
        bSectorFound = FillSectorBox(SelX - 1, SelY + 1)
    Else
        bSectorFound = FillSectorBox(SelX + 1, SelY + 1)
    End If
Case Else
    Exit Sub
End Select

DisplaySelect = udUNKNOWN
rsEnemyUnit.Seek "=", SelX, SelY
If rsEnemyUnit.NoMatch Then
    rsNuke.Seek "=", SelX, SelY
    If rsNuke.NoMatch Then
        rsPlane.Seek "=", SelX, SelY
        If rsPlane.NoMatch Then
            rsLand.Seek "=", SelX, SelY
            If rsLand.NoMatch Then
                rsShip.Seek "=", SelX, SelY
                If Not rsShip.NoMatch Then
                    DisplaySelect = udSHIP
                    FillGrid
                End If
            Else
                DisplaySelect = udLAND
                FillGrid
            End If
        Else
            DisplaySelect = udPLANE
            FillGrid
        End If
    Else
        DisplaySelect = udNUKE
        FillGrid
    End If
Else
    DisplaySelect = udENEMY
    FillGrid
End If
DisplaySectorPanel
enableOrDisableMenus bSectorFound '112103 rjk: Added bSectorFound to better determine what menus are enabled
End Sub

Private Sub picMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim PosX As Single
Dim PosY As Single
Dim bSectorFound As Boolean
' Dim bShift As Boolean   removed efj 8/2003
Dim bCtl As Boolean
Dim n As enumUnitDisplay
Dim str As String
'Dim bAlt As Boolean

' bShift = False   removed efj 8/2003
If (Shift And vbShiftMask) = 1 Then
    ' bShift = True   removed efj 8/2003
End If

bCtl = False
If (Shift And vbCtrlMask) = 2 Then
    bCtl = True
End If

'On Error GoTo picmap_MouseDown_End

GridSelection = False

If DrawingPath Then
    If Button = vbRightButton Then
        RemoveFromPath
    Else
        AddtoPath
    End If
    Exit Sub
End If

'fill the sector box
bSectorFound = FillSectorBox(MxPos, MyPos)

'if you are marking territories
If MarkingTerritory And bSectorFound Then
    On Error GoTo 0
    ToggleTerritory MxPos, MyPos
    PromptForm.ToggleHex MxPos, MyPos, lblDes(1).Caption
    Exit Sub
End If

'if you are marking territories
If WorldBuilding And bSectorFound Then
    On Error GoTo 0
    PromptForm.DesignateHex MxPos, MyPos, Button
    If MapValid Then
        UpdateScreenInfo MxPos, MyPos
        'Get Center position and draw new hex
        Coord MxPos, MyPos, PosX, PosY
        DrawSector Me.picMap, PosX, PosY, MxPos, MyPos
    End If
    Exit Sub
End If



'Set Indexs
rsLand.Index = "location"
rsShip.Index = "location"
rsPlane.Index = "location"
rsEnemyUnit.Index = "location"
rsNuke.Index = "location"
bShips = False
bLands = False
bPlanes = False
benemys = False
bNukes = False

'check for ship/land/plane/nuke
'rjk 05/13/03: Remove the checks for PromptUp so the sector panel is always updated, also
'              remove the checks for ViewUnits so the sector panel is always updated independent
'              of show/hide unit option
'              moved the tab control to DisplaySectorPanel
    
n = udUNKNOWN
rsNuke.Seek "=", MxPos, MyPos
If Not rsNuke.NoMatch Then
    bNukes = True
End If
rsEnemyUnit.Seek "=", MxPos, MyPos
If Not rsEnemyUnit.NoMatch Then
    n = udENEMY
    benemys = True
End If

rsNuke.Seek "=", MxPos, MyPos
If Not rsNuke.NoMatch Then
    n = udNUKE
    bNukes = True
End If
rsPlane.Seek "=", MxPos, MyPos
If Not rsPlane.NoMatch Then
    n = udPLANE
    bPlanes = True
End If

rsLand.Seek "=", MxPos, MyPos
If Not rsLand.NoMatch Then
    n = udLAND
    bLands = True
End If

rsShip.Seek "=", MxPos, MyPos
If Not rsShip.NoMatch Then
    n = udSHIP
    bShips = True
End If

If n <> udUNKNOWN Then
    'if you are on the list option, you don't need to switch unless it enemy only
    If DisplaySelect <> udList Or n = udENEMY Then
        'if whatever you are currently displaying is present in the sector, stay on that
        If Not ((DisplaySelect = udSHIP And bShips) Or _
                (DisplaySelect = udLAND And bLands) Or _
                (DisplaySelect = udPLANE And bPlanes) Or _
                (DisplaySelect = udNUKE And bNukes) Or _
                (DisplaySelect = udENEMY And benemys)) Then
            'if you are already not displaying the first choice, flip the selection
            If DisplaySelect <> n Then
                DisplaySelect = n
                SubType = TYPE_ALL
            End If
        End If
    End If
    DisplaySubset = ssSECT
    ' 092303 rjk: removed secx and secy and switched to SelX and SelY
    'secx = MxPos
    'secy = MyPos
    'FillSubSetFlag = True 092503 rjk: Eliminated
    FillGrid
Else
    DisplaySelect = udUNKNOWN
End If

DisplaySectorPanel 'Updated the Tab display
        
'This subroutine puts the clicked sector designation in the text box of
'the prompt form, and activates the form

If PromptUp Then
    If Not (PromptForm.ActiveControl Is Nothing) Then 'drk@unxsoft.com 7/2/02.  this seems to be required to prevent crashes in vb6?
        If Button = vbLeftButton And (TypeOf PromptForm.ActiveControl Is TextBox) Then 'drk@unxsoft.com: this seems to be required to prevent crashes in vb6?
            Dim tb As TextBox
            Set tb = PromptForm.ActiveControl
            If tb.Name = "txtOrigin" Or tb.Name = "txtDest" Or tb.Name = "txtPath" Or tb.Name = "txtSector" Then '111603 rjk: Added txtSector frmPromptSector
                tb.Text = SectorString(MxPos, MyPos)
            '101803 rjk: Added multiple sectors selection
            ElseIf tb.Name = "txtMultOrigin" Or tb.Name = "txtMultDest" Or tb.Name = "txtMultPath" Then
                If bCtl And InStr(tb.Text, ";") = 0 Then 'Ensure no waypoints present
                    If Len(tb.Text) > 0 Then
                        If InStr(tb.Text, "\") = 0 Then 'one item
                            If (InStr(tb.Text, SectorString(MxPos, MyPos)) = 1) And (Len(tb.Text) = Len(SectorString(MxPos, MyPos))) Then
                                tb.Text = ""
                            Else
                                tb.Text = tb.Text + "\" + SectorString(MxPos, MyPos)
                            End If
                        Else 'multiple items
                            'check left position
                            If (InStr(tb.Text, SectorString(MxPos, MyPos)) = 1) And (InStr(tb.Text, "\") = Len(SectorString(MxPos, MyPos)) + 1) Then
                                tb.Text = Mid$(tb.Text, InStr(tb.Text, "\") + 1)
                            Else
                                If InStr(tb.Text, "\" + SectorString(MxPos, MyPos) + "\") > 0 Then 'check middle positions
                                    tb.Text = Replace(tb.Text, "\" + SectorString(MxPos, MyPos) + "\", "\")
                                Else
                                    If (InStr(InStrRev(tb.Text, "\") + 1, tb.Text, SectorString(MxPos, MyPos)) > 0) And _
                                        ((InStrRev(tb.Text, "\") + Len(SectorString(MxPos, MyPos))) = Len(tb.Text)) _
                                        Then
                                            tb.Text = Left$(tb.Text, InStrRev(tb.Text, "\") - 1)
                                    Else
                                        tb.Text = tb.Text + "\" + SectorString(MxPos, MyPos)
                                    End If
                                End If
                            End If
                        End If
                    Else
                        tb.Text = SectorString(MxPos, MyPos)
                    End If
                Else
                    tb.Text = SectorString(MxPos, MyPos)
                End If
            ElseIf tb.Name = "txtTarget" Then
                If Len(strEnemyid) = 0 Then
                    tb.Text = SectorString(MxPos, MyPos)
                ElseIf InStr(strEnemyid, "/") = 0 Then
                    tb.Text = strEnemyid
                Else
                    DisplaySelect = udENEMY
                    DisplaySubset = ssSECT
                    ' 092303 rjk: removed secx and secy and switched to SelX and SelY
                    'secx = MxPos
                    'secy = MyPos
                    'FillSubSetFlag = True 092503 rjk: Eliminated
                    SubType = TYPE_ALL
                    FillGrid
                    DisplayUnitBox 'rjk 05/13/03: switched over DisplayUnitBox function instead of DisplaySectorPanel 3
                End If
            '092003 rjk: added check to make sure what SHIPs
'            ElseIf tb.Name = "txtUnit" And Len(strShipid) > 0 And LastUnitDisplay = udSHIP Then
'                tb.Text = strShipid
            '092003 rjk: added to get Plane
'            ElseIf tb.Name = "txtUnit" And Len(strPlaneid) > 0 And LastUnitDisplay = udPLANE Then
'                tb.Text = strPlaneid
            '092003 rjk: added to get Land Unit
'            ElseIf tb.Name = "txtUnit" And Len(strLandid) > 0 And LastUnitDisplay = udLAND Then
'                tb.Text = strLandid
            '101903 rjk: Removed frmPromptMove, it now uses txtMultPath
            '103103 rjk: Added back with Multiple Sector Selection capability
            ElseIf PromptForm.Name = "frmPromptMove" Then
                Set tb = PromptForm.txtMultPath
                If bCtl And InStr(tb.Text, ";") = 0 Then 'Ensure no waypoints present
                    If Len(tb.Text) > 0 Then
                        If InStr(tb.Text, "\") = 0 Then 'one item
                            If (InStr(tb.Text, SectorString(MxPos, MyPos)) = 1) And (Len(tb.Text) = Len(SectorString(MxPos, MyPos))) Then
                                tb.Text = ""
                            Else
                                tb.Text = tb.Text + "\" + SectorString(MxPos, MyPos)
                            End If
                        Else 'multiple items
                            'check left position
                            If (InStr(tb.Text, SectorString(MxPos, MyPos)) = 1) And (InStr(tb.Text, "\") = Len(SectorString(MxPos, MyPos)) + 1) Then
                                tb.Text = Mid$(tb.Text, InStr(tb.Text, "\") + 1)
                            Else
                                If InStr(tb.Text, "\" + SectorString(MxPos, MyPos) + "\") > 0 Then 'check middle positions
                                    tb.Text = Replace(tb.Text, "\" + SectorString(MxPos, MyPos) + "\", "\")
                                Else
                                    If (InStr(InStrRev(tb.Text, "\") + 1, tb.Text, SectorString(MxPos, MyPos)) > 0) And _
                                        ((InStrRev(tb.Text, "\") + Len(SectorString(MxPos, MyPos))) = Len(tb.Text)) _
                                        Then
                                            tb.Text = Left$(tb.Text, InStrRev(tb.Text, "\") - 1)
                                    Else
                                        tb.Text = tb.Text + "\" + SectorString(MxPos, MyPos)
                                    End If
                                End If
                            End If
                        End If
                    Else
                        tb.Text = SectorString(MxPos, MyPos)
                    End If
                Else
                    tb.Text = SectorString(MxPos, MyPos)
                End If
            ElseIf PromptForm.Name = "frmPromptLaunch" _
                Or PromptForm.Name = "frmPromptTrans" _
                Or PromptForm.Name = "frmPromptNav" Then
                PromptForm.txtPath.Text = SectorString(MxPos, MyPos)
            End If
            
            PromptForm.SetFocus
        End If
    End If '(PromptForm.ActiveControl <> Nothing)
'if control key is down, add to text line
ElseIf bCtl Then
'    txtEntry.Text = txtEntry.Text + " " + SectorString(MxPos, MyPos)
    str = " " + SectorString(MxPos, MyPos) + " "
    'check for a sector substitute.  If found, the put in sector and execute
    n = InStr(txtEntry.Text, "&s")
    If n = 0 Then n = InStr(txtEntry.Text, "&S")
    If n > 0 Then
        str = Left$(txtEntry.Text, n - 1) + str + Mid$(txtEntry.Text, n + 2)
        frmEmpCmd.SubmitFromCommandLine str
    ElseIf Not CommandCursorFocus Then
        txtEntry.SelText = str
        CommandCursorPos = CommandCursorPos + Len(str)
        SetCommandFocus
    Else
        txtEntry.SelText = str
    End If
End If
'092503 rjk: Do when a new sector is selected, need for top level menu
'112103 rjk: Added bSectorFound to better determine what menus are enabled
enableOrDisableMenus bSectorFound
If Button = vbRightButton Then
    If Not PromptUp Then        'Don't prompt when already prompted
        PopUpMenuSource = DMAP_PUMS_MAP
        'enableOrDisableMenus 'drk 6/5/02 092503 rjk: switched to be when a select is selected
        If Not (bShips Or bLands Or bPlanes Or bNukes Or bDeity) Then
            If bSectorFound Then
                PopupMenu mnuSectCmd
            Else
                rsBmap.Seek "=", MxPos, MyPos
                If rsBmap.NoMatch Then
                    PopupMenu mnuAttack
                ElseIf rsBmap.Fields("des") <> "." And rsBmap.Fields("des") <> "X" Then '112703 rjk: Added check for mined sea sector
                    PopupMenu mnuAttack
                Else
                    mnuSeaCenter.Visible = True
                    PopupMenu mnuSea '091403 rjk: Added Center Map for Sea
                    If Len(strEnemyid) > 0 Then
                        'set enemp ship wake to next in sector
                        Dim varShipNum As Variant
                        Dim NewShip As Integer
                        varShipNum = ParseUnitString(strEnemyid)
                        NewShip = CInt(varShipNum(1))
                        For n = 1 To UBound(varShipNum) - 1
                            If LastShip = Int(varShipNum(n)) Then
                                NewShip = Int(varShipNum(n))
                            End If
                        Next n
                        LastShip = NewShip
                        DrawHexes
                    End If
                End If
            End If
        Else
            '092503 rjk: Moved to enableOrDisableMenus
            'mnuLand.Visible = blands
            'mnuShip.Visible = bships
            'mnuPlane.Visible = bplanes
            'mnuNuke.Visible = bnukes
            'mnuDeity.Visible = bDeity
            '112003 rjk: Suppress Sector menu when on sea sector, 112103 moved to enableOrDisableMenus
            'mnuSectCmd.Visible = bSectorFound
            If bSectorFound Then
                'mnuSectCmd.Visible = True '112003 rjk: Suppress Sector menu when on sea sector
                PopupMenu mnuUnitCommand
            Else '112703 rjk: Changed to show Attack and Unit menus together on enemy sectors
                rsBmap.Seek "=", MxPos, MyPos
                If Not rsBmap.NoMatch Then
                    If rsBmap.Fields("des") = "." Or rsBmap.Fields("des") = "X" Then
                        If Not (bLands Or bPlanes Or bDeity) Then
                            PopupMenu mnuShip
                        ElseIf Not (bLands Or bShips Or bDeity) Then
                            PopupMenu mnuPlane
                        ElseIf Not (bPlanes Or bShips Or bDeity) Then '112003 rjk: Change one bships to bplanes
                            PopupMenu mnuLand
                        Else
                            PopupMenu mnuUnitCommand
                        End If
                    Else
                        PopupMenu mnuUnitCommand
                    End If
                Else
                    PopupMenu mnuUnitCommand
                End If
            End If
        End If
    Else
        'if there is an OK button, then run it.
        On Error Resume Next
        PromptForm.cmdOK_Click
        On Error GoTo picmap_MouseDown_End
    End If
End If
picmap_MouseDown_End:

End Sub

'112103 rjk: Added bSectorFound to better determine what menus should be active.
Private Sub enableOrDisableMenus(bSectorFound As Boolean)
'060502 drk: enable or disable menu items intelligently based on the sector designation
'fixme, this section could be significantly expanded upon.
'091303 rjk: added check for grind,Convert,Demobilize,Enlist,Shoot,Improve,Start,Stop,Spy,Explore
'103003 rjk: Switched to SelX and SelY from MxPos and MyPos.
'rsSectors.Seek "=", MxPos, MyPos
rsSectors.Seek "=", SelX, SelY
If Not rsSectors.NoMatch Then
    With rsSectors.Fields("des")
        mnuSectCmdFire.Enabled = (.Value = "f")
        mnuSectCmdBuild = ((.Value = "h") Or (.Value = "*") Or (.Value = "!") Or _
            (.Value = "n") Or (rsSectors.Fields("coast") = 1))
        mnuSectCmdRadar = (.Value = ")")
        mnuSectCmdCap = (.Value = "c") Or (.Value = "^")
    End With
    Me.mnuSectCmdAssault = rsSectors.Fields("coast") = 1
    mnuSectCmdGrind = rsSectors.Fields("bar") > 0
    '100203 rjk: Changed to Anti to not have to be occuppied but must have mil. and mob.
    'mnuSectCmdAnti = (rsSectors.Fields("*").Value = "*") And (rsSectors.Fields("mob").Value > 0)
    mnuSectCmdAnti = (rsSectors.Fields("mil").Value > 0) And (rsSectors.Fields("mob").Value > 0)
    mnuSectCmdConvert = (rsSectors.Fields("*").Value = "*") And (rsSectors.Fields("civ").Value > 0) And (rsSectors.Fields("mob").Value > 0)
    mnuSectCmdDemob = (rsSectors.Fields("mil").Value > 0) And (rsSectors.Fields("mob").Value > 0)
    If frmOptions.b2K4 Then
        mnuSectCmdEnlist = (MilReserves > 0) And (rsSectors.Fields("mob").Value > 0)
    Else
        mnuSectCmdEnlist = (rsSectors.Fields("civ").Value > 0) And (MilReserves > 0) And (rsSectors.Fields("mob").Value > 0)
    End If
    mnuSectCmdShoot = ((rsSectors.Fields("civ").Value > 0) Or (rsSectors.Fields("uw").Value > 0)) And (rsSectors.Fields("mob").Value > 0)
    mnuSectCmdImprove = (rsSectors.Fields("road").Value < 100) Or (rsSectors.Fields("rail").Value < 100) Or _
        (rsSectors.Fields("defense").Value < 100) And (rsSectors.Fields("mob").Value > 0)
    mnuSectCmdStart = rsSectors.Fields("off").Value = 1
    mnuSectCmdStop = rsSectors.Fields("off").Value <> 1
    mnuSectCmdSpy = rsSectors.Fields("mil").Value > 0
    mnuSectCmdExplore = ((rsSectors.Fields("civ").Value + rsSectors.Fields("mil").Value) > 1) And (rsSectors.Fields("mob").Value > 0)
    If frmOptions.bSPAtlantis Then
        mnuSectCmdImprove = True
        If rsSectors.Fields("gold") > 0 Then
            mnuSectCmdRadar = True
        End If
        If rsSectors.Fields("fert") > 0 Then
            mnuSectCmdFire = True
        End If
    End If
End If
If VersionCheck(4, 3, 6) = VER_LESS Then
    mnuLandStart(0).Visible = False
    mnuLandStart(1).Visible = False
    mnuShipMisc(7).Visible = False
    mnuShipMisc(8).Visible = False
    mnuPlaneOther(4).Visible = False
    mnuPlaneOther(5).Visible = False
    mnuNukeStart(0).Visible = False
    mnuNukeStart(1).Visible = False
End If

'092503 rjk: moved from RightClick to allow modification for top level menu
'112103 rjk: Set Visibility only if there is something to show
If bSectorFound Or bLands Or bShips Or bPlanes Or bNukes Or bDeity Then
    mnuUnitCommand.Visible = True
    mnuDeity.Visible = True 'Need as least one menu visible
    mnuSectCmd.Visible = bSectorFound
    mnuLand.Visible = bLands
    mnuShip.Visible = bShips
    mnuPlane.Visible = bPlanes
    mnuNuke.Visible = bNukes
    mnuDeity.Visible = bDeity
    If Not bSectorFound Then '112703 rjk: Added to display attack and unit menu together
        rsBmap.Seek "=", MxPos, MyPos
        If Not rsBmap.NoMatch Then
            If rsBmap.Fields("des") = "." Or rsBmap.Fields("des") = "X" Then
                mnuUnitAttack.Visible = False
            Else
                mnuUnitAttack.Visible = True
            End If
        Else
            mnuUnitAttack.Visible = True
        End If
    Else
        mnuUnitAttack.Visible = False
    End If
Else
    mnuUnitCommand.Visible = False
End If
End Sub

Public Function FillSectorBox(SxPos As Integer, SyPos As Integer) As Boolean
Dim strOwner As String
On Error GoTo Empire_Error

'set the selected sector indicators
SelX = SxPos
SelY = SyPos

'Unhighlight the old hex and highlight the current one
Dim PosX As Single
Dim PosY As Single
Dim oldx As Integer
Dim oldy As Integer
oldx = CurrentSectorX
oldy = CurrentSectorY
CurrentSectorX = SxPos
CurrentSectorY = SyPos

If MapValid Then
    On Error Resume Next 'might not still be visible
    'Get Center position and redraw old current hex
    Coord oldx, oldy, PosX, PosY
    DrawSector Me.picMap, PosX, PosY, oldx, oldy
    DrawEvent picMap, oldx, oldy
    'Get Center position and draw new hex
    Coord SxPos, SyPos, PosX, PosY
    DrawSector Me.picMap, PosX, PosY, SxPos, SyPos
    DrawEvent picMap, SxPos, SyPos
    On Error GoTo Empire_Error
End If

'Get Sector Information
FillSectorBox = False
lblSect.Caption = "Sector " + SectorString(SxPos, SyPos)


rsSectors.Seek "=", SxPos, SyPos
If Not rsSectors.NoMatch Then
    FillSectorBox = True
    SelSectType = SEL_SECT_OWN
    lblDes(1).Caption = vbNullString
    lblDes(1).Caption = colDes.Item(rsSectors.Fields("des").Value)
    
    'drk 7/22/03
    If (CapSect = SectorString(SxPos, SyPos)) Then lblDes(1).Caption = lblDes(1).Caption + " : National capital"
    lblDes(1).Caption = CStr(rsSectors.Fields("eff").Value) + "% " + lblDes(1).Caption
    If Not bDeity Then
        If Len(Trim$(rsSectors.Fields("sdes").Value)) > 0 Then
            lblNDes.Caption = "new = " + colDes.Item(Trim$(rsSectors.Fields("sdes").Value))
        Else
            lblNDes.Caption = vbNullString
        End If
    Else
        ' for deities - store newdes in top label, put owner in the second
        If Len(Trim$(rsSectors.Fields("sdes").Value)) > 0 Then
            lblDes(1).Caption = lblDes(1).Caption + " (newdes=" + Trim$(rsSectors.Fields("sdes").Value) + ")" '121203 rjk: Added closing bracket
        End If
        lblNDes.Caption = Nations.NationName(rsSectors.Fields("country").Value)
    End If
        
    lblDes(2).Caption = lblNDes.Caption 'rjk 5/12/03: added new designation info to distribution tab
    lblDes(0).Caption = lblDes(1).Caption
    '240104 rjk: Added distribution point title
    If SxPos = rsSectors.Fields("dist_x").Value And SyPos = rsSectors.Fields("dist_y").Value Then
        lblDes(4) = "No Dist. Point"
    Else
        lblDes(4).Caption = SectorString(rsSectors.Fields("dist_x").Value, rsSectors.Fields("dist_y").Value)
    End If
'    Label1(3).Caption = CStr(rsSectors.Fields("civ").Value)
'    Label1(4).Caption = CStr(rsSectors.Fields("mil").Value)
'    Label1(5).Caption = CStr(rsSectors.Fields("uw").Value)
    Label1(3).Caption = CStr(Items.FindByLetter("c").DatabaseValue(rsSectors))
    Label1(0).Caption = Items.FindByLetter("c").SectorPanelLabel
    Label1(4).Caption = CStr(Items.FindByLetter("m").DatabaseValue(rsSectors))
    Label1(1).Caption = Items.FindByLetter("m").SectorPanelLabel
    Label1(5).Caption = CStr(Items.FindByLetter("u").DatabaseValue(rsSectors))
    Label1(2).Caption = Items.FindByLetter("u").SectorPanelLabel
    Label1(9).Caption = CStr(rsSectors.Fields("mob").Value)
    Label1(11).Caption = CStr(rsSectors.Fields("avail").Value)
    Label1(10).Caption = CStr(rsSectors.Fields("work").Value)
    Label1(20).Caption = CStr(rsSectors.Fields("min").Value)
    Label1(19).Caption = CStr(rsSectors.Fields("gold").Value)
    Label1(18).Caption = CStr(rsSectors.Fields("fert").Value)
    Label1(32).Caption = CStr(rsSectors.Fields("ocontent").Value)
    Label1(31).Caption = CStr(rsSectors.Fields("uran").Value)
    Label1(14).Caption = CStr(rsSectors.Fields("road").Value)
    Label1(13).Caption = CStr(rsSectors.Fields("rail").Value)
    Label1(12).Caption = CStr(rsSectors.Fields("defense").Value)
    Label1(26).Caption = CStr(rsSectors.Fields("fallout").Value)
    Label1(25).Caption = CStr(rsSectors.Fields("terr").Value)
    If frmOptions.bSPAtlantis Then
        Label1(23).Caption = "Run:"
        Label1(22).Caption = "Rad:"
        Label1(21).Caption = "Fort:"
        Label1(35).Caption = "Nav:"
    End If
   
    'Commodities
'    Label1(60).Caption = CStr(rsSectors.Fields("food").Value)
'    Label1(61).Caption = CStr(rsSectors.Fields("shell").Value)
'    Label1(62).Caption = CStr(rsSectors.Fields("gun").Value)
'    Label1(63).Caption = CStr(rsSectors.Fields("pet").Value)
'    Label1(64).Caption = CStr(rsSectors.Fields("iron").Value)
'    Label1(65).Caption = CStr(rsSectors.Fields("dust").Value)
'    Label1(66).Caption = CStr(rsSectors.Fields("bar").Value)
'    Label1(67).Caption = CStr(rsSectors.Fields("oil").Value)
'    Label1(68).Caption = CStr(rsSectors.Fields("lcm").Value)
'    Label1(69).Caption = CStr(rsSectors.Fields("hcm").Value)
'    Label1(70).Caption = CStr(rsSectors.Fields("rad").Value)
    Label1(60).Caption = CStr(Items.FindByLetter("f").DatabaseValue(rsSectors))
    Label1(55).Caption = Items.FindByLetter("f").SectorPanelLabel
    Label1(61).Caption = CStr(Items.FindByLetter("s").DatabaseValue(rsSectors))
    DisplayCommodity Label1(54), "s", rsSectors
    Label1(62).Caption = CStr(Items.FindByLetter("g").DatabaseValue(rsSectors))
    DisplayCommodity Label1(53), "g", rsSectors
    Label1(63).Caption = CStr(Items.FindByLetter("p").DatabaseValue(rsSectors))
    DisplayCommodity Label1(24), "p", rsSectors
    Label1(64).Caption = CStr(Items.FindByLetter("i").DatabaseValue(rsSectors))
    DisplayCommodity Label1(27), "i", rsSectors
    Label1(65).Caption = CStr(Items.FindByLetter("d").DatabaseValue(rsSectors))
    DisplayCommodity Label1(30), "d", rsSectors
    Label1(66).Caption = CStr(Items.FindByLetter("b").DatabaseValue(rsSectors))
    DisplayCommodity Label1(49), "b", rsSectors
    Label1(67).Caption = CStr(Items.FindByLetter("o").DatabaseValue(rsSectors))
    DisplayCommodity Label1(48), "o", rsSectors
    Label1(68).Caption = CStr(Items.FindByLetter("l").DatabaseValue(rsSectors))
    DisplayCommodity Label1(47), "l", rsSectors
    Label1(69).Caption = CStr(Items.FindByLetter("h").DatabaseValue(rsSectors))
    DisplayCommodity Label1(38), "h", rsSectors
    Label1(70).Caption = CStr(Items.FindByLetter("r").DatabaseValue(rsSectors))
    DisplayCommodity Label1(39), "r", rsSectors
   
       'Distribution
    Label1(100).Caption = CStr(rsSectors.Fields("c_dist").Value)
    Label1(99).Caption = Items.FindByLetter("c").DistributionPanelLabel
    Label1(101).Caption = CStr(rsSectors.Fields("m_dist").Value)
    Label1(80).Caption = Items.FindByLetter("m").DistributionPanelLabel
    Label1(102).Caption = CStr(rsSectors.Fields("u_dist").Value)
    Label1(94).Caption = Items.FindByLetter("u").DistributionPanelLabel
    Label1(103).Caption = CStr(rsSectors.Fields("f_dist").Value)
    Label1(76).Caption = Items.FindByLetter("f").DistributionPanelLabel
    Label1(104).Caption = CStr(rsSectors.Fields("s_dist").Value)
    Label1(77).Caption = Items.FindByLetter("s").DistributionPanelLabel
    Label1(105).Caption = CStr(rsSectors.Fields("g_dist").Value)
    Label1(78).Caption = Items.FindByLetter("g").DistributionPanelLabel
    Label1(106).Caption = CStr(rsSectors.Fields("p_dist").Value)
    Label1(97).Caption = Items.FindByLetter("p").DistributionPanelLabel
    Label1(107).Caption = CStr(rsSectors.Fields("i_dist").Value)
    Label1(96).Caption = Items.FindByLetter("i").DistributionPanelLabel
    Label1(108).Caption = CStr(rsSectors.Fields("d_dist").Value)
    Label1(95).Caption = Items.FindByLetter("d").DistributionPanelLabel
    Label1(109).Caption = CStr(rsSectors.Fields("b_dist").Value)
    Label1(82).Caption = Items.FindByLetter("b").DistributionPanelLabel
    Label1(110).Caption = CStr(rsSectors.Fields("o_dist").Value)
    Label1(83).Caption = Items.FindByLetter("o").DistributionPanelLabel
    Label1(111).Caption = CStr(rsSectors.Fields("l_dist").Value)
    Label1(84).Caption = Items.FindByLetter("l").DistributionPanelLabel
    Label1(112).Caption = CStr(rsSectors.Fields("h_dist").Value)
    Label1(91).Caption = Items.FindByLetter("h").DistributionPanelLabel
    Label1(113).Caption = CStr(rsSectors.Fields("r_dist").Value)
    Label1(90).Caption = Items.FindByLetter("r").DistributionPanelLabel
        
    'Delivery
    Label1(120).Caption = DeliveryLabel(rsSectors.Fields("c_cut").Value, rsSectors.Fields("c_del").Value)
    Label1(33).Caption = Items.FindByLetter("c").DistributionPanelLabel
    Label1(121).Caption = DeliveryLabel(rsSectors.Fields("m_cut").Value, rsSectors.Fields("m_del").Value)
    Label1(36).Caption = Items.FindByLetter("m").DistributionPanelLabel
    Label1(122).Caption = DeliveryLabel(rsSectors.Fields("u_cut").Value, rsSectors.Fields("u_del").Value)
    Label1(42).Caption = Items.FindByLetter("u").DistributionPanelLabel
    Label1(123).Caption = DeliveryLabel(rsSectors.Fields("f_cut").Value, rsSectors.Fields("f_del").Value)
    Label1(88).Caption = Items.FindByLetter("f").DistributionPanelLabel
    Label1(124).Caption = DeliveryLabel(rsSectors.Fields("s_cut").Value, rsSectors.Fields("s_del").Value)
    Label1(87).Caption = Items.FindByLetter("s").DistributionPanelLabel
    Label1(125).Caption = DeliveryLabel(rsSectors.Fields("g_cut").Value, rsSectors.Fields("g_del").Value)
    Label1(86).Caption = Items.FindByLetter("g").DistributionPanelLabel
    Label1(126).Caption = DeliveryLabel(rsSectors.Fields("p_cut").Value, rsSectors.Fields("p_del").Value)
    Label1(43).Caption = Items.FindByLetter("p").DistributionPanelLabel
    Label1(127).Caption = DeliveryLabel(rsSectors.Fields("i_cut").Value, rsSectors.Fields("i_del").Value)
    Label1(44).Caption = Items.FindByLetter("i").DistributionPanelLabel
    Label1(128).Caption = DeliveryLabel(rsSectors.Fields("d_cut").Value, rsSectors.Fields("d_del").Value)
    Label1(45).Caption = Items.FindByLetter("d").DistributionPanelLabel
    Label1(129).Caption = DeliveryLabel(rsSectors.Fields("b_cut").Value, rsSectors.Fields("b_del").Value)
    Label1(75).Caption = Items.FindByLetter("b").DistributionPanelLabel
    Label1(130).Caption = DeliveryLabel(rsSectors.Fields("o_cut").Value, rsSectors.Fields("o_del").Value)
    Label1(74).Caption = Items.FindByLetter("o").DistributionPanelLabel
    Label1(131).Caption = DeliveryLabel(rsSectors.Fields("l_cut").Value, rsSectors.Fields("l_del").Value)
    Label1(73).Caption = Items.FindByLetter("l").DistributionPanelLabel
    Label1(132).Caption = DeliveryLabel(rsSectors.Fields("h_cut").Value, rsSectors.Fields("h_del").Value)
    Label1(52).Caption = Items.FindByLetter("h").DistributionPanelLabel
    Label1(133).Caption = DeliveryLabel(rsSectors.Fields("r_cut").Value, rsSectors.Fields("r_del").Value)
    Label1(56).Caption = Items.FindByLetter("r").DistributionPanelLabel
    '101703 rjk: Strength labels 147 to 158 - odd numbers the label even numbers the values
    '112003 rjk: Added a new option for controlling strength updates, blue if suppressed
    If frmOptions.bSuppressStrength = True Then
        Label1(148).ForeColor = vbBlue
        Label1(150).ForeColor = vbBlue
        Label1(152).ForeColor = vbBlue
        Label1(154).ForeColor = vbBlue
        Label1(156).ForeColor = vbBlue
        Label1(158).ForeColor = vbBlue
    Else
        Label1(148).ForeColor = vbBlack
        Label1(150).ForeColor = vbBlack
        Label1(152).ForeColor = vbBlack
        Label1(154).ForeColor = vbBlack
        Label1(156).ForeColor = vbBlack
        Label1(158).ForeColor = vbBlack
    End If
    '112503 rjk: If Null display unknown
    If Not IsNull(rsSectors.Fields("units").Value) Then
        Label1(148).Caption = CStr(rsSectors.Fields("units").Value)
    Else
        Label1(148).Caption = "???"
    End If
    If Not IsNull(rsSectors.Fields("runits").Value) Then
        Label1(150).Caption = CStr(rsSectors.Fields("runits").Value)
    Else
        Label1(150).Caption = "???"
    End If
    If Not IsNull(rsSectors.Fields("lmines").Value) Then
        If rsSectors.Fields("lmines").Value = -1 Then '120303 rjk: display enemy mines in red
            Label1(152).ForeColor = vbRed
            Label1(152).Caption = "?"
        Else
            Label1(152).Caption = CStr(rsSectors.Fields("lmines").Value)
        End If
    Else
        Label1(152).Caption = "???"
    End If
    If Not IsNull(rsSectors.Fields("sec_mult").Value) Then
        Label1(154).Caption = CStr(rsSectors.Fields("sec_mult").Value)
    Else
        Label1(154).Caption = "???"
    End If
    If Not IsNull(rsSectors.Fields("sec_def").Value) Then
        Label1(156).Caption = CStr(rsSectors.Fields("sec_def").Value)
    Else
        Label1(156).Caption = "???"
    End If
    If Not IsNull(rsSectors.Fields("tot_def").Value) Then
        Label1(158).Caption = CStr(rsSectors.Fields("tot_def").Value)
    Else
        Label1(158).Caption = "???"
    End If
    
    '102303 rjk: Added Terr1/Terr2/Terr3 fields labels 159 to 164 - odd numbers the label even numbers the values
    Label1(160).Caption = CStr(rsSectors.Fields("terr1").Value)
    Label1(162).Caption = CStr(rsSectors.Fields("terr2").Value)
    Label1(164).Caption = CStr(rsSectors.Fields("terr3").Value)
    
    'calculated fields
    Dim n As Integer
    n = FoodRequired(rsSectors)
    If n >= 0 Then
        Label1(58).Caption = CStr(n)
    Else
        Label1(58).Caption = "??"
    End If
    
    If rsSectors.Fields("food").Value - n < 0 Then
        Label1(58).ForeColor = vbRed
    Else
        Label1(58).ForeColor = vbBlack
    End If
    
    
    Label1(41).ForeColor = vbBlack
    Dim v As productionDataType
    v = Production(rsSectors)
    If (v.prodValidFlag) Then
        Label1(85).Caption = CStr(v.ProdAmount)
        Label1(142).Caption = CStr(v.MaxProdAmount)
        If v.prodMaxMaterialLimit Then
            Label1(142).ForeColor = vbBlue
        Else
            Label1(142).ForeColor = vbBlack
        End If
        n = CInt(v.ExcessCiv)
        If n < 0 Then
            Label1(41).ForeColor = vbRed
            n = 0 - n
        End If
        Label1(41).Caption = CStr(n)
        Label1(46).Caption = CStr(v.NewEff)
    Else
        Label1(85).Caption = vbNullString
        Label1(142).Caption = vbNullString
        Label1(41).Caption = vbNullString
        Label1(46).Caption = vbNullString
    End If
    
    For n = 60 To 70
        If Label1(n).Caption = "0" Then
            Label1(n).Caption = "--"
        End If
    Next n
    For n = 100 To 113
        If Label1(n).Caption = "0" Then
            Label1(n).Caption = "--"
        End If
    Next n
    
    For n = 120 To 133
        If Label1(n).Caption = "0." Then
            Label1(n).Caption = "--"
        End If
    Next n
    
    lblWarn = EventMarkers.Find(SxPos, SyPos)
    If Len(lblWarn) = 0 Then
        If rsSectors.Fields("off").Value = 1 Then
            lblWarn = "Stopped "
        End If
        If rsSectors.Fields("*").Value = "*" Then
            lblWarn = lblWarn + "Not Yours"
        End If
    End If
        
    lblPlayers.Caption = ""
    If UBound(currentPlayersNumber) > 0 Then
        For n = 0 To UBound(currentPlayersNumber) - 1
            lblPlayers.Caption = lblPlayers.Caption + CStr(currentPlayersNumber(n)) + _
                                 ":" + currentPlayersName(n) + " "
        Next n
        lblPlayers.Caption = Left$(lblPlayers.Caption, Len(lblPlayers.Caption) - 1)
    End If

    'removed rjk 05/13/03: Removed frame code as it is done in DisplaySectorPanel
    'If mintCurFrame = 0 Then
    '    mintCurFrame = TabStrip1.SelectedItem.Index
    'End If
    
    'Frame1(0).Visible = False
    'Frame1(3).Visible = False
    'Frame1(mintCurFrame).Visible = True
Else
    'Check for Enemy Sector
    rsEnemySect.Seek "=", SxPos, SyPos
    If Not rsEnemySect.NoMatch Then
        'Frame1(mintCurFrame).Visible = False rjk 05/13/03: Removed frame code as it is done in DisplaySectorPanel
        SelSectType = SEL_SECT_ENEMY
        lblDes(5).Caption = vbNullString
        On Error Resume Next    'in case des is not in the collection
        '110304 rjk: Switched to use IsNull for EnemySector designation, display Unknown designation for null entry
        '            if you can not find the information in the rsBmap.
        '            Show ??? for unknown efficiency.
        If IsNull(rsEnemySect.Fields("des").Value) Then
            rsBmap.Seek "=", SxPos, SyPos
            If rsBmap.NoMatch Then
                lblDes(5).Caption = "Unknown designation"
            Else
                If rsBmap.Fields("des") <> "?" Then
                    lblDes(5).Caption = colDes.Item(rsBmap.Fields("des"))
                    If rsEnemySect.Fields("eff").Value >= 0 Then
                        lblDes(5).Caption = CStr(rsEnemySect.Fields("eff").Value) + "% " + lblDes(5).Caption
                    Else
                        lblDes(5).Caption = "???% " + lblDes(5).Caption
                    End If
                Else
                    lblDes(5).Caption = "Unknown designation"
                End If
            End If
        ElseIf rsEnemySect.Fields("des").Value = "?" Then
            rsBmap.Seek "=", SxPos, SyPos
            If rsBmap.NoMatch Then
                lblDes(5).Caption = "Unknown designation"
            Else
                If rsBmap.Fields("des") <> "?" Then
                    lblDes(5).Caption = colDes.Item(rsBmap.Fields("des"))
                    If rsEnemySect.Fields("eff").Value >= 0 Then
                        lblDes(5).Caption = CStr(rsEnemySect.Fields("eff").Value) + "% " + lblDes(5).Caption
                    Else
                        lblDes(5).Caption = "???% " + lblDes(5).Caption
                    End If
                Else
                    lblDes(5).Caption = "Unknown designation"
                End If
            End If
        Else
            lblDes(5).Caption = colDes.Item(rsEnemySect.Fields("des").Value)
            If rsEnemySect.Fields("eff").Value >= 0 Then
                lblDes(5).Caption = CStr(rsEnemySect.Fields("eff").Value) + "% " + lblDes(5).Caption
            Else
                lblDes(5).Caption = "???% " + lblDes(5).Caption
            End If
        End If
        On Error GoTo Empire_Error
        
        strOwner = Nations.NationName(rsEnemySect.Fields("owner").Value)
        If strOwner = vbNullString Then
            strOwner = CStr(rsEnemySect.Fields("owner").Value)
        End If
        Label7.Caption = "owner = " + strOwner
        n = rsEnemySect.Fields("oldowner").Value
        If n > 0 And n <> rsEnemySect.Fields("owner").Value Then
            Label7.Caption = Label7.Caption + "; oldown = " + CStr(rsEnemySect.Fields("oldowner").Value)
        End If
        
        'Set Enemy Sectors
        Dim strlab As String
        strlab = "098 114 115 138 137 117 118 116 081 079 136"
        For n = 0 To 10
            If rsEnemySect.Fields(n + 6).Value < 0 Then
                Label1(CInt(Mid$(strlab, n * 4 + 1, 3))).Caption = "???"
            Else
                Label1(CInt(Mid$(strlab, n * 4 + 1, 3))).Caption = rsEnemySect.Fields(n + 6).Value
            End If
        Next n
        Label8.Caption = EventMarkers.Find(SxPos, SyPos)
        If Len(Label8.Caption) = 0 Then
            Label8.Caption = Format$(ConvertToLocalTimeFromWinACEUTC(rsEnemySect.Fields("timestamp").Value), "ddd mmm dd hh:mm:ss yyyy")
        End If
        
        'removed rjk 05/13/03: Removed frame code as it is done in DisplaySectorPanel
        'mintCurFrame = 0
        'Frame1(mintCurFrame).Visible = True
    Else 'Not rsEnemySect.NoMatch => not enemy either.  maybe it is a sea sector?  drk 5/13/03
        rsSea.Seek "=", SxPos, SyPos
        If Not rsSea.NoMatch Then
            SelSectType = SEL_SECT_SEA
            '101503 rjk: Switch Oil and Fert fields so they display correctly on the screen
            If rsSea.Fields("ocontent").Value < 0 Then
                Label1(146).Caption = "???"
            Else
                Label1(146).Caption = CStr(rsSea.Fields("ocontent").Value)
            End If
            If rsSea.Fields("fert").Value < 0 Then
                Label1(143).Caption = "???"
            Else
                Label1(143).Caption = CStr(rsSea.Fields("fert").Value)
            End If
            Label4.Caption = EventMarkers.Find(SxPos, SyPos)
        Else
            SelSectType = SEL_SECT_UNKNOWN
        End If
        'removed rjk 05/13/03: Removed frame code as it is done in DisplaySectorPanel
        'If mintCurFrame <> 3 Then
        '    Frame1(mintCurFrame).Visible = False
        'End If
    End If
        
End If

Empire_Error:

End Function

Private Function DeliveryLabel(vCutoff As Variant, vDirection As Variant) As String
If vDirection >= 0 And vDirection <= 7 And frmOptions.bStarWars Then
DeliveryLabel = CStr(vCutoff) + ":" + CStr(vDirection)
Else
DeliveryLabel = CStr(vCutoff) + CStr(vDirection)
End If
End Function

Private Sub DisplayCommodity(lblComm As Label, strComm As String, rsSect As Recordset)
Dim CommodityRatio As String

If frmOptions.bCommodityRatio Then
    If rsSect.Fields("sdes") <> " " Then
        rsSectorType.Seek "=", rsSect.Fields("sdes")
    Else
        rsSectorType.Seek "=", rsSect.Fields("des")
    End If
    If Not rsSectorType.NoMatch Then
        If rsSectorType.Fields("use1s") = strComm Then
            CommodityRatio = " " + CStr(rsSectorType.Fields("use1n"))
        End If
        If rsSectorType.Fields("use2s") = strComm Then
            CommodityRatio = " " + CStr(rsSectorType.Fields("use2n"))
        End If
        If rsSectorType.Fields("use3s") = strComm Then
            CommodityRatio = " " + CStr(rsSectorType.Fields("use3n"))
        End If
    End If
End If
If Len(CommodityRatio) > 0 Then
    lblComm.ForeColor = vbMediumGreen
    lblComm.Caption = Items.FindByLetter(strComm).SectorPanelLabel + CommodityRatio
Else
    lblComm.ForeColor = vbBlack
    lblComm.Caption = Items.FindByLetter(strComm).SectorPanelLabel
End If
End Sub

Private Sub PicMap_DblClick()

Dim Sector As String
' Dim n As Integer   removed efj 8/2003
'Dont bring up box when drawing path
If DrawingPath Then
    Exit Sub
End If

'Get Center position of click
Dim PosX As Single
Dim PosY As Single

Coord MxPos, MyPos, PosX, PosY
If bShips Or bLands Or bPlanes Or benemys Then
    If SubType = TYPE_ALL Then
        DisplayUnitBox 'rjk 05/13/03: Switch to DisplayUnitBox function from DisplaySectorPanel 3
    Else
'        SetUnitType TYPE_ALL 091903 rjk: removed SetUnitType function, do same actions as TYPE_ALL from combo box
        SubType = TYPE_ALL
        FillGrid
    End If
    Exit Sub
End If

Sector = SectorString(MxPos, MyPos)
rsSectors.Seek "=", MxPos, MyPos
If rsSectors.NoMatch Then
    Exit Sub
End If

rsSectorType.Seek "=", rsSectors.Fields("des")
If Not rsSectorType.NoMatch Then
    If rsSectorType.Fields("product") <> " " Then
        If (rsSectorType.Fields("des") <> 0) And (rsSectors.Fields("sdes") <> "") Then
            If (rsSectors.Fields("des") = "g" And rsSectors.Fields("gold").Value = 0) Or _
               (rsSectors.Fields("des") = "o" And rsSectors.Fields("ocontent").Value = 0) Or _
               (rsSectors.Fields("des") = "u" And rsSectors.Fields("uran").Value = 0) Then
                mnuSectCmdDes_Click
                Exit Sub
            End If
        End If
        frmEmpCmd.SubmitEmpireCommand "production " + " " + Sector, True
        Exit Sub
    End If
End If
    

Select Case rsSectors.Fields("des").Value

    Case ")"
        frmEmpCmd.SubmitEmpireCommand "radar " + " " + Sector, True
        frmEmpCmd.SubmitEmpireCommand "db1", False
        frmEmpCmd.SubmitEmpireCommand "bmap *", False
        frmEmpCmd.SubmitEmpireCommand "db2", False
    Case "h", "!", "*", "n"
        mnuUnitBuild_Click
    Case "-", "~"
        mnuSectCmdDes_Click
    Case "f"
        mnuSectCmdFire_Click
    Case "c"
        mnuReportNation_Click
    Case "b"
        mnuReportBudget_Click
    Case "w"
        mnuSectCmdDist_Click
'    Case "j", "k", "l", "t", "d", "i", "m", "e", "%", "p", "a", "r"
'        frmEmpCmd.SubmitEmpireCommand "production " + " " + Sector, True
'    Case "g", "o", "u"
'        If (rsSectors.Fields("des").Value = "g" And _
'            Len(Trim$(rsSectors.Fields("sdes").Value)) = 0 And _
'            rsSectors.Fields("gold").Value = 0) Or _
'            (rsSectors.Fields("des").Value = "o" And _
'            Len(Trim$(rsSectors.Fields("sdes").Value)) = 0 And _
'            rsSectors.Fields("ocontent").Value = 0) Or _
'            (rsSectors.Fields("des").Value = "u" And _
'            Len(Trim$(rsSectors.Fields("sdes").Value)) = 0 And _
'            rsSectors.Fields("rad").Value = 0) Then
'            mnuSectCmdDes_Click
'        Else
'            frmEmpCmd.SubmitEmpireCommand "production " + " " + Sector, True
'        End If
End Select


End Sub

'112903 rjk: Removed ViewUnits check, always display units on bottom status line
Private Sub picMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim oldx As Integer
Dim oldy As Integer
Dim PosX As Single
Dim PosY As Single
Dim AdjX As Integer
Dim AdjY As Integer
Dim HalfLength As Single
Dim HalfWidth As Single
Dim hexWidth As Single
Dim hexHeight As Single

Me.MousePointer = vbDefault
If Not MapValid Then
    Exit Sub
End If

'Save the old position x and y to see if they have changed
oldx = MxPos
oldy = MyPos

HalfLength = 0.5 * GetHexSideLength
HalfWidth = 0.8667 * GetHexSideLength

hexHeight = 2 * GetHexSideLength
hexWidth = 0.8667 * 2 * GetHexSideLength

MyPos = Int(Y / (hexHeight * 1.5))
PosY = Y - ((hexHeight * 1.5) * MyPos)

MyPos = MyPos * 2
If PosY > hexHeight Then
    MyPos = MyPos + 1
    PosY = PosY - (0.75 * hexHeight)
End If

If (MyPos Mod 2) = 0 Then
    MxPos = Int(X / hexWidth)
    PosX = X - (hexWidth * MxPos)
    MxPos = MxPos * 2
Else
    MxPos = Int((X - HalfWidth) / hexWidth)
    PosX = X - (hexWidth * MxPos) - HalfWidth
    MxPos = (MxPos * 2) + 1
End If

If PosY < HalfLength Then
    If PosX <= HalfWidth Then
        If PosY > (HalfLength - (HalfLength * PosX) / HalfWidth) Then
            AdjX = 0
            AdjY = 0
        Else
            AdjX = -1
            AdjY = -1
        End If
    Else
        If PosY > (HalfLength * (PosX - HalfWidth) / HalfWidth) Then
            AdjX = 0
            AdjY = 0
        Else
            AdjX = 1
            AdjY = -1
        End If
    End If

ElseIf PosY > 3 * HalfLength Then
    If PosX <= HalfWidth Then
        If PosY > (3 * HalfLength + (HalfLength * PosX) / HalfWidth) Then
            AdjX = -1
            AdjY = 1
        Else
            AdjX = 0
            AdjY = 0
        End If
    Else
        If PosY > (4 * HalfLength - (HalfLength * (PosX - HalfWidth) / HalfWidth)) Then
            AdjX = 1
            AdjY = 1
        Else
            AdjX = 0
            AdjY = 0
        End If
    End If

Else
    AdjX = 0
    AdjY = 0
End If

'Calc new X and Y position of mouse
MxPos = MxPos + AdjX + OriginX
MyPos = MyPos + AdjY + OriginY

If MyPos > (MapSizeY / 2) - 1 Then
    MyPos = MyPos - MapSizeY
End If

If MxPos > (MapSizeX / 2) - 1 Then
    MxPos = MxPos - MapSizeX
End If

'if we haven't changed hexes, then we're done
If MxPos = oldx And MyPos = oldy Then
    Exit Sub
End If

'Output Label
'sb1.Panels.Item("Hex").Text = "Sect: " + SectorString(MxPos, MyPos) + " "
sb1.Panels.Item("Hex").Text = SectorString(MxPos, MyPos) + " "

'Top Third of Block
'Get ship and plane info
Dim ship As Integer
Dim land As Integer
Dim plane As Integer
Dim nuke As Integer
Dim enemy As Integer
Dim lname As String
Dim strLand As String
Dim sname As String
Dim strShip As String
Dim strEnemy As String
' Dim strSect As String    removed efj 8/2003
Dim strPlane As String
Dim strNuke As String '100603 rjk: Added nuke information.
Dim strOwner As String
Dim strLastOwner As String
Dim lastid As Integer

strEnemyid = vbNullString
'Get Sector Information
'rsSectors.Seek "=", MxPos, MyPos
'If Not rsSectors.NoMatch Then
'    sb1.Panels.Item("Hex").Text = sb1.Panels.Item("Hex").Text _
'        + CStr(rsSectors.Fields("eff").Value) + "% "
'    strSect = CStr(rsSectors.Fields("civ").Value) _
'        + "c/" + CStr(rsSectors.Fields("mil").Value) _
'        + "m/" + CStr(rsSectors.Fields("uw").Value) + "u"
'End If

ship = 0
plane = 0
land = 0
nuke = 0
enemy = 0
strShipid = vbNullString
strPlaneid = vbNullString
strLandid = vbNullString
strNuke = vbNullString '100603 rjk: Added nuke information.
strNukeid = vbNullString '100603 rjk: Added nuke information.
strEnemyid = vbNullString

'Set Indexs
rsLand.Index = "location"
rsShip.Index = "location"
rsPlane.Index = "location"
rsNuke.Index = "location"
rsEnemyUnit.Index = "location"

'Load number of land units
rsLand.Seek "=", MxPos, MyPos
If Not rsLand.NoMatch Then

land_loop:
    If (Not rsLand.EOF) Then
        If rsLand.Fields("x").Value = MxPos And rsLand.Fields("y").Value = MyPos Then
            land = land + 1
            lname = rsLand.Fields("type").Value
            'rsBuild.Seek "=", "l", lname
            'If Not rsBuild.NoMatch Then
            '    If Len(rsBuild.Fields("desc")) > 0 Then
            '        lname = rsBuild.Fields("desc")
            '    End If
            'End If
            strLand = strLand + " " + lname + " #" + CStr(rsLand.Fields("id").Value) + _
                ","
            If land = 1 Then
                strLandid = CStr(rsLand.Fields("id").Value)
            Else
                strLandid = strLandid + "/" + CStr(rsLand.Fields("id").Value)
            End If
            lastid = rsLand.Fields("id").Value
            rsLand.MoveNext
            GoTo land_loop
        End If
    End If
    If land > 2 Then
        strLand = " " + CStr(land) + " units:" + strLand
    End If
    strLand = Left$(strLand, Len(strLand) - 1)
    If land = 1 Then
        strLand = strLand + " " + BuildMissionString("l", lastid)
    End If
End If

'Load number of ships
rsShip.Seek "=", MxPos, MyPos
If Not rsShip.NoMatch Then
ship_loop:
    If (Not rsShip.EOF) Then
        If rsShip.Fields("x").Value = MxPos And rsShip.Fields("y").Value = MyPos Then
            ship = ship + 1
            sname = rsShip.Fields("type").Value
            'rsBuild.Seek "=", "s", sname
            'If Not rsBuild.NoMatch Then
            '    If Len(rsBuild.Fields("desc")) > 0 Then
            '        sname = rsBuild.Fields("desc")
            '    End If
            'End If
            strShip = strShip + " " + sname + " #" + CStr(rsShip.Fields("id").Value) + ","
            If ship = 1 Then
                strShipid = CStr(rsShip.Fields("id").Value)
            Else
                strShipid = strShipid + "/" + CStr(rsShip.Fields("id").Value)
            End If
            lastid = rsShip.Fields("id").Value
            rsShip.MoveNext
            GoTo ship_loop
        End If
    End If
    If ship > 2 Then
        strShip = " " + CStr(ship) + " ships:" + strShip
    End If
    strShip = Left$(strShip, Len(strShip) - 1)
    If ship = 1 Then
        strShip = strShip + " " + BuildMissionString("s", lastid)
    End If
End If

Dim cp1() As Integer
Dim cp2() As String
Dim cpfound As Boolean
ReDim cp1(1)
ReDim cp2(1)

'Load nubmer of planes
rsPlane.Seek "=", MxPos, MyPos
If Not rsPlane.NoMatch Then
plane_loop:
    If (Not rsPlane.EOF) Then
        If rsPlane.Fields("x").Value = MxPos And rsPlane.Fields("y").Value = MyPos Then
            plane = plane + 1
            lname = rsPlane.Fields("type").Value
            cpfound = False
            For X = LBound(cp1) To UBound(cp1)
                If cp2(X) = lname Then
                    cp1(X) = cp1(X) + 1
                    cpfound = True
                    Exit For
                End If
            Next X
            If Not cpfound Then
                ReDim Preserve cp1(UBound(cp1) + 1)
                ReDim Preserve cp2(UBound(cp2) + 1)
                cp2(X) = lname
                cp1(X) = 1
            End If
            
            'rsBuild.Seek "=", "l", lname
            'If Not rsBuild.NoMatch Then
            '   If Len(rsBuild.Fields("desc")) > 0 Then
            '        lname = rsBuild.Fields("desc")
            '    End If
            'End If
            If plane = 1 Then
                strPlaneid = CStr(rsPlane.Fields("id").Value)
            Else
                strPlaneid = strPlaneid + "/" + CStr(rsPlane.Fields("id").Value)
            End If
            lastid = rsPlane.Fields("id").Value
            rsPlane.MoveNext
            GoTo plane_loop
        End If
    End If
End If
' print out summation of planes
For X = LBound(cp1) To UBound(cp1)
    If cp1(X) > 0 Then
        If cp1(X) > 1 Then
            strPlane = strPlane + " " + CStr(cp1(X))
        End If
        lname = cp2(X)
        rsBuild.Seek "=", "p", lname
        If Not rsBuild.NoMatch Then
           If Len(rsBuild.Fields("desc")) > 0 Then
                lname = rsBuild.Fields("desc")
           End If
        End If
        strPlane = strPlane + " " + lname + ","
    End If
Next X
If Len(strPlane) > 1 Then
    strPlane = Left$(strPlane, Len(strPlane) - 1)
End If
If plane = 1 Then
    strPlane = strPlane + " " + BuildMissionString("p", lastid)
End If

'100603 rjk: Added nuke information.
'Look for nukes
rsNuke.Seek "=", MxPos, MyPos
If Not rsNuke.NoMatch Then
nuke_loop:
    If (Not rsNuke.EOF) Then
        If rsNuke.Fields("x").Value = MxPos And rsNuke.Fields("y").Value = MyPos Then
            If Len(strNuke) = 0 Then
                strNuke = " Nukes:"
            End If
            strNuke = strNuke + " " + rsNuke.Fields("type").Value + _
                    "(" + CStr(rsNuke.Fields("num").Value) + "),"
            nuke = nuke + 1
            If nuke = 1 Then
                strNukeid = CStr(rsNuke.Fields("id").Value)
            Else
                strNukeid = strNukeid + "/" + CStr(rsNuke.Fields("id").Value)
            End If
            rsNuke.MoveNext
            GoTo nuke_loop
        End If
    End If
    'remove last comma
    If Len(strNuke) > 0 Then
        strNuke = Left$(strNuke, Len(strNuke) - 1)
    End If
End If

'Look for enemy units
strLastOwner = vbNullString
rsEnemyUnit.Seek "=", MxPos, MyPos
If Not rsEnemyUnit.NoMatch Then
enemy_loop:
    If (Not rsEnemyUnit.EOF) Then
        If rsEnemyUnit.Fields("x").Value = MxPos And rsEnemyUnit.Fields("y").Value = MyPos Then
            sname = rsEnemyUnit.Fields("type").Value
            rsBuild.Seek "=", rsEnemyUnit.Fields("class").Value, sname
            If Not rsBuild.NoMatch Then
                If Len(rsBuild.Fields("desc")) > 0 Then
                    sname = rsBuild.Fields("desc")
                End If
            End If
            
            'for enemy ships, build the enemy id string
            If rsEnemyUnit.Fields("class").Value = "s" Then
                enemy = enemy + 1
                If enemy = 1 Then
                    strEnemyid = CStr(rsEnemyUnit.Fields("id").Value)
                Else
                    strEnemyid = strEnemyid + "/" + CStr(rsEnemyUnit.Fields("id").Value)
                End If
            End If

            'get the owner
            strOwner = Nations.NationName(rsEnemyUnit.Fields("owner").Value)
            If strOwner = vbNullString Then
                strOwner = CStr(rsEnemyUnit.Fields("owner").Value)
            End If
            
            'use last owner string to prevent repeating the owner
            'name on multiple ships
            If strOwner = strLastOwner Then
                strOwner = vbNullString
            Else
                strLastOwner = strOwner
            End If
            
            strEnemy = strEnemy + " " + strOwner + " " + sname + " (#" + CStr(rsEnemyUnit.Fields("id").Value) + _
                "),"
            rsEnemyUnit.MoveNext
            GoTo enemy_loop
        End If
    End If
    'remove first space, last comma
    If Len(strEnemy) > 0 Then
        strEnemy = Trim$(Left$(strEnemy, Len(strEnemy) - 1))
    End If
End If
'Combine strings
'100603 rjk: Added nuke information.
'If strShip = vbnullstring And strLand = vbnullstring And strEnemy = vbnullstring Then
'    sb1.Panels.Item("Hex").Text = sb1.Panels.Item("Hex").Text + strSect
'Else
    sb1.Panels.Item("Hex").Text = sb1.Panels.Item("Hex").Text + strEnemy + strShip + strLand + strNuke + strPlane
'End If
End Sub

Private Sub lstCmdResult_KeyPress(KeyAscii As Integer)
'Trap Control C and Copy to Clipboard
If KeyAscii = 3 Then
    CopyListBoxToClipboard lstCmdResult
End If
End Sub

Private Sub ToggleAutoRead()
If frmEmpCmd.bAutoRead Then
    frmEmpCmd.bAutoRead = False
Else
    frmEmpCmd.bAutoRead = True
End If
UpdateAutoRead
End Sub

Public Sub UpdateAutoRead()
If frmEmpCmd.bAutoRead Then
    tbMain.Buttons(4).ToolTipText = "AutoRead is On"
    tbMain.Buttons(4).Image = 24
Else
    tbMain.Buttons(4).ToolTipText = "AutoRead is Off"
    tbMain.Buttons(4).Image = 25
End If
End Sub

Private Sub ToggleAutoUpdate()
If frmEmpCmd.bAutoUpdate Then
    frmEmpCmd.bAutoUpdate = False
Else
    frmEmpCmd.bAutoUpdate = True
    UpdateCommands
    If VersionCheck(4, 3, 10) <> VER_LESS Then
        frmEmpCmd.SubmitEmpireCommand "xdump game *", False
    End If
End If
UpdateAutoUpdate
End Sub

Public Sub UpdateAutoUpdate()
If frmEmpCmd.bAutoUpdate Then
    tbMain.Buttons(3).ToolTipText = "AutoUpdate is On"
    tbMain.Buttons(3).Image = 2
Else
    tbMain.Buttons(3).ToolTipText = "AutoUpdate is Off"
    tbMain.Buttons(3).Image = 23
End If
End Sub

Public Sub UpdateBars()
tbMain.Visible = frmOptions.bToolbar
sb1.Visible = frmOptions.bStatusBar
ResizePanels
ArrangeControls
End Sub

'100903 rjk: Rename AcceptReport so as to not be confused with Accept from the Reject command
Private Sub mnuDiploAcceptanceReport_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

'Setup Designation dialog and execute
Set PromptForm = frmPromptNation
PromptUp = True
PromptForm.cbNations.Move PromptForm.cbRelations.Left, PromptForm.cbRelations.top
PromptForm.cbRelations.Visible = False
PromptForm.Label1.Visible = False
PromptForm.Label2.Caption = "Accept"

'Put form in proper place
PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height

'Dialog box will take it from here
PromptForm.Show vbModeless

End Sub

Private Sub mnuDiploCollect_Click()
If PromptUp Then
    Exit Sub
End If

'Setup Designation dialog and execute
Set PromptForm = frmPromptCollect
PromptUp = True

'Put form in proper place
PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height

'Dialog box will take it from here
PromptForm.Show vbModeless

End Sub

Private Sub mnuDiploConsider_Click()
If PromptUp Then
    Exit Sub
End If

'Setup Designation dialog and execute
Set PromptForm = frmPromptOffer
PromptUp = True

'100903 rjk: Moved the field logic to the form
PromptForm.Label2.Caption = "Consider"

'Dialog box will take it from here
PromptForm.Show vbModeless
End Sub

Private Sub mnuDiploDeclare_Click()

'check for current prompt
If PromptUp Then
    Exit Sub
End If

'Setup Designation dialog and execute
Set PromptForm = frmPromptNation
PromptUp = True
PromptForm.cbRelations.Visible = True
PromptForm.Label1.Visible = True
PromptForm.Label2.Caption = "Declare"

'Put form in proper place
PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height

'Dialog box will take it from here
PromptForm.Show vbModeless


End Sub

Private Sub mnuDiploOffer_Click()
If PromptUp Then
    Exit Sub
End If

'Setup Designation dialog and execute
Set PromptForm = frmPromptOffer
PromptUp = True

'100903 rjk: Moved field logic to form
'Dialog box will take it from here
PromptForm.Show vbModeless
End Sub

'100903 rjk: Split Reject into Reject and Accept separate menu items
Private Sub mnuDiploReject_Click()
Set PromptForm = frmPromptOffer
PromptUp = True
PromptForm.Label2.Caption = "Reject"

'Dialog box will take it from here
PromptForm.Show vbModeless

'frmEmpCmd.SubmitEmpireCommand "reject", True
End Sub

Private Sub mnuDiploAccept_Click()
Set PromptForm = frmPromptOffer
PromptUp = True
PromptForm.Label2.Caption = "Accept"

'Dialog box will take it from here
PromptForm.Show vbModeless

'frmEmpCmd.SubmitEmpireCommand "reject", True
End Sub

Private Sub mnuDiploRelations_Click()

'check for current prompt
If PromptUp Then
    Exit Sub
End If

'Setup Designation dialog and execute
Set PromptForm = frmPromptNation
PromptUp = True
PromptForm.cbNations.Move PromptForm.cbRelations.Left, PromptForm.cbRelations.top
PromptForm.cbRelations.Visible = False
PromptForm.Label1.Visible = False
PromptForm.Label2.Caption = "Relations"

'Put form in proper place
PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height

'Dialog box will take it from here
PromptForm.Show vbModeless

End Sub

Private Sub mnuDiploRepay_Click()
If PromptUp Then
    Exit Sub
End If

'Setup Designation dialog and execute
Set PromptForm = frmPromptCollect
PromptUp = True
PromptForm.Label2 = "Repay"
'Put form in proper place
PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height

'Dialog box will take it from here
PromptForm.Show vbModeless


End Sub

Private Sub mnuDiploSharebmap_Click()

'check for current prompt
If PromptUp Then
    Exit Sub
End If

'Setup Designation dialog and execute
Set PromptForm = frmPromptNation
PromptUp = True
PromptForm.cbRelations.Visible = False
PromptForm.Label1.Visible = True
PromptForm.Label2.Caption = "Sharebmap"
PromptForm.txtOrigin.Visible = True
PromptForm.txtNum.Visible = True
PromptForm.Label3.Visible = True
PromptForm.txtOrigin.top = PromptForm.cbRelations.top
PromptForm.txtNum.top = PromptForm.cbRelations.top
PromptForm.Label3.top = PromptForm.cbRelations.top

'Put form in proper place
PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height

'Dialog box will take it from here
PromptForm.Show vbModeless

End Sub

Private Sub mnuDiploShark_Click()
If PromptUp Then
    Exit Sub
End If

'Setup Designation dialog and execute
Set PromptForm = frmPromptCollect
PromptUp = True
PromptForm.Label2 = "Shark"
'Put form in proper place
PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height

'Dialog box will take it from here
PromptForm.Show vbModeless



End Sub

Private Sub mnuDiploTreaty_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmEmpCmd.SubmitEmpireCommand "treaty *", True
End Sub

Private Sub mnuLandCargo_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmPromptLand.Label2 = "Land Cargo"
frmPromptLand.cmd = "lcargo"
mnuUnitLand_Click

End Sub

Private Sub mnuLandFire_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

'Setup Designation dialog and execute
Set PromptForm = frmPromptFire

'Show Prompt
ShowUnitPrompt "land"

PromptForm.Option2.Value = True 'need to set after loaded

End Sub

Private Sub mnuLandFuel_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptFuel

PromptForm.Option2.Value = True
'PromptForm.strSector = SectorString(MxPos, MyPos) 091503 rjk: removed, strSector not used in frmPromptFuel
    
'Show Prompt
ShowUnitPrompt "land"

End Sub

Private Sub mnuLandLoad_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptLoad
PromptUp = True
PromptForm.strCmd = "lload"
'PromptForm.strSector = SectorString(MxPos, MyPos) 092103 rjk: Switched to SelX and SelY for top level menu
PromptForm.strSector = SectorString(SelX, SelY)

'Show Prompt
ShowUnitPrompt "land"

End Sub

Private Sub mnuLandLook_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

'If there is only one unit, don't bother prompting
If InStr(strLandid, "/") = 0 And Len(strLandid) > 0 Then
    frmEmpCmd.SubmitEmpireCommand "llook " + strLandid, True
    frmEmpCmd.SubmitEmpireCommand "db2", False  'force map redraw
Else
    Set PromptForm = frmPromptLook
    PromptForm.strCmd = "llook"
        
    'Show Prompt
    ShowUnitPrompt "land"
End If

End Sub

Private Sub mnulandMarch_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptNav
    
PromptForm.Label2 = "March"
'PromptForm.strSector = SectorString(MxPos, MyPos) 092303 rjk: Switched to SelX and SelY for top level menu
PromptForm.strSector = SectorString(SelX, SelY)
    
'Show Prompt
ShowUnitPrompt "land"

If Len(PromptForm.txtUnit.Text) > 0 Then
    PromptForm.txtPath.SetFocus
End If
End Sub

Private Sub mnuLandMission_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptMission
'Setup op sector hex
PromptForm.txtOrigin = "."
PromptForm.Option2.Value = True
    
'Show Prompt
ShowUnitPrompt "land"

End Sub

Private Sub mnuLandMorale_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptSail
PromptForm.strCmd = "morale"
PromptForm.Label2 = "Morale"
PromptForm.Label1 = "retreat %"
    
'Show Prompt
ShowUnitPrompt "land"

End Sub

Private Sub mnuLandRadar_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptLook
PromptForm.strCmd = "lradar"
PromptForm.Label2 = "Radar"
    
'Show Prompt
ShowUnitPrompt "land"

End Sub

Private Sub mnuLandScrap_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptScrap
PromptForm.strCmd = "scrap "
PromptForm.Label2 = "Scrap"
PromptForm.Option2 = True
    
'Show Prompt
ShowUnitPrompt "land"
End Sub

Private Sub mnuLandScuttle_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptScrap
    
PromptForm.strCmd = "scuttle"
PromptForm.Label2 = "Scuttle"
PromptForm.Option2 = True
    
'Show Prompt
ShowUnitPrompt "land"
End Sub

Private Sub mnuLandStat_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmPromptLand.Label2 = "Land Stats"
frmPromptLand.cmd = "lstat"
mnuUnitLand_Click

End Sub

Private Sub mnuLandSupply_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptLook
    
PromptForm.strCmd = "supply"
PromptForm.Label2 = "Supply"
    
'Show Prompt
ShowUnitPrompt "land"

End Sub

Private Sub mnuLandTend_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptLoad
    
PromptForm.strCmd = "ltend"
'Show Prompt
ShowUnitPrompt "ship" '100303 rjk: Switched to ship from land so initial field show ships.
End Sub

Private Sub mnuLandUnload_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptLoad
    
PromptForm.strCmd = "lunload"
'PromptForm.strSector = SectorString(MxPos, MyPos) 092103 rjk: Switched to SelX and SelY for top level menu
PromptForm.strSector = SectorString(SelX, SelY)
'Show Prompt
ShowUnitPrompt "land"
End Sub

Private Sub mnuLandUpgrade_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptScrap
PromptForm.strCmd = "upgrade "
PromptForm.Label2 = "Upgrade"
PromptForm.Option2 = True
    
'Show Prompt
ShowUnitPrompt "land"
End Sub

Private Sub mnuMarketTrade_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmEmpCmd.SubmitEmpireCommand "trade", True
End Sub

Private Sub mnuPlaneBomb_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptBomb
    
'Setup op sector hex
'092303 rjk: Switched to only use SelX and SelY, fixes top level menu access
'If PopUpMenuSource = DMAP_PUMS_MAP Then
'    PromptForm.txtOrigin.Text = SectorString(MxPos, MyPos)
'Else
    PromptForm.txtOrigin.Text = SectorString(SelX, SelY)
'End If
    
'Show Prompt
ShowUnitPrompt "plane"

End Sub

'Private Sub mnuPlaneCargo_Click() 092303 rjk: empire command does not exist
''check for current prompt
'If PromptUp Then
'    Exit Sub
'End If
'
'frmPromptLand.Label2 = "Plane Cargo"
'frmPromptLand.cmd = "pcargo"
'mnuUnitLand_Click
'End Sub

Private Sub mnuPlaneDrop_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptRecon
'Setup op sector hex
'092303 rjk: Switched to only use SelX and SelY, fixes top level menu access
'If PopUpMenuSource = DMAP_PUMS_MAP Then
'    PromptForm.txtOrigin.Text = SectorString(MxPos, MyPos)
'Else
    PromptForm.txtOrigin.Text = SectorString(SelX, SelY)
'End If
PromptForm.Label2.Caption = "Drop"
        
'Show Prompt
ShowUnitPrompt "plane"

End Sub

Private Sub mnuPlaneFly_Click()
Set PromptForm = frmPromptRecon
    
'Setup op sector hex
'092303 rjk: Switched to only use SelX and SelY, fixes top level menu access
'If PopUpMenuSource = DMAP_PUMS_MAP Then
'    PromptForm.txtOrigin.Text = SectorString(MxPos, MyPos)
'Else
    PromptForm.txtOrigin.Text = SectorString(SelX, SelY)
'End If
PromptForm.Label2.Caption = "Fly"
    
'Show Prompt
ShowUnitPrompt "plane"

End Sub

Private Sub mnuPlaneMission_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptMission
    
'Setup op sector hex
PromptForm.txtOrigin = "."
PromptForm.Option3.Value = True
    
'Show Prompt
ShowUnitPrompt "plane"

End Sub

Private Sub mnuPlanePara_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptRecon
    
'Setup op sector hex
'092303 rjk: Switched to only use SelX and SelY, fixes top level menu access
'If PopUpMenuSource = DMAP_PUMS_MAP Then
'    PromptForm.txtOrigin.Text = SectorString(MxPos, MyPos)
'Else
    PromptForm.txtOrigin.Text = SectorString(SelX, SelY)
'End If
PromptForm.Label2.Caption = "Paradrop"
    
'Show Prompt
ShowUnitPrompt "plane"

End Sub

Private Sub mnuPlaneRange_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptSail
    
PromptForm.strCmd = "range"
PromptForm.Label2 = "Range"
PromptForm.Label1 = "radius"
    
'Show Prompt
ShowUnitPrompt "plane"

End Sub

Private Sub mnuPlaneRecon_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptRecon
PromptForm.Label2.Caption = "Recon"
    
'Setup op sector hex
'092303 rjk: Switched to only use SelX and SelY, fixes top level menu access
'If PopUpMenuSource = DMAP_PUMS_MAP Then
'    PromptForm.txtOrigin.Text = SectorString(MxPos, MyPos)
'Else
    PromptForm.txtOrigin.Text = SectorString(SelX, SelY)
'End If

'Show Prompt
ShowUnitPrompt "plane"

End Sub

Private Sub mnuPlaneStats_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmPromptLand.Label2 = "Plane Stats"
frmPromptLand.cmd = "pstat"
mnuUnitLand_Click
End Sub

Private Sub mnuPlaneSweep_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptRecon
    
'Setup op sector hex
'092303 rjk: Switched to only use SelX and SelY, fixes top level menu access
'If PopUpMenuSource = DMAP_PUMS_MAP Then
'    PromptForm.txtOrigin.Text = SectorString(MxPos, MyPos)
'Else
    PromptForm.txtOrigin.Text = SectorString(SelX, SelY)
'End If
PromptForm.Label2.Caption = "Sweep"
    
'Show Prompt
ShowUnitPrompt "plane"

End Sub

Private Sub mnuPlaneTrans_Click()
Set PromptForm = frmPromptTrans
PromptUp = True
    
'Setup op sector hex
'PromptForm.strSector = SectorString(MxPos, MyPos) switched to SelX and SelY for top level menu
PromptForm.strSector = SectorString(SelX, SelY)

'Show Prompt
ShowUnitPrompt "plane"

End Sub

Private Sub mnuReportBudget_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmEmpCmd.SubmitEmpireCommand "budget", True
End Sub

Public Sub mnuReportCensus_Click()
'check for current prompt

If PromptUp Then
    Exit Sub
End If

'Setup Designation dialog and execute
'frmPromptCensus.txtOrigin = SectorString(MxPos, MyPos) 092103 rjk: Switch to SelX and SelY for top level menu
frmPromptCensus.txtOrigin = SectorString(SelX, SelY)

'Put form in proper place
frmPromptCensus.Left = frmDrawMap.Left + (frmDrawMap.Width - frmPromptCensus.Width) \ 2
frmPromptCensus.top = frmDrawMap.top + frmDrawMap.Height - frmPromptCensus.Height
Set PromptForm = frmPromptCensus
PromptUp = True

'Dialog box will take it from here
frmPromptCensus.Show vbModeless

End Sub

Private Sub mnuReportCoastwatch_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmEmpCmd.SubmitEmpireCommand "coastwatch *", True
frmEmpCmd.SubmitEmpireCommand "db2", False  'redraw map
End Sub

'111903 rjk: Added Commodity Total Report
Private Sub mnuReportCommodity_Click(Index As Integer)
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Select Case Index
Case 1
    frmPromptCensus.Label2.Caption = "Commodity"
    mnuReportCensus_Click
Case 2
    Set PromptForm = frmPromptLoad
    PromptForm.strCmd = "commoditytotal"
    PromptForm.strSector = SectorString(SelX, SelY)
    PromptUp = True
    PromptForm.Show vbModeless
    PromptForm.txtMultOrigin.SetFocus
End Select
End Sub

Private Sub mnuReportCountry_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmEmpCmd.SubmitEmpireCommand "country *", True
End Sub

Private Sub mnuReportCutoff_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmPromptCensus.Label2.Caption = "Cutoff"
mnuReportCensus_Click
End Sub

Private Sub mnuReportFinancial_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmEmpCmd.SubmitEmpireCommand "financial", True
End Sub

Private Sub mnuReportLevel_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmPromptCensus.Label2.Caption = "Level"
mnuReportCensus_Click
End Sub

Private Sub mnuReportMap_Click(Index As Integer)
'check for current prompt
If PromptUp Then
    Exit Sub
End If

'Setup Map dialog and execute
Set PromptForm = frmPromptMap
PromptUp = True

'prompt for parameters
Dim strCmd As String
strCmd = Trim$(Me.mnuReportMap.Item(Index).Caption)
'092303 rjk: Remove & from strCmd so submit works
If InStr(strCmd, "&") <> 0 Then
    If InStr(strCmd, "&") = 1 Then
        strCmd = Right$(strCmd, Len(strCmd) - 1)
    Else
        strCmd = Left$(strCmd, InStr(strCmd, "&") - 1) + Mid$(strCmd, InStr(strCmd, "&") + 1)
    End If
End If
frmPromptMap.Label2.Caption = strCmd


'Put form in proper place
PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height

PromptForm.Show vbModeless

End Sub

Private Sub mnuReportMotd_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmEmpCmd.SubmitEmpireCommand "motd", True
End Sub

Private Sub mnuReportNation_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmEmpCmd.SubmitEmpireCommand "nation", True
End Sub

Private Sub mnuReportNewsPaper_Click()
' Dim strCmd As String   removed efj 8/2003
' Dim strConnect As String   removed efj 8/2003

'check for current prompt
If PromptUp Then
    Exit Sub
End If

'Setup Designation dialog and execute
Set PromptForm = frmPromptNews
PromptUp = True

'Put form in proper place
PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height

'prompt for parameters
frmPromptNews.Show vbModeless

End Sub

Private Sub mnuReportNuke_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmPromptLand.Label2 = "Nukes"
frmPromptLand.cmd = "nuke"
mnuUnitLand_Click
End Sub

Private Sub mnuReportPlayers_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmEmpCmd.SubmitEmpireCommand "players", True
End Sub

Private Sub mnuReportPower_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmEmpCmd.SubmitEmpireCommand "power", True
End Sub

Private Sub mnuReportPowerNew_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmEmpCmd.SubmitEmpireCommand "power new", True
End Sub

Private Sub mnuReportProduction_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmPromptCensus.Label2.Caption = "Production"
frmPromptCensus.txtOrigin.Text = "*"
'Put form in proper place
frmPromptCensus.Left = frmDrawMap.Left + (frmDrawMap.Width - frmPromptCensus.Width) \ 2
frmPromptCensus.top = frmDrawMap.top + frmDrawMap.Height - frmPromptCensus.Height
Set PromptForm = frmPromptCensus
PromptUp = True

'Dialog box will take it from here
frmPromptCensus.Show vbModeless
End Sub

'120203 rjk: Added to do multiple sector Start and Stops
Private Sub mnuReportStart_Click(Index As Integer)
'check for current prompt
If PromptUp Then
    Exit Sub
End If

If Index = 1 Then
    frmPromptCensus.Label2.Caption = "Start"
Else
    frmPromptCensus.Label2.Caption = "Stop"
End If
frmPromptCensus.cmbType.ListIndex = 0

frmPromptCensus.txtOrigin = SectorString(SelX, SelY)
'Put form in proper place
frmPromptCensus.Left = frmDrawMap.Left + (frmDrawMap.Width - frmPromptCensus.Width) \ 2
frmPromptCensus.top = frmDrawMap.top + frmDrawMap.Height - frmPromptCensus.Height
Set PromptForm = frmPromptCensus
PromptUp = True
'Dialog box will take it from here
frmPromptCensus.Show vbModeless
End Sub

Private Sub mnuReportRead_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If
  
SubmitTelegramRead False, True
End Sub

Private Sub mnuReportReport_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmEmpCmd.SubmitEmpireCommand "report *", True
End Sub

Private Sub mnuReportResource_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmPromptCensus.Label2.Caption = "Resource"
mnuReportCensus_Click
End Sub

Private Sub mnuReportSky_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmEmpCmd.SubmitEmpireCommand "skywatch *", True
End Sub

Private Sub mnuReportsLedger_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmEmpCmd.SubmitEmpireCommand "ledger *", True
End Sub

Private Sub mnuReportStarve_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmEmpCmd.SubmitEmpireCommand "bf1", False  'run as batch file
frmEmpCmd.SubmitEmpireCommand "starvation s *", True  'ships
frmEmpCmd.SubmitEmpireCommand "starvation l *", True  'lands
frmEmpCmd.SubmitEmpireCommand "starvation *", True    'sectors
frmEmpCmd.SubmitEmpireCommand "bf2", False  'run as batch file
frmEmpCmd.SubmitEmpireCommand "db2", False  'force map redraw
End Sub

Private Sub mnuReportUpdate_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

If VersionCheck(4, 3, 10) <> VER_LESS Then
    frmEmpCmd.SubmitEmpireCommand "show updates", True
Else
    frmEmpCmd.SubmitEmpireCommand "update", True
End If
End Sub

Private Sub mnuReportVersion_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

'If VersionCheck(4, 3, 10) <> VER_LESS Then
'    frmEmpCmd.SubmitEmpireCommand "xdump ver", True
'Else
    frmEmpCmd.SubmitEmpireCommand "version", True
'End If
End Sub

Private Sub mnuReportWire_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptWire
PromptUp = True

'Put form in proper place
PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height

'Dialog box will take it from here
PromptForm.Show vbModeless
End Sub

Private Sub mnuSectCmdAnti_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

'Setup Designation dialog and execute
Set PromptForm = frmPromptAnti
PromptUp = True
'PromptForm.txtOrigin.Text = SectorString(MxPos, MyPos) 092303 rjk: Switch to SelX and SelY for top level menu
PromptForm.txtOrigin.Text = SectorString(SelX, SelY)

'Put form in proper place
PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height

'Dialog box will take it from here
PromptForm.Show vbModeless
End Sub

Private Sub mnuSectCmdBoard_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

'Setup dialog
'frmPromptBoard.txtOrigin.Text = SectorString(MxPos, MyPos) 092203 rjk: switched to SelX and SelY for top level menu
frmPromptBoard.txtOrigin.Text = SectorString(SelX, SelY)
frmPromptBoard.txtOrigin.Visible = True
frmPromptBoard.txtUnit.Visible = False
frmPromptBoard.Frame3.Caption = "Attacking Sector"
frmPromptBoard.txtTarget = strEnemyid
Set PromptForm = frmPromptBoard
PromptUp = True

'Put form in proper place
PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height

'Dialog box will take it from here
PromptForm.Show vbModeless
PromptForm.txtTarget.SetFocus '092203 rjk: Added to set the focus to the enemy list automatically
End Sub

Private Sub mnuSectCmdBuild_Click()
mnuUnitBuild_Click
End Sub

Private Sub mnuSectCmdCap_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

'CapSect = SectorString(MxPos, MyPos) 092303 rjk: Switched to SelX and SelY for top level menu
CapSect = SectorString(SelX, SelY)
'frmEmpCmd.SubmitEmpireCommand "capital " + SectorString(MxPos, MyPos), True
CapSect = SectorString(SelX, SelY)
frmEmpCmd.SubmitEmpireCommand "capital " + SectorString(SelX, SelY), True

End Sub

Private Sub mnuSectCmdConvert_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

'Setup Designation dialog and execute
'frmPromptConvert.txtOrigin.Text = SectorString(MxPos, MyPos) 092203 rjk: switched to SelX and SelY for top level menu
frmPromptConvert.txtOrigin.Text = SectorString(SelX, SelY)
Set PromptForm = frmPromptConvert
PromptUp = True

'Put form in proper place
frmPromptConvert.Left = frmDrawMap.Left + (frmDrawMap.Width - frmPromptConvert.Width) \ 2
frmPromptConvert.top = frmDrawMap.top + frmDrawMap.Height - frmPromptConvert.Height
Set PromptForm = frmPromptConvert
PromptUp = True

'Dialog box will take it from here
frmPromptConvert.Show vbModeless

End Sub

Private Sub mnuSectCmdDel_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

'Setup Designation dialog and execute
Set PromptForm = frmPromptMove
PromptUp = True
PromptForm.Label2.Caption = "Deliver"
'PromptForm.txtOrigin.Text = SectorString(MxPos, MyPos) switched to SelX and SelY for top level menu
'101903 rjk: Switched txtMultOrigin for multiple sector selection
PromptForm.txtMultOrigin.Text = SectorString(SelX, SelY)
'100203 rjk: Moved the field logic to the form.

'Dialog box will take it from here
PromptForm.Show vbModeless
End Sub

Private Sub mnuSectCmdDemob_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmPromptConvert.Label2.Caption = "Demobilize"
frmPromptConvert.Label1.Caption = "military in"
frmPromptConvert.chkReserve.Visible = True
mnuSectCmdConvert_Click
End Sub

Private Sub mnuSectCmdDes_Click()

'Setup Designation dialog and execute
'check for current prompt
If PromptUp Then
    Exit Sub
End If

'frmPromptDes.txtOrigin.Text = SectorString(MxPos, MyPos) 092303 rjk: Switched to SelX and SelY for top level menu
'101903 rjk: Switched form to Multiple Sector selection
'frmPromptDes.txtOrigin.Text = SectorString(SelX, SelY)
'101903 rjk: Switched txtMultOrigin for multiple sector selection
frmPromptDes.txtMultOrigin.Text = SectorString(SelX, SelY)

Set PromptForm = frmPromptDes
PromptUp = True

'Put form in proper place
frmPromptDes.Left = frmDrawMap.Left + (frmDrawMap.Width - frmPromptDes.Width) \ 2
frmPromptDes.top = frmDrawMap.top + frmDrawMap.Height - frmPromptDes.Height


'Dialog box will take it from here
frmPromptDes.Show vbModeless

End Sub

Private Sub mnuSectCmdDist_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

'Setup Designation dialog and execute
'frmPromptDist.txtOrigin = SectorString(MxPos, MyPos) 092303 rjk: Switched to SelX and SelY for top level menu
frmPromptDist.txtOrigin = SectorString(SelX, SelY)

'Put form in proper place
frmPromptDist.Left = frmDrawMap.Left + (frmDrawMap.Width - frmPromptDist.Width) \ 2
frmPromptDist.top = frmDrawMap.top + frmDrawMap.Height - frmPromptDist.Height
Set PromptForm = frmPromptDist
PromptUp = True

'Dialog box will take it from here
PromptForm.Show vbModeless
PromptForm.txtMultDest.SetFocus '101803 rjk: Switched txtMultDest for multiple sectors selection

End Sub

Private Sub mnuSectCmdEnlist_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmPromptConvert.Label2.Caption = "Enlist"
mnuSectCmdConvert_Click
End Sub

Private Sub mnuSectCmdExplore_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

'Setup Designation dialog and execute
'frmPromptExplore.txtOrigin = SectorString(MxPos, MyPos) 0920303 rjk: Switched to SelX and SelY for top level menu
frmPromptExplore.txtOrigin = SectorString(SelX, SelY)

'Put form in proper place
frmPromptExplore.Left = frmDrawMap.Left + (frmDrawMap.Width - frmPromptExplore.Width) \ 2
frmPromptExplore.top = frmDrawMap.top + frmDrawMap.Height - frmPromptExplore.Height
Set PromptForm = frmPromptExplore
PromptUp = True

'Dialog box will take it from here
frmPromptExplore.Show vbModeless

End Sub

Private Sub mnuSectCmdFire_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

'Setup Designation dialog and execute
'frmPromptFire.txtOrigin.Text = SectorString(MxPos, MyPos) 092303 rjk: Switched for top level menu
frmPromptFire.txtOrigin.Text = SectorString(SelX, SelY)
Set PromptForm = frmPromptFire
PromptUp = True

'Put form in proper place
PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height

'Dialog box will take it from here
PromptForm.Show vbModeless

PromptForm.Option3.Value = True

On Error Resume Next
PromptForm.txtTarget.SetFocus
End Sub

Private Sub mnuSectCmdGrind_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmPromptConvert.Label2.Caption = "Grind"
frmPromptConvert.Label1.Caption = "gold bars in"
mnuSectCmdConvert_Click
End Sub

Private Sub mnuSectCmdImprove_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

'Setup Designation dialog and execute
'frmPromptImprove.txtOrigin.Text = SectorString(MxPos, MyPos) 092103 rjk: Switched to SelX and SelY for top level menu access
frmPromptImprove.txtOrigin.Text = SectorString(SelX, SelY)
Set PromptForm = frmPromptImprove
PromptUp = True

'Put form in proper place
PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height

'Dialog box will take it from here
PromptForm.Show vbModeless

End Sub

Private Sub mnuSectCmdMove_Click()

Dim n As Integer

'check for current prompt
If PromptUp Then
    Exit Sub
End If

'Setup Designation dialog and execute
Set PromptForm = frmPromptMove
PromptUp = True
'PromptForm.txtOrigin.Text = SectorString(MxPos, MyPos) switched to SelX and SelY for top level menu
'101903 rjk: Switched txtMultOrigin for multiple sector selection
PromptForm.txtMultOrigin.Text = SectorString(SelX, SelY)
'100203 rjk: Moved the field logic to the form.

'Dialog box will take it from here
PromptForm.Show vbModeless
PromptForm.txtNum.SetFocus

End Sub

Private Sub mnuSectCmdNeweff_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmPromptCensus.Label2.Caption = "Neweff"
mnuReportCensus_Click
End Sub

Private Sub mnuSectCmdOrigin_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

'Setup Designation dialog and execute
'frmPromptOrigin.txtOrigin.Text = SectorString(MxPos, MyPos) 092203 rjk: Switched to SelX and SelY for top level menu
frmPromptOrigin.txtOrigin.Text = SectorString(SelX, SelY)
Set PromptForm = frmPromptOrigin
PromptUp = True

'Put form in proper place
frmPromptOrigin.Left = frmDrawMap.Left + (frmDrawMap.Width - frmPromptOrigin.Width) \ 2
frmPromptOrigin.top = frmDrawMap.top + frmDrawMap.Height - frmPromptOrigin.Height

'Dialog box will take it from here
frmPromptOrigin.Show vbModeless

End Sub

Private Sub mnuSectCmdProduction_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmPromptCensus.Label2.Caption = "Production"
mnuReportCensus_Click
frmPromptCensus.cmdOK_Click
End Sub

Private Sub mnuSectCmdRadar_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

'Setup Designation dialog and execute
Set PromptForm = frmPromptRadar
PromptUp = True
'PromptForm.txtOrigin.Text = SectorString(MxPos, MyPos) 092203 rjk: Switched to SelX and SelY for toplevel menu
PromptForm.txtOrigin.Text = SectorString(SelX, SelY)
PromptForm.Label2.Caption = "Radar"

'Put form in proper place
PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height

'Dialog box will take it from here
frmPromptRadar.Show vbModeless

End Sub

Private Sub mnuSectCmdShoot_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

'Setup Designation dialog and execute
'frmPromptShoot.txtOrigin.Text = SectorString(MxPos, MyPos) 092203 rjk: Changed to SelX and SelY for top level menu
frmPromptShoot.txtOrigin.Text = SectorString(SelX, SelY)
Set PromptForm = frmPromptShoot
PromptUp = True

'Put form in proper place
frmPromptShoot.Left = frmDrawMap.Left + (frmDrawMap.Width - frmPromptShoot.Width) \ 2
frmPromptShoot.top = frmDrawMap.top + frmDrawMap.Height - frmPromptShoot.Height

'default = number of civs in sector
'092203 rjk: Changed to SelX and SelY for top level menu
'rsSectors.Seek "=", MxPos, MyPos
rsSectors.Seek "=", SelX, SelY
If Not rsSectors.NoMatch Then
    frmPromptShoot.txtNum = CStr(rsSectors.Fields("civ"))
End If
'Dialog box will take it from here
frmPromptShoot.Show vbModeless

End Sub

Private Sub mnuSectCmdSpy_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

' 092303 rjk: Switched to SelX and SelY for top level menu
'frmEmpCmd.SubmitEmpireCommand "spy " + SectorString(MxPos, MyPos), True
frmEmpCmd.SubmitEmpireCommand "spy " + SectorString(SelX, SelY), True

'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
'101603 rjk: Does not need a dump as none of our sectors change
'GetSectors
'101703 rjk: Added Strength fields to Sector database
'frmEmpCmd.SubmitEmpireCommand "strength * ?timestamp>" + CStr(tsSectors), False
frmEmpCmd.SubmitEmpireCommand "db2", False

End Sub

Private Sub mnuSectCmdStart_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmPromptCensus.Label2.Caption = "Start"
frmPromptCensus.cmbType.ListIndex = 0
mnuReportCensus_Click
frmPromptCensus.cmdOK_Click
End Sub

Private Sub mnuSectCmdStop_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmPromptCensus.Label2.Caption = "Stop"
frmPromptCensus.cmbType.ListIndex = 0
mnuReportCensus_Click
frmPromptCensus.cmdOK_Click
End Sub

Private Sub mnuSectCmdStrength_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmPromptCensus.Label2.Caption = "Strength"
mnuReportCensus_Click
End Sub

Private Sub mnuSectCmdThr_Click()

'check for current prompt
If PromptUp Then
    Exit Sub
End If

'Setup Designation dialog and execute
Set PromptForm = frmPromptThresh
PromptUp = True
'PromptForm.txtOrigin.Text = SectorString(MxPos, MyPos) 092303 rjk: Switched to SelX and SelY for top level menu
'101903 rjk: Switched txtMultOrigin for multiple sector selection
PromptForm.txtMultOrigin.Text = SectorString(SelX, SelY)

'Put form in proper place
PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height

'Dialog box will take it from here
PromptForm.Show vbModeless

End Sub

Private Sub mnuSectCmdThrAll_Click()
' Dim n As Integer 092103 rjk: functionality moved to form
' Dim thresh As Integer   removed efj 8/2003

'check for current prompt
If PromptUp Then
    Exit Sub
End If

'Setup Designation dialog and execute
Set PromptForm = frmPromptThreshAll
PromptUp = True
'PromptForm.txtOrigin.Text = SectorString(MxPos, MyPos) 092303 rjk: Switched to SelX and SelY for top level menu
PromptForm.txtMultOrigin.Text = SectorString(SelX, SelY)

PromptUp = True 'rjk 5/12/03: moved this line until we know it is valid sector and we will display the dialog box

'Dialog box will take it from here
PromptForm.Show vbModeless

End Sub

Private Sub mnuSectCmdWipe_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

'Setup Designation dialog and execute
'frmPromptWipe.txtOrigin.Text = SectorString(MxPos, MyPos) 092303 rjk: Switched to SelX and SelY for top level menu
frmPromptWipe.txtOrigin.Text = SectorString(SelX, SelY)
Set PromptForm = frmPromptWipe
PromptUp = True

'Put form in proper place
frmPromptWipe.Left = frmDrawMap.Left + (frmDrawMap.Width - frmPromptWipe.Width) \ 2
frmPromptWipe.top = frmDrawMap.top + frmDrawMap.Height - frmPromptWipe.Height

'Dialog box will take it from here
frmPromptWipe.Show vbModeless
End Sub

Private Sub mnuShipAssult_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

'Setup dialog
frmPromptAttack.Label2.Caption = "Assault"
Set PromptForm = frmPromptAttack

'Show Prompt
'Put form in proper place
PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height
PromptUp = True

'Fill the combo box with ships in the hex
On Error GoTo PromptExit
Dim slist As Variant
Dim n As Integer
If Len(strShipid) > 0 Then
    slist = ParseUnitString(strShipid)
    If IsArray(slist) Then
        For n = LBound(slist) To UBound(slist)
            If Len(slist(n)) > 0 Then
                frmPromptAttack.cbShips.AddItem slist(n)
            End If
        Next n
        If frmPromptAttack.cbShips.ListCount > 0 Then
            frmPromptAttack.cbShips.ListIndex = 0
        End If
    End If
End If
If (InStr(strShipid, "/") = 0 And Len(strShipid) > 0) Then
    PromptForm.cbShips.Add strShipid
End If

PromptExit:
'Dialog box will take it from here
PromptForm.Show vbModeless

End Sub

Private Sub mnuShipBoard_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

'Setup dialog
frmPromptBoard.txtOrigin.Visible = False
frmPromptBoard.txtUnit.Visible = True
frmPromptBoard.Frame3.Caption = "Attacking Ship"
frmPromptBoard.txtTarget = strEnemyid
Set PromptForm = frmPromptBoard

'Show Prompt
ShowUnitPrompt "ship"

End Sub

Private Sub mnuShipCargo_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmPromptLand.Label2 = "Ship Cargo"
frmPromptLand.cmd = "cargo"
mnuUnitLand_Click

End Sub


Private Sub mnuShipFire_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

'Setup Designation dialog and execute
Set PromptForm = frmPromptFire
PromptUp = True

'Show Prompt
ShowUnitPrompt "ship"

PromptForm.Option1.Value = True ' 091803 rjk ensure form is load, because there is SetFocus in Click event
'PromptForm.txtUnit.Text = strShipid


'if a firing ship has been filled in, switch to target
If Len(PromptForm.txtUnit.Text) > 0 Then
    PromptForm.txtTarget.SetFocus
End If
End Sub

Private Sub mnuShipLoad_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptLoad
PromptForm.strCmd = "load"
'PromptForm.strSector = SectorString(MxPos, MyPos) 092103 rjk: Switched to SelX and SelY for top level menu
PromptForm.strSector = SectorString(SelX, SelY)
'Show Prompt
ShowUnitPrompt "ship"
End Sub

Private Sub mnuShipLook_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

If InStr(strShipid, "/") = 0 And Len(strShipid) > 0 Then
    frmEmpCmd.SubmitEmpireCommand "look " + strShipid, True
    frmEmpCmd.SubmitEmpireCommand "db2", False  'force map redraw
Else
    Set PromptForm = frmPromptLook
    PromptUp = True
    PromptForm.strCmd = "look"
    ShowUnitPrompt "ship"
End If

End Sub

Private Sub mnuShipMine_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptLook
PromptForm.strCmd = "mine"
PromptForm.Label2 = "Mine"

'Show Prompt
ShowUnitPrompt "ship"

End Sub

Private Sub mnuShipMission_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptMission
'Setup op sector hex
PromptForm.txtOrigin = "."
PromptForm.Option1.Value = True

'Show Prompt
ShowUnitPrompt "ship"

End Sub

Private Sub mnuShipMquota_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptSail
PromptForm.strCmd = "mquota"
PromptForm.Label2 = "Mquota"
PromptForm.Label1 = "level"
    
'Show Prompt
ShowUnitPrompt "ship"

End Sub

Private Sub mnuShipNav_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set frmDrawMap.PromptForm = frmPromptNav
frmDrawMap.PromptUp = True
'Setup op sector hex
'092303 rjk: Switched to SelX and SelY for top level menu access
'If PopUpMenuSource = DMAP_PUMS_MAP Then
'    PromptForm.strSector = SectorString(MxPos, MyPos)
'Else
    PromptForm.strSector = SectorString(SelX, SelY)
'End If

'Show Prompt
ShowUnitPrompt "ship"

If Len(PromptForm.txtUnit.Text) > 0 Then
    PromptForm.txtPath.SetFocus
End If
End Sub

Private Sub mnuShipNavMark_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptLook
PromptUp = True
PromptForm.strCmd = "marker"
PromptForm.Label2 = "Marker"

'for single ships, just show the markers and put out the prompt
If InStr(strShipid, "/") = 0 And Len(strShipid) > 0 Then
    'Put form in proper place
    PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
    PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height
    PromptForm.txtUnit = strShipid
    
    'Dialog box will take it from here
    PromptForm.Show vbModeless
    NavMarkerShip = strShipid
    DrawHexes
Else
    ShowUnitPrompt "ship"
End If

End Sub

Private Sub mnuShipOrder_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptOrder
PromptUp = True

'Show Prompt
ShowUnitPrompt "ship"

End Sub

Private Sub mnuShipRetreat_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptSail
    
PromptForm.strCmd = "retreat"
PromptForm.Label2 = "Retreat"
PromptForm.Label1 = "path"
PromptForm.Label3.Visible = True
PromptForm.Combo1.Visible = True
LoadRetreatBox PromptForm.Combo1
PromptForm.Combo1.ListIndex = 0

'Show Prompt
ShowUnitPrompt "ship"

End Sub

Private Sub mnuShipSail_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptSail
    
PromptForm.strCmd = "sail"
PromptForm.Label2 = "Sail"
    
'Show Prompt
ShowUnitPrompt "ship"

End Sub

Private Sub mnuShipSonar_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptLook
    
PromptForm.strCmd = "sonar"
PromptForm.Label2 = "Sonar"

'Show Prompt
ShowUnitPrompt "ship"

End Sub

Private Sub mnuShipStat_Click()
 'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmPromptLand.Label2 = "Ship Stats"
frmPromptLand.cmd = "sstat"
mnuUnitLand_Click
End Sub

Private Sub mnuShipTorp_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

'Setup dialog and execute
Set PromptForm = frmPromptTorp

'Show Prompt
ShowUnitPrompt "ship"

End Sub

Private Sub mnuShipUnload_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptLoad

PromptForm.strCmd = "unload"
'PromptForm.strSector = SectorString(MxPos, MyPos) 092103 rjk: Switched to SelX and SelY for top level menu
PromptForm.strSector = SectorString(SelX, SelY)

'Show Prompt
ShowUnitPrompt "ship"
End Sub

Private Sub mnuShipUnsail_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

Set PromptForm = frmPromptLook
    
PromptForm.strCmd = "unsail"
PromptForm.Label2 = "Unsail"

'Show Prompt
ShowUnitPrompt "ship"

End Sub

Private Sub mnuUnitBuild_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

'Setup Designation dialog and execute
'092203 rjk: Switched to SelX, SelY for top level menu
'frmPromptBuild.txtOrigin = SectorString(MxPos, MyPos)
'rsSectors.Seek "=", MxPos, MyPos
frmPromptBuild.txtOrigin = SectorString(SelX, SelY)
rsSectors.Seek "=", SelX, SelY

'Put form in proper place
frmPromptBuild.Left = frmDrawMap.Left + (frmDrawMap.Width - frmPromptBuild.Width) \ 2
frmPromptBuild.top = frmDrawMap.top + frmDrawMap.Height - frmPromptBuild.Height
Set PromptForm = frmPromptBuild
PromptUp = True

'Dialog box will take it from here
frmPromptBuild.Show vbModeless

End Sub

Private Sub mnuUnitLand_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

'Setup Designation dialog and execute
frmPromptLand.txtOrigin = vbNullString

'Put form in proper place
frmPromptLand.Left = frmDrawMap.Left + (frmDrawMap.Width - frmPromptLand.Width) \ 2
frmPromptLand.top = frmDrawMap.top + frmDrawMap.Height - frmPromptLand.Height
Set PromptForm = frmPromptLand
PromptUp = True

'Dialog box will take it from here
frmPromptLand.Show vbModeless

End Sub

Private Sub mnuUnitPlanes_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmPromptLand.Label2 = "Plane"
frmPromptLand.cmd = "plane"
mnuUnitLand_Click
End Sub

Private Sub mnuUnitShip_Click()
'check for current prompt
If PromptUp Then
    Exit Sub
End If

frmPromptLand.Label2 = "Ship"
frmPromptLand.cmd = "ship"
mnuUnitLand_Click
End Sub


Private Sub Msg_Timer_Timer()

Dim Msg As String
Dim Done As Boolean
Dim n As Integer
Dim i As Integer
Dim bwidth As Integer
Static lstlen As Integer

'If you have a pending message, then add it to the display box
If MsgQ.Count = 0 Then
    frmDrawMap.Msg_Timer.Interval = 0
    Exit Sub
End If

Msg = MsgQ(1)

'add msg to the message box
'break into multiple lines if too long
n = 1
Done = False
bwidth = rtbTelegram.Width * 0.9

While Not Done
    If Me.TextWidth(Msg) > bwidth Then
        i = 1
        Do
            n = i
            i = InStr(n + 1, Msg, " ")
        Loop Until (i = 0 Or Me.TextWidth(Left$(Msg, i)) > bwidth)
    End If
    
    rtbTelegram.SelStart = 999999
    rtbTelegram.SelLength = 0
    If n = 1 Then
        If Not (lstlen = 0 And Len(Msg) = 0) Then
            rtbTelegram.SelText = vbCrLf + Msg
            lstlen = Len(Msg)
        End If
        Done = True
    Else
        rtbTelegram.SelText = vbCrLf + Left$(Msg, n)
        Msg = Mid$(Msg, n + 1)
        lstlen = 99
        n = 1
    End If
Wend

MsgQ.Remove 1

End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = vbDefault
End Sub
Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = vbDefault
End Sub
Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = vbDefault
End Sub
Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = vbDefault
End Sub

Private Sub PlayersTimer_Timer()
Static playersMinutes As Integer
If playersMinutes <= 0 Then
    frmEmpCmd.SubmitEmpireCommand "db1", False
    frmEmpCmd.SubmitEmpireCommand "players", False
    frmEmpCmd.SubmitEmpireCommand "db2", False
    playersMinutes = frmOptions.playersTimeInterval - 1
Else
    playersMinutes = playersMinutes - 1
End If
End Sub

Private Sub rtbTelegram_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lStartPos As Long
Dim lBegPos As Long
Dim lEndPos As Long
Dim strAddress As String
Dim strUnit As String

If Button = vbRightButton Then
    lStartPos = rtbTelegram.SelStart
    rtbTelegram.UpTo " (:@" + vbLf, False, False
    lBegPos = rtbTelegram.SelStart + 1
    rtbTelegram.UpTo " ):!" + vbLf, True, False
    lEndPos = rtbTelegram.SelStart + 1
    rtbTelegram.SelStart = lStartPos
    strAddress = Mid$(rtbTelegram.Text, lBegPos, lEndPos - lBegPos)
    If InStr(strAddress, ",") > 0 Then
        If ParseSectors(iTelegramSelX, iTelegramSelY, strAddress) Then
            mnuTelegramCenter.Visible = True
        Else
            mnuTelegramCenter.Visible = False
        End If
    Else
        mnuTelegramCenter.Visible = False
    End If
    PopupMenu mnuTelegram
End If
End Sub

Private Sub rtbTelegram_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = vbDefault
End Sub

Private Sub sb1_PanelClick(ByVal Panel As MSComCtlLib.Panel)
    Select Case Panel.key
        Case "Timer"
            GetUpdate True
            If VersionCheck(4, 3, 10) <> VER_LESS Then
                frmEmpCmd.SubmitEmpireCommand "xdump game *", False
            End If

        Case "Mail", "Mail Count"
            SubmitTelegramRead True, False
            sb1.Panels.Item("Mail").Text = "..."
            sb1.Panels.Item("Mail Count").Text = "..."
        Case "Anno"
            frmEmpCmd.SubmitEmpireCommand "wire y", True
        'Case "Repeat" 102503 rjk: No Repeat Panel exists
        '    txtEntry.Text = frmEmpCmd.GetPreviousCommand
    End Select
End Sub

Private Sub TabStrip1_Click()
'rjk 05/13/03: Moved the tab selection to DisplaySectorPanel
'              Moved the frame code to SelectTabContent
'    If TabStrip1.SelectedItem.Index = mintCurFrame _
'        Then Exit Sub ' No need to change frame.
'
'   Select Case SelSectType
'        Case SEL_SECT_OWN
'            ' Otherwise, hide old frame, show new.
'            Frame1(TabStrip1.SelectedItem.Index).Visible = True
'            Frame1(mintCurFrame).Visible = False
'            mintCurFrame = TabStrip1.SelectedItem.Index
'        Case SEL_SECT_ENEMY
'            If TabStrip1.SelectedItem.Index = 3 Then
'                Frame1(3).Visible = True
'                Frame1(mintCurFrame).Visible = False
'                mintCurFrame = 3
'            Else
'                Frame1(mintCurFrame).Visible = False
'                Frame1(0).Visible = True
'                mintCurFrame = 0
'            End If
'        Case SEL_SECT_UNKNOWN
'            If TabStrip1.SelectedItem.Index = 3 Then
'                Frame1(3).Visible = True
'                Frame1(mintCurFrame).Visible = False
'                mintCurFrame = 3
'            Else
'                Frame1(mintCurFrame).Visible = False
'                mintCurFrame = 0
'            End If
'    End Select
SelectTabContent
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'If F2 is pressed, repeat last command
If KeyCode = vbKeyF2 Then
    frmEmpCmd.SubmitLastCommand
    Exit Sub
End If

'If F5 is pressed, Zoom or UnZoom
If KeyCode = vbKeyF5 Then
    If Shift > 0 Then
        'mnuMapUnzoom_Click
    Else
        'munMapZoom_Click
    End If
    Form_Resize
    Exit Sub
End If

'If F9 is pressed, move back thru command list
'102503 rjk: Added shift F9 moves forward thru the command list
If KeyCode = vbKeyF9 Then
    If Not ((Shift And vbShiftMask) = vbShiftMask) Then
        txtEntry.Text = frmEmpCmd.GetPreviousCommand
    Else
        txtEntry.Text = frmEmpCmd.GetNextCommand
    End If
    txtEntry.SelStart = 999
    CommandCursorPos = 999
    SetCommandFocus
    Exit Sub
End If

Dim n As Integer
If KeyCode = vbKeyPageUp Then
    n = lstCmdResult.TopIndex - CInt(lstCmdResult.Height / Me.txtEntry.Height)
    If n > 0 Then
        lstCmdResult.TopIndex = n
    Else
        lstCmdResult.TopIndex = 0
    End If
End If
If KeyCode = vbKeyPageDown Then
    n = lstCmdResult.TopIndex + CInt(lstCmdResult.Height / Me.txtEntry.Height)
    If n < lstCmdResult.ListCount Then
        lstCmdResult.TopIndex = n
    ElseIf lstCmdResult.ListCount > 0 Then
        lstCmdResult.TopIndex = lstCmdResult.ListCount - 1
    End If
End If

'check and see if the control key is down
If Not ((Shift And vbCtrlMask) = 2) Then
    Exit Sub
End If
On Error Resume Next

If KeyCode = vbKey1 Then
    frmEmpCmd.WindowState = vbNormal
    frmEmpCmd.SetFocus
End If
If KeyCode = vbKey2 Then
    frmTelegram.SetFocus
End If
If KeyCode = vbKey3 Then
    frmTelegram.WindowState = vbNormal
    frmTelegram.SetFocus
End If

'Shift Fonts
If KeyCode = vbKey4 Then
    On Error GoTo Font_exit
    Load frmCommonDialog
    ' Set CancelError is True
    frmCommonDialog.cdb.CancelError = True
    ' Set flags
    frmCommonDialog.cdb.Flags = cdlCFScreenFonts
    frmCommonDialog.cdb.FontName = lstCmdResult.FontName
    frmCommonDialog.cdb.FontSize = lstCmdResult.FontSize
    
    ' Display the font dialog box
    frmCommonDialog.cdb.ShowFont
    lstCmdResult.FontName = frmCommonDialog.cdb.FontName
    lstCmdResult.FontSize = frmCommonDialog.cdb.FontSize
    Form_Resize
End If
Font_exit:


End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)

' Dim n As Integer   removed efj 8/2003
' When Enter is pressed, process everything
If KeyAscii = 13 Then
    If frmEmpCmd.ServerQuery Then
        frmEmpCmd.ExecuteEmpireCmd txtEntry.Text
        lstCmdResult.List(lstCmdResult.ListCount - 1) = lstCmdResult.List(lstCmdResult.ListCount - 1) + txtEntry.Text
    Else
        'now submit the command
        frmEmpCmd.SubmitFromCommandLine txtEntry.Text
    End If
    txtEntry.Text = vbNullString
    KeyAscii = 0
End If

'if an escape key is hit when a query is up
If KeyAscii = 27 And frmEmpCmd.ServerQuery Then
    'frmEmpCmd.ExecuteEmpireCmd "ctld"
    frmEmpCmd.ExecuteEmpireCmd "aborted"
    lstCmdResult.List(lstCmdResult.ListCount - 1) = lstCmdResult.List(lstCmdResult.ListCount - 1) + " aborted"
End If

'Any normal key - route to command box
If KeyAscii >= 8 Or (KeyAscii >= 32 And KeyAscii <= 125) Then
    SetCommandFocus KeyAscii
End If

End Sub

Private Sub TabStrip1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = vbDefault
End Sub

Private Sub tbMain_ButtonClick(ByVal Button As MSComCtlLib.Button)

Select Case Button.key
    Case "SetColors"
        mnuViewOptions_Click 1  'Colors off the menu screen
    Case "AutoUpdate"
        ToggleAutoUpdate '120203 rjk: Switch to frmOptions and boolean options.
    Case "AutoRead"
        ToggleAutoRead '120203 rjk: Switch to frmOptions and boolean options.
    Case "UpdateSector"
        mnuRefresh_Click 1
    Case "UpdateUnits"
        mnuRefresh_Click 3
    Case "Budget"
        mnuReportBudget_Click
    Case "Coastwatch"
        mnuReportCoastwatch_Click
    Case "Nation"
        mnuReportNation_Click
    Case "Newspaper"
        frmEmpCmd.SubmitEmpireCommand "newspaper 1"
    Case "PowerNew"
        mnuReportPowerNew_Click
    Case "Report"
        mnuReportReport_Click
    Case "Starvation"
        mnuReportStarve_Click
    Case "Version"
        mnuReportVersion_Click
    Case "Wire"
        frmEmpCmd.SubmitEmpireCommand "wire 1"
    Case "Declare"
        mnuDiploDeclare_Click
    Case "Relations"
        mnuDiploRelations_Click
    Case "Script"
        mnuToolsOption_Click 1
    Case "Build"
        mnuToolsOption_Click 2
    Case "Famine"
        mnuToolsOption_Click 3
    Case "NationLevels"
        mnuToolsOption_Click 4
    Case "Chi"
        mnuToolsOption_Click 5
    Case "ReportShow"
        mnuReportShow_Click
    Case "UpdateBmap"
        If frmOptions.bSPAtlantis Then
            frmEmpCmd.ForceUpdates = True
            If frmOptions.bSuppressBmapRefresh Then
                MsgBox ("Bmap updates are suppressed. Change option and then request again")
                Exit Sub
            End If
            'Send Code to start database update
            frmEmpCmd.SubmitEmpireCommand "db1", False
            'Request an update of map information
            frmEmpCmd.SubmitEmpireCommand "map *", False
            frmEmpCmd.SubmitEmpireCommand "bmap *", False
            'send Code to end database update
            frmEmpCmd.SubmitEmpireCommand "db2", False
        Else
            mnuRefresh_Click 2
        End If
    Case Else
        If InStr(Button.key, "Script") = 1 Then
            ExecuteFile CInt(Mid$(Button.key, 7))
        End If
End Select
End Sub

Private Sub txtEntry_GotFocus()
On Error Resume Next
CommandCursorFocus = True
txtEntry.SelStart = CommandCursorPos
End Sub

Private Sub txtEntry_LostFocus()
On Error Resume Next
CommandCursorPos = txtEntry.SelStart
CommandCursorFocus = False
End Sub

Private Sub txtEntry_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = vbIbeam
End Sub

Private Sub UpdateTimer_Timer()

Dim totsec As Long
Dim totmin As Long
Dim sec As Long
Dim Min As Long
Dim hour As Long
Dim strMinus As String

If TimerAtUpdate = 0 And SecondsToUpdate = 0 Then
    frmDrawMap.sb1.Panels.Item("Timer").Text = "No Update"
Else
    totsec = Timer - TimerAtUpdate
    While totsec < 0
        totsec = totsec + 86400
    Wend
    totsec = SecondsToUpdate - totsec
    If totsec <= -10 Then
        If totsec Mod 10 = 0 Then
            If frmEmpCmd.bAutoUpdate And frmEmpCmd.bUpdateEnabled Then
                If VersionCheck(4, 3, 10) = VER_LESS Then
                    UpdateCommands
                Else
                    frmEmpCmd.SubmitEmpireCommand "xdump game *", False
                End If
            Else
                GetUpdate False
                'Ensure it is still disabled
                If VersionCheck(4, 3, 10) <> VER_LESS Then
                    frmEmpCmd.SubmitEmpireCommand "xdump game *", False
                End If
            End If
        End If
    End If
    sec = Abs(totsec Mod 60)
    totmin = (totsec - sec) \ 60
    Min = Abs(totmin Mod 60)
    hour = Abs(totmin \ 60)
    If hour > 99 Then
        hour = 99
    End If
    If totsec < 0 Then
        strMinus = "-"
    Else
        strMinus = ""
    End If
    frmDrawMap.sb1.Panels.Item("Timer").Text = strMinus + Format$(hour, "00") + ":" _
            + Format$(Min, "00") + ":" + Format$(sec, "00")
End If

'if its time to run a pre-update script, run it now
If BatchScript1Time > 0 And totsec <= BatchScript1Time Then
    frmScript.ExectuteBatchScript BatchScript1File
    BatchScript1Time = 0
    BatchScript1File = vbNullString
End If

'if you are idle for 10 minutes or more, submit a command to keep the line open
TimeIdle = TimeIdle + 1
If TimeIdle >= 600 Then
    GetUpdate False
    TimeIdle = 0
End If

'flash the lights if server query is up
Static ToggleLights As Boolean
If frmEmpCmd.ServerQuery Then
    sb1.Panels.Item(3).ToolTipText = "Empire Server Waiting for Response"
    If ToggleLights Then
        Set frmDrawMap.sb1.Panels.Item("Lights").Picture = picBothLights
        Set frmEmpCmd.imgLights.Picture = picBothLights
    Else
        Set frmDrawMap.sb1.Panels.Item("Lights").Picture = picGreenLight
        Set frmEmpCmd.imgLights.Picture = picGreenLight
    End If
    ToggleLights = Not ToggleLights
End If

'If a prompt help box needs positioned - do it
If PositionHelp Then
    If PromptUp Then
        If frmReport.Visible Then
            If frmReport.WindowState = vbNormal Then
                
                'Size Report
                frmReport.top = Me.top
                If frmReport.Height > (PromptForm.top - Me.top) Then
                    frmReport.Height = PromptForm.top - Me.top
                End If
                
                'move report and shut off timer
                PositionHelp = False
            End If
        End If
    Else
        PositionHelp = False
    End If
End If
End Sub

Private Sub vsMap_Change()

Static ChangeSkip As Boolean

'Prevents second pass if we have to manually reset below due to
'world wrap
If ChangeSkip Then
    ChangeSkip = False
    Exit Sub
End If

'Reset Origin
OriginY = Int(vsMap.Value / 2) * 2
If OriginY < MapSizeY / -2 Then
    OriginY = OriginY + MapSizeY
End If
If OriginY > (MapSizeY / 2) - 1 Then
    OriginY = OriginY - MapSizeY
End If

'Update index value if necessary
If vsMap.Value <> OriginY Then
    ChangeSkip = True
    vsMap.Value = OriginY
End If

'Don't do a draw if the map's not valid
If Not MapValid Then
    Exit Sub
End If

DrawHexes

End Sub
Private Sub hsMap_Change()

Static ChangeSkip As Boolean

'Prevents second pass if we have to manually reset below due to
'world wrap
If ChangeSkip Then
    ChangeSkip = False
    Exit Sub
End If

'Reset Origin
OriginX = Int(hsMap.Value / 2) * 2
If OriginX < MapSizeX / -2 Then
    OriginX = OriginX + MapSizeX
End If
If OriginX > (MapSizeX / 2) - 1 Then
    OriginX = OriginX - MapSizeX
End If

'Update index value if necessary
If hsMap.Value <> OriginX Then
    ChangeSkip = True
    hsMap.Value = OriginX
End If

'Don't do a draw if the map's not valid
If Not MapValid Then
    Exit Sub
End If

'Redraw hex Centers
DrawHexes

End Sub


Public Sub DrawHexes()
' Dim Y As Integer   removed efj 8/2003
Dim X As Integer
' Dim nbrcol As Integer   removed efj 8/2003
' Dim nbrrow As Integer   removed efj 8/2003
Dim oldcolor
Dim oldfontsize
Dim oldWidth
Dim PosX As Single
Dim PosY As Single
Dim n As Integer
Dim strPath As String
Dim startindex As Integer

'Don't do a draw if the map's not valid
If Not MapValid Then
    Exit Sub
End If

picMap.AutoRedraw = True
picMap.Cls

DrawGrid picMap

If ViewShipWake Then
    DrawShipWake -1
ElseIf LastShip >= 0 Then
    DrawShipWake LastShip
End If

'handle event markers
EventMarkers.RemoveExpired
If EventMarkers.Count > 0 Then
    DrawEvent picMap
End If

'If you were drawing path hexes - redraw them
If DrawingPath Or MarkingTerritory Or DisplayingPath Then
    'Draw Hex
    oldWidth = picMap.DrawWidth
    picMap.DrawWidth = 3   ' Set DrawWidth
    oldcolor = picMap.ForeColor
    picMap.ForeColor = vbRed
    
    startindex = LBound(WorkingPath, 2)
    If MarkingTerritory Then
        startindex = startindex + 1
    End If
    
    For X = startindex To UBound(WorkingPath, 2)
        Coord WorkingPath(1, X), WorkingPath(2, X), PosX, PosY
        picMap.ForeColor = vbRed
        DrawHex picMap, PosX, PosY
        
        'draw center lines if appropriate
        If X > LBound(WorkingPath, 2) And DrawingPath Then
            picMap.ForeColor = vbWhite
            'Compute Direction
            strPath = EmpirePathDirection(WorkingPath(1, X - 1) - WorkingPath(1, X), WorkingPath(2, X - 1) - WorkingPath(2, X))
            n = InStr(HEX_DIR, strPath)
            DrawHexCenterLine picMap, PosX, PosY, GetHexSideLength, (InStr(HEX_DIR, strPath))
            n = n - 3
            If n < 1 Then
                n = n + 6
            End If
            Coord WorkingPath(1, X - 1), WorkingPath(2, X - 1), PosX, PosY
            DrawHexCenterLine picMap, PosX, PosY, GetHexSideLength, n
        End If
    Next X
   
    picMap.ForeColor = oldcolor
    picMap.DrawWidth = oldWidth   ' Set DrawWidth
End If

'if nav markers are up, redraw them
If Len(NavMarkerShip) > 0 Then
    oldfontsize = picMap.Font.Size
    NavMarkers NavMarkerShip, bNavMarkerShipIncludeUpdateMob, bNavMarkerShipIncludeUpdateEff, _
        iNavMarkerShipUpdates, bNavMarkerShipLimitMobility
    picMap.Font.Size = oldfontsize
End If

If Len(NukeDamageType) > 0 Then
    oldfontsize = picMap.Font.Size
    ShowNukeDamage picMap
    picMap.Font.Size = oldfontsize
ElseIf Len(PlaneRangeType) > 0 Then
    oldfontsize = picMap.Font.Size
    ShowPlaneRange picMap
    picMap.Font.Size = oldfontsize
End If

'hsMap.Max = (MapSizeX / 2) - 1 - ((NbrHexWide - 1) * 2)
'vsMap.Max = (MapSizeY / 2) - 1 - (NbrHexTall - 1)
hsMap.Max = (MapSizeX / 2)
vsMap.Max = (MapSizeY / 2)
hsMap.LargeChange = (NbrHexWide - 1)
vsMap.LargeChange = (NbrHexTall - 1) / 2

End Sub

Public Sub StartPath(strStartSect As String)
Dim x1 As Integer
Dim y1 As Integer
If Not ParseSectors(x1, y1, strStartSect) Then
    Exit Sub
End If

ReDim WorkingPath(0 To 2, 0)
WorkingPath(1, 0) = x1
WorkingPath(2, 0) = y1
DrawingPath = True

'Draw Starting Hex

Dim oldcolor
Dim oldWidth
oldWidth = picMap.DrawWidth
picMap.DrawWidth = 3   ' Set DrawWidth
oldcolor = picMap.ForeColor
picMap.ForeColor = vbRed

Dim PosX As Single
Dim PosY As Single

'Get Center position and draw hex
Coord x1, y1, PosX, PosY
DrawHex picMap, PosX, PosY

picMap.ForeColor = oldcolor
picMap.DrawWidth = oldWidth   ' Set DrawWidth
End Sub

'111803 rjk: Created one copy of the DisplayPath code
Public Sub DisplayPath(strStart As String, strPath As String)
Dim i As Integer
Dim ch As String

StartPath strStart
For i = 1 To Len(strPath)
    ch = Mid$(strPath, i, 1)
    ExtendPath ch
Next i
DrawingPath = False
DisplayingPath = True
End Sub

Public Sub AddtoPath()

Dim lasthex As Integer
Dim DeltaX As Integer
Dim DeltaY As Integer
' Dim Index As Integer   removed efj 8/2003
Dim strPath As String
Dim n As Integer

On Error GoTo 0

lasthex = UBound(WorkingPath, 2)

'Make sure its an adjacent hex
DeltaX = MxPos - WorkingPath(1, lasthex)
DeltaY = MyPos - WorkingPath(2, lasthex)
'111303 rjk: Added World wraparound checks
If DeltaX > rsVersion.Fields("World X") / 2 Then
    DeltaX = DeltaX - rsVersion.Fields("World X")
ElseIf DeltaX < -rsVersion.Fields("World X") / 2 Then
    DeltaX = DeltaX + rsVersion.Fields("World X")
End If
If DeltaY > rsVersion.Fields("World Y") / 2 Then
    DeltaY = DeltaY - rsVersion.Fields("World Y")
ElseIf DeltaY < -rsVersion.Fields("World Y") / 2 Then
    DeltaY = DeltaY + rsVersion.Fields("World Y")
End If
If Not ((Abs(DeltaX) = 2 And DeltaY = 0) Or (Abs(DeltaX) = 1 And Abs(DeltaY) = 1)) Then
    Exit Sub
End If

'Compute Direction
strPath = EmpirePathDirection(DeltaX, DeltaY)

'Add path to prompt text box
If PromptForm.Name = "frmPromptPath" Then
    PromptForm.txtPath.Text = PromptForm.txtPath.Text + strPath
    PromptForm.lblEndSector.Caption = SectorString(MxPos, MyPos)
End If

ReDim Preserve WorkingPath(0 To 2, lasthex + 1)
WorkingPath(1, lasthex + 1) = MxPos
WorkingPath(2, lasthex + 1) = MyPos

'Draw Hex
Dim oldcolor
Dim oldWidth
oldWidth = picMap.DrawWidth
picMap.DrawWidth = 3   ' Set DrawWidth
oldcolor = picMap.ForeColor
picMap.ForeColor = vbRed

Dim PosX As Single
Dim PosY As Single

'Get Center position and draw hex
Coord MxPos, MyPos, PosX, PosY
DrawHex picMap, PosX, PosY

picMap.ForeColor = vbWhite

n = InStr(HEX_DIR, strPath)
n = n - 3
If n < 1 Then
    n = n + 6
End If

DrawHexCenterLine picMap, PosX, PosY, GetHexSideLength, n

Coord WorkingPath(1, lasthex), WorkingPath(2, lasthex), PosX, PosY
DrawHexCenterLine picMap, PosX, PosY, GetHexSideLength, (InStr(HEX_DIR, strPath))

picMap.ForeColor = oldcolor
picMap.DrawWidth = oldWidth   ' Set DrawWidth

End Sub

Public Function FinishPath() As String
DrawingPath = False
DrawHexes
End Function

Public Sub ExtendPath(ByVal strDir As String)

Dim lasthex As Integer
Dim x1 As Integer
Dim y1 As Integer
Dim X2 As Integer
Dim Y2 As Integer

'if we are drawing a path
If Not DrawingPath Then
    Exit Sub
End If

'get the last hex
lasthex = UBound(WorkingPath, 2)
x1 = WorkingPath(1, lasthex)
y1 = WorkingPath(2, lasthex)

'Comput the new hex
ComputeDestination x1, y1, X2, Y2, strDir

'Set the new hex
MxPos = X2
MyPos = Y2

'Now add to path
AddtoPath

End Sub

Public Sub RemoveFromPath()
Dim lasthex As Integer
' Dim DeltaX As Integer   removed efj 8/2003
' Dim DeltaY As Integer   removed efj 8/2003
' Dim Index As Integer   removed efj 8/2003
' Dim strPath As String   removed efj 8/2003
On Error GoTo 0

'Get the last hex
lasthex = UBound(WorkingPath, 2)
If lasthex <= LBound(WorkingPath, 2) Then
    Exit Sub
End If

Dim PosX As Single
Dim PosY As Single

'Get Center position and draw hex
Coord WorkingPath(1, lasthex), WorkingPath(2, lasthex), PosX, PosY
DrawSector picMap, PosX, PosY, WorkingPath(1, lasthex), WorkingPath(2, lasthex)

'Remove path from prompt text box
If Len(PromptForm.txtPath.Text) > 0 Then
    PromptForm.txtPath.Text = Left$(PromptForm.txtPath.Text, Len(PromptForm.txtPath.Text) - 1)
End If
PromptForm.lblEndSector.Caption = CStr(WorkingPath(1, lasthex)) + "," + CStr(WorkingPath(2, lasthex))

'remove hex from working path
ReDim Preserve WorkingPath(0 To 2, lasthex - 1)

End Sub

Public Sub StartTerritory()
MarkingTerritory = True
ReDim WorkingPath(0 To 2, 0)
DrawHexes '100103 rjk: Clear Markers
End Sub

Public Sub ToggleTerritory(sx As Integer, sy As Integer)
Dim lasthex As Integer
Dim found As Boolean
Dim n As Integer

On Error Resume Next

'if not drawing territory - 'quit
If Not MarkingTerritory Then
    Exit Sub
End If

'Get the last hex
lasthex = UBound(WorkingPath, 2)

'search for the x,y coords
found = False
For n = LBound(WorkingPath, 2) + 1 To lasthex
    If WorkingPath(1, n) = sx And _
       WorkingPath(2, n) = sy Then
        found = True
        WorkingPath(1, n) = WorkingPath(1, lasthex)
        WorkingPath(2, n) = WorkingPath(2, lasthex)
    End If
Next n

Dim PosX As Single
Dim PosY As Single
Dim oldcolor
Dim oldWidth
oldWidth = picMap.DrawWidth
picMap.DrawWidth = 3   ' Set DrawWidth
oldcolor = picMap.ForeColor

If Not found Then
    'add hex
    ReDim Preserve WorkingPath(0 To 2, lasthex + 1)
    WorkingPath(1, lasthex + 1) = sx
    WorkingPath(2, lasthex + 1) = sy
    picMap.ForeColor = vbRed
Else
    'remove hex from working path
    ReDim Preserve WorkingPath(0 To 2, lasthex - 1)
    picMap.ForeColor = vbBlack
End If

'Get Center position and draw hex
Coord sx, sy, PosX, PosY
DrawHex picMap, PosX, PosY

'Redraw all the hexs in case we messed up a border
picMap.ForeColor = vbRed
For n = LBound(WorkingPath, 2) + 1 To UBound(WorkingPath, 2)
    Coord WorkingPath(1, n), WorkingPath(2, n), PosX, PosY
    DrawHex picMap, PosX, PosY
Next n

picMap.ForeColor = oldcolor
picMap.DrawWidth = oldWidth   ' Set DrawWidth

End Sub

'Public Sub ShowCommandHistory()   removed efj 8/2003
'
''Shows the previous commands
'Load frmRepeatCmd
'frmRepeatCmd.Move Me.Left + Me.Width - frmRepeatCmd.Width, _
'                        Me.top + Me.Height - frmRepeatCmd.Height
'frmRepeatCmd.Show vbModal
'
'End Sub

Public Sub ShowUnitPrompt(Optional UnitType As Variant)

Dim FillBox As Boolean
On Error Resume Next

'Put form in proper place
PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - PromptForm.Width) \ 2
PromptForm.top = frmDrawMap.top + frmDrawMap.Height - PromptForm.Height
PromptUp = True

'if there is only one unit, put it in the text box.  Otherwise,
'show the unit dialog box
' Dim oneunit As Boolean   removed efj 8/2003
If (UnitType = "ship" And InStr(strShipid, "/") = 0 And Len(strShipid) > 0) Or _
    (UnitType = "plane" And InStr(strPlaneid, "/") = 0 And Len(strPlaneid) > 0) Or _
    (UnitType = "land" And InStr(strLandid, "/") = 0 And Len(strLandid) > 0) Or _
    (UnitType = "nuke" And InStr(strNukeid, "/") = 0 And Len(strNukeid) > 0) Or _
    GridSelection = True Then
    Select Case UnitType
    Case "ship"
        PromptForm.txtUnit.Text = strShipid
    Case "land"
        PromptForm.txtUnit.Text = strLandid
    Case "plane"
        PromptForm.txtUnit.Text = strPlaneid
    Case "nuke"
        PromptForm.txtUnit.Text = strNukeid
    End Select
Else
    'Bring up Unit dialog box
    DisplayUnitBox
    
    'If specific unit requested then make sure that unit is showing in the unit box
    FillBox = False
    If Not IsMissing(UnitType) Then
        Select Case UnitType
            Case "ship"
                If DisplaySelect <> udSHIP Then
                    DisplaySelect = udSHIP
                    FillBox = True
                End If
            
            Case "land"
                If DisplaySelect <> udLAND Then
                    DisplaySelect = udLAND
                    FillBox = True
                End If
            
            Case "plane"
                If DisplaySelect <> udPLANE Then
                    DisplaySelect = udPLANE
                    FillBox = True
                End If
            Case "nuke"
                If DisplaySelect <> udNUKE Then
                    DisplaySelect = udNUKE
                    FillBox = True
                End If
        End Select
        
        'refill the box if necessary
        If FillBox Then
            DisplaySubset = ssSECT
            ' 092303 rjk: Removed secx and secy and switched to SelX and SelY
            'secx = MxPos
            'secy = MyPos
            'FillSubSetFlag = True 092503 rjk: Eliminate
            FillGrid
        End If
    End If
End If


'Dialog box will take it from here
PromptForm.Show vbModeless
On Error Resume Next
PromptForm.txtUnit.SetFocus
End Sub

Private Sub NavMarkers(ByVal Shipnum As String, bIncludeUpdateMob As Boolean, bIncludeUpdateEff As Boolean, iUpdates As Integer, bLimitMobility As Boolean)
Const MaxRadius = 20
Dim mcost(0 To MaxRadius * 2 + 1, 0 To MaxRadius * 2 + 1) As Single
Dim path(0 To MaxRadius * 2 + 1, 0 To MaxRadius * 2 + 1) As String
Dim colCheckList As New Collection
Dim StartX As Integer
Dim StartY As Integer
Dim iDistance As Integer
Dim iDirection As Integer
Dim iSpoke As Integer
Dim iAdjacent As Integer
Dim iAdjX As Integer
Dim iAdjY As Integer
Dim sx As Integer
Dim sy As Integer
Dim X As Integer
Dim Y As Integer
Dim str As String
Dim PosX As Single
Dim PosY As Single
Dim movecost As Single
Dim hldcost As Single
Dim hldIndex As String
Dim shipid As Integer
Dim strSearchSector As String
Dim xa, ya, xas, yas, pa As Variant
Dim iEfficiency As Integer

xa = Array(-2, -1, 1, 2, 1, -1)
ya = Array(0, -1, -1, 0, 1, 1)
xas = Array(1, 2, 1, -1, -2, -1)
yas = Array(-1, 0, 1, 1, 0, -1)
pa = Array("g", "y", "u", "j", "n", "b")

'Save ship string to aid in redraw
NavMarkerShip = Shipnum

hldIndex = rsShip.Index
rsShip.Index = "PrimaryKey"

movecost = 0
While Len(Shipnum) > 0
    X = InStr(Shipnum, "/")
    If X > 0 Then
        shipid = CInt(Left$(Shipnum, X - 1))
        Shipnum = Mid$(Shipnum, X + 1)
    Else
        shipid = CInt(Shipnum)
        Shipnum = vbNullString
    End If

    rsShip.Seek "=", shipid
    If rsShip.NoMatch Then
        rsShip.Index = hldIndex
        Exit Sub
    End If
    iEfficiency = rsShip.Fields("eff")
    If iEfficiency <> 100 And bIncludeUpdateEff Then
        iEfficiency = iEfficiency + iUpdates * rsVersion("Eff gain - Ship")
        If iEfficiency > 100 Then
            iEfficiency = 100
        End If
    End If
    hldcost = ShipMoveCost(rsShip.Fields("spd"), iEfficiency, rsShip.Fields("tech"), _
        rsShip.Fields("type"))
    If hldcost > movecost Then
        movecost = hldcost
        mcost(MaxRadius, MaxRadius) = rsShip.Fields("mob")
        If bIncludeUpdateMob Then
            mcost(MaxRadius, MaxRadius) = mcost(MaxRadius, MaxRadius) + iUpdates * rsVersion("Mob gain - Ship")
            If mcost(MaxRadius, MaxRadius) > rsVersion("Max mob - Ship") And _
               bLimitMobility Then
                mcost(MaxRadius, MaxRadius) = rsVersion("Max mob - Ship")
            End If
        End If
        StartX = rsShip.Fields("x")
        StartY = rsShip.Fields("y")
    End If
Wend

colCheckList.Add SectorString(0, 0)
path(MaxRadius, MaxRadius) = ""

While colCheckList.Count > 0
    strSearchSector = colCheckList.Item(1)
    ParseSectors X, Y, strSearchSector
    If Len(path(Int(X / 2) + MaxRadius, Y + MaxRadius)) < MaxRadius Then
        For iAdjacent = 0 To 5
            iAdjX = X + xa(iAdjacent)
            iAdjY = Y + ya(iAdjacent)
            If mcost(Int(iAdjX / 2) + MaxRadius, iAdjY + MaxRadius) = 0 Then
                path(Int(iAdjX / 2) + MaxRadius, iAdjY + MaxRadius) = path(Int(X / 2) + MaxRadius, Y + MaxRadius) + pa(iAdjacent)
                mcost(Int(iAdjX / 2) + MaxRadius, iAdjY + MaxRadius) = mcost(Int(X / 2) + MaxRadius, Y + MaxRadius) - (SectorMobCost(StartX + iAdjX, StartY + iAdjY, MT_SHIP) * movecost)
                If mcost(Int(iAdjX / 2) + MaxRadius, iAdjY + MaxRadius) >= 0 Then
                    colCheckList.Add SectorString(iAdjX, iAdjY)
                End If
            Else
                If mcost(Int(iAdjX / 2) + MaxRadius, iAdjY + MaxRadius) < mcost(Int(X / 2) + MaxRadius, Y + MaxRadius) - (SectorMobCost(StartX + iAdjX, StartY + iAdjY, MT_SHIP) * movecost) Then
                    path(Int(iAdjX / 2) + MaxRadius, iAdjY + MaxRadius) = path(Int(X / 2) + MaxRadius, Y + MaxRadius) + pa(iAdjacent)
                    mcost(Int(iAdjX / 2) + MaxRadius, iAdjY + MaxRadius) = mcost(Int(X / 2) + MaxRadius, Y + MaxRadius) - (SectorMobCost(StartX + iAdjX, StartY + iAdjY, MT_SHIP) * movecost)
                    If mcost(Int(iAdjX / 2) + MaxRadius, iAdjY + MaxRadius) >= 0 Then
                        colCheckList.Add SectorString(iAdjX, iAdjY)
                    End If
                End If
            End If
        Next iAdjacent
    End If
    colCheckList.Remove 1
Wend

'Now put out the nav markers
Dim HldColor As Long
Dim HldFontSize As Variant
HldColor = picMap.ForeColor
picMap.ForeColor = vbWhite
HldFontSize = picMap.Font.Size
picMap.Font.Size = picMap.Font.Size * 0.7

For iDistance = 1 To MaxRadius - 1
    For iDirection = 0 To 5
        For iSpoke = 0 To MaxRadius - 2
            X = (iDistance * xa(iDirection)) + (iSpoke * xas(iDirection))
            Y = (iDistance * ya(iDirection)) + (iSpoke * yas(iDirection))
            Coord StartX + X, StartY + Y, PosX, PosY
            If mcost(Int(X / 2) + MaxRadius, Y + MaxRadius) > 0 Then
                str = CInt(mcost(Int(X / 2) + MaxRadius, Y + MaxRadius))
                sx = PosX - (picMap.TextHeight(str) / 2)
                sy = PosY - (picMap.TextWidth(str) / 2)
                If sx >= 0 And sy >= 0 Then
                    picMap.CurrentX = sx
                    picMap.CurrentY = sy
                    picMap.Print str
                End If
            End If
        Next iSpoke
    Next iDirection
Next iDistance

picMap.ForeColor = HldColor
picMap.Font.Size = HldFontSize
rsShip.Index = hldIndex
End Sub

Private Sub DrawShipWake(Shipnum As Integer)
'ship number less than zero means draw all
Dim found As Boolean
Dim shippath As Variant
found = False
With rsEnemyUnit
    If .BOF And .EOF Then
        Exit Sub
    End If
    .MoveFirst
    While Not .EOF And Not found
        If .Fields("class") = "s" Then
            If .Fields("id") = Shipnum Then
                found = True
            End If
               
            If found Or Shipnum < 0 Then
            
                'can't draw a wake if the ship does not have one.
                If Len(.Fields("wake")) > 0 Then
                
                    shippath = ParseWaypoints(.Fields("wake"))
                    If Not IsArray(shippath) Then
                        Exit Sub
                    End If
                    ReDim Preserve shippath(UBound(shippath) + 1)
                    shippath(UBound(shippath)) = SectorString(.Fields("x"), .Fields("y"))
                    DrawWake .Fields("id"), shippath
                End If
            End If
        End If
                        
        .MoveNext
    Wend
    
End With
End Sub

Private Sub DrawWake(ByVal Shipnum As Integer, shippath As Variant)

'find the enemy ship unit
On Error GoTo DrawWake_Error


Dim PosX As Single
Dim PosY As Single
Dim n As Integer
Dim X As Integer
Dim i As Integer
Dim strPath As String

Dim sx As Integer
Dim sy As Integer
Dim sx2 As Integer
Dim sy2 As Integer
Dim dx As Integer
Dim dy As Integer

Dim strTemp As String
    
'Draw Hex
Dim oldcolor
Dim oldWidth
oldWidth = picMap.DrawWidth
oldcolor = picMap.ForeColor
picMap.DrawWidth = 3   'Set DrawWidth
picMap.ForeColor = BasicText(2)

For X = LBound(shippath) To UBound(shippath)
    strTemp = CStr(shippath(X))
    If ParseSectors(sx, sy, strTemp) Then
        'draw center lines if appropriate
        If X > LBound(shippath) Then
            'Compute Direction
            strTemp = CStr(shippath(X - 1))
            If ParseSectors(sx2, sy2, strTemp) Then
                dx = sx - sx2
                dy = sy - sy2
                'check for world wrap
                If dx > MapSizeX / 2 Then dx = dx - MapSizeX
                If dx < -1 * MapSizeX / 2 Then dx = dx + MapSizeX
                If dy > MapSizeY / 2 Then dy = dy - MapSizeY
                If dy < -1 * MapSizeY / 2 Then dy = dy + MapSizeY '112703 rjk: Fixed
                strPath = EmpirePathDirection(dx, dy)
                
                For i = 1 To Len(strPath)
                    ComputeDestination sx2, sy2, sx, sy, Mid$(strPath, i, 1)
                    n = InStr(HEX_DIR, Mid$(strPath, i, 1))
                    Coord sx2, sy2, PosX, PosY
                    DrawHexCenterLine picMap, PosX, PosY, GetHexSideLength, n
                    n = n - 3
                    If n < 1 Then
                        n = n + 6
                    End If
                    If Not (X = UBound(shippath) And i = Len(strPath)) Then
                        Coord sx, sy, PosX, PosY
                        DrawHexCenterLine picMap, PosX, PosY, GetHexSideLength, n
                    End If
                    sx2 = sx: sy2 = sy
                Next i
            End If
        End If
    End If
Next X

DrawWake_Error:
picMap.ForeColor = oldcolor
picMap.DrawWidth = oldWidth   ' Set DrawWidth

End Sub


' Calculate index for use with TextArray property.
Function gridIndex(row As Integer, col As Integer) As Long
gridIndex = row * grid1.Cols + col
End Function

Public Sub FillGrid()
Dim rs As Recordset
Dim fieldcnt As Integer
Dim X As Integer
Dim iIndex As Integer
Dim substype As String
Dim substypekey As String
Dim BuildType As String '101403 rjk: Moved to be private variable of FillGrid
Dim hldIndex As String

Dim PlaneRange As Integer
Static GridUpdateinProgress As Boolean

'Airport vars
Dim Airports() As String
Dim i As Long
Dim strLast As String
Dim strLoc As String
Dim airport As Long
Dim found As Boolean
Dim iCalculatedFields As Integer

ReDim Airports(0)
'Prevent recursion
If GridUpdateinProgress Then
    Exit Sub
End If
GridUpdateinProgress = True

On Error Resume Next

'Reset sort variables
SortCol = -1
SortDecending = True
PlaneRange = 0

FillTypeBox

'Determain which type of unit we are displaying
Select Case DisplaySelect
    Case udSHIP
        Set rs = rsShipList
        rs.Requery
        substype = "Fleet "
        substypekey = "flt"
        BuildType = "s"
        If VersionCheck(4, 3, 0) <> VER_LESS Then
            ComputeUnitCountsForShips
        End If
    Case udLAND
        Set rs = rsLandList
        rs.Requery
        substype = "Army "
        substypekey = "army"
        BuildType = "l"
        If VersionCheck(4, 3, 0) <> VER_LESS Then
            ComputeUnitCountsForLandUnits
        End If
    Case udPLANE
        Set rs = rsPlaneList
        rs.Requery
        substype = "Wing "
        substypekey = "wing"
        BuildType = "p"
    Case udNUKE
        Set rs = rsNuke
        substype = "Type "
        substypekey = "type"
        BuildType = "n"
    Case udENEMY
        Set rs = rsEnemyUnit
        'rs.Requery 100603 rjk: Removed, does not work this table.  Need to investigate more later.
        substype = "Nation "
        substypekey = "owner"
    
    Case udList
        FillUnitList
        GridUpdateinProgress = False
        Exit Sub
    Case Else
        GridUpdateinProgress = False
        Exit Sub
End Select
If bDeity And DisplaySelect <> udENEMY Then
    substype = "Nation "
    substypekey = "country"
End If

'rs.Index = "Primary Key" 100603 rjk: Removed, does not work these tables.  Need to investigate more later.

'092503 rjk: Rearrange the code to eliminate FillSubSetFlag
ClearSubCombo
'if this is a "range" display, make that the only entry in sub box
If DisplaySubset = ssPLANE_RANGE Or DisplaySubset = ssMISSILE_RANGE _
   Or DisplaySubset = ssATTACK_RANGE Then
    cmbSub.AddItem "In Range"
    cmbSub.ListIndex = 0
Else
    cmbSub.AddItem "All"
End If

fieldcnt = rs.Fields.Count
rtbUnitList.Visible = False
grid1.Visible = False
grid1.Clear
grid1.Rows = 2
grid1.Cols = fieldcnt
grid1.FixedCols = 2

'The following is to trap the size of the column width for column zero
'so that if manually changed it is not always reset.
Static ColZeroWidth1 As Single
Static ColZeroWidth2 As Single
Static ColZeroWidth3 As Single
Static LastDisplayFlag As enumUnitDisplay
Static bOldShortNameUnitGrid As Boolean

If LastDisplayFlag <> udUNKNOWN Then
    If Not (frmOptions.bShortNameUnitGrid Xor bOldShortNameUnitGrid) Then
        If frmOptions.bShortNameUnitGrid Then
            ColZeroWidth3 = grid1.ColWidth(0)
        ElseIf LastDisplayFlag = udPLANE Or LastDisplayFlag = udENEMY Then
            ColZeroWidth1 = grid1.ColWidth(0)
        Else
            ColZeroWidth2 = grid1.ColWidth(0)
        End If
    Else
        bOldShortNameUnitGrid = frmOptions.bShortNameUnitGrid
    End If
Else
    grid1.ColWidth(0) = -1      'return to default
    grid1.ColWidth(-1) = grid1.ColWidth(0) / 2.5
    ColZeroWidth1 = grid1.ColWidth(0) * 3.9
    ColZeroWidth2 = grid1.ColWidth(0) * 2.8
    ColZeroWidth3 = grid1.ColWidth(0)
    If bDeity Then
        ColZeroWidth1 = ColZeroWidth1 + 400
        ColZeroWidth2 = ColZeroWidth2 + 400
        ColZeroWidth3 = ColZeroWidth3 + 400
    End If
    bOldShortNameUnitGrid = Not frmOptions.bShortNameUnitGrid
End If
LastDisplayFlag = DisplaySelect
    
grid1.ColWidth(0) = -1      'return to default
grid1.ColWidth(-1) = grid1.ColWidth(0) / 2.5

If frmOptions.bShortNameUnitGrid Then
    grid1.ColWidth(0) = ColZeroWidth3
ElseIf DisplaySelect = udPLANE Or DisplaySelect = udENEMY Then
    grid1.ColWidth(0) = ColZeroWidth1
Else
    grid1.ColWidth(0) = ColZeroWidth2
End If

If DisplaySelect = udENEMY Then
    grid1.ColWidth(8) = grid1.ColWidth(0) * 1.25
End If

grid1.ColData(0) = 0
grid1.ColData(1) = 1
grid1.row = 0
'Fill in row headers
grid1.col = 1
grid1.Text = "id"
grid1.CellAlignment = flexAlignLeftCenter
If DisplaySelect = udPLANE Then
    iCalculatedFields = iCalculatedFields + 1
    grid1.ColWidth(2) = grid1.ColWidth(1)
    grid1.col = 2
    grid1.Text = "mis"
    grid1.CellAlignment = flexAlignLeftCenter
    grid1.ColData(2) = 0 'Text
End If
If DisplaySelect = udNUKE Then
    X = 1
    iCalculatedFields = 1
Else
    X = 2
End If
Do While X < fieldcnt
    If DisplaySelect <> udNUKE Or X <> 4 Then
        grid1.col = X + iCalculatedFields
        grid1.Text = rs.Fields(X).Name
        grid1.CellAlignment = flexAlignLeftCenter
        If rs.Fields(X).Type = dbText Then
            grid1.ColData(X + iCalculatedFields) = 0 'Text
        Else
            'rjk 8/11/03 Change the range column type to Airport Code so the grid can be sorted
            If (DisplaySubset = ssPLANE_RANGE Or DisplaySubset = ssMISSILE_RANGE) _
                    And grid1.Text = "range" Then
                grid1.ColData(X + iCalculatedFields) = 2  'Airport Code
            Else
                grid1.ColData(X + iCalculatedFields) = 1  'Number
            End If
        End If
        If DisplaySelect = udSHIP And rs.Fields(X).Name = "rng" Then
            For iIndex = 1 To 4
                iCalculatedFields = iCalculatedFields + 1
                grid1.Cols = fieldcnt + iCalculatedFields
                grid1.ColWidth(grid1.Cols - 1) = grid1.ColWidth(grid1.Cols - 2)
                grid1.col = X + iCalculatedFields
                Select Case (iIndex)
                Case 1:
                    grid1.Text = "RFR"
                Case 2:
                    grid1.Text = "SMC"
                Case 3:
                    grid1.Text = "RPT"
                Case 4:
                    grid1.Text = "RMM"
                End Select
                grid1.CellAlignment = flexAlignLeftCenter
                grid1.ColData(X + iCalculatedFields) = 1 'Number
            Next iIndex
        End If
        If DisplaySelect = udLAND And rs.Fields(X).Name = "frg" Then
            iCalculatedFields = iCalculatedFields + 1
            grid1.Cols = fieldcnt + iCalculatedFields
            grid1.ColWidth(grid1.Cols - 1) = grid1.ColWidth(grid1.Cols - 2)
            grid1.col = X + iCalculatedFields
            grid1.Text = "RFR"
            grid1.CellAlignment = flexAlignLeftCenter
            grid1.ColData(X + iCalculatedFields) = 1 'Number
        End If
        grid1.CellFontBold = False
    Else
        iCalculatedFields = iCalculatedFields - 1
    End If
    X = X + 1
Loop

'Fill in rows
If Not (rs.BOF And rs.EOF) Then
    rs.MoveFirst
End If
While Not rs.EOF

    'Add to the sub combo box
    '092503 rjk: Eliminated FillSubSetFlag
    'If FillSubSetFlag Then
    If DisplaySubset <> ssPLANE_RANGE And DisplaySubset <> ssMISSILE_RANGE _
       And DisplaySubset <> ssATTACK_RANGE Then
        AddSubBox substype + CStr(rs.Fields(substypekey)), rs.Fields("x"), rs.Fields("y") '100803 rjk: Added CStr for nation number
    End If
    
    If DisplaySelect = udPLANE Then
        PlaneRange = rs.Fields("range")
    End If
    
    'Check to see if we are displaying this fleet
        ' 092303 rjk: Removed secx and secy and switched to SelX and SelY
        '(DisplaySubset = DSP_SECT And secx = rs.Fields("x") And secy = rs.Fields("y")) Or _
        '(DisplaySubset = DSP_ATTACK_RANGE And ((SectorDistance(secx, secy, rs.Fields("x"), rs.Fields("y"))) = 1)) Or _
        '(DisplaySubset = DSP_PLANE_RANGE And ((SectorDistance(secx, secy, rs.Fields("x"), rs.Fields("y")) * 2) <= planerange)) Or _
        '(DisplaySubset = DSP_MISSILE_RANGE And ((SectorDistance(secx, secy, rs.Fields("x"), rs.Fields("y"))) <= planerange)) Then
        '092503 rjk: Switched Fleet to have whole string, allows separation ship fleets from plane wings (FillUnitList)
        '(DisplaySubset = ssFLEET And Fleet = Trim$(rs.Fields(substypekey))) Or
    If (DisplaySubset = ssALL) Or _
       (DisplaySubset = ssFLEET And Fleet = substype + Trim$(rs.Fields(substypekey))) Or _
       (DisplaySubset = ssSECT And SelX = rs.Fields("x") And SelY = rs.Fields("y")) Or _
       (DisplaySubset = ssATTACK_RANGE And ((SectorDistance(SelX, SelY, rs.Fields("x"), rs.Fields("y"))) = 1)) Or _
       (DisplaySubset = ssPLANE_RANGE And ((SectorDistance(SelX, SelY, rs.Fields("x"), rs.Fields("y")) * 2) <= PlaneRange)) Or _
       (DisplaySubset = ssMISSILE_RANGE And ((SectorDistance(SelX, SelY, rs.Fields("x"), rs.Fields("y"))) <= PlaneRange)) Then

        'check to see if it is in the selected type subset
        If SubType = TYPE_ALL Or InStr(strType(SubType), " " + rs.Fields(1) + " ") Or _
            ((DisplaySelect = udLAND Or DisplaySelect = udPLANE) And SubType = TYPE_SHIP_LOADED And rs.Fields("ship") <> -1) Or _
            ((DisplaySelect = udLAND Or DisplaySelect = udPLANE) And SubType = TYPE_SHIP_UNLOADED And rs.Fields("ship") = -1) Or _
            ((DisplaySelect = udLAND Or DisplaySelect = udPLANE) And SubType = TYPE_LAND_LOADED And rs.Fields("land") <> -1) Or _
            ((DisplaySelect = udLAND Or DisplaySelect = udPLANE) And SubType = TYPE_LAND_UNLOADED And rs.Fields("land") = -1) Then
            'check the mob and eff requirement
            If DisplaySelect = udENEMY Or _
            (((Not NeedMob) Or rs.Fields("mob") > 0) And _
            ((Not Needeff) Or rs.Fields("eff") > 59 Or _
            (DisplaySelect = udPLANE And rs.Fields("eff") > 39))) Then
        
            'check if we are only displaying warships
            '092403 rjk: Switched to UnitCharacteristicCheck to remove hard coding
            '(ShowOnlyWarships And (rs.Fields("fir") > 0 Or rs.Fields("type") = "ls")) Then
            If DisplaySelect <> udSHIP Or _
             Not frmOptions.bShowOnlyWarships Or _
             (frmOptions.bShowOnlyWarships And (rs.Fields("fir") > 0 Or _
              UnitCharacteristicCheck(TYPE_S_LAND, rsShip.Fields("type")))) Then
      
                grid1.row = grid1.row + 1
                grid1.col = 0
                grid1.CellForeColor = vbBlack
                
                If DisplaySelect = udENEMY Then
                    BuildType = rs.Fields(9)    'Check the class
                Else
                    'check for mission - use red if found
                    rsMissions.Seek "=", BuildType, rs.Fields(0)
                    If Not rsMissions.NoMatch Then
                        grid1.CellForeColor = vbRed
                    End If
                    
                    If DisplaySelect = udSHIP Then
                        rsShipOrders.Seek "=", rs.Fields(0)
                        If Not rsShipOrders.NoMatch Then
                            If grid1.CellForeColor = vbBlack Then
                                grid1.CellForeColor = vbBlue
                            Else
                                grid1.CellForeColor = vbMagenta
                            End If
                       End If
                    End If
                End If
                
                'Get the unit description
                If DisplaySelect = udNUKE Then
                    grid1.Text = rs.Fields("type")
                Else
                    grid1.Text = rs.Fields(1) ' Set in case lookup fails
                    If Not frmOptions.bShortNameUnitGrid Then
                        rsBuild.Seek "=", BuildType, rs.Fields(1)
                        If Not rsBuild.NoMatch Then
                            If Len(rsBuild.Fields("desc")) > 0 Then
                                grid1.Text = rsBuild.Fields("desc")
                            End If
                        End If
                    End If
                End If
                
                grid1.CellAlignment = flexAlignLeftCenter
                If bDeity Then
                    grid1.Text = Left$(Nations.NationName(rs.Fields("country")), 4) + " " + grid1.Text
                End If
                grid1.col = 1
                grid1.CellForeColor = vbBlack
                Select Case DisplaySelect
                Case udSHIP
                    hldIndex = rsShip.Index
                    rsShip.Index = "PrimaryKey"
                    rsShip.Seek "=", rs.Fields("id")
                    If Not rsShip.NoMatch Then
                        If rsShip.Fields("off") Then
                            grid1.CellForeColor = vbRed
                        End If
                    End If
                    rsShip.Index = hldIndex
                Case udLAND
                    hldIndex = rsLand.Index
                    rsLand.Index = "PrimaryKey"
                    rsLand.Seek "=", rs.Fields("id")
                    If Not rsLand.NoMatch Then
                        If rsLand.Fields("off") Then
                            grid1.CellForeColor = vbRed
                        End If
                    End If
                    rsLand.Index = hldIndex
                Case udPLANE
                    hldIndex = rsPlane.Index
                    rsPlane.Index = "PrimaryKey"
                    rsPlane.Seek "=", rs.Fields("id")
                    If Not rsPlane.NoMatch Then
                        If rsPlane.Fields("off") Then
                            grid1.CellForeColor = vbRed
                        End If
                    End If
                    rsPlane.Index = hldIndex
                Case udNUKE
                    If rs.Fields("off") Then
                        grid1.CellForeColor = vbRed
                    End If
                End Select
                grid1.Text = CStr(rs.Fields(0))
                grid1.CellAlignment = flexAlignLeftCenter
                
                iCalculatedFields = 0
                If DisplaySelect = udPLANE Then
                    iCalculatedFields = iCalculatedFields + 1
                    grid1.col = 2
                    grid1.CellForeColor = vbBlack
                    grid1.Text = ""
                    If Not rsMissions.NoMatch Then
                        If Len(rsMissions.Fields("mission")) > 0 Then
                            grid1.Text = StrConv(Left$(rsMissions.Fields("mission"), 1), vbUpperCase) + _
                                "-" + CStr(rsMissions.Fields("radius"))
                        End If
                    End If
                End If
                If DisplaySelect = udNUKE Then
                    X = 1
                    iCalculatedFields = 1
                Else
                    X = 2
                End If
                Do While X < fieldcnt
                    If DisplaySelect <> udNUKE Or X <> 4 Then
                        grid1.col = X + iCalculatedFields
                        grid1.CellForeColor = vbBlack
    
                        '101803 rjk: Added check for x,y fields as -1 is valid for these fields.
                        If (DisplaySelect = udENEMY) And (grid1.Text = "-1") And _
                            (rs.Fields(X).Name <> "x" And rs.Fields(X).Name <> "y") Then
                            grid1.Text = "???"
                        ElseIf rs.Fields(X).Name = "timestamp" Then
                            grid1.Text = Format$(ConvertToLocalTimeFromWinACEUTC(rs.Fields("timestamp").Value), "ddd mmm dd hh:mm:ss yyyy")
                        ElseIf (DisplaySubset = ssPLANE_RANGE Or DisplaySubset = ssMISSILE_RANGE) And _
                                rs.Fields(X).Name = "range" Then
                            grid1.CellForeColor = vbRed
                            On Error GoTo 0
                            strLoc = SectorString(rs.Fields("x"), rs.Fields("y"))
                            If strLoc <> strLast Then
                                i = 0
                                found = False
                                strLast = strLoc
                                While i < UBound(Airports) And Not found
                                    i = i + 1
                                    If Airports(i) = strLoc Then
                                        found = True
                                    End If
                                Wend
                                'then add if necessary
                                If Not found Then
                                    i = UBound(Airports) + 1
                                    ReDim Preserve Airports(i)
                                    Airports(i) = strLoc
                                End If
                                airport = 96 + i
                            End If
                            'Save last airport, as planes usually come out in order
                            strLast = strLoc
                            ' 092303 rjk: Removed secx and secy and switched to SelX and SelY
                            'grid1.Text = Chr$(airport) + "-" + CStr(SectorDistance(secx, secy, rs.Fields("x"), rs.Fields("y")))
                            grid1.Text = Chr$(airport) + "-" + CStr(SectorDistance(SelX, SelY, rs.Fields("x"), rs.Fields("y")))
                            On Error Resume Next
                        Else
                            grid1.Text = rs.Fields(X).Value
                        End If
                        grid1.CellAlignment = flexAlignRightCenter
                        If (DisplaySelect = udSHIP And rs.Fields(X).Name = "rng") Then
                            Dim movecost As Single
                                   
                            movecost = ShipMoveCost(rs.Fields("spd"), rs.Fields("eff"), _
                                CInt(rs.Fields("tech")), rs.Fields("type"))
                            For iIndex = 1 To 4
                                iCalculatedFields = iCalculatedFields + 1
                                grid1.col = X + iCalculatedFields
                                grid1.CellForeColor = vbBlack
                                Select Case iIndex
                                Case 1:
                                    grid1.Text = Format$(UnitFireRange(rs.Fields("tech"), rs.Fields("rng")), "#0.00")
                                Case 2:
                                    grid1.Text = Format$(movecost, "#0.0")
                                Case 3:
                                    grid1.Text = Format$((rsVersion("Mob gain - Ship") / movecost), "#0.0")
                                Case 4:
                                    grid1.Text = Format$((rsVersion("Max mob - Ship") / movecost), "#0.0")
                                End Select
                            Next iIndex
                        End If
                        If (DisplaySelect = udLAND And rs.Fields(X).Name = "frg") Then
                            iCalculatedFields = iCalculatedFields + 1
                            grid1.col = X + iCalculatedFields
                            grid1.CellForeColor = vbBlack
                            grid1.Text = Format$(UnitFireRange(rs.Fields("tech"), rs.Fields("frg")), "#0.00")
                        End If
                    Else
                        iCalculatedFields = iCalculatedFields - 1
                    End If
                    X = X + 1
                Loop
                grid1.Rows = grid1.Rows + 1
            End If
        End If
        End If
    End If
    rs.MoveNext
Wend

'Set the count
If grid1.Rows > 1 Then
    grid1.row = grid1.row + 1
    grid1.col = 0
    grid1.Text = "Total: " + CStr(grid1.Rows - 2) + " units"
    grid1.CellAlignment = flexAlignLeftCenter
    grid1.CellFontBold = True
    grid1.CellFontName = "Arial"
End If
    
'Reshow the grid
grid1.Visible = True

'Set the option buttons to the correct entry
Toolbar1.Buttons(DisplaySelect).Value = tbrPressed

'Set the cmbSelect if box was refilled
'092503 rjk: Moved to SetIndexSubCombo
SetIndexSubCombo

'FillSubSetFlag = False  092503 rjk: Eliminated
GridUpdateinProgress = False
End Sub

Private Sub ClearSubCombo()
cmbSub.Clear
strSubSect = vbNullString
strSubFleet = vbNullString
End Sub

Private Sub SetIndexSubCombo()
Dim tstx As Long
Dim tsty As Long
Dim X As Integer

For X = 0 To cmbSub.ListCount
    Select Case DisplaySubset
    Case ssALL '112503 rjk: Moved All inside, as All is not always at the top
        If cmbSub.List(X) = "All" Then
            cmbSub.ListIndex = X
            Exit Sub
        End If
    Case ssFLEET
        If cmbSub.List(X) = Fleet Then
            cmbSub.ListIndex = X
            Exit Sub
        End If
    Case ssSECT
        If InStr(cmbSub.List(X), ",") > 0 Then
            tstx = cmbSub.ItemData(X) / 10000
            tsty = cmbSub.ItemData(X) - tstx * 10000
            If tstx >= 1000 Then
                tstx = 0 - tstx + 1000
            End If
            If tsty >= 1000 Then
                tsty = 0 - tsty + 1000
            End If
            If tstx = SelX And tsty = SelY Then
                cmbSub.ListIndex = X
                Exit Sub
            End If
        End If
    End Select
Next
End Sub

Private Sub cmbSub_Click()

Dim strSelect As String
Dim bUpdate As Boolean
Dim n As Integer
' Dim n2 As Integer   removed efj 8/2003
Dim NewSecX As Integer
Dim NewSecY As Integer

'Boolean keeps track if an update of grid is needed
bUpdate = False

'First check if all was selected
strSelect = cmbSub.List(cmbSub.ListIndex)
If strSelect = "All" Then
    If DisplaySubset <> ssALL Then
        bUpdate = True
        DisplaySubset = ssALL
    End If

'Then See if a range is up - this never changes
ElseIf Left$(strSelect, 3) = "In " Then
    Exit Sub
    
'Then See if a sector was selected
ElseIf InStr(strSelect, ",") > 0 Then
    n = InStr(strSelect, ",")
    NewSecX = CInt(Trim$(Left$(strSelect, n - 1)))
    NewSecY = CInt(Trim$(Mid$(strSelect, n + 1)))
    
    'Make sure selection criteria has changed
    If DisplaySubset <> ssSECT Then
        bUpdate = True
        DisplaySubset = ssSECT
        ' 092303 rjk: Removed secx and secy and switched to SelX and SelY
        FillSectorBox NewSecX, NewSecY
'        secx = NewSecX
'        secy = NewSecY
    Else
        ' 092303 rjk: Removed secx and secy and switched to SelX and SelY
'        If secx <> NewSecX Or secy <> NewSecY Then
        If SelX <> NewSecX Or SelY <> NewSecY Then
            bUpdate = True
            ' 092303 rjk: Removed secx and secy and switched to SelX and SelY
            FillSectorBox NewSecX, NewSecY
'            secx = NewSecX
 '           secy = NewSecY
        End If
    End If
Else

' Finally, must be a fleet/army/wing
    '092503 rjk: Switched Fleet to have whole name, used in FillUnitList to separate ship fleets from plane wings
    If DisplaySubset <> ssFLEET Then
        bUpdate = True
        DisplaySubset = ssFLEET
        Fleet = strSelect
    Else
        If Fleet <> strSelect Then
            Fleet = strSelect
            bUpdate = True
        End If
    End If
End If

If bUpdate Then
    FillGrid
End If

End Sub

'092503 rjk: Remove stype as not needed any more
Private Sub AddSubBox(Fleet As String, X As Integer, Y As Integer)
Dim strSec As String
Dim n As Long

'If fleet has not been added
'092503 rjk: Switched Fleet to have whole name, used for FillUnitList to separate ship fleets from plane wings
If bDeity Or InStr(Fleet, "Nation") = 1 Then
    If InStr(strSubFleet, Mid$(Fleet, 7)) = 0 Then
        strSubFleet = strSubFleet + Mid$(Fleet, 7) + " "
        cmbSub.AddItem Fleet
    End If
ElseIf InStr(Fleet, "Type") = 1 Then
    If InStr(strSubFleet, Fleet) = 0 Then
        strSubFleet = strSubFleet + Fleet + " "
        cmbSub.AddItem Fleet
    End If
ElseIf InStr(strSubFleet, Right$(Fleet, 1)) = 0 Then
    strSubFleet = strSubFleet + Right$(Fleet, 1)
    cmbSub.AddItem Fleet
End If

strSec = " " + SectorString(X, Y) + " "

If InStr(strSubSect, strSec) = 0 Then
    strSubSect = strSubSect + strSec
    'Encode sector numbers in item data
    If X < 0 Then
        n = 0 - X + 1000
    Else
        n = X
    End If
    n = n * 10000
    If Y < 0 Then
        n = n - Y + 1000
    Else
        n = n + Y
    End If
    
    cmbSub.AddItem Trim$(strSec)
    cmbSub.ItemData(cmbSub.NewIndex) = n
End If
End Sub

Private Sub DeleteUnits()
Dim StartX As Integer
Dim Finishx As Integer
Dim x1 As Integer
Dim gi As Integer
Dim id As Integer
Dim strOwner As String
Dim strClass As String
Dim hldIndex As String
Dim rs As Recordset

StartX = grid1.row
Finishx = grid1.RowSel

If Finishx < StartX Then
    x1 = StartX
    StartX = Finishx
    Finishx = x1
End If
If Finishx > grid1.Rows - 2 Then
    Finishx = grid1.Rows - 2
End If

'set recordset
Select Case DisplaySelect
Case udSHIP
    Set rs = rsShip
Case udLAND
    Set rs = rsLand
Case udPLANE
    Set rs = rsPlane
Case udNUKE
    Set rs = rsNuke
Case udENEMY
    Set rs = rsEnemyUnit
Case Else
    Exit Sub
End Select

'set index
hldIndex = rs.Index
rs.Index = "PrimaryKey"

'go thru selected rows, delete units
For x1 = StartX To Finishx
    gi = gridIndex(x1, 1)
    id = Val(grid1.TextArray(gi))
    If DisplaySelect = udENEMY Then
        gi = gridIndex(x1, 4)
        strOwner = grid1.TextArray(gi)
        gi = gridIndex(x1, 9)
        strClass = grid1.TextArray(gi)
        If strOwner = "???" Then '091603 rjk: CInt("???") gives run-time, it is trying to select '-1' row
            rs.Seek "=", -1, strClass, id
        Else
            rs.Seek "=", CInt(strOwner), strClass, id
        End If
    Else
        rs.Seek "=", id
    End If
    If Not rs.NoMatch Then
        rs.Delete
    End If
Next x1
rs.Index = hldIndex
Set rs = Nothing

'redraw the map
frmDrawMap.DrawHexes
FillGrid

End Sub

'Private Sub SetUnitType(n As enumUnitType, Optional ReqMobility As Boolean, Optional ReqMinEff As Boolean)

'Static SaveMobility As Variant
'Static SaveEfficiency As Variant
'Static SaveSubtype As Variant
'Static SkipSave As Boolean

'Save values if this is an option box request
'If n = TYPE_RESTORE Then
'    n = SaveSubtype
'    Toolbar1.Buttons(6).Value = SaveMobility
'    Toolbar1.Buttons(7).Value = SaveEfficiency
'    NeedMob = (SaveMobility = tbrPressed)
'    Needeff = (SaveEfficiency = tbrPressed)
'    SkipSave = False
'ElseIf (ReqMobility Or ReqMinEff) And (Not SkipSave) Then
'    SaveMobility = Toolbar1.Buttons(6).Value
'    SaveEfficiency = Toolbar1.Buttons(7).Value
'    SaveSubtype = SubType
'    SkipSave = True
'End If

'SubType = n
'DisplayUnitBox 'rjk 05/13/03: Switched to DisplayUnitBox function from DisplaySectorPanel 3

'NeedMob = ReqMobility
'If ReqMobility Then
'    Toolbar1.Buttons(6).Value = tbrPressed
'Else
'    Toolbar1.Buttons(6).Value = tbrUnpressed
'End If

'Needeff = ReqMinEff
'If ReqMinEff = True Then
'    Toolbar1.Buttons(7).Value = tbrPressed
'Else
'    Toolbar1.Buttons(7).Value = tbrUnpressed
'End If

'FillGrid
'End Sub

Private Sub grid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim tb As TextBox
Dim StartX As Integer
Dim Finishx As Integer
Dim x1 As Integer
Dim n As Integer
Dim bCtl As Boolean
Dim found As Boolean
Dim strSel As String ' 092503 rjk: moved to here, only place used
Dim CurrUnit As Variant

On Error Resume Next

strTextBox = vbNullString

'If clicking in top row, resort current box
If grid1.MouseRow = 0 Then
    SortGrid grid1.MouseCol
    Exit Sub
End If

'this routine is used to fill prompt boxes
If Not PromptUp Then
    Exit Sub
End If

'See if this is a box you add units to
Set tb = PromptForm.ActiveControl
If Not (tb.Name = "txtUnit" Or tb.Name = "txtUnit2" _
        Or tb.Name = "txtEscort" Or tb.Name = "txtUnit3" Or _
        (tb.Name = "txtTarget" And DisplaySelect = udENEMY)) Then
    Exit Sub
End If

' bShift = False   removed efj 8/2003
If (Shift And vbShiftMask) = 1 Then
'    bShift = True   removed efj 8/2003
End If

bCtl = False
If (Shift And vbCtrlMask) = 2 Then
    bCtl = True
End If

'use a temp string to avoid firing text boxs "change"
'event muliple times
strTextBox = tb.Text

'If control key is down, clear existing units else
'Get current units
If Len(strTextBox) > 0 And Not bCtl Then
    CurrUnit = ParseUnitString(tb.Text)
Else
    ReDim CurrUnit(0)
    CurrUnit(0) = vbNullString
    strTextBox = vbNullString
End If

StartX = grid1.row
Finishx = grid1.RowSel

If Finishx < StartX Then
    x1 = StartX
    StartX = Finishx
    Finishx = x1
End If

strSel = vbNullString

'go thru selected rows, and add units to select String
For x1 = StartX To Finishx
    strSel = Trim$(grid1.TextMatrix(x1, 1))
    found = False
    For n = LBound(CurrUnit) To UBound(CurrUnit)
        If CurrUnit(n) = strSel Then
            found = True
        End If
    Next n
    If Not found Then
        If Len(strTextBox) > 0 Then
            strTextBox = strTextBox + "/"
        End If
        strTextBox = strTextBox + strSel
    End If
Next x1

tb.Text = strTextBox
frmDrawMap.PromptForm.SetFocus

End Sub

Private Sub SortGrid(col As Integer)

'This does a simple bubble sort of the grid in place.  Due to the small size
'of the sample, a more efficent sort is usually not necessary.

Dim n As Integer
Dim row As Integer
Dim var1 As Variant
Dim var2 As Variant
Dim sng1 As Single
Dim sng2 As Single

'If clicked multiple times, change the sort order
If SortCol = col Then
    SortDecending = Not SortDecending
End If
SortCol = col

grid1.Visible = False

n = grid1.Rows - 3
'rjk 8/11/03 Added sorting for Airport Code
Select Case grid1.ColData(col)
Case 0 'Text
    While n > 0
        For row = 1 To n
            var1 = grid1.TextArray(gridIndex(row, col))
            var2 = grid1.TextArray(gridIndex(row + 1, col))
            
            If SortDecending Then
                If var1 > var2 Then
                    grid1.RowPosition(row + 1) = row
                End If
            Else
                If var1 < var2 Then
                    grid1.RowPosition(row + 1) = row
                End If
            End If
        Next row
        n = n - 1
    Wend
Case 1 'Number
    While n > 0
        For row = 1 To n
            If Len(grid1.TextArray(gridIndex(row, col))) = 0 Then
                sng1 = 0
            Else
                sng1 = CSng(grid1.TextArray(gridIndex(row, col)))
            End If
            If Len(grid1.TextArray(gridIndex(row + 1, col))) = 0 Then
                sng2 = 0
            Else
                sng2 = CSng(grid1.TextArray(gridIndex(row + 1, col)))
            End If
            
            If SortDecending Then
                If sng1 > sng2 Then
                    grid1.RowPosition(row + 1) = row
                End If
            Else
                If sng1 < sng2 Then
                    grid1.RowPosition(row + 1) = row
                End If
            End If
        Next row
        n = n - 1
    Wend
Case 2 'Airport Code (letter 'dash' range (number))
    While n > 0
        'Sort by range first
        For row = 1 To n
            If Len(grid1.TextArray(gridIndex(row, col))) < 3 Then
                sng1 = 0
            Else
                sng1 = CSng(Mid(grid1.TextArray(gridIndex(row, col)), 3))
            End If
            If Len(grid1.TextArray(gridIndex(row + 1, col))) < 3 Then
                sng2 = 0
            Else
                sng2 = CSng(Mid(grid1.TextArray(gridIndex(row + 1, col)), 3))
            End If
            
            'If the range is equal then sort by the airport letter
            If sng1 = sng2 Then
                var1 = Mid(grid1.TextArray(gridIndex(row, col)), 1, 1)
                var2 = Mid(grid1.TextArray(gridIndex(row + 1, col)), 1, 1)
                
                If SortDecending Then
                    If var1 > var2 Then
                        grid1.RowPosition(row + 1) = row
                    End If
                Else
                    If var1 < var2 Then
                        grid1.RowPosition(row + 1) = row
                    End If
                End If
            Else
                If SortDecending Then
                    If sng1 > sng2 Then
                        grid1.RowPosition(row + 1) = row
                    End If
                Else
                    If sng1 < sng2 Then
                        grid1.RowPosition(row + 1) = row
                    End If
                End If
            End If
        Next row
        n = n - 1
    Wend
End Select

grid1.Visible = True

End Sub

'091903 rjk: Removed ChangeUnitDisplay combined with Toolbar1_ButtonClick
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
'   ChangeUnitDisplay Button.key, (Button.Value = tbrPressed)
'End Sub

'Public Sub ChangeUnitDisplay(ByVal key As String, ByVal v As Boolean)
' added by thomas lullier
' revised by daniel to easily manage unit selection

Dim n As enumUnitDisplay
Select Case Button.key
    Case "Ship"
        n = udSHIP
    Case "Land"
        n = udLAND
    Case "Plane"
        n = udPLANE
    Case "Nuke"
        n = udNUKE
    Case "Enemy"
        n = udENEMY
    Case "List"
        n = udList
    
    Case "Mob"
        NeedMob = (Button.Value = tbrPressed)
        FillGrid
        Exit Sub

    Case "Delete"
        DeleteUnits
        Exit Sub
    
    Case "Eff"
        Needeff = (Button.Value = tbrPressed)
        FillGrid
        Exit Sub
End Select

'FillTypeBox - already done in FillGrid, and should not need to be redone, it just
' reclicking button

'If option Changed - Refill the grid
If DisplaySelect <> n Then
    DisplaySelect = n
    SubType = TYPE_ALL
    
    'if showing fleet/wing/army, shift to all
    If DisplaySubset = ssFLEET Then
        DisplaySubset = ssALL
    End If

    'Refill
    'FillSubSetFlag = True 092503 rjk: Eliminated
    FillGrid
End If

End Sub

Private Sub FillTypeBox()
Static CurrentType As enumUnitDisplay
'Dim hldSubtype As enumUnitType
Dim Index As Integer

'loading the box can change the subtype, so we need
'to store it and reload it
'hldSubtype = SubType

If CurrentType <> DisplaySelect Then
    CurrentType = DisplaySelect
    'optSelect(0).Tag = CStr(TYPE_ALL)
    'optSelect(0).Caption = "All"
    'optSelect(1).ToolTipText = "Show All"
    
    'show all
    'For n = 1 To 7
    '    optSelect(n).Visible = True
    'Next n

    cmbUnitFilter.Clear
    cmbUnitFilter.AddItem strTypeTitle(TYPE_ALL)
    cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_ALL
    
    Select Case DisplaySelect
    Case udSHIP
        cmbUnitFilter.AddItem strTypeTitle(TYPE_SC_FIRE)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_SC_FIRE
        cmbUnitFilter.AddItem strTypeTitle(TYPE_S_FISH)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_S_FISH
        cmbUnitFilter.AddItem strTypeTitle(TYPE_S_TORP)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_S_TORP
        cmbUnitFilter.AddItem strTypeTitle(TYPE_S_DCHRG)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_S_DCHRG
        cmbUnitFilter.AddItem strTypeTitle(TYPE_S_PLANE)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_S_PLANE
        cmbUnitFilter.AddItem strTypeTitle(TYPE_S_MISS)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_S_MISS
        cmbUnitFilter.AddItem strTypeTitle(TYPE_S_OIL)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_S_OIL
        cmbUnitFilter.AddItem strTypeTitle(TYPE_S_OILER)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_S_OILER
        cmbUnitFilter.AddItem strTypeTitle(TYPE_S_SONAR)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_S_SONAR
        cmbUnitFilter.AddItem strTypeTitle(TYPE_S_MINE)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_S_MINE
        cmbUnitFilter.AddItem strTypeTitle(TYPE_S_SWEEP)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_S_SWEEP
        cmbUnitFilter.AddItem strTypeTitle(TYPE_S_SUB)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_S_SUB
        cmbUnitFilter.AddItem strTypeTitle(TYPE_S_SPY)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_S_SPY
        cmbUnitFilter.AddItem strTypeTitle(TYPE_S_LAND)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_S_LAND
        cmbUnitFilter.AddItem strTypeTitle(TYPE_S_SEMI_LAND)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_S_SEMI_LAND
        cmbUnitFilter.AddItem strTypeTitle(TYPE_S_SUB_TORP)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_S_SUB_TORP
        cmbUnitFilter.AddItem strTypeTitle(TYPE_S_TRADE)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_S_TRADE
        cmbUnitFilter.AddItem strTypeTitle(TYPE_S_ANTI_MISSILE)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_S_ANTI_MISSILE
        cmbUnitFilter.Visible = True
    Case udLAND
        cmbUnitFilter.AddItem strTypeTitle(TYPE_LC_FIRE)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_LC_FIRE
        cmbUnitFilter.AddItem strTypeTitle(TYPE_L_XLIGHT)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_L_XLIGHT
        cmbUnitFilter.AddItem strTypeTitle(TYPE_L_ENGINEER)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_L_ENGINEER
        cmbUnitFilter.AddItem strTypeTitle(TYPE_L_SUPPLY)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_L_SUPPLY
        cmbUnitFilter.AddItem strTypeTitle(TYPE_L_SECURITY)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_L_SECURITY
        cmbUnitFilter.AddItem strTypeTitle(TYPE_L_LIGHT)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_L_LIGHT
        cmbUnitFilter.AddItem strTypeTitle(TYPE_L_MARINE)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_L_MARINE
        cmbUnitFilter.AddItem strTypeTitle(TYPE_L_RECON)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_L_RECON
        cmbUnitFilter.AddItem strTypeTitle(TYPE_L_RADAR)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_L_RADAR
        cmbUnitFilter.AddItem strTypeTitle(TYPE_L_ASSAULT)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_L_ASSAULT
        cmbUnitFilter.AddItem strTypeTitle(TYPE_LAND_UNLOADED)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_LAND_UNLOADED
        cmbUnitFilter.AddItem strTypeTitle(TYPE_LAND_LOADED)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_LAND_LOADED
        cmbUnitFilter.AddItem strTypeTitle(TYPE_SHIP_UNLOADED)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_SHIP_UNLOADED
        cmbUnitFilter.AddItem strTypeTitle(TYPE_SHIP_LOADED)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_SHIP_LOADED
        cmbUnitFilter.Visible = True
    Case udPLANE
        cmbUnitFilter.AddItem strTypeTitle(TYPE_PC_BOMB)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_PC_BOMB
        cmbUnitFilter.AddItem strTypeTitle(TYPE_PG_ESCORTS)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_PG_ESCORTS
        cmbUnitFilter.AddItem strTypeTitle(TYPE_P_TACTICAL)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_P_TACTICAL
        cmbUnitFilter.AddItem strTypeTitle(TYPE_P_BOMBER)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_P_BOMBER
        cmbUnitFilter.AddItem strTypeTitle(TYPE_P_INTERCEPT)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_P_INTERCEPT
        cmbUnitFilter.AddItem strTypeTitle(TYPE_P_ESCORT)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_P_ESCORT
        cmbUnitFilter.AddItem strTypeTitle(TYPE_PC_DROP)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_PC_DROP
        cmbUnitFilter.AddItem strTypeTitle(TYPE_P_CARGO)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_P_CARGO
        cmbUnitFilter.AddItem strTypeTitle(TYPE_P_MINE)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_P_MINE
        cmbUnitFilter.AddItem strTypeTitle(TYPE_P_SWEEP)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_P_SWEEP
        cmbUnitFilter.AddItem strTypeTitle(TYPE_P_PARA)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_P_PARA
        cmbUnitFilter.AddItem strTypeTitle(TYPE_P_SPY)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_P_SPY
        cmbUnitFilter.AddItem strTypeTitle(TYPE_PC_LAUNCH)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_PC_LAUNCH
        cmbUnitFilter.AddItem strTypeTitle(TYPE_P_MISSILE)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_P_MISSILE
        cmbUnitFilter.AddItem strTypeTitle(TYPE_P_MARINE)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_P_MARINE
        cmbUnitFilter.AddItem strTypeTitle(TYPE_P_SDI)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_P_SDI
        cmbUnitFilter.AddItem strTypeTitle(TYPE_P_SATELLITE)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_P_SATELLITE
        cmbUnitFilter.AddItem strTypeTitle(TYPE_P_IMAGE)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_P_IMAGE
        cmbUnitFilter.AddItem strTypeTitle(TYPE_P_ANTISUB)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_P_ANTISUB
        cmbUnitFilter.AddItem strTypeTitle(TYPE_P_CHOPPER)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_P_CHOPPER
        cmbUnitFilter.AddItem strTypeTitle(TYPE_P_LIGHT)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_P_LIGHT
        cmbUnitFilter.AddItem strTypeTitle(TYPE_P_X_LIGHT)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_P_X_LIGHT
        cmbUnitFilter.AddItem strTypeTitle(TYPE_P_STEALTH)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_P_STEALTH
        cmbUnitFilter.AddItem strTypeTitle(TYPE_P_HALF_STEALTH)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_P_HALF_STEALTH
        cmbUnitFilter.AddItem strTypeTitle(TYPE_P_VTOL)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_P_VTOL
        cmbUnitFilter.AddItem strTypeTitle(TYPE_LAND_UNLOADED)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_LAND_UNLOADED
        cmbUnitFilter.AddItem strTypeTitle(TYPE_LAND_LOADED)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_LAND_LOADED
        cmbUnitFilter.AddItem strTypeTitle(TYPE_SHIP_UNLOADED)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_SHIP_UNLOADED
        cmbUnitFilter.AddItem strTypeTitle(TYPE_SHIP_LOADED)
        cmbUnitFilter.ItemData(cmbUnitFilter.NewIndex) = TYPE_SHIP_LOADED
        cmbUnitFilter.Visible = True
    Case udNUKE
        cmbUnitFilter.Visible = False
    Case udENEMY, udList
        cmbUnitFilter.Visible = False
    End Select
End If

For Index = 0 To cmbUnitFilter.ListCount - 1
    If SubType = cmbUnitFilter.ItemData(Index) Then
        cmbUnitFilter.ListIndex = Index
        Exit For
    End If
Next Index
End Sub

Private Sub FillUnitList()
Dim NeedHeader As Boolean
Dim strText As String
Dim ListType As enumUnitDisplay

Dim rs As Recordset
Dim X As Integer
Dim substype As String
Dim substypekey As String
Dim hldIndex As String

rtbUnitList.Visible = False
rtbUnitList.Text = vbNullString
grid1.Visible = False

ClearSubCombo
cmbSub.AddItem "All"

For ListType = udSHIP To udENEMY '100603 rjk: changed to udENEMY to include enemy list
    Select Case ListType
    Case udSHIP
        Set rs = rsShipList
        rs.Requery
        substype = "Fleet "
        substypekey = "flt"
        'BuildType = "s" 100603 rjk: Not used in this procedure
        strText = PadString("Ship", 5, False)
        strText = strText + "   id mil gun  sh  fd"
        
    Case udLAND
        Set rs = rsLandList
        rs.Requery
        substype = "Army "
        substypekey = "army"
        'BuildType = "l" 100603 rjk: Not used in this procedure
        strText = PadString("Land", 5, False)
        strText = strText + "   id mil gun  sh  fd"
    
    Case udPLANE
        Set rs = rsPlaneList
        rs.Requery
        substype = "Wing "
        substypekey = "wing"
        'BuildType = "p" 100603 rjk: Not used in this procedure
        strText = PadString("Plane", 5, False)
        strText = strText + "   id range armed sts"
    
    Case udNUKE
        Set rs = rsNuke
        'rs.Requery 100606 rjk: Removed, does not work this table.  The others do queries.
        substype = "Type "
        substypekey = "type"
        strText = PadString("Nuke", 5, False)
        strText = strText + "   id range armed sts"
    
    Case udENEMY
        Set rs = rsEnemyUnit
        'rs.Requery 100603 rjk: Removed, does not work this table.  Need to investigate more later.
        substype = "Nation "
        substypekey = "owner"
        strText = PadString("Enemy", 5, False)
        strText = strText + "   id mil eff tech  type owner"
    End Select
    
    NeedHeader = True
    
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveFirst
    End If
    While Not rs.EOF

        'Add to the sub combo box
        'If FillSubSetFlag Then 092503 rjk: always include
        '100603 rjk: Added CStr for enemy nation to be converted to string.
            AddSubBox substype + CStr(rs.Fields(substypekey)), rs.Fields("x"), rs.Fields("y")
        'End If
        
        'Check to see if we are displaying this fleet
        ' 092303 rjk: Removed secx and secy and switched to SelX and SelY
        '(DisplaySubset = DSP_SECT And secx = rs.Fields("x") And secy = rs.Fields("y")) Then
        ' 092503 rjk: Switched Fleet to have the whole fleet name otherwise can separate ship from planes
        '(DisplaySubset = ssFLEET And Fleet = Trim$(rs.Fields(substypekey))) Or
           If (DisplaySubset = ssALL) Or _
           (DisplaySubset = ssFLEET And Fleet = substype + Trim$(rs.Fields(substypekey))) Or _
           (DisplaySubset = ssSECT And SelX = rs.Fields("x") And SelY = rs.Fields("y")) Then
            'check to see if it is in the selected type subset
            'If SubType = 0 Or InStr(strType(SubType), " " + rs.Fields(1) + " ") Then
            
                'check the mob and eff requirement
                If ListType = udENEMY Then '100603 rjk: Updated udENEMY so it produced output
                    If NeedHeader Then
                        NeedHeader = False
                        rtbUnitList.SelBold = True
                        rtbUnitList.SelStart = 99999
                        rtbUnitList.SelLength = 0
                        rtbUnitList.SelText = strText
                        rtbUnitList.SelBold = False
                    End If
'                    If ListType = udENEMY Then 100603 rjk: not class in this procedure
'                        BuildType = rs.Fields(9)    'Check the class
'                    End If
                    strText = rs.Fields(1)
                    strText = PadString(strText, 5, False)
                    strText = strText + PadString("#" + CStr(rs.Fields(0)), 5)
                    If rs.Fields("mil") <> -1 Then
                        strText = strText + PadString(rs.Fields("mil"), 4)
                    Else
                        strText = strText + " ???"
                    End If
                    If rs.Fields("eff") <> -1 Then
                        strText = strText + PadString(rs.Fields("eff"), 4)
                    Else
                        strText = strText + " ???"
                    End If
                    If rs.Fields("tech") <> -1 Then
                        strText = strText + PadString(rs.Fields("tech"), 5)
                    Else
                        strText = strText + "  ???"
                    End If
                    Select Case rs.Fields("class")
                    Case "p"
                        strText = strText + " Plane"
                    Case "s"
                        strText = strText + "  Ship"
                    Case "l"
                        strText = strText + "  Land"
                    End Select
                    strText = strText + " " + Nations.NationName(rs.Fields("owner")) + "(" + CStr(rs.Fields("owner")) + ")"
                    rtbUnitList.SelStart = 99999
                    rtbUnitList.SelLength = 0
                    rtbUnitList.SelText = vbCrLf + strText
                Else
                    If (((Not NeedMob) Or rs.Fields("mob") > 0) And _
                       ((Not Needeff) Or rs.Fields("eff") > 59 Or _
                       (ListType = udPLANE And rs.Fields("eff") > 39))) Then
                
                        If NeedHeader Then
                            NeedHeader = False
                            rtbUnitList.SelBold = True
                            rtbUnitList.SelStart = 99999
                            rtbUnitList.SelLength = 0
                            rtbUnitList.SelText = strText
                            rtbUnitList.SelBold = False
                        End If
        
                        'Get the unit description
                        strText = rs.Fields(1)
                        strText = PadString(strText, 5, False)
                        strText = strText + PadString("#" + CStr(rs.Fields(0)), 5)
                        If ListType = udSHIP Or ListType = udLAND Then
                            strText = strText + PadString(rs.Fields("mil"), 4) + PadString(rs.Fields("gun"), 4) _
                                + PadString(rs.Fields("shell"), 4) + PadString(rs.Fields("food"), 4)
                        End If
                        If ListType = udPLANE Then
                            strText = strText + PadString(rs.Fields("range"), 6)
                            If rs.Fields("nuke") <> "N/A" Then
                                strText = strText + PadString(rs.Fields("nuke"), 6)
                            Else
                                strText = strText + "      "
                            End If
                            If rs.Fields("eff") < 39 Or rs.Fields("mob") < 0 Then
                                 strText = strText + " out  "
                            ElseIf rs.Fields("eff") > 89 Then
                                 strText = strText + " OK   "
                            Else
                                strText = strText + " damgd"
                            End If
                            
                        End If
                        rtbUnitList.SelStart = 99999
                        rtbUnitList.SelLength = 0
                        rtbUnitList.SelText = vbCrLf + strText
                        
                        If ListType = udSHIP Then
                        
                        'list cargo
                        For X = 12 To 24
                            If InStr("pln civ uw hcm lcm oil petrol rad iron dust bar ", rs.Fields(X).Name) > 0 Then
                                If rs.Fields(X).Value > 0 Then
                                    strText = "  cargo: " + CStr(rs.Fields(X).Value) + " " + rs.Fields(X).Name
                                    rtbUnitList.SelItalic = True
                                    rtbUnitList.SelStart = 99999
                                    rtbUnitList.SelLength = 0
                                    rtbUnitList.SelText = vbCrLf + strText
                                    rtbUnitList.SelItalic = False
                                End If
                            End If
                        Next X
                        
                        'check for lands on ship
                        If rs.Fields("land") > 0 Then
                            hldIndex = rsLand.Index
                            rsLand.Index = "location"
                            rsLand.Seek "=", rs.Fields("x"), rs.Fields("y")
                            If Not rsLand.NoMatch Then
land_loop:
                                If (Not rsLand.EOF) Then
                                    If rsLand.Fields("x").Value = rs.Fields("x") And rsLand.Fields("y").Value = rs.Fields("y").Value Then
                                        If rsLand.Fields("ship") = rs.Fields(0) Then
                                            strText = "  cargo: " + rsLand.Fields("type").Value + " #" + CStr(rsLand.Fields(0))
                                            rtbUnitList.SelItalic = True
                                            rtbUnitList.SelStart = 99999
                                            rtbUnitList.SelLength = 0
                                            rtbUnitList.SelText = vbCrLf + strText
                                            rtbUnitList.SelItalic = False
                                        End If
                                        rsLand.MoveNext
                                        GoTo land_loop
                                    End If
                                End If
                            End If
                            rsLand.Index = hldIndex
                        End If
                        
                        'check for planes on ship
                        If rs.Fields("pln") > 0 Then
                            hldIndex = rsPlane.Index
                            rsPlane.Index = "location"
                            rsPlane.Seek "=", rs.Fields("x"), rs.Fields("y")
                            If Not rsPlane.NoMatch Then
plane_loop:
                                If (Not rsPlane.EOF) Then
                                    If rsPlane.Fields("x").Value = rs.Fields("x") And rsPlane.Fields("y").Value = rs.Fields("y").Value Then
                                        If rsPlane.Fields("ship") = rs.Fields(0) Then
                                            strText = "  cargo: " + rsPlane.Fields("type").Value + " #" + CStr(rsPlane.Fields(0))
                                            If rsPlane.Fields("nuke") <> "N/A" Then
                                                strText = strText + " armed: " + rsPlane.Fields("nuke")
                                            End If
                                            rtbUnitList.SelItalic = True
                                            rtbUnitList.SelStart = 99999
                                            rtbUnitList.SelLength = 0
                                            rtbUnitList.SelText = vbCrLf + strText
                                            rtbUnitList.SelItalic = False
                                        End If
                                        rsPlane.MoveNext
                                        GoTo plane_loop
                                    End If
                                End If
                            End If
                            rsPlane.Index = hldIndex
                        End If
                    End If
                End If
           End If
        End If
        rs.MoveNext
    Wend
    If Not NeedHeader Then
        rtbUnitList.SelStart = 99999
        rtbUnitList.SelLength = 0
        rtbUnitList.SelText = vbCrLf + vbCrLf
    End If

Next ListType

SetIndexSubCombo '092503 rjk: Added to ensure the combo is reflective of the grid contents.

rtbUnitList.Visible = True
End Sub

Sub SetCommandFocus(Optional KeyAscii As Integer)
'check for current prompt
If CommandCursorFocus Or PromptUp Then
    Exit Sub
End If

On Error Resume Next
Me.txtEntry.SetFocus

' if hitting a key caused the command box to get focus, then
' insert the key hit in the proper place
If KeyAscii > 0 Then
    If KeyAscii = 8 Then 'backspace
        CommandCursorPos = CommandCursorPos - 1
        If CommandCursorPos < 1 Then
            CommandCursorPos = 1
        Else
            txtEntry.SelStart = CommandCursorPos
            txtEntry.SelLength = 1
            txtEntry.SelText = vbNullString
        End If
    Else
        txtEntry.SelStart = CommandCursorPos
        txtEntry.SelText = Chr$(KeyAscii)
        CommandCursorPos = CommandCursorPos + 1
        txtEntry.SelStart = CommandCursorPos
    End If
End If
End Sub

Public Sub MoveMap(sx As Integer, sy As Integer)
CenterMap picMap, sx, sy
MapValid = False        'suppress screen redraws
vsMap.Value = OriginY
hsMap.Value = OriginX
MapValid = True
DrawHexes
FillSectorBox sx, sy
End Sub

Public Function BuildMissionString(BuildType As String, id As Integer) As String

Dim n As Integer
Dim strTemp As String
Dim strTemp2 As String

rsMissions.Seek "=", BuildType, id
If Not rsMissions.NoMatch Then
    strTemp = "Mission: " + rsMissions.Fields("mission") + ", "
    If Len(rsMissions.Fields("op sector")) > 0 Then
        strTemp = strTemp + "Center: " + CStr(rsMissions.Fields("op sector")) + ",  "
    End If
    strTemp = strTemp + " Radius: " + CStr(rsMissions.Fields("radius"))
End If

'check orders if ship
If BuildType = "s" Then
    rsShipOrders.Seek "=", id
    If Not rsShipOrders.NoMatch Then
        strTemp = strTemp + "  Orders - dest: " + rsShipOrders.Fields("start sector")
        If Len(rsShipOrders.Fields("end sector")) > 0 Then
            strTemp = strTemp + " => " + rsShipOrders.Fields("end sector")
        End If
        For n = 1 To 6
            If Len(rsShipOrders.Fields("start cargo " + CStr(n))) > 0 Then
                strTemp2 = strTemp2 + " S " + rsShipOrders.Fields("start cargo " + CStr(n)) + " " _
                    + CStr(rsShipOrders.Fields("start level " + CStr(n)))
            End If
            If Len(rsShipOrders.Fields("end cargo " + CStr(n))) > 0 Then
                strTemp2 = strTemp2 + " E " + rsShipOrders.Fields("end cargo " + CStr(n)) + " " _
                    + CStr(rsShipOrders.Fields("end level " + CStr(n))) + " "
            End If
        Next n
        If Len(strTemp2) > 0 Then
            strTemp = strTemp + "  Holds " + strTemp2
        End If
    End If
End If

BuildMissionString = strTemp
End Function

Public Sub DisplayPromptHelp(strSubject As String)
If frmOptions.bLocalHelp Then
    HtmlHelp hwnd, App.path + "/help/winace.chm", HH_DISPLAY_TOPIC, ByVal strSubject + ".html"
Else
    PositionHelp = True
    frmEmpCmd.SubmitEmpireCommand "info " + LCase$(strSubject), True
End If
End Sub

Private Sub BuildUnitFilters()
Dim n As Integer
Dim bAdd As Boolean

For n = TYPE_START To TYPE_END
    strType(n) = " "
Next n
strTypeTitle(TYPE_START) = "Start - Should never see"
strTypeTitle(TYPE_ALL) = "All"
strTypeTitle(TYPE_SHIP_UNLOADED) = "Not on Ship"
strTypeTitle(TYPE_LAND_UNLOADED) = "Not on Land"
strTypeTitle(TYPE_SHIP_LOADED) = "Loaded on Ship"
strTypeTitle(TYPE_LAND_LOADED) = "Loaded on Land"
'ship capabilities/abilities
strTypeTitle(TYPE_S_FISH) = "Fish"
strTypeTitle(TYPE_S_TORP) = "Torpedo"
strTypeTitle(TYPE_S_DCHRG) = "Depth Charges"
strTypeTitle(TYPE_S_PLANE) = "Planes"
strTypeTitle(TYPE_S_MISS) = "Missiles"
strTypeTitle(TYPE_S_OIL) = "Drill Oil"
strTypeTitle(TYPE_S_OILER) = "Haul Oil"
strTypeTitle(TYPE_S_SONAR) = "Sonar"
strTypeTitle(TYPE_S_MINE) = "Mine"
strTypeTitle(TYPE_S_SWEEP) = "Sweep"
strTypeTitle(TYPE_S_SUB) = "Submarine"
strTypeTitle(TYPE_S_SPY) = "Spy"
strTypeTitle(TYPE_S_LAND) = "Land"
strTypeTitle(TYPE_S_SEMI_LAND) = "Semi-Land"
strTypeTitle(TYPE_S_SUB_TORP) = "Torpedo a Sub"
strTypeTitle(TYPE_S_TRADE) = "Trade"
strTypeTitle(TYPE_S_CANAL) = "Canal"
strTypeTitle(TYPE_S_ANTI_MISSILE) = "Anti Missile"
'ship commands
strTypeTitle(TYPE_SC_FIRE) = "fire command"
'plane capabilities/abilities
strTypeTitle(TYPE_P_TACTICAL) = "Pinpoint (Tactical) Bombers"
strTypeTitle(TYPE_P_BOMBER) = "Strategic Bombers"
strTypeTitle(TYPE_P_INTERCEPT) = "Interceptors"
strTypeTitle(TYPE_P_CARGO) = "Cargo"
strTypeTitle(TYPE_P_VTOL) = "VTOL"
strTypeTitle(TYPE_P_MISSILE) = "Missile"
strTypeTitle(TYPE_P_LIGHT) = "Light"
strTypeTitle(TYPE_P_SPY) = "Recon."
strTypeTitle(TYPE_P_IMAGE) = "Image"
strTypeTitle(TYPE_P_SATELLITE) = "Satellite"
strTypeTitle(TYPE_P_STEALTH) = "Stealth"
strTypeTitle(TYPE_P_SDI) = "SDI"
strTypeTitle(TYPE_P_HALF_STEALTH) = "Half-Stealth"
strTypeTitle(TYPE_P_X_LIGHT) = "x-light"
strTypeTitle(TYPE_P_CHOPPER) = "Chopper"
strTypeTitle(TYPE_P_ANTISUB) = "Anti-Submarine"
strTypeTitle(TYPE_P_PARA) = "Paratroop"
strTypeTitle(TYPE_P_ESCORT) = "Escort"
strTypeTitle(TYPE_P_MINE) = "Mine"
strTypeTitle(TYPE_P_SWEEP) = "Sweep"
strTypeTitle(TYPE_P_MARINE) = "Marine Missile"
'plane groups
strTypeTitle(TYPE_PG_ESCORTS) = "escorts (intercept and escort)"
'plane commands
strTypeTitle(TYPE_PC_BOMB) = "bomb command"
strTypeTitle(TYPE_PC_LAUNCH) = "Launch command"
strTypeTitle(TYPE_PC_DROP) = "Drop command"
'land units capabilities/abilities
strTypeTitle(TYPE_L_XLIGHT) = "x-light"
strTypeTitle(TYPE_L_ENGINEER) = "Engineer"
strTypeTitle(TYPE_L_SUPPLY) = "Supply"
strTypeTitle(TYPE_L_SECURITY) = "Security"
strTypeTitle(TYPE_L_LIGHT) = "Light"
strTypeTitle(TYPE_L_MARINE) = "Marine"
strTypeTitle(TYPE_L_RECON) = "Recon."
strTypeTitle(TYPE_L_RADAR) = "Radar"
strTypeTitle(TYPE_L_ASSAULT) = "Assault"
strTypeTitle(TYPE_L_TRAIN) = "Train"
'land unit commands
strTypeTitle(TYPE_LC_FIRE) = "fire command"
strTypeTitle(TYPE_END) = "End - Should never see"

If rsBuild.RecordCount = 0 Then
    Exit Sub
End If
rsBuild.MoveFirst
While Not (rsBuild.EOF)
    For n = 32 To 51
        Select Case rsBuild.Fields("type").Value
        Case "s"
            Select Case rsBuild.Fields(n).Value
            'ship capabilities and abilities
            Case "fish"
                strType(TYPE_S_FISH) = strType(TYPE_S_FISH) + rsBuild.Fields("id").Value + " "
            Case "torp"
                strType(TYPE_S_TORP) = strType(TYPE_S_TORP) + rsBuild.Fields("id").Value + " "
            Case "dchrg"
                strType(TYPE_S_DCHRG) = strType(TYPE_S_DCHRG) + rsBuild.Fields("id").Value + " "
            Case "plane"
                strType(TYPE_S_PLANE) = strType(TYPE_S_PLANE) + rsBuild.Fields("id").Value + " "
            Case "miss"
                strType(TYPE_S_MISS) = strType(TYPE_S_MISS) + rsBuild.Fields("id").Value + " "
            Case "oil"
                strType(TYPE_S_OIL) = strType(TYPE_S_OIL) + rsBuild.Fields("id").Value + " "
            Case "oiler"
                strType(TYPE_S_OILER) = strType(TYPE_S_OILER) + rsBuild.Fields("id").Value + " "
            Case "sonar"
                strType(TYPE_S_SONAR) = strType(TYPE_S_SONAR) + rsBuild.Fields("id").Value + " "
            Case "mine"
                strType(TYPE_S_MINE) = strType(TYPE_S_MINE) + rsBuild.Fields("id").Value + " "
            Case "sweep"
                strType(TYPE_S_SWEEP) = strType(TYPE_S_SWEEP) + rsBuild.Fields("id").Value + " "
            Case "sub"
                strType(TYPE_S_SUB) = strType(TYPE_S_SUB) + rsBuild.Fields("id").Value + " "
            Case "spy"
                strType(TYPE_S_SPY) = strType(TYPE_S_SPY) + rsBuild.Fields("id").Value + " "
            Case "land"
                strType(TYPE_S_LAND) = strType(TYPE_S_LAND) + rsBuild.Fields("id").Value + " "
            Case "semi-land"
                strType(TYPE_S_SEMI_LAND) = strType(TYPE_S_SEMI_LAND) + rsBuild.Fields("id").Value + " "
            Case "sub-torp"
                strType(TYPE_S_SUB_TORP) = strType(TYPE_S_SUB_TORP) + rsBuild.Fields("id").Value + " "
            Case "trade"
                strType(TYPE_S_TRADE) = strType(TYPE_S_TRADE) + rsBuild.Fields("id").Value + " "
            Case "canal"
                strType(TYPE_S_CANAL) = strType(TYPE_S_CANAL) + rsBuild.Fields("id").Value + " "
            Case "anti-missile"
                strType(TYPE_S_ANTI_MISSILE) = strType(TYPE_S_ANTI_MISSILE) + rsBuild.Fields("id").Value + " "
            End Select
        Case "p" 'plane capabilities and abilities
            Select Case rsBuild.Fields(n).Value
            Case "tactical"
                strType(TYPE_P_TACTICAL) = strType(TYPE_P_TACTICAL) + rsBuild.Fields("id").Value + " "
            Case "bomber"
                strType(TYPE_P_BOMBER) = strType(TYPE_P_BOMBER) + rsBuild.Fields("id").Value + " "
            Case "intercept"
                strType(TYPE_P_INTERCEPT) = strType(TYPE_P_INTERCEPT) + rsBuild.Fields("id").Value + " "
            Case "cargo"
                strType(TYPE_P_CARGO) = strType(TYPE_P_CARGO) + rsBuild.Fields("id").Value + " "
            Case "VTOL"
                strType(TYPE_P_VTOL) = strType(TYPE_P_VTOL) + rsBuild.Fields("id").Value + " "
            Case "missile"
                strType(TYPE_P_MISSILE) = strType(TYPE_P_MISSILE) + rsBuild.Fields("id").Value + " "
            Case "light"
                strType(TYPE_P_LIGHT) = strType(TYPE_P_LIGHT) + rsBuild.Fields("id").Value + " "
            Case "spy"
                strType(TYPE_P_SPY) = strType(TYPE_P_SPY) + rsBuild.Fields("id").Value + " "
            Case "image"
                strType(TYPE_P_IMAGE) = strType(TYPE_P_IMAGE) + rsBuild.Fields("id").Value + " "
            Case "satellite"
                strType(TYPE_P_SATELLITE) = strType(TYPE_P_SATELLITE) + rsBuild.Fields("id").Value + " "
            Case "stealth"
                strType(TYPE_P_STEALTH) = strType(TYPE_P_STEALTH) + rsBuild.Fields("id").Value + " "
            Case "SDI"
                strType(TYPE_P_SDI) = strType(TYPE_P_SDI) + rsBuild.Fields("id").Value + " "
            Case "half-stealth"
                strType(TYPE_P_HALF_STEALTH) = strType(TYPE_P_HALF_STEALTH) + rsBuild.Fields("id").Value + " "
            Case "x-light"
                strType(TYPE_P_X_LIGHT) = strType(TYPE_P_X_LIGHT) + rsBuild.Fields("id").Value + " "
            Case "helo"
                strType(TYPE_P_CHOPPER) = strType(TYPE_P_CHOPPER) + rsBuild.Fields("id").Value + " "
            Case "ASW"
                strType(TYPE_P_ANTISUB) = strType(TYPE_P_ANTISUB) + rsBuild.Fields("id").Value + " "
            Case "para"
                strType(TYPE_P_PARA) = strType(TYPE_P_PARA) + rsBuild.Fields("id").Value + " "
            Case "escort"
                strType(TYPE_P_ESCORT) = strType(TYPE_P_ESCORT) + rsBuild.Fields("id").Value + " "
            Case "mine"
                strType(TYPE_P_MINE) = strType(TYPE_P_MINE) + rsBuild.Fields("id").Value + " "
            Case "sweep"
                strType(TYPE_P_SWEEP) = strType(TYPE_P_SWEEP) + rsBuild.Fields("id").Value + " "
            Case "marine"
                strType(TYPE_P_MARINE) = strType(TYPE_P_MARINE) + rsBuild.Fields("id").Value + " "
            End Select
        Case "l" 'land unit capabilities
            Select Case rsBuild.Fields(n).Value
            Case "xlight"
                strType(TYPE_L_XLIGHT) = strType(TYPE_L_XLIGHT) + rsBuild.Fields("id").Value + " "
            Case "engineer"
                strType(TYPE_L_ENGINEER) = strType(TYPE_L_ENGINEER) + rsBuild.Fields("id").Value + " "
            Case "supply"
                strType(TYPE_L_SUPPLY) = strType(TYPE_L_SUPPLY) + rsBuild.Fields("id").Value + " "
            Case "security"
                strType(TYPE_L_SECURITY) = strType(TYPE_L_SECURITY) + rsBuild.Fields("id").Value + " "
            Case "LIGHT"
                strType(TYPE_L_LIGHT) = strType(TYPE_L_LIGHT) + rsBuild.Fields("id").Value + " "
            Case "marine"
                strType(TYPE_L_MARINE) = strType(TYPE_L_MARINE) + rsBuild.Fields("id").Value + " "
            Case "recon"
                strType(TYPE_L_RECON) = strType(TYPE_L_RECON) + rsBuild.Fields("id").Value + " "
            Case "radar"
                strType(TYPE_L_RADAR) = strType(TYPE_L_RADAR) + rsBuild.Fields("id").Value + " "
            Case "assault"
                strType(TYPE_L_ASSAULT) = strType(TYPE_L_ASSAULT) + rsBuild.Fields("id").Value + " "
            Case "train"
                strType(TYPE_L_TRAIN) = strType(TYPE_L_TRAIN) + rsBuild.Fields("id").Value + " "
            End Select
        End Select
    Next n
    rsBuild.MoveNext
Wend

rsBuild.MoveFirst
Do While Not (rsBuild.EOF)
    Select Case rsBuild.Fields("type").Value
    Case "s" 'Ships
        If rsBuild.Fields(17).Value > 0 Then 'the number of guns can fire
            strType(TYPE_SC_FIRE) = strType(TYPE_SC_FIRE) + rsBuild.Fields("id").Value + " "
        End If
    Case "p" 'Planes
        bAdd = True
        If InStr(strType(TYPE_P_INTERCEPT), rsBuild.Fields("id").Value) > 0 Or _
           InStr(strType(TYPE_P_ESCORT), rsBuild.Fields("id").Value) > 0 Then
            strType(TYPE_PG_ESCORTS) = strType(TYPE_PG_ESCORTS) + rsBuild.Fields("id").Value + " "
        End If
        If InStr(strType(TYPE_P_MISSILE), rsBuild.Fields("id").Value) > 0 Or _
           InStr(strType(TYPE_P_SATELLITE), rsBuild.Fields("id").Value) > 0 Then
            strType(TYPE_PC_LAUNCH) = strType(TYPE_PC_LAUNCH) + rsBuild.Fields("id").Value + " "
        End If
        If InStr(strType(TYPE_P_CARGO), rsBuild.Fields("id").Value) > 0 Or _
           InStr(strType(TYPE_P_MINE), rsBuild.Fields("id").Value) > 0 Then
            strType(TYPE_PC_DROP) = strType(TYPE_PC_DROP) + rsBuild.Fields("id").Value + " "
        End If
        If InStr(strType(TYPE_P_CARGO), rsBuild.Fields("id").Value) = 0 Or _
           InStr(strType(TYPE_P_TACTICAL), rsBuild.Fields("id").Value) > 0 Then
            strType(TYPE_PC_BOMB) = strType(TYPE_PC_BOMB) + rsBuild.Fields("id").Value + " "
        End If
    Case "l" 'Land Units
        If rsBuild.Fields(21).Value > 0 Then 'the number of guns can fire
            strType(TYPE_LC_FIRE) = strType(TYPE_LC_FIRE) + rsBuild.Fields("id").Value + " "
        End If
    End Select
    rsBuild.MoveNext
Loop
End Sub

Public Sub SetUnitDisplay(UnitDisplay As enumUnitDisplay, UnitFilter As enumUnitType, bMob As Boolean, bEff As Boolean, bRange As Boolean, Optional strTarget As String)
Static LastUnitFilter As enumUnitType
Static bLastRange As Boolean
Static bLastMob As Boolean
Static bLastEff As Boolean
Static strLastTarget As String
Dim tstSectX As Integer
Dim tstSectY As Integer

If TabStrip1.Tabs(TabStrip1.SelectedItem.Index).key <> "Unit" Or _
   LastUnitDisplay <> UnitDisplay Or _
   bLastRange <> bRange Or _
   LastUnitFilter <> UnitFilter Or _
   bLastMob <> bMob Or _
   bLastEff <> bEff Then
    'Mob Button
    NeedMob = bMob
    If bMob Then
        Toolbar1.Buttons(7).Value = tbrPressed
    Else
        Toolbar1.Buttons(7).Value = tbrUnpressed
    End If
    'Eff Button
    Needeff = bEff
    If bEff Then
        Toolbar1.Buttons(8).Value = tbrPressed
    Else
        Toolbar1.Buttons(8).Value = tbrUnpressed
    End If
    
    SubType = UnitFilter
    
    DisplaySubset = ssSECT
    
    Select Case UnitDisplay
    Case udSHIP
        DisplaySelect = udSHIP
    Case udLAND
        If bRange Then
            If strLastTarget <> strTarget Then
                If ParseSectors(tstSectX, tstSectY, strTarget) Then
                    DisplaySelect = udLAND
                    DisplaySubset = ssATTACK_RANGE
                End If
                strTarget = strLastTarget
            End If
        Else
            DisplaySelect = udLAND
        End If
    Case udPLANE
        If bRange Then
            If strLastTarget <> strTarget Then
                If ParseSectors(tstSectX, tstSectY, strTarget) Then
                    DisplaySelect = udPLANE
                    If UnitFilter = TYPE_P_MISSILE Then
                        DisplaySubset = ssMISSILE_RANGE
                    Else
                        DisplaySubset = ssPLANE_RANGE
                    End If
                End If
                strTarget = strLastTarget
            End If
        Else
            DisplaySelect = udPLANE
        End If
    Case udNUKE
        DisplaySelect = udNUKE
    Case udENEMY
        DisplaySelect = udENEMY
    Case udList
        DisplaySelect = udList
    End Select
    
    FillGrid
    DisplayUnitBox
    If bRange Then
        SortGrid 3
    Else
        SortGrid 1
    End If
    'Save current system system (not user settings)
    LastUnitDisplay = UnitDisplay
    bLastRange = bRange
    LastUnitFilter = UnitFilter
    bLastMob = bMob
    bLastEff = bEff
End If
End Sub

Public Function UnitCharacteristicCheck(UnitType As enumUnitType, strUnit As String) As Boolean
If InStr(strType(UnitType), " " + strUnit + " ") > 0 Then
    UnitCharacteristicCheck = True
Else
    UnitCharacteristicCheck = False
End If
End Function

Public Sub SetPlayerTimeInterval()
If frmOptions.playersTimeInterval = 0 Then
    PlayersTimer.Enabled = False
Else
    PlayersTimer.Interval = 60000
    PlayersTimer.Enabled = True
End If
End Sub

Private Sub DrawEvent(picMap As Object, Optional secx As Integer = 0, Optional secy As Integer = -1)
Dim oldWidth
Dim oldcolor
Dim natColor As Integer
Dim iIndex As Integer
Dim bSectorCheck As Boolean
Dim PosX As Single
Dim PosY As Single


If secx = 0 And secy = -1 Then
    bSectorCheck = False
Else
    bSectorCheck = True
End If

oldWidth = picMap.DrawWidth
picMap.DrawWidth = 3   ' Set DrawWidth
oldcolor = picMap.ForeColor
For iIndex = 1 To EventMarkers.Count
    If (EventMarkers.X(iIndex) = secx And EventMarkers.Y(iIndex) = secy) Or Not bSectorCheck Then
        natColor = EventMarkers.Nation(iIndex)
        If natColor >= LBound(PlayerColors) And natColor <= UBound(PlayerColors) Then
            picMap.ForeColor = PlayerColors(natColor)
        ElseIf natColor = -2 Then
            picMap.ForeColor = BasicText(2)  'water text color
        ElseIf natColor = -3 Then
            picMap.ForeColor = vbYellow
        Else
            picMap.ForeColor = vbRed
        End If
        Coord EventMarkers.X(iIndex), EventMarkers.Y(iIndex), PosX, PosY
        DrawHex picMap, PosX, PosY
        If bSectorCheck Then
            picMap.ForeColor = oldcolor
            picMap.DrawWidth = oldWidth   ' Set DrawWidth
            Exit Sub
        End If
    End If
Next iIndex
picMap.ForeColor = oldcolor
picMap.DrawWidth = oldWidth   ' Set DrawWidth
End Sub

Private Sub ExecuteFile(iIndex As Integer)
If tScriptButtons(iIndex).bJumpPoint Then
    Dim sx As Integer
    Dim sy As Integer
    
    rsNation.MoveFirst
    If Not IsNull(rsNation.Fields("JumpPoint" + CStr(tScriptButtons(iIndex).iJumpPoint + 1))) And _
        tScriptButtons(iIndex).iJumpPoint >= 0 Then
        If ParseSectors(sx, sy, rsNation.Fields("JumpPoint" + CStr(tScriptButtons(iIndex).iJumpPoint + 1))) Then
            MoveMap sx, sy
        End If
    End If
Else
    Dim strCommand As String
    Dim fs As Object
    Dim fileObject As Object
    
    If Len(tScriptButtons(iIndex).strFileName) = 0 Then
        Exit Sub
    End If
    If Not tScriptButtons(iIndex).bSendOutputReportWindow Then
        frmEmpCmd.SubmitEmpireCommand "bf1", False
    Else
        frmEmpCmd.SubmitEmpireCommand "br1", False
    End If
    On Error GoTo Empire_Error
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set fileObject = fs.OpenTextFile(tScriptButtons(iIndex).strFileName)
    While fileObject.AtEndOfStream <> True
        strCommand = fileObject.readline
        frmEmpCmd.SubmitEmpireCommand CStr(strCommand), False
    Wend
    fileObject.Close
    
    'filenumber = FreeFile
    'On Error GoTo Empire_Error
    'Open tScriptButtons(iIndex).strFileName For Input As #filenumber
    'While Not EOF(filenumber)
    '    Input #filenumber, vCommand
    '    frmEmpCmd.SubmitEmpireCommand CStr(vCommand), False
    'Wend
    'Close #filenumber
    
    If Not tScriptButtons(iIndex).bSendOutputReportWindow Then
        frmEmpCmd.SubmitEmpireCommand "bf2", False
    Else
        frmEmpCmd.SubmitEmpireCommand "br2", False
    End If
End If
Exit Sub

Empire_Error:
EmpireError "Failed to run script file for custom script button", CStr(iIndex), tScriptButtons(iIndex).strFileName
End Sub

