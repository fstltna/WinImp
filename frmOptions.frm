VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   3345
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6150
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picPlaying 
      BorderStyle     =   0  'None
      Height          =   2220
      Left            =   240
      ScaleHeight     =   2220
      ScaleWidth      =   5685
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   5685
      Begin VB.Frame frmPathLength 
         Caption         =   "Path Length"
         Height          =   615
         Left            =   4560
         TabIndex        =   118
         Top             =   1560
         Width           =   1092
         Begin VB.TextBox txtPathLength 
            Height          =   285
            Left            =   240
            TabIndex        =   119
            Text            =   "20"
            ToolTipText     =   "Maximum Path Length Checked"
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame frmFriendly 
         Caption         =   "Friendly Relation"
         Height          =   612
         Left            =   0
         TabIndex        =   24
         Top             =   1560
         Width           =   1932
         Begin VB.OptionButton optFriendlyAlly 
            Caption         =   "Allied"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton optFriendlyNeutral 
            Caption         =   "Neutral"
            Height          =   255
            Left            =   960
            TabIndex        =   25
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame frmAutoAttack 
         Caption         =   "Auto Attack"
         Height          =   615
         Left            =   2040
         TabIndex        =   56
         Top             =   1560
         Width           =   2412
         Begin VB.TextBox txtDefenseScaling 
            Height          =   285
            Left            =   1680
            TabIndex        =   57
            Text            =   "20"
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblDefenseScaling 
            Caption         =   "DefenseScaling (%)"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame fraGame 
         Caption         =   "Game"
         Height          =   1475
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   5652
         Begin VB.CheckBox chkDumpHeaders 
            Caption         =   "Display Dump Headers"
            Height          =   375
            Left            =   2040
            TabIndex        =   116
            Top             =   1080
            Width           =   2172
         End
         Begin VB.CheckBox chkCapital 
            Caption         =   "Use Capital as Start"
            Height          =   195
            Left            =   120
            TabIndex        =   65
            Top             =   1140
            Width           =   1695
         End
         Begin VB.CheckBox chkLocalHelp 
            Caption         =   "Use Local Help"
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   840
            Width           =   1575
         End
         Begin VB.CheckBox chkSound 
            Caption         =   "Use Sound"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   1335
         End
         Begin VB.CheckBox chkNoIron 
            Caption         =   "No Iron Game"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   540
            Width           =   1335
         End
         Begin VB.CheckBox chkRetreat 
            Caption         =   "Set Retreat on Nav/March"
            Height          =   195
            Left            =   2040
            TabIndex        =   21
            Top             =   840
            Width           =   2295
         End
         Begin VB.CheckBox chkMaxProd 
            Caption         =   "Max. Prod. instead of Pop."
            Height          =   195
            Left            =   2040
            TabIndex        =   20
            Top             =   549
            Width           =   2295
         End
      End
   End
   Begin VB.PictureBox picView 
      BorderStyle     =   0  'None
      Height          =   2220
      Left            =   240
      ScaleHeight     =   2220
      ScaleWidth      =   5685
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraView 
         Caption         =   "Viewing"
         Height          =   1320
         Left            =   0
         TabIndex        =   6
         Top             =   120
         Width           =   3735
         Begin VB.CheckBox chkShortName 
            Caption         =   "Use Short Names in Unit Grid"
            Height          =   255
            Left            =   1200
            TabIndex        =   117
            Top             =   240
            Width           =   2412
         End
         Begin VB.CheckBox chkClear 
            Caption         =   "Clear Result Box on New Command"
            Height          =   375
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   2895
         End
         Begin VB.CheckBox chkToolbar 
            Caption         =   "Toolbar"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox chkStatusBar 
            Caption         =   "Status Bar"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   540
            Width           =   1092
         End
         Begin VB.CheckBox chkWarship 
            Caption         =   "Show Only Warships"
            Height          =   255
            Left            =   1200
            TabIndex        =   7
            Top             =   540
            Width           =   1935
         End
      End
      Begin VB.Frame frmFlash 
         Caption         =   "Flash"
         Height          =   615
         Left            =   0
         TabIndex        =   59
         Top             =   1560
         Width           =   5655
         Begin VB.CheckBox chkFlashLogging 
            Caption         =   "Logging"
            Height          =   255
            Left            =   4200
            TabIndex        =   104
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox chkFlashPlayerColors 
            Caption         =   "Use Player Colors"
            Height          =   255
            Left            =   2160
            TabIndex        =   61
            Top             =   240
            Width           =   1695
         End
         Begin VB.CheckBox chkFlashChat 
            Caption         =   "Window Enabled"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame frmImage 
         Caption         =   "Background Image"
         Height          =   1335
         Left            =   3840
         TabIndex        =   50
         Top             =   120
         Width           =   1815
         Begin VB.PictureBox picBorderColor 
            AutoRedraw      =   -1  'True
            FillStyle       =   0  'Solid
            FontTransparent =   0   'False
            Height          =   255
            Left            =   1200
            ScaleHeight     =   195
            ScaleWidth      =   435
            TabIndex        =   53
            Top             =   960
            Width           =   495
         End
         Begin VB.CommandButton cmdImage 
            Caption         =   "Browse"
            Height          =   255
            Left            =   480
            TabIndex        =   52
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtImageName 
            Height          =   315
            Left            =   240
            TabIndex        =   51
            Text            =   "Text1"
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblBorderColor 
            Caption         =   "Border Color"
            Height          =   255
            Left            =   240
            TabIndex        =   54
            Top             =   960
            Width           =   975
         End
      End
   End
   Begin VB.PictureBox picUpdate 
      BorderStyle     =   0  'None
      Height          =   2220
      Left            =   240
      ScaleHeight     =   2220
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   5685
      Begin VB.Frame fraPlayer 
         Caption         =   "Display Other Players"
         Height          =   615
         Left            =   3000
         TabIndex        =   15
         Top             =   1560
         Width           =   2535
         Begin VB.ComboBox cmbPlayer 
            Height          =   315
            ItemData        =   "frmOptions.frx":000C
            Left            =   120
            List            =   "frmOptions.frx":0020
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame frmDeity 
         Caption         =   "Deity"
         Height          =   615
         Left            =   0
         TabIndex        =   62
         Top             =   1560
         Width           =   2535
         Begin VB.CheckBox chkFullDumpwithoutSea 
            Caption         =   "Full Dump without Sea"
            Height          =   195
            Left            =   240
            TabIndex        =   63
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame fraUpdate 
         Caption         =   "Updating"
         Height          =   1500
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   5535
         Begin VB.CheckBox chkSuppressUnitDamageReports 
            Caption         =   "Suppress Unit Damage Reports"
            Height          =   255
            Left            =   120
            TabIndex        =   101
            ToolTipText     =   "Suppress the displaying of Telegram Notification with INORM OFF"
            Top             =   960
            Width           =   2655
         End
         Begin VB.CheckBox chkSuppressTelegramNotification 
            Caption         =   "Suppress Telegrams Notifiations"
            Height          =   255
            Left            =   120
            TabIndex        =   100
            ToolTipText     =   "Suppress the displaying of Telegram Notification with INORM OFF"
            Top             =   1200
            Width           =   2655
         End
         Begin VB.CheckBox chkSuppressUpdateCommands 
            Caption         =   "Suppress Update Commands"
            Height          =   255
            Left            =   2760
            TabIndex        =   99
            ToolTipText     =   "Suppress the displaying of the update commands"
            Top             =   960
            Width           =   2415
         End
         Begin VB.CheckBox chkSuppressOilContentAtSea 
            Caption         =   "Suppress Oil Content Refresh"
            Height          =   255
            Left            =   2760
            TabIndex        =   98
            ToolTipText     =   "Suppress Oil Content Refresh For Sea Sectors"
            Top             =   240
            Width           =   2415
         End
         Begin VB.CheckBox chkCommodityRatio 
            Caption         =   "Show Commodity Ratio"
            Height          =   255
            Left            =   120
            TabIndex        =   97
            Top             =   720
            Width           =   2055
         End
         Begin VB.CheckBox chkRead 
            Caption         =   "Automatic Read"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   480
            Width           =   2295
         End
         Begin VB.CheckBox chkSuppressBmap 
            Caption         =   "Suppress Bmap Refresh"
            Height          =   255
            Left            =   2760
            TabIndex        =   13
            Top             =   720
            Width           =   2055
         End
         Begin VB.CheckBox chkSuppressStrength 
            Caption         =   "Suppress Strength Refresh"
            Height          =   255
            Left            =   2760
            TabIndex        =   12
            Top             =   480
            Width           =   2295
         End
         Begin VB.CheckBox chkUpdate 
            Caption         =   "Automatic Update"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   2295
         End
      End
   End
   Begin VB.PictureBox picTTS 
      BorderStyle     =   0  'None
      Height          =   2220
      Left            =   240
      ScaleHeight     =   2220
      ScaleWidth      =   5685
      TabIndex        =   110
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   5685
      Begin VB.Frame frmTTS 
         Caption         =   "Text To Speech"
         Height          =   2040
         Left            =   120
         TabIndex        =   111
         Top             =   120
         Width           =   2055
         Begin VB.CheckBox chkTTSFlashes 
            Caption         =   "Flashes"
            Height          =   255
            Left            =   240
            TabIndex        =   115
            ToolTipText     =   "Enable TTS for Flashes"
            Top             =   1440
            Width           =   1575
         End
         Begin VB.CheckBox chkTTSBulletins 
            Caption         =   "Bulletins"
            Height          =   255
            Left            =   240
            TabIndex        =   114
            ToolTipText     =   "Enable TTS for Bulletins"
            Top             =   1080
            Width           =   1575
         End
         Begin VB.CheckBox chkTTSTelegrams 
            Caption         =   "Telegram"
            Height          =   255
            Left            =   240
            TabIndex        =   113
            ToolTipText     =   "Enable TTS for Telegrams"
            Top             =   720
            Width           =   1575
         End
         Begin VB.CheckBox chkTTSReportWindow 
            Caption         =   "Report Window"
            Height          =   255
            Left            =   240
            TabIndex        =   112
            ToolTipText     =   "Enable TTS for Report Window"
            Top             =   360
            Width           =   1575
         End
      End
   End
   Begin VB.PictureBox picThemeGame 
      BorderStyle     =   0  'None
      Height          =   2220
      Left            =   240
      ScaleHeight     =   2220
      ScaleWidth      =   5685
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   5685
      Begin VB.Frame fraThemeGame 
         Caption         =   "Theme Game"
         Height          =   2040
         Left            =   120
         TabIndex        =   47
         Top             =   120
         Width           =   3855
         Begin VB.OptionButton optThemeGames 
            Caption         =   "Heavy Metal"
            Height          =   255
            Index           =   7
            Left            =   1680
            TabIndex        =   109
            ToolTipText     =   "South Pacific: The Search for Atlantis"
            Top             =   960
            Width           =   1455
         End
         Begin VB.OptionButton optThemeGames 
            Caption         =   "SP: Search for Atlantis"
            Height          =   255
            Index           =   6
            Left            =   1680
            TabIndex        =   105
            ToolTipText     =   "South Pacific: The Search for Atlantis"
            Top             =   600
            Width           =   2055
         End
         Begin VB.OptionButton optThemeGames 
            Caption         =   "SP 2005"
            Height          =   255
            Index           =   5
            Left            =   1680
            TabIndex        =   103
            ToolTipText     =   "South Pacific 2005"
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optThemeGames 
            Caption         =   "EE9"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   102
            Top             =   1680
            Width           =   1455
         End
         Begin VB.OptionButton optThemeGames 
            Caption         =   "2K4"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   64
            Top             =   1320
            Width           =   1455
         End
         Begin VB.OptionButton optThemeGames 
            Caption         =   "Retro"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   55
            Top             =   960
            Width           =   1455
         End
         Begin VB.OptionButton optThemeGames 
            Caption         =   "St@r W@rs"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   49
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton optThemeGames 
            Caption         =   "None"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   48
            Top             =   240
            Width           =   1455
         End
      End
   End
   Begin VB.PictureBox picScriptButton 
      BorderStyle     =   0  'None
      Height          =   2220
      Left            =   240
      ScaleHeight     =   2220
      ScaleWidth      =   5685
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   5685
      Begin VB.OptionButton optCustomScript 
         Caption         =   "Custom Script"
         Height          =   435
         Left            =   4680
         TabIndex        =   107
         Top             =   60
         Width           =   975
      End
      Begin VB.OptionButton optJumpPoint 
         Caption         =   "Jump Point"
         Height          =   435
         Left            =   3840
         TabIndex        =   106
         Top             =   60
         Width           =   735
      End
      Begin VB.Frame fraScriptButton 
         Caption         =   "Custom Script Button"
         Height          =   1575
         Left            =   0
         TabIndex        =   68
         Top             =   600
         Width           =   5655
         Begin VB.ComboBox cmbJumpPoint 
            Height          =   315
            ItemData        =   "frmOptions.frx":0072
            Left            =   1440
            List            =   "frmOptions.frx":0085
            TabIndex        =   108
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdScriptImageBrowse 
            Caption         =   "Image Browse"
            Height          =   375
            Left            =   4200
            TabIndex        =   75
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox txtScriptImageFileName 
            Height          =   285
            Left            =   1200
            TabIndex        =   74
            Text            =   "Text1"
            Top             =   1080
            Width           =   2895
         End
         Begin VB.CheckBox chkScriptReport 
            Caption         =   "Send Output to a Report Window"
            Height          =   255
            Left            =   240
            TabIndex        =   72
            Top             =   600
            Width           =   2775
         End
         Begin VB.TextBox txtScriptFileName 
            Height          =   285
            Left            =   1440
            TabIndex        =   70
            Text            =   "Text1"
            Top             =   240
            Width           =   4095
         End
         Begin VB.CommandButton cmdScriptBrowse 
            Caption         =   "File Name Browse"
            Height          =   375
            Left            =   3960
            TabIndex        =   69
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lblScriptImage 
            Alignment       =   1  'Right Justify
            Caption         =   "Button Image"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblScriptFileName 
            Alignment       =   1  'Right Justify
            Caption         =   "Scrit File Name"
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.ComboBox cmbScriptButton 
         Height          =   315
         ItemData        =   "frmOptions.frx":00CF
         Left            =   0
         List            =   "frmOptions.frx":00D1
         Style           =   2  'Dropdown List
         TabIndex        =   67
         Top             =   120
         Width           =   3735
      End
   End
   Begin VB.PictureBox picItem 
      BorderStyle     =   0  'None
      Height          =   2220
      Left            =   240
      ScaleHeight     =   2220
      ScaleWidth      =   5685
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   5685
      Begin VB.Frame frmItemLabels 
         Caption         =   "Labels"
         Height          =   1335
         Left            =   120
         TabIndex        =   38
         Top             =   600
         Width           =   2655
         Begin VB.TextBox txtItemDistributionPanelLabel 
            Height          =   285
            Left            =   1440
            TabIndex        =   43
            Text            =   "Text1"
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtItemSectorPanelLabel 
            Height          =   285
            Left            =   1440
            TabIndex        =   41
            Text            =   "Text1"
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtItemFormLabel 
            Height          =   285
            Left            =   1440
            TabIndex        =   39
            Text            =   "Text1"
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblItemDistributionPanelLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "Distibution Panel"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblItemSectorPanelLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "Sector Panel"
            Height          =   255
            Left            =   360
            TabIndex        =   42
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lblItemFormLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "Form"
            Height          =   255
            Left            =   600
            TabIndex        =   40
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.ComboBox cbItem 
         Height          =   315
         Left            =   120
         TabIndex        =   29
         Text            =   "Combo1"
         Top             =   120
         Width           =   2655
      End
      Begin VB.Frame frmItem 
         Caption         =   "Names"
         Height          =   1785
         Left            =   3120
         TabIndex        =   28
         Top             =   120
         Width           =   2415
         Begin VB.TextBox txtItemConditionalName 
            Height          =   285
            Left            =   1200
            TabIndex        =   36
            Text            =   "Text1"
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox txtItemIntelligenceNames 
            Height          =   285
            Left            =   1200
            TabIndex        =   34
            Text            =   "Text1"
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox txtItemDBName 
            Height          =   285
            Left            =   1200
            TabIndex        =   32
            Text            =   "Text1"
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtItemFormName 
            Height          =   285
            Left            =   1200
            TabIndex        =   30
            Text            =   "Text1"
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblItemConditionalName 
            Alignment       =   1  'Right Justify
            Caption         =   "Conditional"
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   960
            Width           =   855
         End
         Begin VB.Label lblItemIntelligenceNames 
            Alignment       =   1  'Right Justify
            Caption         =   "Intelligence"
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label lblItemDBName 
            Alignment       =   1  'Right Justify
            Caption         =   "Database"
            Height          =   255
            Left            =   360
            TabIndex        =   33
            Top             =   600
            Width           =   735
         End
         Begin VB.Label lblItemFormName 
            Alignment       =   1  'Right Justify
            Caption         =   "Form"
            Height          =   255
            Left            =   720
            TabIndex        =   31
            Top             =   240
            Width           =   375
         End
      End
   End
   Begin VB.PictureBox picTelegram 
      BorderStyle     =   0  'None
      Height          =   2220
      Left            =   240
      ScaleHeight     =   2220
      ScaleWidth      =   5685
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   5685
      Begin VB.Frame fraTelegram 
         Caption         =   "Hard Limit"
         Height          =   1755
         Index           =   2
         Left            =   3840
         TabIndex        =   83
         ToolTipText     =   "Set the column where automatic line where new line is added"
         Top             =   480
         Width           =   1815
         Begin VB.ComboBox cmbTelegramSound 
            Height          =   315
            Index           =   2
            ItemData        =   "frmOptions.frx":00D3
            Left            =   120
            List            =   "frmOptions.frx":00E3
            Style           =   2  'Dropdown List
            TabIndex        =   96
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox txtTelegram 
            Height          =   285
            Index           =   2
            Left            =   840
            TabIndex        =   95
            Top             =   720
            Width           =   855
         End
         Begin VB.CheckBox chkTelegram 
            Caption         =   "Enabled"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   86
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblTelegram 
            Caption         =   "Column"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   89
            Top             =   720
            Width           =   615
         End
      End
      Begin VB.Frame fraTelegram 
         Caption         =   "Soft Limit"
         Height          =   1755
         Index           =   1
         Left            =   1920
         TabIndex        =   82
         ToolTipText     =   "Set the column where the automatic new line"
         Top             =   480
         Width           =   1815
         Begin VB.ComboBox cmbTelegramSound 
            Height          =   315
            Index           =   1
            ItemData        =   "frmOptions.frx":010B
            Left            =   120
            List            =   "frmOptions.frx":011B
            Style           =   2  'Dropdown List
            TabIndex        =   94
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox txtTelegram 
            Height          =   285
            Index           =   1
            Left            =   840
            TabIndex        =   93
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   2040
            TabIndex        =   91
            Top             =   120
            Width           =   615
         End
         Begin VB.CheckBox chkTelegram 
            Caption         =   "Enabled"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   85
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblTelegram 
            Caption         =   "Column"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   88
            Top             =   720
            Width           =   615
         End
      End
      Begin VB.ComboBox cmbTelegram 
         Height          =   315
         ItemData        =   "frmOptions.frx":0143
         Left            =   0
         List            =   "frmOptions.frx":0145
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   120
         Width           =   5655
      End
      Begin VB.Frame fraTelegram 
         Caption         =   "Warning"
         Height          =   1755
         Index           =   0
         Left            =   0
         TabIndex        =   80
         ToolTipText     =   "Set the column where the warning starts"
         Top             =   480
         Width           =   1815
         Begin VB.ComboBox cmbTelegramSound 
            Height          =   315
            Index           =   0
            ItemData        =   "frmOptions.frx":0147
            Left            =   120
            List            =   "frmOptions.frx":0149
            Style           =   2  'Dropdown List
            TabIndex        =   92
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox txtTelegram 
            Height          =   285
            Index           =   0
            Left            =   840
            TabIndex        =   90
            Top             =   720
            Width           =   855
         End
         Begin VB.CheckBox chkTelegram 
            Caption         =   "Enabled"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   84
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblTelegram 
            Caption         =   "Column"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   87
            Top             =   720
            Width           =   615
         End
      End
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   2220
      Left            =   240
      ScaleHeight     =   2220
      ScaleWidth      =   5685
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   5685
      Begin VB.Frame fraToolbar 
         Caption         =   "Toolbar Button Visiblity"
         Height          =   2220
         Left            =   0
         TabIndex        =   77
         Top             =   0
         Width           =   5655
         Begin VB.ListBox lstToolbar 
            Columns         =   2
            Height          =   1635
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   78
            Top             =   240
            Width           =   5414
         End
      End
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   2685
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4736
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   9
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Viewing"
            Key             =   "view"
            Object.ToolTipText     =   "Set Options for Viewing"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Updating"
            Key             =   "update"
            Object.ToolTipText     =   "Set Options for Updating"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Playing"
            Key             =   "playing"
            Object.ToolTipText     =   "Set Options for Playing"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Telegram"
            Key             =   "telegram"
            Object.ToolTipText     =   "Set Options for Telegram Window"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Item Names"
            Key             =   "item"
            Object.ToolTipText     =   "Set Options for Item Names"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Theme Games"
            Key             =   "theme"
            Object.ToolTipText     =   "Set Options for Theme Games"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Custom Buttons"
            Key             =   "custombuttons"
            Object.ToolTipText     =   "Set Options for Script Buttons"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Toolbar"
            Key             =   "toolbar"
            Object.ToolTipText     =   "Setting the Visiblity for the Toolbar buttons"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Speak"
            Key             =   "tts"
            Object.ToolTipText     =   "Text to Speech Options"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   2895
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   2895
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   2895
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'120203 rjk: Created when switched to frmOptions and boolean options
'120203 rjk: Switched to frmOptions and boolean options.
'121203 rjk: Added item options
'122903 rjk: Added Local Help option
'260104 rjk: Added Theme Game option (St@rW@rs)
'140204 rjk: Added ImageFileName to store the file name for Background Image
'150204 rjk: Added to store the border color for the map
'270304 rjk: Added Theme Game option for Retro
'300304 rjk: Added RolloverScaling
'080404 rjk: Added DefenseScaling Option
'220504 rjk: Added Show Flash Chat
'010704 rjk: Added FullDumpwithoutSea for deity option.
'240804 rjk: Added 2K4 Theme Game Option
'270804 rjk: Added Xdump option to allow use select of either xdump or dump
'270904 rjk: Added saving the display mode (DD_CURRENT, DD_NEW and DD_BOTH).
'151104 rjk: Added the ability to start with your Capital instead of 0,0
'190105 rjk: Set UTF8 in the server when the option is changed or loaded.
'100405 rjk: UTF8 encapsulated in server version check for 4.2.21 or newer.
'240405 rjk: Added Custom Script Buttons.  Added Toolbar visibility controls.
'100505 rjk: Switch Toolbar to listbox and custom buttons images.
'250505 rjk: Added Telegram Warning, Soft Limit and Hard Limit.
'260505 rjk: Switched UTF8 from an registry option to login option.
'290505 rjk: Added ability show Commodity Ratio and an option to control it.
'120705 rjk: Added option to the control display of Telegram Notification with INFORM OFF in the cmd window
'            Added option to the control display of Update Commands in the cmd window
'160705 rjk: Added option to getting of oil content from the sea.
'210705 rjk: Added option to the control the suppress of the unit damage reports.
'230705 rjk: Added EE9 Theme Game Option
'120905 rjk: Changed the default for Suppress Get Oil Content to true.
'171005 rjk: Fixed Scripts buttons crashing when there invalid image name.
'281205 rjk: Added South Pacific Theme Game Option for building loaded ships.
'190206 rjk: Remove xdump, use VersionCheck instead.
'220406 rjk: Added option to control the logging of flash conversions.
'050506 rjk: Added South Pacific Search for Atlantis Theme Game Option
'110606 rjk: Enhanced script buttons to support jump points.
'090706 rjk: Added Heavy Metal Theme Game Option
'170906 rjk: Added options for controlling TTS
'311206 rjk: Added short name options for unit grid.
'090607 rjk: Removed the RolloverAvail scaling by sector efficiency.
'301007 rjk: Added option the maximum path length when determining path costs.
'090311 kab: Modified the date parses to deal more regional effects.

Public bSound As Boolean
Public bToolbar As Boolean
Public bStatusBar As Boolean
Public bClearResultNewCommand As Boolean
Public bNoIron As Boolean
Public bSuppressBmapRefresh As Boolean
Public bShowOnlyWarships As Boolean
Public bDefaultRetreat As Boolean
Public bMaxProd As Boolean
Public bSuppressStrength As Boolean
Public playersTimeInterval As Integer '0 is disabled
Public bDisplayDumpHeaders As Boolean
Public friendlyUnit As enumFriendly '120303 rjk: Moved from unitView
Public bLocalHelp As Boolean  '122903 rjk: Added Help
Public bStarWars As Boolean
Public bRetro As Boolean '270304 rjk: Added for Retro game
Public b2K4 As Boolean '250804 rjk: Added for 2K4 game
Public bEE9 As Boolean '230705 rjk: Added EE9 Theme Game Option
Public bSP2005 As Boolean '281205 rjk: Added South Pacific Theme Game Option
Public bSPAtlantis As Boolean '050506 rjk: Added South Pacific Search for Atlantis Theme Game Option
Public bHeavyMetal As Boolean '090706 rjk: Added Heavy Metal Theme Game Option
Public strImageFileName As String '140204 rjk: Added to store the file name for Background Image
Public lBorderColor As Long '150204 rjk: Added to store the border color for the map
Public iDefenseScaling As Integer '080404 rjk: Added DefenseScaling Option
Public bFlashChat As Boolean '220504 rjk: Added Flash Chat
Public bFlashPlayerColors As Boolean '060604 rjk: Added Player colors for Flash Window
Public bFullDumpwithoutSea As Boolean '010704 rjk: Added for only non-sea sectors for Deities.
Public dspDesignationSave As enumDisplayDesignation '250904 rjk: Save the operator selection.
Public bCapital As Boolean '151104 rjk: Added the ability to start at your capital instead of 0,0
Public bCommodityRatio As Boolean '290505 rjk: Added to control whether Commodity Ratio is shown
Public bOilContentForSeaSectors As Boolean '110605 rjk: Added to control the Updates requests ocontent
'120705 rjk: Added to the control display of Telegram Notification with INFORM OFF in the cmd window
Public bSuppressTelegramNotification As Boolean
'120705 rjk: Added to the control display of Update Commands in the cmd window
Public bSuppressUpdateCommands As Boolean
'210705 rjk: Added to the control the suppress of the unit damage reports
Public bSuppressUnitDamageReports As Boolean
Public bFLashLogging As Boolean '220406 rjk: Added to control the logging of flash conversions.
'170906 rjk: Added options for controlling TTS
Public bShortNameUnitGrid As Boolean '311206 rjk: Added to display short names in the unit grid.
Public iMaximumPathLength As Integer '301007 rjk: Added to control the maximum path length when
'determining path costs.
Public bTTSReportWindow As Boolean
Public bTTSTelegrams As Boolean
Public bTTSBulletins As Boolean
Public bTTSFlashes As Boolean

Enum enumFriendly
    FRIEND_ALLIED
    FRIEND_NEUTRAL
End Enum

Private Sub cbItem_Click()
FillItemFrame Items.FindByLetter(Left$(cbItem.Text, 1))
End Sub

Private Sub chkTelegram_Click(Index As Integer)
If chkTelegram(Index).Value <> vbUnchecked Then
    txtTelegram(Index).Enabled = True
    cmbTelegramSound(Index).Enabled = True
Else
    txtTelegram(Index).Enabled = False
    cmbTelegramSound(Index).Enabled = False
End If
End Sub

Private Sub cmbScriptButton_Click()
FillScriptButtonFrame cmbScriptButton.ItemData(cmbScriptButton.ListIndex)
End Sub

Private Sub cmbTelegram_Click()
LoadTelegramFrames
End Sub

Private Sub cmdApply_Click()
ApplyOptions
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdImage_Click()
ChDir App.path

Load frmCommonDialog

frmCommonDialog.cdb.DialogTitle = "Image File for Background"
' Set CancelError is True
frmCommonDialog.cdb.CancelError = False
' Set flags
frmCommonDialog.cdb.Flags = cdlOFNFileMustExist
' Set filters
frmCommonDialog.cdb.Filter = "All Files (*.*)|*.*|Image Files (*.jpg;*.gif;*.bmp)|*.jpg;*.gif;*.bmp"
' Specify default filter
frmCommonDialog.cdb.FilterIndex = 2
' Display name of selected file
frmCommonDialog.cdb.FileName = strImageFileName

frmCommonDialog.cdb.InitDir = App.path

' Display the open dialog box
frmCommonDialog.cdb.ShowOpen

strImageFileName = frmCommonDialog.cdb.FileName
txtImageName = strImageFileName
bMapBitMapLoaded = False
Unload frmCommonDialog
End Sub

Private Sub cmdScriptBrowse_Click()
ChDir App.path

Load frmCommonDialog

frmCommonDialog.cdb.DialogTitle = "Script File for Custom Script Button"
' Set CancelError is True
frmCommonDialog.cdb.CancelError = False
' Set flags
frmCommonDialog.cdb.Flags = cdlOFNFileMustExist
' Set filters
frmCommonDialog.cdb.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt"
' Specify default filter
frmCommonDialog.cdb.FilterIndex = 2
' Display name of selected file
frmCommonDialog.cdb.FileName = txtScriptFileName.Text

frmCommonDialog.cdb.InitDir = App.path + "\Scripts"

' Display the open dialog box
frmCommonDialog.cdb.ShowOpen

txtScriptFileName.Text = frmCommonDialog.cdb.FileName

Unload frmCommonDialog
End Sub

Private Sub cmdScriptImageBrowse_Click()
ChDir App.path

Load frmCommonDialog

frmCommonDialog.cdb.DialogTitle = "Image File for Custom Script Button"
' Set CancelError is True
frmCommonDialog.cdb.CancelError = False
' Set flags
frmCommonDialog.cdb.Flags = cdlOFNFileMustExist
' Set filters
frmCommonDialog.cdb.Filter = "All Files (*.*)|*.*|Text Files (*.bmp)|*.bmp"
' Specify default filter
frmCommonDialog.cdb.FilterIndex = 2
' Display name of selected file
frmCommonDialog.cdb.FileName = txtScriptImageFileName.Text

frmCommonDialog.cdb.InitDir = App.path

' Display the open dialog box
frmCommonDialog.cdb.ShowOpen

txtScriptImageFileName.Text = frmCommonDialog.cdb.FileName

Unload frmCommonDialog
End Sub

Private Sub cmdOK_Click()
ApplyOptions
SaveOptions
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
'handle ctrl+tab to move to the next tab
If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
    i = tbsOptions.SelectedItem.Index
    If i = tbsOptions.Tabs.Count Then
        'last tab so we need to wrap to tab 1
        Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
    Else
        'increment the tab
        Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
    End If
End If
End Sub

Private Sub Form_Load()
Dim Index As Integer

'center the form
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

If frmEmpCmd.bAutoUpdate Then
    chkUpdate = vbChecked
Else
    chkUpdate = vbUnchecked
End If
If frmEmpCmd.bAutoRead Then
    chkRead = vbChecked
Else
    chkRead = vbUnchecked
End If
If bSound Then
    chkSound = vbChecked
Else
    chkSound = vbUnchecked
End If
If bToolbar Then
    chkToolbar = vbChecked
Else
    chkToolbar = vbUnchecked
End If
If bStatusBar Then
    chkStatusBar = vbChecked
Else
    chkStatusBar = vbUnchecked
End If
If bClearResultNewCommand Then
    chkClear = vbChecked
Else
    chkClear = vbUnchecked
End If
If bNoIron Then
    chkNoIron = vbChecked
Else
    chkNoIron = vbUnchecked
End If
If bSuppressBmapRefresh Then
    chkSuppressBmap = vbChecked
Else
    chkSuppressBmap = vbUnchecked
End If
If bShowOnlyWarships Then
    chkWarship = vbChecked
Else
    chkWarship = vbUnchecked
End If
If bDefaultRetreat Then
    chkRetreat = vbChecked
Else
    chkRetreat = vbUnchecked
End If
If bMaxProd Then
    chkMaxProd = vbChecked
Else
    chkMaxProd = vbUnchecked
End If
If bSuppressStrength Then
    chkSuppressStrength = vbChecked
Else
    chkSuppressStrength = vbUnchecked
End If
For Index = 0 To cmbPlayer.ListCount - 1
    If cmbPlayer.ItemData(Index) = playersTimeInterval Then
        cmbPlayer.ListIndex = Index
        Exit For
    End If
Next Index
If bDisplayDumpHeaders Then
    chkDumpHeaders = vbChecked
Else
    chkDumpHeaders = vbUnchecked
End If
If friendlyUnit = FRIEND_NEUTRAL Then
    optFriendlyNeutral = True
Else
    optFriendlyAlly = True
End If
If bLocalHelp Then
    chkLocalHelp = vbChecked
Else
    chkLocalHelp = vbUnchecked
End If
If bStarWars Then
    optThemeGames(1).Value = True
ElseIf bRetro Then
    optThemeGames(2).Value = True
ElseIf b2K4 Then
    optThemeGames(3).Value = True
ElseIf bEE9 Then
    optThemeGames(4).Value = True
ElseIf bSP2005 Then
    optThemeGames(5).Value = True
ElseIf bSPAtlantis Then
    optThemeGames(6).Value = True
ElseIf bHeavyMetal Then
    optThemeGames(7).Value = True
Else
    optThemeGames(0).Value = True
End If
txtImageName = strImageFileName
picBorderColor.BackColor = lBorderColor
txtDefenseScaling = CStr(iDefenseScaling)
If bFlashChat Then
    chkFlashChat = vbChecked
Else
    chkFlashChat = vbUnchecked
End If
If bFLashLogging Then
    chkFlashLogging = vbChecked
Else
    chkFlashLogging = vbUnchecked
End If
If bFlashPlayerColors Then
    chkFlashPlayerColors = vbChecked
Else
    chkFlashPlayerColors = vbUnchecked
End If
If bFullDumpwithoutSea Then
    chkFullDumpwithoutSea = vbChecked
Else
    chkFullDumpwithoutSea = vbUnchecked
End If
If bCapital Then
    chkCapital = vbChecked
Else
    chkCapital = vbUnchecked
End If
If bCommodityRatio Then
    chkCommodityRatio = vbChecked
Else
    chkCommodityRatio = vbUnchecked
End If
If bOilContentForSeaSectors Then
    chkSuppressOilContentAtSea = vbUnchecked
Else
    chkSuppressOilContentAtSea = vbChecked
End If
If bSuppressTelegramNotification Then
    chkSuppressTelegramNotification = vbChecked
Else
    chkSuppressTelegramNotification = vbUnchecked
End If
If bSuppressUpdateCommands Then
    chkSuppressUpdateCommands = vbChecked
Else
    chkSuppressUpdateCommands = vbUnchecked
End If
If bSuppressUnitDamageReports Then
    chkSuppressUnitDamageReports = vbChecked
Else
    chkSuppressUnitDamageReports = vbUnchecked
End If
If bTTSReportWindow Then
    chkTTSReportWindow = vbChecked
Else
    chkTTSReportWindow = vbUnchecked
End If
If bTTSTelegrams Then
    chkTTSTelegrams = vbChecked
Else
    chkTTSTelegrams = vbUnchecked
End If
If bTTSBulletins Then
    chkTTSBulletins = vbChecked
Else
    chkTTSBulletins = vbUnchecked
End If
If bTTSFlashes Then
    chkTTSFlashes = vbChecked
Else
    chkTTSFlashes = vbUnchecked
End If
If bShortNameUnitGrid Then
    chkShortName = vbChecked
Else
    chkShortName = vbUnchecked
End If
txtPathLength = CStr(iMaximumPathLength)

LoadTelegramPicture
LoadItemPicture
LoadToolbarPicture
LoadScriptButtonPicture
End Sub

Private Sub optCustomScript_Click()
UpdateCustomButtonDisplay False
End Sub

Private Sub optJumpPoint_Click()
UpdateCustomButtonDisplay True
End Sub

Private Sub picBorderColor_Click()
frmCommonDialog.cdb.Color = picBorderColor.BackColor
frmCommonDialog.cdb.ShowColor
picBorderColor.BackColor = frmCommonDialog.cdb.Color
End Sub

Private Sub tbsOptions_Click()
picView.Visible = False
picUpdate.Visible = False
picPlaying.Visible = False
picTelegram.Visible = False
picItem.Visible = False
picThemeGame.Visible = False
picScriptButton.Visible = False
picToolbar.Visible = False
picTTS.Visible = False
Select Case tbsOptions.SelectedItem.key
Case "view"
    picView.Visible = True
Case "update"
    picUpdate.Visible = True
Case "playing"
    picPlaying.Visible = True
Case "telegram"
    picTelegram.Visible = True
Case "item"
    picItem.Visible = True
Case "theme"
    picThemeGame.Visible = True
Case "custombuttons"
    picScriptButton.Visible = True
Case "toolbar"
    picToolbar.Visible = True
Case "tts"
    picTTS.Visible = True
End Select
End Sub

Private Sub ApplyOptions()
If Not CheckTelegramOptions Then
    Exit Sub
End If
'AutoUpdate
If chkUpdate <> vbUnchecked Then
    frmEmpCmd.bAutoUpdate = True
Else
    frmEmpCmd.bAutoUpdate = False
End If
frmDrawMap.UpdateAutoUpdate
'AutoRead
If chkRead <> vbUnchecked Then
    frmEmpCmd.bAutoRead = True
Else
    frmEmpCmd.bAutoRead = False
End If
frmDrawMap.UpdateAutoRead
'Use Sound
If chkSound <> vbUnchecked Then
    bSound = True
Else
    bSound = False
End If
'Show Toolbar
If chkToolbar <> vbUnchecked Then
    bToolbar = True
Else
    bToolbar = False
End If
'Show Status Bar
If chkStatusBar <> vbUnchecked Then
    bStatusBar = True
Else
    bStatusBar = False
End If
frmDrawMap.UpdateBars
'Clear Result Box on New Command
If chkClear <> vbUnchecked Then
    If Not bClearResultNewCommand Then
        bClearResultNewCommand = True
        frmDrawMap.lstCmdResult.Clear
    End If
Else
    bClearResultNewCommand = False
End If
'No Iron Game
If chkNoIron <> vbUnchecked Then
    If Not bNoIron Then
        bNoIron = True
        If rsNation.BOF Or rsNation.EOF Then
            rsNation.MoveFirst
            rsNation.Edit
            rsNation.Fields("NoIron").Value = True
            rsNation.Update
        End If
    End If
Else
    If bNoIron Then
        bNoIron = False
        If rsNation.BOF Or rsNation.EOF Then
            rsNation.MoveFirst
            rsNation.Edit
            rsNation.Fields("NoIron").Value = False
            rsNation.Update
        End If
    End If
End If
'Suppress Bmap Refresh
If chkSuppressBmap <> vbUnchecked Then
    If Not bSuppressBmapRefresh Then
        bSuppressBmapRefresh = True
        If rsNation.BOF Or rsNation.EOF Then
            rsNation.MoveFirst
            rsNation.Edit
            rsNation.Fields("SuppressBmap").Value = True
            rsNation.Update
        End If
    End If
Else
    If bSuppressBmapRefresh Then
        bSuppressBmapRefresh = False
        If rsNation.BOF Or rsNation.EOF Then
            rsNation.MoveFirst
            rsNation.Edit
            rsNation.Fields("SuppressBmap").Value = False
            rsNation.Update
        End If
    End If
End If
'Show Only Warships
If chkWarship <> vbUnchecked Then
    If Not bShowOnlyWarships Then
        bShowOnlyWarships = True
        frmDrawMap.DrawHexes
    End If
Else
    If bShowOnlyWarships Then
        bShowOnlyWarships = False
        frmDrawMap.DrawHexes
    End If
End If
'Set Retreat on Nav/March
If chkRetreat <> vbUnchecked Then
    bDefaultRetreat = True
Else
    bDefaultRetreat = False
End If
'Max. Prod instead of Pop.
If chkMaxProd <> vbUnchecked Then
    bMaxProd = True
Else
    bMaxProd = False
End If
'SuppressStrength
If chkSuppressStrength <> vbUnchecked Then
    If Not bSuppressStrength Then
        bSuppressStrength = True
        frmEmpCmd.SubmitEmpireCommand "db1", False
        GetCurrentStrength
        frmEmpCmd.SubmitEmpireCommand "db2", False
    End If
Else
    bSuppressStrength = False
End If
playersTimeInterval = cmbPlayer.ItemData(cmbPlayer.ListIndex)
frmDrawMap.SetPlayerTimeInterval
If chkDumpHeaders <> vbUnchecked Then
    bDisplayDumpHeaders = True
Else
    bDisplayDumpHeaders = False
End If
If optFriendlyNeutral Then
    friendlyUnit = FRIEND_NEUTRAL
Else
    friendlyUnit = FRIEND_ALLIED
End If
If chkLocalHelp <> vbUnchecked Then
    bLocalHelp = True
Else
    bLocalHelp = False
End If

bStarWars = False
bRetro = False
b2K4 = False
bEE9 = False
bSP2005 = False
bSPAtlantis = False
bHeavyMetal = False
If optThemeGames(1).Value Then
    bStarWars = True
ElseIf optThemeGames(2).Value Then
    bRetro = True
ElseIf optThemeGames(3).Value Then
    b2K4 = True
ElseIf optThemeGames(4).Value Then
    bEE9 = True
ElseIf optThemeGames(5).Value Then
    bSP2005 = True
ElseIf optThemeGames(6).Value Then
    bSPAtlantis = True
ElseIf optThemeGames(7).Value Then
    bHeavyMetal = True
End If

lBorderColor = picBorderColor.BackColor
bMapBitMapLoaded = False
iDefenseScaling = Val(txtDefenseScaling)
If chkFlashChat <> vbUnchecked Then
    bFlashChat = True
Else
    bFlashChat = False
End If
If chkFlashPlayerColors <> vbUnchecked Then
    bFlashPlayerColors = True
Else
    bFlashPlayerColors = False
End If
If chkFlashLogging <> vbUnchecked Then
    bFLashLogging = True
Else
    bFLashLogging = False
End If
If chkFullDumpwithoutSea <> vbUnchecked Then
    bFullDumpwithoutSea = True
Else
    bFullDumpwithoutSea = False
End If
If chkCapital <> vbUnchecked Then
    bCapital = True
Else
    bCapital = False
End If
If chkCommodityRatio <> vbUnchecked Then
    If Not bCommodityRatio Then
        bCommodityRatio = True
        frmDrawMap.FillSectorBox frmDrawMap.SelX, frmDrawMap.SelY
    End If
Else
    If bCommodityRatio Then
        bCommodityRatio = False
        frmDrawMap.FillSectorBox frmDrawMap.SelX, frmDrawMap.SelY
    End If
End If
If chkSuppressOilContentAtSea <> vbUnchecked Then
    bOilContentForSeaSectors = False
Else
    bOilContentForSeaSectors = True
End If
If chkSuppressTelegramNotification <> vbUnchecked Then
    bSuppressTelegramNotification = True
Else
    bSuppressTelegramNotification = False
End If
If chkSuppressUpdateCommands <> vbUnchecked Then
    bSuppressUpdateCommands = True
Else
    bSuppressUpdateCommands = False
End If
If chkSuppressUnitDamageReports <> vbUnchecked Then
    bSuppressUnitDamageReports = True
Else
    bSuppressUnitDamageReports = False
End If
If chkTTSReportWindow <> vbUnchecked Then
    bTTSReportWindow = True
Else
    bTTSReportWindow = False
End If
If chkTTSTelegrams <> vbUnchecked Then
    bTTSTelegrams = True
Else
    bTTSTelegrams = False
End If
If chkTTSBulletins <> vbUnchecked Then
    bTTSBulletins = True
Else
    bTTSBulletins = False
End If
If chkTTSFlashes <> vbUnchecked Then
    bTTSFlashes = True
Else
    bTTSFlashes = False
End If
If chkShortName <> vbUnchecked Then
    bShortNameUnitGrid = True
Else
    bShortNameUnitGrid = False
End If
iMaximumPathLength = CInt(txtPathLength)

frmChat.ControlFlashForm
StoreItemNames Items.FindByLetter(Left$(cbItem.Text, 1))
StoreScriptButtonFrame cmbScriptButton.ItemData(cmbScriptButton.ListIndex)
StoreToolbarFrame
SaveTelegramFrames
End Sub

Public Sub LoadOptions()
'AutoUpdate and AutoRead are set during initial connection setup.
frmDrawMap.UpdateAutoUpdate
frmDrawMap.UpdateAutoRead

bSound = GetSetting(APPNAME, "Options", "Use Sound", True)
bToolbar = GetSetting(APPNAME, "Options", "Show Toolbar", True)
bStatusBar = GetSetting(APPNAME, "Options", "Show StatusBar", True)
frmDrawMap.UpdateBars

bClearResultNewCommand = GetSetting(APPNAME, "Options", "Clear Result Box", True)
bDefaultRetreat = GetSetting(APPNAME, "Options", "Default Retreat", True)
bMaxProd = GetSetting(APPNAME, "Options", "Max Prod", True)
bSuppressStrength = GetSetting(APPNAME, "Options", "Suppress Strength", False)
playersTimeInterval = GetSetting(APPNAME, "Options", "Player Time Interval", 0)
InitializePlayers
frmDrawMap.SetPlayerTimeInterval

bDisplayDumpHeaders = GetSetting(APPNAME, "Options", "Display Dump Headers", False)
friendlyUnit = GetSetting(APPNAME, "Options", "Friendly", FRIEND_ALLIED)

If Not (rsNation.BOF And rsNation.EOF) Then
    rsNation.MoveFirst
    bNoIron = rsNation.Fields("NoIron").Value
    bSuppressBmapRefresh = rsNation.Fields("SuppressBmap").Value
Else
    bNoIron = False
    bSuppressBmapRefresh = False
End If
bLocalHelp = GetSetting(APPNAME, "Options", "Display Local Help", True)
bStarWars = GetSetting(APPNAME, "Options", "Star Wars", False)
bRetro = GetSetting(APPNAME, "Options", "Retro", False)
b2K4 = GetSetting(APPNAME, "Options", "2K4", False)
bEE9 = GetSetting(APPNAME, "Options", "EE9", False)
bSP2005 = GetSetting(APPNAME, "Options", "SP2005", False)
bSPAtlantis = GetSetting(APPNAME, "Options", "SPAtlantis", False)
bHeavyMetal = GetSetting(APPNAME, "Options", "HeavyMetal", False)
strImageFileName = GetSetting(APPNAME, "Options", "Map Image File Name", "")
lBorderColor = GetSetting(APPNAME, "Options", "Map Border Color", vbBlack)
iDefenseScaling = GetSetting(APPNAME, "Options", "Defense Scaling", 20)
bFlashChat = GetSetting(APPNAME, "Options", "Show Flash Chat", True)
bFlashPlayerColors = GetSetting(APPNAME, "Options", "Show Flash Player Colors", False)
bFLashLogging = GetSetting(APPNAME, "Options", "Flash Logging", False)
bFullDumpwithoutSea = GetSetting(APPNAME, "Options", "FullDump without Sea", False)
bCapital = GetSetting(APPNAME, "Options", "Capital", False)
bCommodityRatio = GetSetting(APPNAME, "Options", "CommodityRatio", True)
bOilContentForSeaSectors = GetSetting(APPNAME, "Options", "Oil Content Updates", True)
bSuppressTelegramNotification = GetSetting(APPNAME, "Options", "Suppress Telegram Notifications", False)
bSuppressUpdateCommands = GetSetting(APPNAME, "Options", "Suppress Update Commands", False)
dspDesignationSave = GetSetting(APPNAME, "Options", "Display Designation", DD_BOTH)
bSuppressUnitDamageReports = GetSetting(APPNAME, "Options", "Suppress Unit Damage Reports", False)
bTTSReportWindow = GetSetting(APPNAME, "Options", "TTS Report Window", False)
bTTSTelegrams = GetSetting(APPNAME, "Options", "TTS Telegrams", False)
bTTSBulletins = GetSetting(APPNAME, "Options", "TTS Bulletins", False)
bTTSFlashes = GetSetting(APPNAME, "Options", "TTS Flashes", False)
bShortNameUnitGrid = GetSetting(APPNAME, "Options", "Short Name Unit Grid", False)
iMaximumPathLength = GetSetting(APPNAME, "Options", "Maximum Path Length", 20)

Select Case dspDesignationSave
Case DD_CURRENT
    dspDesignation = DD_CURRENT
Case DD_NEW
    dspDesignation = DD_NEW
Case DD_BOTH, DD_BOTH_CURRENT, DD_BOTH_NEW
    dspDesignation = DD_BOTH
End Select
LoadScriptButtons
LoadToolbarVisibility
LoadTelegramOptions
End Sub

Public Sub SaveOptions()
'AutoUpdate and AutoRead are saved during initial connection setup.
SaveSetting APPNAME, "Options", "Use Sound", bSound
SaveSetting APPNAME, "Options", "Show Toolbar", bToolbar
SaveSetting APPNAME, "Options", "Show StatusBar", bStatusBar
SaveSetting APPNAME, "Options", "Clear Result Box", bClearResultNewCommand
SaveSetting APPNAME, "Options", "Default Retreat", bDefaultRetreat
SaveSetting APPNAME, "Options", "Max Prod", bMaxProd
SaveSetting APPNAME, "Options", "Suppress Strength", bSuppressStrength
SaveSetting APPNAME, "Options", "Player Time Interval", playersTimeInterval
SaveSetting APPNAME, "Options", "Display Dump Headers", bDisplayDumpHeaders
SaveSetting APPNAME, "Options", "Friendly", friendlyUnit
    
If Not (rsNation.BOF And rsNation.EOF) Then
    rsNation.MoveFirst
    rsNation.Edit
    rsNation.Fields("NoIron").Value = bNoIron
    rsNation.Fields("SuppressBmap").Value = bSuppressBmapRefresh
    rsNation.Update
End If
SaveSetting APPNAME, "Options", "Display Local Help", bLocalHelp
SaveSetting APPNAME, "Options", "Star Wars", bStarWars
SaveSetting APPNAME, "Options", "Retro", bRetro
SaveSetting APPNAME, "Options", "2K4", b2K4
SaveSetting APPNAME, "Options", "EE9", bEE9
SaveSetting APPNAME, "Options", "SP2005", bSP2005
SaveSetting APPNAME, "Options", "SPAtlantis", bSPAtlantis
SaveSetting APPNAME, "Options", "HeavyMetal", bHeavyMetal
SaveSetting APPNAME, "Options", "Map Image File Name", strImageFileName
SaveSetting APPNAME, "Options", "Map Border Color", lBorderColor
SaveSetting APPNAME, "Options", "Defense Scaling", iDefenseScaling
SaveSetting APPNAME, "Options", "Show Flash Chat", bFlashChat
SaveSetting APPNAME, "Options", "Show Flash Player Colors", bFlashPlayerColors
SaveSetting APPNAME, "Options", "Flash Logging", bFLashLogging
SaveSetting APPNAME, "Options", "FullDump without Sea", bFullDumpwithoutSea
SaveSetting APPNAME, "Options", "Capital", bCapital
SaveSetting APPNAME, "Options", "CommodityRatio", bCommodityRatio
SaveSetting APPNAME, "Options", "Oil Content Updates", bOilContentForSeaSectors
SaveSetting APPNAME, "Options", "Suppress Telegram Notifications", bSuppressTelegramNotification
SaveSetting APPNAME, "Options", "Suppress Update Commands", bSuppressUpdateCommands
SaveSetting APPNAME, "Options", "Display Designation", dspDesignationSave
SaveSetting APPNAME, "Options", "Suppress Unit Damage Reports", bSuppressUnitDamageReports
SaveSetting APPNAME, "Options", "TTS Report Window", bTTSReportWindow
SaveSetting APPNAME, "Options", "TTS Telegrams", bTTSTelegrams
SaveSetting APPNAME, "Options", "TTS Bulletins", bTTSBulletins
SaveSetting APPNAME, "Options", "TTS Flashes", bTTSFlashes
SaveSetting APPNAME, "Options", "Short Name Unit Grid", bShortNameUnitGrid
SaveSetting APPNAME, "Options", "Maximum Path Length", iMaximumPathLength

SaveScriptButtons
SaveToolbarVisibility
SaveTelegramOptions
End Sub

Private Sub LoadItemPicture()
Dim Index As Integer

cbItem.Clear
For Index = 1 To Items.Count
    cbItem.AddItem Items(Index).Letter + " - " + Items(Index).ItemName
Next Index
cbItem.ListIndex = 0
FillItemFrame Items.FindByLetter(Left$(cbItem.Text, 1))
End Sub

Private Sub FillItemFrame(eiItem As EmpItem)
If eiItem Is Nothing Then
    txtItemFormName = ""
    txtItemSectorPanelLabel = ""
    txtItemDistributionPanelLabel = ""
    txtItemFormLabel = ""
    txtItemDBName = ""
    txtItemIntelligenceNames = ""
    txtItemConditionalName = ""
Else
    txtItemFormName = eiItem.FormName
    txtItemSectorPanelLabel = eiItem.SectorPanelLabel
    txtItemDistributionPanelLabel = eiItem.DistributionPanelLabel
    txtItemFormLabel = eiItem.FormLabel
    txtItemDBName = eiItem.DatabaseName
    txtItemIntelligenceNames = eiItem.IntelligenceNames
    txtItemConditionalName = eiItem.ConditionalName
End If
End Sub

Private Sub StoreItemNames(eiItem As EmpItem)
If Not (eiItem Is Nothing) Then
    eiItem.FormName = txtItemFormName
    eiItem.SectorPanelLabel = txtItemSectorPanelLabel
    eiItem.DistributionPanelLabel = txtItemDistributionPanelLabel
    eiItem.FormLabel = txtItemFormLabel
    eiItem.DatabaseName = txtItemDBName
    eiItem.IntelligenceNames = txtItemIntelligenceNames
    eiItem.ConditionalName = txtItemConditionalName
End If
eiItem.SaveItem
End Sub

Private Sub txtDefenseScaling_Change()
Dim iDefScaling As Integer

iDefScaling = Val(txtDefenseScaling)
If iDefScaling < 0 Then
    iDefScaling = 0
ElseIf iDefScaling > 500 Then
    iDefScaling = 500
End If
txtDefenseScaling = CStr(txtDefenseScaling)
End Sub

Private Sub txtImageName_Change()
strImageFileName = txtImageName
bMapBitMapLoaded = False
End Sub

Private Sub LoadScriptButtonPicture()
Dim iIndex As Integer

cmbScriptButton.Clear
For iIndex = 1 To NUMBER_SCRIPT_BUTTONS
    cmbScriptButton.AddItem frmDrawMap.tbMain.Buttons(28 + iIndex).ToolTipText
    cmbScriptButton.ItemData(cmbScriptButton.ListCount - 1) = iIndex
Next iIndex
cmbScriptButton.ListIndex = 0
FillScriptButtonFrame 1
End Sub

Private Sub FillScriptButtonFrame(iIndex As Integer)
txtScriptFileName.Text = tScriptButtons(iIndex).strFileName
If tScriptButtons(iIndex).bSendOutputReportWindow Then
    chkScriptReport.Value = vbChecked
Else
    chkScriptReport.Value = vbUnchecked
End If
txtScriptImageFileName.Text = tScriptButtons(iIndex).strImageFileName
cmbJumpPoint.ListIndex = tScriptButtons(iIndex).iJumpPoint

If tScriptButtons(iIndex).bJumpPoint Then
    If tScriptButtons(iIndex).iJumpPoint >= 0 Then
        frmDrawMap.tbMain.Buttons(28 + iIndex).Visible = True
    Else
        frmDrawMap.tbMain.Buttons(28 + iIndex).Visible = False
    End If
    optJumpPoint.Value = vbChecked
    UpdateCustomButtonDisplay True
Else
    If Len(tScriptButtons(iIndex).strFileName) > 0 Then
        frmDrawMap.tbMain.Buttons(28 + iIndex).Visible = True
    Else
        frmDrawMap.tbMain.Buttons(28 + iIndex).Visible = False
    End If
    UpdateCustomButtonDisplay False
    optCustomScript.Value = vbChecked
End If

End Sub

Private Sub StoreScriptButtonFrame(iIndex As Integer)
tScriptButtons(iIndex).strFileName = txtScriptFileName.Text
UpdateToolTipText iIndex
If chkScriptReport.Value <> vbUnchecked Then
    tScriptButtons(iIndex).bSendOutputReportWindow = True
Else
    tScriptButtons(iIndex).bSendOutputReportWindow = False
End If
If tScriptButtons(iIndex).strImageFileName <> txtScriptImageFileName.Text Then
    If Len(txtScriptImageFileName.Text) = 0 Then
        frmDrawMap.tbMain.Buttons(28 + iIndex).Image = 3
    Else
        On Error GoTo Error_Exit
        frmDrawMap.ImageTB2.ListImages.Add , , LoadPicture(txtScriptImageFileName.Text)
        frmDrawMap.tbMain.Buttons(28 + iIndex).Image = frmDrawMap.ImageTB2.ListImages.Count
    End If
    tScriptButtons(iIndex).strImageFileName = txtScriptImageFileName.Text
End If
tScriptButtons(iIndex).iJumpPoint = cmbJumpPoint.ListIndex
If optJumpPoint.Value = True Then
    If tScriptButtons(iIndex).iJumpPoint >= 0 Then
        frmDrawMap.tbMain.Buttons(28 + iIndex).Visible = True
    Else
        frmDrawMap.tbMain.Buttons(28 + iIndex).Visible = False
    End If
    tScriptButtons(iIndex).bJumpPoint = True
Else
    If Len(tScriptButtons(iIndex).strFileName) > 0 Then
        frmDrawMap.tbMain.Buttons(28 + iIndex).Visible = True
    Else
        frmDrawMap.tbMain.Buttons(28 + iIndex).Visible = False
    End If
    tScriptButtons(iIndex).bJumpPoint = False
End If
UpdateToolTipText iIndex
cmbScriptButton.List(iIndex - 1) = frmDrawMap.tbMain.Buttons(28 + iIndex).ToolTipText
cmbScriptButton.ListIndex = iIndex - 1
Exit Sub

Error_Exit:
EmpireError "Failed to load image for custom script button", CStr(iIndex), txtScriptImageFileName.Text
End Sub

Private Sub SaveScriptButtons()
Dim iIndex As Integer

For iIndex = 1 To NUMBER_SCRIPT_BUTTONS
    SaveSetting APPNAME, "Options", "Script Button " + CStr(iIndex) + " File Name", _
        tScriptButtons(iIndex).strFileName
    SaveSetting APPNAME, "Options", "Script Button " + CStr(iIndex) + " Output Report Window", _
        tScriptButtons(iIndex).bSendOutputReportWindow
    SaveSetting APPNAME, "Options", "Script Button " + CStr(iIndex) + " Image File Name", _
        tScriptButtons(iIndex).strImageFileName
    SaveSetting APPNAME, "Options", "Script Button " + CStr(iIndex) + " Jump Point", _
        tScriptButtons(iIndex).bJumpPoint
    SaveSetting APPNAME, "Options", "Script Button " + CStr(iIndex) + " Jump Point Index", _
        tScriptButtons(iIndex).iJumpPoint
Next iIndex
End Sub

Private Sub LoadScriptButtons()
Dim iIndex As Integer

For iIndex = 1 To NUMBER_SCRIPT_BUTTONS
    tScriptButtons(iIndex).strFileName = GetSetting(APPNAME, "Options", _
        "Script Button " + CStr(iIndex) + " File Name", "")
    tScriptButtons(iIndex).bSendOutputReportWindow = GetSetting(APPNAME, "Options", _
        "Script Button " + CStr(iIndex) + " Output Report Window", True)
    tScriptButtons(iIndex).strImageFileName = GetSetting(APPNAME, "Options", _
        "Script Button " + CStr(iIndex) + " Image File Name", "")
    tScriptButtons(iIndex).bJumpPoint = GetSetting(APPNAME, "Options", _
        "Script Button " + CStr(iIndex) + " Jump Point", False)
    tScriptButtons(iIndex).iJumpPoint = GetSetting(APPNAME, "Options", _
        "Script Button " + CStr(iIndex) + " Jump Point Index", -1)

    If Len(tScriptButtons(iIndex).strImageFileName) > 0 Then
        On Error GoTo skip_image
        frmDrawMap.ImageTB2.ListImages.Add , , LoadPicture(tScriptButtons(iIndex).strImageFileName)
        frmDrawMap.tbMain.Buttons.Add 28 + iIndex, "Script" + CStr(iIndex), , _
            tbrDefault, frmDrawMap.ImageTB2.ListImages.Count
    Else
skip_image:
        frmDrawMap.tbMain.Buttons.Add (28 + iIndex), "Script" + CStr(iIndex), , _
            tbrDefault, 3
    End If
    UpdateToolTipText iIndex
    If tScriptButtons(iIndex).bJumpPoint Then
        If tScriptButtons(iIndex).iJumpPoint >= 0 Then
            frmDrawMap.tbMain.Buttons(28 + iIndex).Visible = True
        Else
            frmDrawMap.tbMain.Buttons(28 + iIndex).Visible = False
        End If
    Else
        If Len(tScriptButtons(iIndex).strFileName) > 0 Then
            frmDrawMap.tbMain.Buttons(28 + iIndex).Visible = True
        Else
            frmDrawMap.tbMain.Buttons(28 + iIndex).Visible = False
        End If
    End If
Next iIndex
End Sub

Private Sub UpdateToolTipText(iIndex As Integer)
frmDrawMap.tbMain.Buttons(28 + iIndex).ToolTipText = GetCustomButtonDescription(iIndex)
End Sub

Private Function GetCustomButtonDescription(iIndex) As String
Dim iPos As Integer

If tScriptButtons(iIndex).bJumpPoint Then
    If tScriptButtons(iIndex).iJumpPoint >= 0 Then
        GetCustomButtonDescription = "Jump Point " + CStr(tScriptButtons(iIndex).iJumpPoint + 1)
    Else
        GetCustomButtonDescription = "Jump Point"
    End If
Else
    iPos = InStrRev(tScriptButtons(iIndex).strFileName, "\")
    If iPos > 0 Then
        GetCustomButtonDescription = "Custom Script " + _
            " - " + Mid$(tScriptButtons(iIndex).strFileName, iPos + 1)
    Else
        GetCustomButtonDescription = "Custom Script " + _
            " - " + tScriptButtons(iIndex).strFileName
    End If
End If
End Function

Private Sub UpdateCustomButtonDisplay(bJumpPoint As Boolean)
If bJumpPoint Then
    lblScriptFileName = "Jump Point"
    txtScriptFileName.Visible = False
    chkScriptReport.Visible = False
    cmdScriptBrowse.Visible = False
    cmbJumpPoint.Visible = True
    fraScriptButton.Caption = "Jump Point Button"
Else
    lblScriptFileName = "Script File Name"
    txtScriptFileName.Visible = True
    chkScriptReport.Visible = True
    cmdScriptBrowse.Visible = True
    cmbJumpPoint.Visible = False
    fraScriptButton.Caption = "Custom Script Button"
End If
End Sub

Private Sub LoadToolbarPicture()
Dim iIndex As Integer

lstToolbar.Clear
For iIndex = 1 To 28
    If frmDrawMap.tbMain.Buttons.Item(iIndex).Style = tbrDefault Then
        lstToolbar.AddItem frmDrawMap.tbMain.Buttons.Item(iIndex).ToolTipText
        lstToolbar.ItemData(lstToolbar.NewIndex) = iIndex
        If frmDrawMap.tbMain.Buttons.Item(iIndex).Visible Then
            lstToolbar.Selected(lstToolbar.NewIndex) = True
        Else
            lstToolbar.Selected(lstToolbar.NewIndex) = False
        End If
    End If
Next iIndex
End Sub

Private Sub StoreToolbarFrame()
Dim iIndex As Integer

For iIndex = 0 To lstToolbar.ListCount - 1
    If lstToolbar.Selected(iIndex) Then
        frmDrawMap.tbMain.Buttons.Item(lstToolbar.ItemData(iIndex)).Visible = True
    Else
        frmDrawMap.tbMain.Buttons.Item(lstToolbar.ItemData(iIndex)).Visible = False
    End If
Next iIndex
End Sub

Private Sub LoadToolbarVisibility()
Dim iIndex As Integer
For iIndex = 1 To 28
    frmDrawMap.tbMain.Buttons.Item(iIndex).Visible = GetSetting(APPNAME, "Options", _
        "Toolbar " + CStr(iIndex) + " Visibility", True)
Next iIndex
End Sub

Private Sub SaveToolbarVisibility()
Dim iIndex As Integer

For iIndex = 1 To 28
    SaveSetting APPNAME, "Options", "Toolbar " + CStr(iIndex) + " Visibility", _
        frmDrawMap.tbMain.Buttons.Item(iIndex).Visible
Next iIndex
End Sub

Private Sub LoadTelegramOptions()
Dim iIndex As Integer
Dim iIndexFrame As Integer

For iIndex = totTelegram To totMOTD
    For iIndexFrame = tlWarning To tlHardLimit
        tTelegramOptions(iIndex, iIndexFrame).bEnabled = GetSetting(APPNAME, "Options", _
            "Telegram " + CStr(iIndex) + " " + CStr(iIndexFrame) + " Enabled", False)
        tTelegramOptions(iIndex, iIndexFrame).iColumn = GetSetting(APPNAME, "Options", _
            "Telegram " + CStr(iIndex) + " " + CStr(iIndexFrame) + " Column", 0)
        tTelegramOptions(iIndex, iIndexFrame).eTelegramSound = GetSetting(APPNAME, "Options", _
            "Telegram " + CStr(iIndex) + " " + CStr(iIndexFrame) + " Sound", tsNoSound)
    Next iIndexFrame
Next iIndex
End Sub

Private Sub LoadTelegramPicture()
Dim iIndex As Integer
Dim iIndexSound As Integer

cmbTelegram.Clear
For iIndex = totTelegram To totMOTD
    Select Case iIndex
    Case totTelegram:
        cmbTelegram.AddItem "Telegram"
        cmbTelegram.ItemData(cmbTelegram.NewIndex) = totTelegram
    Case totAnnouncement:
        cmbTelegram.AddItem "Announcement"
        cmbTelegram.ItemData(cmbTelegram.NewIndex) = totAnnouncement
    Case totFlash:
        cmbTelegram.AddItem "Flash"
        cmbTelegram.ItemData(cmbTelegram.NewIndex) = totFlash
    Case totMOTD:
        If bDeity Then
            cmbTelegram.AddItem "Message of the Day (motd)"
            cmbTelegram.ItemData(cmbTelegram.NewIndex) = totMOTD
        End If
    End Select
Next iIndex
For iIndex = tlWarning To tlHardLimit
    cmbTelegramSound(iIndex).Clear
    For iIndexSound = tsNoSound To ts3Beeps
        Select Case iIndexSound
        Case tsNoSound:
            cmbTelegramSound(iIndex).AddItem "No Sound"
        Case ts1Beep:
            cmbTelegramSound(iIndex).AddItem "1 Beep"
        Case ts2Beeps:
            cmbTelegramSound(iIndex).AddItem "2 Beeps"
        Case ts3Beeps:
            cmbTelegramSound(iIndex).AddItem "3 Beeps"
        End Select
        cmbTelegramSound(iIndex).ItemData(cmbTelegramSound(iIndex).NewIndex) = iIndexSound
    Next iIndexSound
Next iIndex
cmbTelegram.ListIndex = 0
LoadTelegramFrames
End Sub

Private Sub LoadTelegramFrames()
Dim iIndex As Integer

For iIndex = tlWarning To tlHardLimit
    If tTelegramOptions(cmbTelegram.ItemData(cmbTelegram.ListIndex), iIndex).bEnabled Then
        chkTelegram(iIndex).Value = vbChecked
        txtTelegram(iIndex).Enabled = True
        cmbTelegramSound(iIndex).Enabled = True
    Else
        chkTelegram(iIndex).Value = vbUnchecked
        txtTelegram(iIndex).Enabled = False
        cmbTelegramSound(iIndex).Enabled = False
    End If
    txtTelegram(iIndex).Text = CStr( _
        tTelegramOptions(cmbTelegram.ItemData(cmbTelegram.ListIndex), iIndex).iColumn)
    cmbTelegramSound(iIndex).ListIndex = _
        tTelegramOptions(cmbTelegram.ItemData(cmbTelegram.ListIndex), iIndex).eTelegramSound
Next iIndex
End Sub

Private Sub SaveTelegramFrames()
Dim iIndex As Integer

For iIndex = tlWarning To tlHardLimit
    If chkTelegram(iIndex).Value <> vbUnchecked Then
        tTelegramOptions(cmbTelegram.ItemData(cmbTelegram.ListIndex), iIndex).bEnabled = True
    Else
        tTelegramOptions(cmbTelegram.ItemData(cmbTelegram.ListIndex), iIndex).bEnabled = False
    End If
    tTelegramOptions(cmbTelegram.ItemData(cmbTelegram.ListIndex), iIndex).iColumn = _
        CInt(txtTelegram(iIndex).Text)
    tTelegramOptions(cmbTelegram.ItemData(cmbTelegram.ListIndex), iIndex).eTelegramSound = _
        cmbTelegramSound(iIndex).ListIndex
Next iIndex
End Sub

Private Sub SaveTelegramOptions()
Dim iIndex As Integer
Dim iIndexFrame As Integer

For iIndex = totTelegram To totMOTD
    For iIndexFrame = tlWarning To tlHardLimit
        SaveSetting APPNAME, "Options", "Telegram " + CStr(iIndex) + " " + CStr(iIndexFrame) + " Enabled", _
            tTelegramOptions(iIndex, iIndexFrame).bEnabled
        SaveSetting APPNAME, "Options", "Telegram " + CStr(iIndex) + " " + CStr(iIndexFrame) + " Column", _
            tTelegramOptions(iIndex, iIndexFrame).iColumn
        SaveSetting APPNAME, "Options", "Telegram " + CStr(iIndex) + " " + CStr(iIndexFrame) + " Sound", _
            tTelegramOptions(iIndex, iIndexFrame).eTelegramSound
    Next iIndexFrame
Next iIndex
End Sub

Private Function CheckTelegramOptions() As Boolean
Dim iPrevious As Integer
Dim iIndex As Integer

On Error GoTo Telegram_Error
iPrevious = 0
For iIndex = tlWarning To tlHardLimit
    If chkTelegram(iIndex).Value <> vbUnchecked Then
        If iPrevious >= CInt(txtTelegram(iIndex).Text) Then
            MsgBox "Values for Enabled Telegram Columns must be of increasing value, 0 < Warning < Soft Limit < Hard Limit", vbOKOnly
            CheckTelegramOptions = False
            Exit Function
        End If
        iPrevious = CInt(txtTelegram(iIndex).Text)
    End If
Next iIndex
CheckTelegramOptions = True
Exit Function

Telegram_Error:
MsgBox "Enabled Columns must be Numeric", vbOKOnly
CheckTelegramOptions = False
End Function

