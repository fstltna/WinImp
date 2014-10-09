VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTelegram 
   Caption         =   "Telegrams"
   ClientHeight    =   6615
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8595
   Icon            =   "frmTelegram.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   4560
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTelegram.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTelegram.frx":075C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTelegram.frx":0A76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTelegram.frx":0D90
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTelegram.frx":10AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTelegram.frx":13C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTelegram.frx":16DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTelegram.frx":19F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTelegram.frx":1D12
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTelegram.frx":2164
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   960
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   1693
      ButtonWidth     =   1773
      ButtonHeight    =   1535
      Appearance      =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "First"
            Key             =   "First"
            Object.ToolTipText     =   "Oldest"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Previous"
            Key             =   "Previous"
            Object.ToolTipText     =   "Previous"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Next"
            Key             =   "Next"
            Object.ToolTipText     =   "Next"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Last"
            Key             =   "Last"
            Object.ToolTipText     =   "Newest"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "New"
            Object.ToolTipText     =   "New Telegram"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reply"
            Key             =   "Reply"
            Object.ToolTipText     =   "Reply"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Forward"
            Key             =   "Forward"
            Object.ToolTipText     =   "Forward"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Import"
            Key             =   "Import"
            Object.ToolTipText     =   "Import Intelligence"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Send"
            Key             =   "Send"
            Object.ToolTipText     =   "Send Telegram"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvTreeView 
      Height          =   4680
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   2016
      _ExtentX        =   3545
      _ExtentY        =   8255
      _Version        =   393217
      Style           =   4
      Appearance      =   1
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
   Begin VB.Timer Timer1 
      Left            =   5280
      Top             =   5880
   End
   Begin VB.ListBox lbNation 
      Height          =   1815
      Left            =   5640
      MultiSelect     =   1  'Simple
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   840
      Width           =   2535
   End
   Begin VB.CheckBox chkSave 
      Caption         =   "Save a copy of sent mail"
      Height          =   495
      Left            =   5640
      TabIndex        =   9
      Top             =   3960
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.CheckBox chkInclude 
      Caption         =   "Include text from received telegram in your reply"
      Height          =   615
      Left            =   5640
      TabIndex        =   8
      Top             =   3360
      Width           =   2535
   End
   Begin RichTextLib.RichTextBox txtBody 
      Height          =   4695
      Left            =   2040
      TabIndex        =   6
      Top             =   1080
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   8281
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmTelegram.frx":25B6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   5400
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   72
   End
   Begin VB.PictureBox picTitles 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   8595
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Width           =   8595
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ListView:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   2040
         TabIndex        =   3
         Tag             =   " ListView:"
         Top             =   0
         Width           =   3210
      End
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Received Telegrams"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Tag             =   " TreeView:"
         Top             =   0
         Width           =   2010
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   2280
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   7
      Top             =   6345
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7391
            MinWidth        =   3528
            Text            =   "xx Telegrams"
            TextSave        =   "xx Telegrams"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3545
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3545
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
   Begin VB.Frame frameSend 
      Caption         =   "What to Send"
      Height          =   1815
      Left            =   5640
      TabIndex        =   10
      Top             =   4440
      Width           =   2175
      Begin VB.OptionButton optMotd 
         Caption         =   "Motd"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   1440
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Flash"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Announcement"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Telegram"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Label labNation 
      Alignment       =   2  'Center
      Caption         =   "<< ANNOUNCEMENT >>"
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
      Left            =   5640
      TabIndex        =   14
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Image imgSplitter 
      Height          =   4788
      Left            =   1965
      MousePointer    =   9  'Size W E
      Top             =   705
      Width           =   150
   End
End
Attribute VB_Name = "frmTelegram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'08xx03 efj: removed dead variables
'093003 rjk: general reformatting and cleanup
'101503 rjk: Fixed so telegram does not delete blank lines.
'103003 rjk: Added Import Intelligence functionality
'110303 rjk: Modified ButtonSend to account for the server not supporting up 1024 characters,
'            it only supports 1022 characters in a telegram line.
'            Fixed to support lines longer than 2044.
'111403 rjk: Clear Telegram window after removing all the telegrams.
'111803 rjk: Switched to a Yes/No question box for clearing telegrams.
'112903 rjk: Remove the question on clearing the telegrams as there is a separate form now.
'270303 rjk: Switched over to DeleteAllRecords for clearing tables.
'110804 rjk: Added the motd to telegram window for deity mode.
'290804 rjk: Always save telegrams, unless it is a duplicate.
'250505 rjk: Added Telegram Warning, Soft Limit and Hard Limit.
'180206 rjk: Removed "The deity" from the telegram list for 4.3.0.  The 4.3.0 country list now
'            contains the deities in it.
'181006 rjk: Added Import Offset to ImportIntelligence.
'090607 rjk: Added forward ability for telegrams.
'010108 rjk: Fixed incorrect telegram destinations.

Const NAME_COLUMN = 0
Const TYPE_COLUMN = 1
Const SIZE_COLUMN = 2
Const DATE_COLUMN = 3

Const BUTTON_FIRST = 1
Const BUTTON_BACK = 2
Const BUTTON_NEXT = 3
Const BUTTON_LATEST = 4
'Separator
Const BUTTON_NEW = 6
Const BUTTON_REPLY = 7
Const BUTTON_FORWARD = 8
Const BUTTON_IMPORT = 9 '103003 rjk: Added Import Intelligence functionality
'Separator
Const BUTTON_DELETE = 11
Const BUTTON_SEND = 12

Public TreeDirty As Boolean

Dim NewTelegram As Boolean
Dim SelectedTelegram As Integer
' Dim hldNation As String    8/2003 efj  removed
Dim CurTele As Integer

Dim mbMoving As Boolean
Const sglSplitLimit = 500

Enum eMessageCreate
    MESS_NEW
    MESS_REPLY
    MESS_FORWARD
End Enum
Dim bForwardMessage As Boolean

Private Sub chkInclude_Click()
Dim hldtxt As String
Dim strBuffer As String
Dim n As Integer
Dim strPlayerName As String

If chkInclude.Value = vbChecked Then
    'copy text of letter into the top of telegram
    rsTeleBody.Seek "=", SelectedTelegram, 1
    
    'Load telegram body into text box
    hldtxt = txtBody.Text
    txtBody.Text = vbNullString
    If Not rsTeleBody.NoMatch Then
        If bForwardMessage Then
            txtBody.Text = txtBody.Text + "}" + vbLf
            strPlayerName = Nations.NationName(rsTeleHead.Fields("From"))
            If Len(strPlayerName) = 0 Then
                strPlayerName = "Nation #" + CStr(rsTeleHead.Fields("From"))
            End If
            txtBody.Text = txtBody.Text + "}From: " + strPlayerName + vbLf + "}" + vbLf
        End If
        Do While (Not rsTeleBody.EOF)
            If rsTeleBody.Fields("ID") <> SelectedTelegram Then
                Exit Do
            End If
            txtBody.Text = txtBody.Text + "}" + rsTeleBody.Fields("Body") + vbLf
            rsTeleBody.MoveNext
        Loop
    End If
    txtBody.Text = txtBody.Text + vbLf + hldtxt
Else
    'remove all the copied lines in the text
    strBuffer = txtBody.Text
    txtBody.Text = vbNullString
    While Len(strBuffer) > 0
        'Buffer could contain several lines, so seperate at line feeds
        n = InStr(strBuffer, vbLf)
        If n = 0 Then
            hldtxt = strBuffer
            strBuffer = vbNullString
        Else
            hldtxt = Left$(strBuffer, n - 1)
            If Len(strBuffer) = n Then
                strBuffer = vbNullString
            Else
                strBuffer = Mid$(strBuffer, n + 1)
            End If
        End If
        If Left$(hldtxt, 1) <> "}" Then
            txtBody.Text = txtBody.Text + hldtxt + vbLf
        End If
    Wend
End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
If TreeDirty And Not NewTelegram Then
    LoadTree
    'Set selection to last item viewed
    If Not (rsTeleHead.BOF And rsTeleHead.EOF) Then
        rsTeleHead.MoveLast
        CurTele = rsTeleHead.Fields("ID")
        Set tvTreeView.SelectedItem = tvTreeView.Nodes.Item("T" + CStr(CurTele))
        FillTextBox tvTreeView.SelectedItem
        tvTreeView.SelectedItem.EnsureVisible
    End If
End If
If bDeity Then
    optMotd.Visible = True
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
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
End Sub

Private Sub Form_Load()
Me.Width = Screen.Width * 0.75
Me.Height = Screen.Height * 0.75
Me.Left = (Screen.Width - Me.Width) \ 2
Me.top = (Screen.Height - Me.Height) \ 2
imgSplitter.Left = Me.ScaleWidth \ 3
lbNation.Visible = False
chkInclude.Visible = False
TreeDirty = True
End Sub

' 091903 rjk: Removed as there is no code inside
'Private Sub Form_Paint()
    'txtBody.View = Val(GetSetting(APPNAME, "Settings", "ViewMode", "0"))
    'tbToolBar.Buttons(txtBody.View + LISTVIEW_BUTTON).Value = tbrPressed
    'mnuListViewMode(txtBody.View).Checked = True
'End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Only unload if form Draw says OK
If UnloadMode = vbFormControlMenu Or UnloadMode = vbFormCode Then
    If Not frmDrawMap.ShutDown Then
        Cancel = 1
        Me.WindowState = vbMinimized
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer

'close all sub forms
For i = Forms.Count - 1 To 1 Step -1
    Unload Forms(i)
Next
End Sub


'Private Sub mnuViewOptions_Click()
'    frmOptions.Show vbModal, Me
'End Sub

Private Sub Form_Resize()
'If Me.WindowState <> vbNormal Then
If Me.WindowState = vbMinimized Then
    Exit Sub
End If

On Error Resume Next
If Me.Width < 3000 Then Me.Width = 3000
SizeControls imgSplitter.Left
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With imgSplitter
    picSplitter.Move .Left, .top, .Width \ 2, .Height - 20
End With
picSplitter.Visible = True
mbMoving = True
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim sglPos As Single

If mbMoving Then
    sglPos = X + imgSplitter.Left
    If sglPos < sglSplitLimit Then
        picSplitter.Left = sglSplitLimit
    ElseIf sglPos > Me.Width - sglSplitLimit Then
        picSplitter.Left = Me.Width - sglSplitLimit
    Else
        picSplitter.Left = sglPos
    End If
End If
End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
SizeControls picSplitter.Left
picSplitter.Visible = False
mbMoving = False
End Sub

Private Sub SizeControls(X As Single)
On Error Resume Next

'set the width
If X < 1500 Then X = 1500
If X > (Me.Width - 1500) Then X = Me.Width - 1500
tvTreeView.Width = X
imgSplitter.Left = X
txtBody.Left = X + 40
txtBody.Width = Me.Width - (tvTreeView.Width + 140)
lblTitle(0).Width = tvTreeView.Width
lblTitle(1).Left = txtBody.Left + 20
lblTitle(1).Width = txtBody.Width - 40

'set the top
If tbToolBar.Visible Then
    tvTreeView.top = tbToolBar.Height + picTitles.Height
Else
    tvTreeView.top = picTitles.Height
End If
txtBody.top = tvTreeView.top

'set the height
If sbStatusBar.Visible Then
    tvTreeView.Height = Me.ScaleHeight - (picTitles.top + picTitles.Height + sbStatusBar.Height)
Else
    tvTreeView.Height = Me.ScaleHeight - (picTitles.top + picTitles.Height)
End If

'Move the hidden forms
lbNation.Move tvTreeView.Left, tvTreeView.top, tvTreeView.Width
labNation.Move tvTreeView.Left, tvTreeView.top + lbNation.Height / 2, tvTreeView.Width
chkInclude.Move tvTreeView.Left, lbNation.top + lbNation.Height, tvTreeView.Width
chkSave.Move tvTreeView.Left, chkInclude.top + chkInclude.Height, tvTreeView.Width
frameSend.Move tvTreeView.Left, chkSave.top + chkSave.Height

txtBody.Height = tvTreeView.Height
imgSplitter.top = tvTreeView.top
imgSplitter.Height = tvTreeView.Height
End Sub

Private Sub lbNation_Click()
SetRecipientLabel
End Sub

Private Sub Option1_Click()
lbNation.Visible = True
labNation.Visible = False
Timer1.Interval = 0
'Set titles
lblTitle(0).Caption = "Send Telegram To:"
SetRecipientLabel
End Sub

Private Sub Option2_Click()
lbNation.Visible = False
labNation.Visible = True
Timer1.Interval = 1000
'Set titles
lblTitle(0).Caption = "Make Announcement:"
lblTitle(1).Caption = "Enter New Anno"
labNation.Caption = "<< ANNOUNCEMENT >>"
End Sub

Private Sub Option3_Click()
lbNation.Visible = True
labNation.Visible = False
Timer1.Interval = 0
'Set titles
lblTitle(0).Caption = "Send Flash To:"
SetRecipientLabel
End Sub


Private Sub optMotd_Click()
lbNation.Visible = False
labNation.Visible = True
Timer1.Interval = 1000
'Set titles
lblTitle(0).Caption = ""
lblTitle(1).Caption = "Enter New motd"
labNation.Caption = "<< motd >>"
End Sub

Private Sub Timer1_Timer()
If Option2.Value Or optMotd.Value Then
    If labNation.Visible Then
        labNation.Visible = False
        Timer1.Interval = 500
    Else
        labNation.Visible = True
    Timer1.Interval = 1000
    End If
End If
End Sub

Private Sub tvTreeView_DragDrop(Source As Control, X As Single, Y As Single)
If Source = imgSplitter Then
    SizeControls X
End If
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComCtlLib.Button)
' Dim strKey As String    8/2003 efj  removed
' Dim teleID As Integer    8/2003 efj  removed

rsTeleHead.Index = "PrimaryKey"

Select Case Button.key
Case "First"
        ButtonFirst
Case "Last"
        ButtonLast
Case "Previous"
        ButtonBack
Case "Next"
        ButtonNext
Case "Delete"
    If NewTelegram Then
        ButtonCancel
    Else
        ButtonDelete
    End If
Case "New"
    ButtonNew MESS_NEW
Case "Reply"
    ButtonNew MESS_REPLY
Case "Forward"
    ButtonNew MESS_FORWARD
Case "Send"
    ButtonSend
Case "Import" '103003 rjk: Added Import Intelligence
    ButtonImport
End Select
End Sub

Private Sub ButtonBack()
Dim strKey As String
Dim teleID As Integer

On Error GoTo TelegramError:

strKey = tvTreeView.SelectedItem.key
If Left$(strKey, 1) <> "T" Then
    Exit Sub
End If
teleID = CInt(Mid$(strKey, 2))
rsTeleHead.Seek "=", teleID
If Not rsTeleHead.NoMatch Then
    rsTeleHead.MovePrevious
End If
If rsTeleHead.NoMatch Or rsTeleHead.BOF Then
    MsgBox "No More Telegrams", vbExclamation + vbOKOnly, "Back Error"
Else
    teleID = rsTeleHead.Fields("ID")
    Set tvTreeView.SelectedItem = tvTreeView.Nodes.Item("T" + CStr(teleID))
    FillTextBox tvTreeView.SelectedItem
End If

Exit Sub
TelegramError:
TelegramError
End Sub

Private Sub ButtonNext()
Dim strKey As String
Dim teleID As Integer

On Error GoTo TelegramError:

strKey = tvTreeView.SelectedItem.key
If Left$(strKey, 1) <> "T" Then
    Exit Sub
End If
teleID = CInt(Mid$(strKey, 2))
rsTeleHead.Seek "=", teleID
If Not rsTeleHead.NoMatch Then
    rsTeleHead.MoveNext
End If
If rsTeleHead.NoMatch Or rsTeleHead.EOF Then
    MsgBox "No More Telegrams", vbExclamation + vbOKOnly, "Forward Error"
Else
    teleID = rsTeleHead.Fields("ID")
    Set tvTreeView.SelectedItem = tvTreeView.Nodes.Item("T" + CStr(teleID))
    FillTextBox tvTreeView.SelectedItem
End If

Exit Sub
TelegramError:
TelegramError
End Sub

Private Sub ButtonFirst()
' Dim strKey As String    8/2003 efj  removed
Dim teleID As Integer

On Error GoTo TelegramError:

If rsTeleHead.BOF And rsTeleHead.EOF Then
    Exit Sub
End If
rsTeleHead.MoveFirst
teleID = rsTeleHead.Fields("ID")
Set tvTreeView.SelectedItem = tvTreeView.Nodes.Item("T" + CStr(teleID))
FillTextBox tvTreeView.SelectedItem

Exit Sub
TelegramError:
TelegramError
End Sub

Private Sub ButtonLast()
' Dim strKey As String    8/2003 efj  removed
Dim teleID As Integer

On Error GoTo TelegramError:

If rsTeleHead.BOF And rsTeleHead.EOF Then
    Exit Sub
End If
rsTeleHead.MoveLast
teleID = rsTeleHead.Fields("ID")
Set tvTreeView.SelectedItem = tvTreeView.Nodes.Item("T" + CStr(teleID))
FillTextBox tvTreeView.SelectedItem

Exit Sub
TelegramError:
TelegramError
End Sub

Private Sub ButtonDelete()
Dim strKey As String
Dim teleID As Integer

On Error GoTo TelegramError:

strKey = tvTreeView.SelectedItem.key
If Left$(strKey, 1) <> "T" Then
    Exit Sub
End If

If tvTreeView.SelectedItem.key = tvTreeView.SelectedItem.FirstSibling.key Then
    tvTreeView.SelectedItem = tvTreeView.SelectedItem.Next
Else
    tvTreeView.SelectedItem = tvTreeView.SelectedItem.Previous
End If

teleID = CInt(Mid$(strKey, 2))
tvTreeView.Nodes.Remove strKey

'Delete the records
rsTeleHead.Index = "PrimaryKey"
rsTeleHead.Seek "=", teleID
If Not rsTeleHead.NoMatch Then
    rsTeleHead.Delete
    rsTeleBody.Index = "PrimaryKey"
    rsTeleBody.Seek "=", teleID
    While Not rsTeleBody.NoMatch
        rsTeleBody.Delete
        rsTeleBody.Seek "=", teleID
    Wend
End If

'rsTeleHead.MoveNext
'Now move to the next node
'If rsTeleHead.EOF Then
'    If Not rsTeleHead.BOF Then
'        rsTeleHead.MovePrevious
'    Else
'        Exit Sub
'    End If
'End If

'Fill the box in the new mode
If rsTeleHead.BOF And rsTeleHead.EOF Then
    txtBody.Text = vbNullString
Else
    'teleID = rsTeleHead.Fields("ID")
    'Set tvTreeView.SelectedItem = tvTreeView.Nodes.Item("T" + CStr(teleID))
    FillTextBox tvTreeView.SelectedItem
End If

Exit Sub

TelegramError:
TelegramError
End Sub

Private Sub ButtonNew(MessageType As eMessageCreate)
Dim strKey As String
Dim teleID As Integer
Dim X As Integer
' Dim ReplyNation As String    8/2003 efj  removed

'Fill the combo box
Nations.FillListBox lbNation
If VersionCheck(4, 3, 0) = VER_LESS Then
    lbNation.AddItem "The Deity", 0
    lbNation.ItemData(lbNation.NewIndex) = 0
End If

'Set to telegram
Option1.Value = True

'Set controls
chkSave.Visible = True
lbNation.Visible = True
labNation.Visible = False

'Find the current telegram
SelectedTelegram = 0
If MessageType = MESS_REPLY Or MessageType = MESS_FORWARD Then
    'if no telegram is selected, quit
    On Error GoTo ButtonNew_Exit
    strKey = tvTreeView.SelectedItem.key
    On Error GoTo 0
    If Left$(strKey, 1) = "T" Then
        teleID = CInt(Mid$(strKey, 2))
        rsTeleHead.Seek "=", teleID
        If Not rsTeleHead.NoMatch Then
            For X = 0 To lbNation.ListCount - 1
                If MessageType = MESS_REPLY And lbNation.ItemData(X) = CStr(rsTeleHead.Fields("From")) Then
                    lbNation.Selected(X) = True
                End If
            Next X
            chkInclude.Visible = True
            SelectedTelegram = teleID
            chkInclude.Value = vbChecked
        End If
    End If
End If
tvTreeView.Visible = False

NewTelegram = True
txtBody.Text = vbNullString
txtBody.Locked = False

'Set buttons
tbToolBar.Buttons(BUTTON_FIRST).Enabled = False    'All back
tbToolBar.Buttons(BUTTON_BACK).Enabled = False    'Back 0ne
tbToolBar.Buttons(BUTTON_NEXT).Enabled = False    'forward 0ne
tbToolBar.Buttons(BUTTON_LATEST).Enabled = False    'All forward
tbToolBar.Buttons(BUTTON_NEW).Enabled = False    'New Tele
tbToolBar.Buttons(BUTTON_REPLY).Enabled = False    'Reply
tbToolBar.Buttons(BUTTON_FORWARD).Enabled = False  'Forward
tbToolBar.Buttons(BUTTON_IMPORT).Enabled = False    'Import Intelligence 103003 rjk: Added

tbToolBar.Buttons(BUTTON_DELETE).Enabled = True     'Delete
tbToolBar.Buttons(BUTTON_SEND).Enabled = True     'Send

'Set titles
lblTitle(0).Caption = "Send Telegram To:"

'Fill in text if its a reply or a forward
If MessageType = MESS_FORWARD Then
    bForwardMessage = True
Else
    bForwardMessage = False
End If
If MessageType = MESS_REPLY Or MessageType = MESS_FORWARD Then
    chkInclude_Click
End If

ButtonNew_Exit:
End Sub

Private Sub ButtonCancel()
' Dim strKey As String    8/2003 efj  removed
' Dim teleID As Integer    8/2003 efj  removed

'Set controls
lbNation.Visible = False
labNation.Visible = False
chkSave.Visible = False
chkInclude.Visible = False
tvTreeView.Visible = True
NewTelegram = False
txtBody.Text = vbNullString
txtBody.Locked = True
FillTextBox tvTreeView.SelectedItem
lblTitle(0).Caption = "Received Telegrams"
Timer1.Interval = 0
End Sub

Private Sub ButtonSend()
Dim n As Integer
Dim strTel As String
Dim hldNatNumber As Integer
Dim hldNation As String
Dim X As Integer

Dim strTemp As String
Dim strSource As String
Dim strTarget() As String
Dim numTelegrams As Integer

'Must have a nation !
If lbNation.SelCount = 0 And Option1.Value Then
    MsgBox "Must enter a Nation to send a telegram.", , "Error"
    Exit Sub
End If

If chkSave.Value = vbChecked And Option1.Value Then
    hldNatNumber = -1
    For X = 0 To lbNation.ListCount - 1
        If lbNation.Selected(X) Then
            hldNatNumber = lbNation.ItemData(X)
            hldNation = lbNation.List(X)
        End If
    Next X

    If hldNatNumber < 0 Then 'still couldn't figure out number, so abort
        chkSave.Value = vbUnchecked
    Else
        IncomingTelegramLine "> Telegram To " + hldNation + " (#" + CStr(hldNatNumber) + ")" _
            + " dated " + Format$(DateAdd("n", rsVersion.Fields("Time Zone Adj"), Now), "ddd mmm dd hh:mm:ss yyyy"), True
    End If
    
End If

''fixme! drk@unxsoft.com 2/26/03 don't want to generate telegrams with lines over 80 chars
strSource = txtBody.Text
ReDim strTarget(1 To 1)
numTelegrams = 1
Do
    n = InStr(strSource, vbLf)
    If n = 0 Then n = Len(strSource) + 1
    strTemp = Left$(strSource, n - 1)
    strSource = Mid$(strSource, n + 1)
    
    'remove carrage returns ''why? drk@unxsoft.com 2/26
    n = InStr(strTemp, vbCr)
    While n > 0
        strTemp = Left$(strTemp, n - 1) + Mid$(strTemp, n + 1)
        n = InStr(strTemp, vbCr)
    Wend
    
    If chkSave.Value = vbChecked And Option1.Value Then
         IncomingTelegramLine strTemp, False
    End If
    
    'just in case we get a real long line, split it up
    '110303 rjk: Modified to account for the server not support up 1024 characters,
    '            it only supports 1022 characters in a telegram line.
    While Len(strTemp) > 1021
        n = 1021 - Len(strTarget(numTelegrams))
        strTarget(numTelegrams) = strTarget(numTelegrams) + Left$(strTemp, n)
        strTemp = Mid$(strTemp, n + 1)
        numTelegrams = numTelegrams + 1 '110303 rjk: Fixed to support lines longer than 2044
        ReDim Preserve strTarget(1 To numTelegrams)
    Wend
    
    If (Len(strTarget(numTelegrams)) + Len(strTemp)) >= 1022 Then
        numTelegrams = numTelegrams + 1
        ReDim Preserve strTarget(1 To numTelegrams)
    End If
    
    strTarget(numTelegrams) = strTarget(numTelegrams) + strTemp + vbLf

Loop While Len(strSource) > 0

'run thru multiple telegrams if necessay
For n = 1 To numTelegrams
    frmEmpCmd.SubmitEmpireCommand "te1 " + strTarget(n), False
    
    If Option1.Value Then
        'Build recipient string
        strTel = vbNullString
        For X = 0 To lbNation.ListCount - 1
            If lbNation.Selected(X) Then
                strTel = strTel + CStr(lbNation.ItemData(X)) + " "
            End If
        Next X
        frmEmpCmd.SubmitEmpireCommand "telegram " + strTel, False
    ElseIf Option2.Value Then
        frmEmpCmd.SubmitEmpireCommand "announce ", False
    ElseIf Option3.Value Then
        'Build recipient string
        strTel = vbNullString
        For X = 0 To lbNation.ListCount - 1
            If lbNation.Selected(X) Then
                strTel = CStr(lbNation.ItemData(X))
            End If
        Next X
        frmEmpCmd.SubmitEmpireCommand "flash " + strTel, False
    ElseIf optMotd.Value Then
        frmEmpCmd.SubmitEmpireCommand "turn mess ", False
    End If
Next n

'Set controls
lbNation.Visible = False
labNation.Visible = False
chkSave.Visible = False
chkInclude.Visible = False
tvTreeView.Visible = True
NewTelegram = False
txtBody.Text = vbNullString
txtBody.Locked = True
FillTextBox tvTreeView.SelectedItem
lblTitle(0).Caption = "Received Telegrams"
Timer1.Interval = 0
Me.WindowState = vbMinimized
End Sub

'103003 rjk: Added Import Intelligence
Private Sub ButtonImport()
Dim strKey As String
Dim teleID As Integer
Dim strTelegram As String

'Find the current telegram
SelectedTelegram = 0
On Error GoTo ButtonImport_Exit
strKey = tvTreeView.SelectedItem.key
On Error GoTo 0

strTelegram = vbNullString

If Left$(strKey, 1) = "T" Then
    teleID = CInt(Mid$(strKey, 2))
    rsTeleHead.Seek "=", teleID
    If Not rsTeleHead.NoMatch Then
        rsTeleBody.Seek "=", teleID, 1
        If Not rsTeleBody.NoMatch Then
            Do While (Not rsTeleBody.EOF)
                If rsTeleBody.Fields("ID") <> teleID Then
                    Exit Do
                End If
                If Len(rsTeleBody.Fields("Body")) > 0 Then
                    strTelegram = strTelegram + rsTeleBody.Fields("Body") + vbLf
                End If
                rsTeleBody.MoveNext
            Loop
        End If
    End If
End If

ImportIntelligenceOffset strTelegram
ButtonImport_Exit:
End Sub

Private Sub LoadTree()
Dim title As String
Dim nodX As Node
' Dim nodP As Node    8/2003 efj  removed
Dim X As Integer
Dim c As Integer
Dim ntele As Integer
Dim nteleexp As Integer

On Error Resume Next

'Clear out old nodes
While tvTreeView.Nodes.Count > 0
    tvTreeView.Nodes.Remove 1
Wend

Set nodX = tvTreeView.Nodes.Add(, , "R", "Telegrams")

'Set the class nodes
For X = TELEGRAM_CLASS_INBOX + 1 To NUMBER_OF_TELEGRAM_CLASSES
    Select Case X
        Case TELEGRAM_CLASS_INBOX
            title = "In Box"
        Case TELEGRAM_CLASS_COMBAT
            title = "Combat Reports"
        Case TELEGRAM_CLASS_BULLETIN
            title = "Bulletins"
        Case TELEGRAM_CLASS_PLAYER
            title = "Player Telegrams"
        Case TELEGRAM_CLASS_INTELLIGENCE '103003 rjk: Added Import Intelligence
            title = "Intelligence Reports"
        Case TELEGRAM_CLASS_PRODUCTION
             title = "Production Reports"
        Case TELEGRAM_CLASS_REPORTS
             title = "Saved Reports"
        Case TELEGRAM_CLASS_GENERAL
             title = "Misc."
        Case Else
            title = "Other"
    End Select
    Set nodX = tvTreeView.Nodes.Add("R", tvwChild, "C" + CStr(X), title)
Next X
nodX.EnsureVisible

ntele = 0
nteleexp = 0

Dim hldIndex As String
hldIndex = rsTeleHead.Index
rsTeleHead.Index = "Received"
If Not (rsTeleHead.BOF And rsTeleHead.EOF) Then
    rsTeleHead.MoveLast
    While Not rsTeleHead.BOF
        ntele = ntele + 1
        c = rsTeleHead.Fields("Class")
        If rsTeleHead.Fields("Exported") <> "Y" Then
            nteleexp = nteleexp + 1
        End If
            
        'Each player gets a subclass
        If c = TELEGRAM_CLASS_PLAYER Then
            X = rsTeleHead.Fields("From")
            AddPlayerNode X, "P"
            Set nodX = tvTreeView.Nodes.Add("P" + CStr(X), tvwChild, _
                    "T" + CStr(rsTeleHead.Fields("ID")), rsTeleHead.Fields("Subject"))
        ElseIf c = TELEGRAM_CLASS_INTELLIGENCE Then '103003 rjk: Added Intelligence Reports
            X = rsTeleHead.Fields("From")
            AddPlayerNode X, "I"
            Set nodX = tvTreeView.Nodes.Add("I" + CStr(X), tvwChild, _
                    "T" + CStr(rsTeleHead.Fields("ID")), rsTeleHead.Fields("Subject"))
        Else
            Set nodX = tvTreeView.Nodes.Add("C" + CStr(c), tvwChild, _
                    "T" + CStr(rsTeleHead.Fields("ID")), rsTeleHead.Fields("Subject"))
        End If
        rsTeleHead.MovePrevious
    Wend
    rsTeleHead.MoveLast
End If

rsTeleHead.Index = hldIndex

tvTreeView.Style = tvwTreelinesText ' Style 4.
tvTreeView.BorderStyle = vbFixedSingle
sbStatusBar.Panels.Item(1).Text = CStr(ntele) + " Telegrams, "
sbStatusBar.Panels.Item(2).Text = CStr(nteleexp) + " not exported"
TreeDirty = False
lblTitle(1).Caption = vbNullString
End Sub

'103103 rjk: Added strClass to separate the Intelligence Class from Player Class
Private Sub AddPlayerNode(X As Integer, strClass As String)
Dim Node As MSComCtlLib.Node
Dim PlayerName As String

On Error GoTo PlayerNode_New_Node
Set Node = tvTreeView.Nodes.Item(strClass + CStr(X))
Exit Sub

PlayerNode_New_Node:
On Error Resume Next
'Create a new node
PlayerName = Nations.NationName(X)
If Len(PlayerName) = 0 Then
    PlayerName = "Nation #" + CStr(X)
End If

Set Node = tvTreeView.Nodes.Add("C" + CStr(rsTeleHead.Fields("Class")), tvwChild, _
    strClass + CStr(X), PlayerName)
End Sub

Private Sub tvTreeView_NodeClick(ByVal Node As MSComCtlLib.Node)
'Put the selected telegram in the text box
FillTextBox Node
End Sub

Private Sub FillTextBox(ByVal Node As MSComCtlLib.Node)
Dim teleID As Integer

On Error Resume Next
tbToolBar.Buttons(BUTTON_FIRST).Enabled = True     'All back
tbToolBar.Buttons(BUTTON_BACK).Enabled = False    'Back 0ne
tbToolBar.Buttons(BUTTON_NEXT).Enabled = False    'forward 0ne
tbToolBar.Buttons(BUTTON_LATEST).Enabled = True     'All forward
tbToolBar.Buttons(BUTTON_NEW).Enabled = True     'New Tele
tbToolBar.Buttons(BUTTON_REPLY).Enabled = False    'Reply
tbToolBar.Buttons(BUTTON_FORWARD).Enabled = False    'Forward
tbToolBar.Buttons(BUTTON_IMPORT).Enabled = False    'Import Intelligence 103003 rjk: Added
tbToolBar.Buttons(BUTTON_DELETE).Enabled = False    'Delete
tbToolBar.Buttons(BUTTON_SEND).Enabled = False    'Send


' Dim hldFont As String    8/2003 efj  removed
Dim hldText As String

txtBody.Text = vbNullString
hldText = vbNullString
If Left$(Node.key, 1) = "T" Then
    teleID = CInt(Mid$(Node.key, 2))
    rsTeleBody.Index = "PrimaryKey"
    rsTeleBody.Seek "=", teleID, 1
    
    'Load telegram body into text box
    If Not rsTeleBody.NoMatch Then
        CurTele = teleID
        Do While (Not rsTeleBody.EOF)
            If rsTeleBody.Fields("ID") <> teleID Then
                Exit Do
            End If
'0901903 rjk: removed linefeed for the reports, as data is not save on per line
' bases anymore but in 255 byte chucks.  However, bulletins are still saved one line at
' time with no linefeeds.
            '101503 rjk: If the record is full then no new line is required.
            '            Blank added at insertion time.
            If Len(rsTeleBody.Fields("Body")) = 255 Then
            'If InStr(rsTeleBody.Fields("Body"), vbLf) > 0 Then 'Reports
                hldText = hldText + rsTeleBody.Fields("Body")
            ElseIf Right$(rsTeleBody.Fields("Body"), 1) <> vbLf Then
                hldText = hldText + rsTeleBody.Fields("Body") + vbLf
            Else
                hldText = hldText + rsTeleBody.Fields("Body")
            End If
            'txtBody.Text = txtBody.Text + rsTeleBody.Fields("Body") + vbLf
            rsTeleBody.MoveNext
        Loop
    End If
    txtBody.Text = hldText
    tbToolBar.Buttons(BUTTON_BACK).Enabled = True     'Back 0ne
    tbToolBar.Buttons(BUTTON_NEXT).Enabled = True     'forward 0ne
    tbToolBar.Buttons(BUTTON_DELETE).Enabled = True   'Delete

    'Set title
    rsTeleHead.Index = "PrimaryKey"
    rsTeleHead.Seek "=", CInt(Mid$(Node.key, 2))
    If Not rsTeleHead.NoMatch Then
        lblTitle(1).Caption = rsTeleHead.Fields("Title")
        If rsTeleHead.Fields("Read") <> "Y" Then
            rsTeleHead.Edit
            rsTeleHead.Fields("Read") = "Y"
            rsTeleHead.Update
        End If
        If rsTeleHead.Fields("Class") = TELEGRAM_CLASS_PLAYER Then
            tbToolBar.Buttons(BUTTON_REPLY).Enabled = True    'Reply
        End If
        If rsTeleHead.Fields("Class") = TELEGRAM_CLASS_INTELLIGENCE Then
            tbToolBar.Buttons(BUTTON_IMPORT).Enabled = True    'Import Intelligence 103003 rjk: Added
        End If
        tbToolBar.Buttons(BUTTON_FORWARD).Enabled = True    'Forward
    End If
Else
    lblTitle(1).Caption = vbNullString
End If
End Sub

'112903 rjk: Remove the question on clearing the telegrams as there is a separate form now.
Public Sub ClearTelegrams()
'if no telegrams, then quit
If rsTeleHead.BOF And rsTeleHead.EOF Then
    Exit Sub
End If

DeleteAllRecords rsTeleHead
DeleteAllRecords rsTeleBody

LoadTree '111403 rjk: Clear Telegram window after removing all the telegrams.
End Sub

Private Sub txtBody_Change()
If Not txtBody.Locked Then
    sbStatusBar.Panels.Item(2).Text = CStr(Len(txtBody.Text)) + " characters"
    ColorLength
End If
End Sub

Private Sub ColorLength()
Dim lIndex As Long
Dim lCharCount As Long
Dim lStartPos As Long
Dim lCharPos As Long
Dim lLength As Long
Dim iTelegramTypeIndex As Integer
Dim iSoundIndex As Integer

If Option1.Value = True Then
    iTelegramTypeIndex = totTelegram
ElseIf Option2.Value = True Then
    iTelegramTypeIndex = totAnnouncement
ElseIf Option3.Value = True Then
    iTelegramTypeIndex = totFlash
ElseIf optMotd.Value = True Then
    iTelegramTypeIndex = totMOTD
Else
    Exit Sub
End If

lCharCount = 0
lStartPos = 0
lLength = txtBody.SelLength
lCharPos = txtBody.SelStart
If Not tTelegramOptions(iTelegramTypeIndex, tlWarning).bEnabled And _
    Not tTelegramOptions(iTelegramTypeIndex, tlSoftLimit).bEnabled Then
    txtBody.SelStart = 0
    txtBody.SelLength = Len(txtBody.Text)
    txtBody.SelColor = vbBlack
    txtBody.SelStart = lCharPos
    txtBody.SelLength = lLength
    If Not tTelegramOptions(iTelegramTypeIndex, tlHardLimit) Then
        Exit Sub
    End If
End If
For lIndex = 0 To Len(txtBody.Text) - 1
    lCharCount = lCharCount + 1
    If tTelegramOptions(iTelegramTypeIndex, tlHardLimit).bEnabled And _
        lCharCount = tTelegramOptions(iTelegramTypeIndex, tlHardLimit).iColumn + 1 And _
        Mid$(txtBody.Text, lIndex + 1, 1) <> vbCr Then
        txtBody.Text = Left$(txtBody.Text, lIndex) + vbCrLf + _
            Right$(txtBody.Text, Len(txtBody.Text) - lIndex)
        If lCharPos > lIndex Then
            lCharPos = lCharPos + 2 'Advance the current point past the vbCrLf
        End If
    End If
    If Mid$(txtBody.Text, lIndex + 1, 1) = vbCr Then
        txtBody.SelStart = lStartPos
        txtBody.SelLength = lIndex - lStartPos + 1
        If tTelegramOptions(iTelegramTypeIndex, tlSoftLimit).bEnabled And _
            lCharCount > tTelegramOptions(iTelegramTypeIndex, tlSoftLimit).iColumn Then
            txtBody.SelColor = vbRed
        ElseIf tTelegramOptions(iTelegramTypeIndex, tlWarning).bEnabled And _
            lCharCount > tTelegramOptions(iTelegramTypeIndex, tlWarning).iColumn Then
            txtBody.SelColor = vbBlue
        ElseIf lIndex = Len(txtBody.Text) - 1 Then
            txtBody.SelColor = vbBlack
        End If
        lCharCount = 0
        lStartPos = lIndex + 1
    ElseIf Mid$(txtBody.Text, lIndex + 1, 1) = vbLf Then
        lCharCount = 0
        lStartPos = lIndex + 1
    ElseIf lIndex = Len(txtBody.Text) - 1 Then
        txtBody.SelStart = lStartPos
        txtBody.SelLength = lIndex - lStartPos + 1
        If tTelegramOptions(iTelegramTypeIndex, tlSoftLimit).bEnabled And _
            lCharCount > tTelegramOptions(iTelegramTypeIndex, tlSoftLimit).iColumn Then
            txtBody.SelColor = vbRed
            If lCharCount = tTelegramOptions(iTelegramTypeIndex, tlSoftLimit).iColumn + 1 Then
                For iSoundIndex = 1 To tTelegramOptions(iTelegramTypeIndex, tlSoftLimit).eTelegramSound
                    Beep
                Next iSoundIndex
            End If
        ElseIf tTelegramOptions(iTelegramTypeIndex, tlWarning).bEnabled And _
            lCharCount > tTelegramOptions(iTelegramTypeIndex, tlWarning).iColumn Then
            txtBody.SelColor = vbBlue
            If lCharCount = tTelegramOptions(iTelegramTypeIndex, tlWarning).iColumn + 1 Then
                For iSoundIndex = 1 To tTelegramOptions(iTelegramTypeIndex, tlWarning).eTelegramSound
                    Beep
                Next iSoundIndex
            End If
        ElseIf lIndex = Len(txtBody.Text) - 1 Then
            txtBody.SelColor = vbBlack
        End If
    ElseIf tTelegramOptions(iTelegramTypeIndex, tlWarning).bEnabled And _
            lCharCount = tTelegramOptions(iTelegramTypeIndex, tlWarning).iColumn Then
        txtBody.SelStart = lStartPos
        txtBody.SelLength = lIndex - lStartPos + 1
        txtBody.SelColor = vbBlack
        lStartPos = lIndex + 1
    ElseIf tTelegramOptions(iTelegramTypeIndex, tlSoftLimit).bEnabled And _
            lCharCount = tTelegramOptions(iTelegramTypeIndex, tlSoftLimit).iColumn Then
        txtBody.SelStart = lStartPos
        txtBody.SelLength = lIndex - lStartPos + 1
        If Not tTelegramOptions(iTelegramTypeIndex, tlWarning).bEnabled Then
            txtBody.SelColor = vbBlack
        Else
            txtBody.SelColor = vbBlue
        End If
        lStartPos = lIndex + 1
    End If
Next lIndex
txtBody.SelStart = lCharPos
txtBody.SelLength = lLength
End Sub

Private Sub SetRecipientLabel()
Dim X As Integer
Dim strNationList As String
Dim strTemp As String

'Look for what is selected
For X = 0 To lbNation.ListCount - 1
    If lbNation.Selected(X) Then
        strNationList = strNationList + lbNation.List(X) + ", "
    End If
Next X

'Now set label
If Option1.Value Then
    strTemp = "Telegram"
Else
    strTemp = "Flash"
End If

If Len(strNationList) = 0 Then
    lblTitle(1).Caption = "Enter New " + strTemp + ", Select Recipient"
Else
    lblTitle(1).Caption = strTemp + " For: " + Left$(strNationList, Len(strNationList) - 2)
End If

End Sub

Private Sub TelegramError()
MsgBox "ERROR - Your telegram database is corrupted.  WinACE was not able" + _
" to move to the requested Telegram.  You should take Export, then Clear, then Import off" + _
" the File Menu to repair the database", vbOKOnly + vbCritical, "Database Error"
End Sub
