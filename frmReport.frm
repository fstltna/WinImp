VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReport 
   Caption         =   "Reports"
   ClientHeight    =   5808
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   8496
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5808
   ScaleWidth      =   8496
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "Command1"
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   4680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   648
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8496
      _ExtentX        =   14986
      _ExtentY        =   1143
      ButtonWidth     =   1990
      ButtonHeight    =   1016
      Appearance      =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save to Disk"
            Key             =   "Save"
            Object.ToolTipText     =   "Save to Disk"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Copy"
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy to Clipboard"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save to Log"
            Key             =   "File"
            Object.ToolTipText     =   "Copy to Telegram File"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit "
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   3855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   720
      Width           =   8175
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7920
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   7920
      Top             =   5040
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   22
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":0542
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":0A84
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReport.frx":0C1E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'300903 rjk: General reformatting.
'290804 rjk: Always save telegrams unless it is duplicate.
'170906 rjk: Added TTS report window.
'020907 rjk: Switched to database field for TimeZoneAdj.

Private TextMaxWidth As Integer
Private TextLines As Integer
Private SavedOnce As Boolean

Private Sub Form_Activate()
Dim nLen As Long

'If mimimized for some reason, pop it up
If Me.WindowState = vbMinimized Then
    Me.WindowState = vbNormal
    Exit Sub
End If

If TextMaxWidth < 5340 Then
    TextMaxWidth = 5340
End If

If TextMaxWidth > 0 Then
    If TextMaxWidth > ((Screen.Width * 4) / 5) Then
        TextMaxWidth = (Screen.Width * 4) / 5
    End If
            
    If Me.WindowState <> vbMaximized Then
        Me.Width = TextMaxWidth * 1.15
        Me.Left = (Screen.Width - Me.Width) \ 2
    End If
End If

If TextLines > 1 Then
    nLen = Me.TextHeight("XX") * (TextLines + 2)
    
    
    Text1.Height = nLen
    If Text1.Height > ((Screen.Height * 3) / 4) Then
        Text1.Height = (Screen.Height * 3) / 4
    End If
    
    Me.Height = Text1.Height + tbToolBar.Height * 2
    Me.top = (Screen.Height - Me.Height) \ 2
End If
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) \ 2
Me.top = (Screen.Height - Me.Height) \ 2
TextLines = 0
SavedOnce = False
End Sub

Private Sub Form_Resize()
If Me.WindowState <> vbNormal Then
    Exit Sub
End If

'Check for minimum width, height
If Me.Width < 4005 Then
    Me.Width = 4005
End If

If Me.Height < 6675 Then
    Me.Height = 6675
End If

'Position tool bar
tbToolBar.Move 0, 0

'Position text Box
'lstReport.Move 0, tbToolBar.Height, Me.ScaleWidth, Me.ScaleHeight - tbToolBar.Height
Text1.Move 0, tbToolBar.Height, Me.ScaleWidth, Me.ScaleHeight - tbToolBar.Height
End Sub

Public Sub ClearReport()
'clear text box
'lstReport.Clear
Text1.Text = vbNullString
TextMaxWidth = 0
TextLines = 0
End Sub

Public Sub AddReportLine(strLine As String)
' Dim S() As String    8/2003 efj  removed
Dim X As Integer, nLen As Integer

On Error GoTo AddReportLine_exit

If frmEmpCmd.SubmittedFromCommandLine Then
    Exit Sub
End If


'Add line to text box
'lstReport.AddItem strLine
Text1.Text = Text1.Text + vbCrLf + strLine
TextLines = TextLines + 1

X = frmReport.TextWidth(strLine)

If X > TextMaxWidth Then
    TextMaxWidth = X
End If

If Me.WindowState = vbMinimized Then
    Me.WindowState = vbNormal
End If

If TextLines > 1 Then
    nLen = Me.TextHeight("XX") * (TextLines + 2)
    'If lstReport.Height < nLen Then
    If Text1.Height < nLen Then
    
        'lstReport.Height = nLen
        Text1.Height = nLen
        'If lstReport.Height > ((Screen.Height * 3) / 4) Then
        '    lstReport.Height = (Screen.Height * 3) / 4
        'End If
        If Text1.Height > ((Screen.Height * 3) / 4) Then
            Text1.Height = (Screen.Height * 3) / 4
        End If
        
        'Me.Height = lstReport.Height + tbToolBar.Height * 2
        Me.Height = Text1.Height + tbToolBar.Height * 2
        Me.top = (Screen.Height - Me.Height) \ 2
    End If
End If
    
If frmOptions.bTTSReportWindow And Len(strLine) > 0 Then
'    ttsReport.Tag = "Normal"
'    ttsReport.Speak strLine
End If

AddReportLine_exit:
End Sub

'Private Sub lstReport_KeyPress(KeyAscii As Integer)    8/2003 efj  removed
'
'Dim txtstring As String
'
''Trap Control C and Copy to Clipboard
'If KeyAscii = 3 Then
'    Clipboard.SetText Text1.Text
'End If
'
'End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComCtlLib.Button)
Select Case Button.key
Case "Save"
        ButtonSave
Case "Copy"
        Clipboard.SetText Me.Text1
Case "File"
        ButtonFile
Case "Exit"
        Unload Me
End Select
End Sub

Private Sub ButtonFile()
If SavedOnce Then
    Exit Sub
End If

' Copy Selected text into a sring
IncomingTelegramLine "> Saved " + Me.Caption + "  dated " _
    + Format$(DateAdd("n", rsVersion.Fields("Time Zone Adj"), Now), "ddd mmm dd hh:mm:ss yyyy"), True
' Dim x As Integer    8/2003 efj  removed
'For x = 0 To (lstReport.ListCount - 1)
'    IncomingTelegramLine lstReport.List(x), False, True
'Next x
IncomingTelegramLine Text1, False
SavedOnce = True
End Sub

Private Sub ButtonSave()
cd1.CancelError = True
On Error GoTo ErrHandler
' Set flags
cd1.FLAGS = cdlOFNHideReadOnly
' Set filters
cd1.Filter = "All Files (*.*)|*.*|Text Files" & _
"(*.txt)|*.txt|Batch Files (*.bat)|*.bat"
' Specify default filter
cd1.FilterIndex = 2
' Set default file name
cd1.FileName = Me.Caption

If VerifySubDirectory("Saved Reports", True) Then
    cd1.InitDir = App.path + "\Saved Reports"
End If

' Display the Open dialog box
cd1.ShowSave
' Display name of selected file

On Error GoTo 0

Dim strFile As String 'drk@unxsoft.com 7/03/02
Dim Reply As Integer

strFile = cd1.FileName

'Append file suffix if necessary
If InStr(strFile, ".") = 0 Then
    strFile = strFile + ".txt"
End If

'If already exists, then prompt
Reply = vbNo
If Len(Dir(strFile)) > 0 Then

    Reply = MsgBox("File " + strFile + " alreay exists.  Do you wish to " + _
        "append this report to the file ?  (Take Yes to Append, No to Replace, " + _
        "Cancel to Exit", vbYesNoCancel, "File Exists")
        
    If Reply = vbCancel Then
        Exit Sub
    End If
End If

Dim filenum As Integer
' Dim x As Integer    8/2003 efj  removed

filenum = FreeFile
If Reply = vbYes Then
    Open strFile For Append As #filenum
Else
    Open strFile For Output As #filenum
End If

Print #filenum, Text1.Text

Close filenum
Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
Exit Sub
End Sub
