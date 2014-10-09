VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmScript 
   Caption         =   "Script File Editor"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   7680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmScript.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   7680
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   645
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   1138
      ButtonWidth     =   1111
      ButtonHeight    =   979
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "New"
            Object.ToolTipText     =   "Start New File"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Key             =   "Open"
            Object.ToolTipText     =   "Open File"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "Save"
            Object.ToolTipText     =   "Save File "
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit Script Editor"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rec."
            Key             =   "Record"
            Object.ToolTipText     =   "Start Script Recorder"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Run"
            Key             =   "Run"
            Object.ToolTipText     =   "Execute Script File"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Timer"
            Key             =   "Timer"
            Object.ToolTipText     =   "Attach to Timer"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   3120
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".txt"
      DialogTitle     =   "Save Script File"
   End
   Begin RichTextLib.RichTextBox rt1 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   6588
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmScript.frx":0442
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2040
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":04C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":0B3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":11B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":1832
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":1944
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":1C5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":1F78
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":2292
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'08xx03 efj: removed dead procedures and variables
'093003 rjk: general reformatting.
'110103 rjk: Form cleanup - removed left and width from toolbar.

Dim Dirty As Boolean

'Private Sub cmdCancel_Click()    8/2003 efj  removed
'If CheckDirty Then
'    Unload Me
'End If
'End Sub

Private Sub ButtonRecord()
'if currently recording - process as a stop
If frmEmpCmd.RecordingScriptFile Then
    frmEmpCmd.RecordingScriptFile = False
    tbToolBar.Buttons(6).Caption = "Rec."
    tbToolBar.Buttons(6).Image = 5
    tbToolBar.Buttons(6).ToolTipText = "Start Script Recorder"
    tbToolBar.Buttons(8).Enabled = True
Else
    'start recording
    frmEmpCmd.RecordingScriptFile = True
    tbToolBar.Buttons(6).Caption = "Stop"
    tbToolBar.Buttons(6).Image = 6
    tbToolBar.Buttons(6).ToolTipText = "Stop Script Recorder"
    tbToolBar.Buttons(8).Enabled = False
End If
End Sub

Private Sub Form_Load()
' Set parent for the toolbar to display on top of:
' Dim success As Long    8/2003 efj  removed
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

Me.Width = Me.tbToolBar.Left + tbToolBar.ButtonWidth * 8
Me.rt1.Width = Me.ScaleWidth
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then
    Exit Sub
End If

Me.rt1.Width = Me.ScaleWidth
Me.rt1.Height = Me.ScaleHeight - Me.rt1.top
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Make sure scripting is turned off
frmEmpCmd.RecordingScriptFile = False
End Sub

Private Function CheckDirty() As Boolean
Dim Reply As Integer
'if file has been changed since last save, update
If Dirty Then
    Reply = MsgBox("File has changed.  Do you want to save changes ?", vbYesNoCancel, "Confirm Clear")
    If Reply = vbCancel Then
        CheckDirty = False
        Exit Function
    End If
    
    If Reply = vbNo Then
        CheckDirty = True
        Exit Function
    End If
    
    'save changes
    CheckDirty = ButtonSave
Else
    CheckDirty = True
End If
End Function

Private Sub rt1_Change()
Dirty = True
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComCtlLib.Button)
' Dim success As Long
'turn off topmost
Call SetWindowPos(Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
Select Case Button.key
Case "New"
        ButtonNew
Case "Open"
        ButtonOpen
Case "Save"
        ButtonSave
Case "Exit"
        Unload Me
Case "Record"
        ButtonRecord
Case "Run"
        ButtonRun
Case "Timer"
        ButtonTimer
End Select

'Pop topmost back up
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Private Function ButtonSave() As Boolean
Dim strFile As String

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
'cd1.FileName = "Empire Script.txt" 'drk@unxsoft.com.  this annoyed me. did it annoy anyone else, or should I put it back?

If VerifySubDirectory("Scripts", True) Then
    cd1.InitDir = App.path + "\Scripts"
End If

' Display the Open dialog box
cd1.DialogTitle = "Save Script File"
cd1.ShowSave
' Display name of selected file

On Error GoTo 0

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
        "replace it?", vbOKCancel + vbQuestion, "File Exists")
        
    If Reply = vbCancel Then
        ButtonSave = False
        Exit Function
    End If
End If

rt1.SaveFile strFile, rtfText
Dirty = False
ButtonSave = True
Exit Function
    
ErrHandler:
    'User pressed the Cancel button
ButtonSave = False
End Function


Private Sub ButtonOpen()
Dim strFile As String

If Not CheckDirty Then
    Exit Sub
End If

cd1.CancelError = True
On Error GoTo ErrHandler
' Set flags
cd1.FLAGS = cdlOFNHideReadOnly
' Set filters
cd1.Filter = "All Files (*.*)|*.*|Text Files" & _
"(*.txt)|*.txt|Batch Files (*.bat)|*.bat"
' Specify default filter
cd1.FilterIndex = 2

' Display the Open dialog box
cd1.DialogTitle = "Open Script File"
cd1.FileName = vbNullString

If VerifySubDirectory("Scripts", True) Then
    cd1.InitDir = App.path + "\Scripts"
End If

cd1.ShowOpen

On Error GoTo 0

' Dim strReply As Integer    8/2003 efj  removed

strFile = cd1.FileName

rt1.LoadFile strFile, rtfText
Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
End Sub
Private Sub ButtonNew()
If Not CheckDirty Then
    Exit Sub
End If

'Clear the text file
rt1.Text = vbNullString
Dirty = False
End Sub

Private Sub ButtonRun()
Dim strSource As String
Dim strTemp As String
Dim n As Integer

frmEmpCmd.SubmitEmpireCommand "bf1", False
strSource = rt1.Text
Do
    n = InStr(strSource, vbLf)
    If n = 0 Then n = Len(strSource) + 1
    strTemp = Left$(strSource, n - 1)
    strSource = Mid$(strSource, n + 1)
    
    'remove carrage returns
    n = InStr(strTemp, vbCr)
    While n > 0
        strTemp = Left$(strTemp, n - 1) + Mid$(strTemp, n + 1)
        n = InStr(strTemp, vbCr)
    Wend
    
    'Execute the command
    frmEmpCmd.SubmitEmpireCommand strTemp, False
    
Loop While Len(strSource) > 0

frmEmpCmd.SubmitEmpireCommand "bf2", False
End Sub

Private Sub ButtonTimer()
Dim strFile As String
Dim nReply As Integer
Dim strReply As String

strReply = TimerHeader
strReply = strReply + vbCr + vbCr + "Do you wish to change selections? " _
    + vbCr + vbCr + "Select Yes to Change" + vbCr + "Select No to Exit without Changes" + vbCr + "Select Cancel to Delete Current Selections"

nReply = MsgBox(strReply, vbYesNoCancel + vbQuestion, "Current Timed Script Settings")

'If cancel, wipe out scripts
If nReply = vbCancel Then
    If Len(BatchScript1File) > 0 Or Len(BatchScriptUpdate) > 0 Then
        BatchScript1File = vbNullString
        BatchScriptUpdate = vbNullString
        BatchScript1Time = 0
        strReply = TimerHeader + vbCr + vbCr + "Scipts cleared"
        MsgBox strReply, vbOKOnly, "Timed Script Settings"
    End If
End If

If nReply <> vbYes Then
    Exit Sub
End If

cd1.CancelError = True
On Error GoTo ButtonTimerExit
' Set flags
cd1.FLAGS = cdlOFNHideReadOnly
' Set filters
cd1.Filter = "All Files (*.*)|*.*|Text Files" & _
"(*.txt)|*.txt|Batch Files (*.bat)|*.bat"
' Specify default filter
cd1.FilterIndex = 2

' Display the Open dialog box
cd1.DialogTitle = "Select Script File"
cd1.FileName = vbNullString
If VerifySubDirectory("Scripts", True) Then
    cd1.InitDir = App.path + "\Scripts"
End If

cd1.ShowOpen

On Error GoTo 0

strFile = cd1.FileName

nReply = MsgBox("Run script " + strFile + " right after update?", vbYesNoCancel, "Script Timer")
If nReply = vbCancel Then
    GoTo ButtonTimerExit
End If

If nReply = vbYes Then
    BatchScriptUpdate = strFile
    GoTo ButtonTimerExit
End If

Do
    strReply = InputBox("Run time prior to update (in seconds)")
    If Val(strReply) <= 0 Then
        nReply = MsgBox("Input invalid - time script cancelled", vbRetryCancel + vbExclamation)
        If nReply = vbCancel Then
            Exit Sub
        Else
            strReply = vbNullString
        End If
    End If
Loop Until Val(strReply) > 0

BatchScript1File = strFile
BatchScript1Time = Val(strReply)

    
'User pressed the Cancel button
ButtonTimerExit:

'display final settings
strReply = "Current Settings:" + vbCr + vbCr + TimerHeader + vbCr
MsgBox strReply, vbOKOnly, "Timed Script Settings"
End Sub

Private Function TimerHeader() As String
Dim strTemp As String
strTemp = "Script to run "
If Len(BatchScript1File) > 0 Then
    strTemp = strTemp + "at update minus " + CStr(BatchScript1Time) + " seconds: " + vbCr + BatchScript1File
Else
    strTemp = strTemp + "prior to update: None"
End If
strTemp = strTemp + vbCr + vbCr + "Script to run post update: "
If Len(BatchScriptUpdate) > 0 Then
    strTemp = strTemp + vbCr + BatchScriptUpdate
Else
    strTemp = strTemp + "None"
End If

TimerHeader = strTemp
End Function

Public Sub AddCommand(strCmd As String)
rt1.Text = rt1.Text + strCmd + vbLf
End Sub

Public Sub ExectuteBatchScript(strFile As String)
' Dim FileName As String    8/2003 efj  removed

rt1.LoadFile strFile, rtfText
ButtonRun
Unload Me
End Sub
