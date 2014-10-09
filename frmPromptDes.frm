VERSION 5.00
Begin VB.Form frmPromptDes 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2760
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9120
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkClearOldThresholds 
      Caption         =   "Clear Old Thresholds when designating"
      Height          =   975
      Left            =   7800
      TabIndex        =   19
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   18
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Set Thresholds when you designate"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.ListBox cbType 
      Columns         =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2525
      IntegralHeight  =   0   'False
      Left            =   3480
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   4080
   End
   Begin VB.Frame Frame1 
      Caption         =   "Thresholds"
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   3255
      Begin VB.CheckBox chkSave 
         Caption         =   "Save as New Defaults"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtThresh 
         Height          =   285
         Index           =   3
         Left            =   2520
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtThresh 
         Height          =   285
         Index           =   2
         Left            =   2520
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtThresh 
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtThresh 
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblThresh 
         Caption         =   "Label3"
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   15
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblThresh 
         Caption         =   "Label3"
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblThresh 
         Caption         =   "Label3"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblThresh 
         Caption         =   "Label3"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.TextBox txtMultOrigin 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   7800
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7800
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblDesc 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   45
   End
   Begin VB.Label Label2 
      Caption         =   "Designate"
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
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "as"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmPromptDes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'081203 efj: removed dead variables
'092803 rjk: general reformatting.
'101403 rjk: Added check box for clearing old thresholds from old designation.
'101703 rjk: Added Strength fields to Sector database
'101903 rjk: Switched txtOrigin to txtMultOrigin to support multiple sector selection.
'            Added bFirstLoad to only do initial field selection once.
'102303 rjk: Added Sector count for Multiple Sector Selection
'110703 rjk: Removed the threshold textbox when the label is blank or n/a
'            Fixed to correctly save the threshold labels and values.
'112003 rjk: Added option to control strength updates
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
'300505 rjk: Fixed defaults 'd' oil 10 and 'j,k' iron 600.
'210306 rjk: Switched SendFullDumpCommand to GetSectors

Private strCodes As String
Private strNewDes As String
Private bAutoFill As Boolean
Private bFirstLoad As Boolean '101903 rjk: Added bFirstLoad to prevent the focus from being drawn away from the txtMultOrigin field

Private Sub cbType_Click()
If cbType.ListIndex < 0 Then
    Exit Sub
End If
strNewDes = Mid$(strCodes, cbType.ItemData(cbType.ListIndex), 1)

'Reset threshholds
SetThreshholds
End Sub

Private Sub cbType_DblClick()
cmdOK.Value = True
End Sub

Private Sub Check1_Click()
'Reset threshholds
SetThreshholds
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Public Sub cmdOK_Click()
Dim X As Integer
Dim beginPos As Integer
Dim endPos As Integer

'Save the new defaults, if indicated
If chkSave.Visible And chkSave.Value = vbChecked Then
    For X = 0 To 3 '110703 rjk: Fixed to save the correct threshold labels
        SaveSetting APPNAME, "Thresholds", strNewDes + CStr(X + 1) + " type", lblThresh(X).Caption
        SaveSetting APPNAME, "Thresholds", strNewDes + CStr(X + 1) + " level", txtThresh(X).Text
    Next X
End If

'101803 rjk: Logic redone and support function SetDesignation added for Multiple Sectors Selection
If InStr(txtMultOrigin.Text, "\") > 0 Then
    beginPos = 1
    endPos = InStr(txtMultOrigin.Text, "\")
    While endPos > 0
        SetDesignation (Mid$(txtMultOrigin.Text, beginPos, endPos - beginPos))
        beginPos = endPos + 1
        endPos = InStr(beginPos, txtMultOrigin.Text, "\")
    Wend
    SetDesignation (Mid$(txtMultOrigin.Text, beginPos))
Else
    SetDesignation (txtMultOrigin.Text)
End If

'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
GetSectors
'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
GetCurrentStrength tsSectors
frmEmpCmd.SubmitEmpireCommand "db2", False
Unload Me
End Sub

'101903 rjk: Added Multiple Sector Selection
Private Sub SetDesignation(strSector As String)
Dim sx As Integer
Dim sy As Integer
Dim strCmd As String
Dim strComm As String
Dim X As Integer

'101403 rjk: Clear old thresholds from old designation when setting new thresholds
If ParseSectors(sx, sy, strSector) And chkClearOldThresholds Then
    'Get Sector Information
    rsSectors.Seek "=", sx, sy
    If Not rsSectors.NoMatch Then
        For X = 0 To 3 '110703 rjk: Switch to 0 to 3 to standardize with the rest of the code
            strComm = GetSetting(APPNAME, "Thresholds", rsSectors.Fields("des").Value + _
                                 CStr(X + 1) + " type", vbNullString)
            If Len(strComm) > 0 And strComm <> "n/a" Then
                strCmd = "threshold " + strComm + " " + txtMultOrigin.Text + " " + " 0"
                frmEmpCmd.SubmitEmpireCommand strCmd, False
            End If
        Next X
    End If
End If

'Set Thresholds first (since des can change optional params)
For X = 0 To 3
    If txtThresh(X).Visible And Len(txtThresh(X).Text) > 0 Then
        strCmd = "threshold " + lblThresh(X).Caption + " " + strSector + " " + txtThresh(X).Text
        frmEmpCmd.SubmitEmpireCommand strCmd, True
    End If
Next X

'now do the designation
frmEmpCmd.SubmitEmpireCommand "designate " + strSector + " " + strNewDes, True
End Sub

Private Sub cbType_KeyPress(KeyAscii As Integer)
Dim ch As String
Dim n As Integer
Dim i As Integer

ch = Chr$(KeyAscii)
n = InStr(strCodes, ch)
If n > 0 Then
strNewDes = ch
For i = 0 To cbType.ListCount - 1
    If cbType.ItemData(i) = n Then
        cbType.ListIndex = i
    End If
Next i
End If
End Sub

Private Sub Form_Activate()
If bFirstLoad Then '101903 rjk: Added bFirstLoad to prevent the focus from being drawn away from the txtMultOrigin field
    bFirstLoad = False
    cbType.SetFocus
End If
End Sub

Private Sub Form_Load()
bFirstLoad = True
strCodes = LoadTypebox(cbType)
cbType.ListIndex = 0
chkSave.Visible = False

'Make Stay always on top
' Dim success As Long  removed 8/12/03 efj
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

Private Sub lblThresh_Click(Index As Integer)
Dim newthresh As String

'Allow a thresh title to be changed by clicking
newthresh = InputBox("Set threshold type", "Change Threshold", lblThresh(Index).Caption)
If newthresh <> lblThresh(Index).Caption Then
    lblThresh(Index).Caption = newthresh
    If Len(newthresh) > 0 And newthresh <> "n/a" Then
        txtThresh(Index).Visible = True
    Else '110703 rjk: Removed the textbox when the label is blank or n/a
        txtThresh(Index).Visible = False
    End If
    If bAutoFill = False And Not chkSave.Visible Then
        chkSave.Visible = True
    End If
End If
End Sub

Private Sub txtMultOrigin_Change()
Dim sx As Integer
Dim sy As Integer
On Error Resume Next

If ParseSectors(sx, sy, txtMultOrigin.Text) Then
    'Get Sector Information
    rsSectors.Seek "=", sx, sy
    If Not rsSectors.NoMatch Then
        lblDesc = "Sector " + txtMultOrigin.Text + " is currently a "
        lblDesc = lblDesc + CStr(rsSectors.Fields("eff").Value) + "% "
        lblDesc = lblDesc + colDes.Item(rsSectors.Fields("des").Value)
    End If
'102303 rjk: Added Sector count for Multiple Sector Selection
ElseIf InStr(txtMultOrigin.Text, "\") > 0 Then
    lblDesc = CStr(CountMultipleSectors(txtMultOrigin.Text)) + " Sectors"
Else
    lblDesc = ""
End If
End Sub

Private Sub txtMultOrigin_DblClick()
Load frmPromptSectors
frmPromptSectors.strSectors = txtMultOrigin.Text
frmPromptSectors.SetControls
frmPromptSectors.Caption = "Select Sectors"
frmPromptSectors.Left = frmDrawMap.Left + frmDrawMap.Width - frmPromptSectors.Width
frmPromptSectors.top = frmDrawMap.top + (frmDrawMap.Height - frmPromptSectors.Height) \ 2
frmPromptSectors.Show vbModeless
End Sub

Private Sub SetThreshholds()
Dim thresh As String
Dim n As Integer

bAutoFill = True
'Clear all thresholds
lblThresh(0) = "n/a"
txtThresh(0).Visible = False
txtThresh(0).Text = vbNullString
lblThresh(1) = "n/a"
txtThresh(1).Visible = False
txtThresh(1).Text = vbNullString
lblThresh(2) = "n/a"
txtThresh(2).Visible = False
txtThresh(2).Text = vbNullString
lblThresh(3) = "n/a"
txtThresh(3).Visible = False
txtThresh(3).Text = vbNullString

chkSave.Visible = False

'If thresh use is off, ignore
If Not Check1.Value = vbChecked Then
    Exit Sub
End If

'get threshhold setting from the registry
If Len(GetSetting(APPNAME, "Thresholds", strNewDes + CStr(1) + " type", vbNullString)) = 0 Then
    SetDefaultThresholds strNewDes
Else
    For n = 0 To 3
        thresh = GetSetting(APPNAME, "Thresholds", strNewDes + CStr(n + 1) + " type", vbNullString)
        If Len(thresh) > 0 And thresh <> "n/a" Then
            lblThresh(n) = thresh
            txtThresh(n).Visible = True
            txtThresh(n).Text = GetSetting(APPNAME, "Thresholds", strNewDes + CStr(n + 1) + " level", vbNullString)
        Else '110703 rjk: Removed the textbox when the label is blank or n/a
            txtThresh(n).Visible = False
        End If
    Next n
End If
    
bAutoFill = False
End Sub

Private Sub SetDefaultThresholds(strNewDes As String)
'now turn on the thresholds based on designations
Select Case strNewDes
    Case "m"
        lblThresh(0) = "iron"
        txtThresh(0).Visible = True
        txtThresh(0).Text = "1"
    Case "k"
        lblThresh(0) = "iron"
        txtThresh(0).Visible = True
        txtThresh(0).Text = "600"
        lblThresh(1) = "hcm"
        txtThresh(1).Visible = True
        txtThresh(1).Text = "1"
    Case "j"
        lblThresh(0) = "iron"
        txtThresh(0).Visible = True
        txtThresh(0).Text = "600"
        lblThresh(1) = "lcm"
        txtThresh(1).Visible = True
        txtThresh(1).Text = "1"
    Case "g"
        lblThresh(0) = "dust"
        txtThresh(0).Visible = True
        txtThresh(0).Text = "1"
    Case "i"
        lblThresh(0) = "lcm"
        txtThresh(0).Visible = True
        txtThresh(0).Text = "300"
        lblThresh(1) = "hcm"
        txtThresh(1).Visible = True
        txtThresh(1).Text = "150"
        lblThresh(2) = "shell"
        txtThresh(2).Visible = True
        txtThresh(2).Text = "1"
    Case "d"
        lblThresh(0) = "oil"
        txtThresh(0).Visible = True
        txtThresh(0).Text = "10"
        lblThresh(1) = "lcm"
        txtThresh(1).Visible = True
        txtThresh(1).Text = "50"
        lblThresh(2) = "hcm"
        txtThresh(2).Visible = True
        txtThresh(2).Text = "100"
        lblThresh(3) = "gun"
        txtThresh(3).Visible = True
        txtThresh(3).Text = "1"
    Case "o"
        lblThresh(0) = "oil"
        txtThresh(0).Visible = True
        txtThresh(0).Text = "1"
    Case "f"
        lblThresh(0) = "hcm"
        txtThresh(0).Visible = True
        txtThresh(0).Text = "100"
        lblThresh(1) = "gun"
        txtThresh(1).Visible = True
        txtThresh(1).Text = "7"
        lblThresh(2) = "shell"
        txtThresh(2).Visible = True
        txtThresh(2).Text = "50"
        lblThresh(3) = "military"
        txtThresh(3).Visible = True
        txtThresh(3).Text = "60"
       
    Case "*"
        lblThresh(0) = "petrol"
        txtThresh(0).Visible = True
        txtThresh(0).Text = "1000"
        lblThresh(1) = "lcm"
        txtThresh(1).Visible = True
        txtThresh(1).Text = "200"
        lblThresh(2) = "hcm"
        txtThresh(2).Visible = True
        txtThresh(2).Text = "100"
        lblThresh(3) = "shell"
        txtThresh(3).Visible = True
        txtThresh(3).Text = "200"
    Case "!"
        lblThresh(0) = "lcm"
        txtThresh(0).Visible = True
        txtThresh(0).Text = "200"
        lblThresh(1) = "hcm"
        txtThresh(1).Visible = True
        txtThresh(1).Text = "100"
        lblThresh(2) = "gun"
        txtThresh(2).Visible = True
        txtThresh(2).Text = vbNullString
        lblThresh(3) = "shell"
        txtThresh(3).Visible = True
        txtThresh(3).Text = vbNullString
    Case "e"
        lblThresh(0) = "civs"
        txtThresh(0).Visible = True
        txtThresh(0).Text = "768"
        lblThresh(1) = "mil"
        txtThresh(1).Visible = True
        txtThresh(1).Text = "118"
    Case "t"
        lblThresh(0) = "oil"
        txtThresh(0).Visible = True
        txtThresh(0).Text = vbNullString
        lblThresh(1) = "lcm"
        txtThresh(1).Visible = True
        txtThresh(1).Text = vbNullString
        lblThresh(2) = "dust"
        txtThresh(2).Visible = True
        txtThresh(2).Text = vbNullString
    Case "l"
        lblThresh(0) = "lcm"
        txtThresh(0).Visible = True
        txtThresh(0).Text = vbNullString
    Case "n"
        lblThresh(0) = "rad"
        txtThresh(0).Visible = True
        txtThresh(0).Text = "999"
    Case "p"
        lblThresh(0) = "lcm"
        txtThresh(0).Visible = True
        txtThresh(0).Text = vbNullString
    Case "%"
        lblThresh(0) = "petrol"
        txtThresh(0).Visible = True
        txtThresh(0).Text = "1"
        lblThresh(1) = "oil"
        txtThresh(1).Visible = True
        txtThresh(1).Text = vbNullString
    Case "b"
        lblThresh(0) = "dust"
        txtThresh(0).Visible = True
        txtThresh(0).Text = vbNullString
    Case "a"
        lblThresh(0) = "food"
        txtThresh(0).Visible = True
        txtThresh(0).Text = "60"
    Case "u"
        lblThresh(0) = "rads"
        txtThresh(0).Visible = True
        txtThresh(0).Text = "1"
    End Select
End Sub

Private Sub txtThresh_Change(Index As Integer)
If bAutoFill = True Or chkSave.Visible Then
    Exit Sub
End If

chkSave.Visible = True
End Sub
