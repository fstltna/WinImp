VERSION 5.00
Begin VB.Form frmPromptThresh 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2205
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkFoodRequired 
      Caption         =   "Use Food Required to determine Threshold"
      Height          =   375
      Left            =   1440
      TabIndex        =   14
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CheckBox chkSupply 
      Caption         =   "Auto Push/Pull supply from warehouse to meet new threshold"
      Height          =   375
      Left            =   1440
      TabIndex        =   12
      Top             =   1680
      Width           =   2775
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
      Left            =   7320
      TabIndex        =   11
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.ListBox cbCom 
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
      Height          =   1740
      ItemData        =   "frmPromptThresh.frx":0000
      Left            =   4320
      List            =   "frmPromptThresh.frx":0002
      TabIndex        =   9
      Top             =   120
      Width           =   2880
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtMultOrigin 
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblFoodPercent 
      Caption         =   "Percent Above Food Required"
      Height          =   375
      Left            =   3120
      TabIndex        =   15
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblDesc 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   1605
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblThresh 
      Caption         =   "Current Threshold: 999"
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
      Left            =   1920
      TabIndex        =   10
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Set"
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
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "for"
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
      Left            =   3840
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "at"
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
      Left            =   1800
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "in"
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
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Threshold"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmPromptThresh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log
'270903 rjk: General reformatting
'171003 rjk: Added Strength fields to Sector database
'191003 rjk: Added Multiple Sector Selection
'231003 rjk: Added Sector count for Multiple Sectors
'            Added sector description field
'201103 rjk: Added option to control strength updates
'251103 rjk: Added the ability to set the food levels based on food required
'240104 rjk: Renamed the Supply check, switch to item object, added multiple sector selection,
'            Removed the field transaction table, update the registry saving (add food required)
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
'210306 rjk: Switched SendFullDumpCommand to GetSectors

Private ftranslation(13) As String

Private Sub cbCom_Click()
SetThreshLabel
End Sub

Private Sub cbCom_KeyPress(KeyAscii As Integer)
Dim i As Integer, ch As String
On Error Resume Next

ch = LCase$(Chr(KeyAscii))
If ch >= "a" And ch < "z" Then
    For i = 0 To cbCom.ListCount - 1
        If Left$(cbCom.List(i), 1) = ch Then
            cbCom.ListIndex = i
        End If
    Next i
    KeyAscii = 0
End If
End Sub

'112503 rjk: Added the ability to set the food levels based on food required
Private Sub chkFoodRequired_Click()
SetThreshLabel
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Public Sub cmdOK_Click()
Dim strCmd As String
Dim iComm As EmpItem

Set iComm = Items.FindByFormName(cbCom.Text)

SetMultipleSectorThreshold iComm, txtMultOrigin.Text, Val(txtNum.Text), chkFoodRequired <> vbUnchecked, chkSupply <> vbUnchecked

'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
'frmEmpCmd.SubmitEmpireCommand "dump " + txtMultOrigin.Text, False
GetSectors
'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
GetCurrentStrength tsSectors
frmEmpCmd.SubmitEmpireCommand "db2", False

Unload Me
End Sub

Public Sub SetMultipleSectorThreshold(iComm As EmpItem, strMultOrigin As String, lThres As Long, bFoodRequired As Boolean, bSupply As Boolean)
Dim beginPos As Integer
Dim endPos As Integer

If InStr(strMultOrigin, "\") > 0 Then '101903 Multiple Sectors Selected
    beginPos = 1
    endPos = InStr(strMultOrigin, "\")
    While endPos > 0
        SetThreshold Mid$(strMultOrigin, beginPos, endPos - beginPos), iComm, lThres, bFoodRequired, bSupply
        beginPos = endPos + 1
        endPos = InStr(beginPos, strMultOrigin, "\")
    Wend
    SetThreshold Mid$(strMultOrigin, beginPos), iComm, lThres, bFoodRequired, bSupply
Else
    SetThreshold strMultOrigin, iComm, lThres, bFoodRequired, bSupply
End If
End Sub

Private Sub SetThreshold(strStart As String, iComm As EmpItem, lThresh As Long, bFoodRequired As Boolean, bSupply As Boolean)
'112503 rjk: Added the ability to set the food levels based on food required
If iComm.Letter = "f" And bFoodRequired Then
    AdjustedFoodLevel strStart, lThresh
End If
frmEmpCmd.SubmitEmpireCommand "threshold " + iComm.Letter + " " + strStart + " " + CStr(lThresh), True

If bSupply Then
    MeetThreshold strStart, lThresh, iComm
End If
End Sub

Public Sub MeetThreshold(txtOrigin As String, lNewThresh As Long, iComm As EmpItem)
Dim sx As Integer, sy As Integer, DistX As Integer, DistY As Integer
Dim lShortage As Long
Dim strCmd As String
    
ParseSectors sx, sy, txtOrigin
rsSectors.Seek "=", sx, sy
lShortage = lNewThresh - iComm.DatabaseValue(rsSectors)

DistX = rsSectors.Fields("dist_x")
DistY = rsSectors.Fields("dist_y")

rsSectors.Seek "=", DistX, DistY

If Not rsSectors.NoMatch Then
    If (lShortage > 0) Then
        lShortage = Min(lShortage, iComm.DatabaseValue(rsSectors))
        strCmd = "move " + iComm.Letter + "  " + SectorString(DistX, DistY) + " " + CStr(lShortage) + " " + txtOrigin
        frmEmpCmd.SubmitEmpireCommand strCmd, False
    ElseIf (lShortage < 0) Then
        strCmd = "move " + iComm.Letter + "  " + txtOrigin + " " + CStr(0 - lShortage) + " " + SectorString(DistX, DistY)
        frmEmpCmd.SubmitEmpireCommand strCmd, False
    End If
End If
End Sub

Private Sub Form_Load()
LoadCommodityBox cbCom
cbCom.ListIndex = 1 'set to civilians

'Make Stay always on top
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags
 
If GetSetting(APPNAME, "frmPromptThresh", "Food Required", False) Then
    chkFoodRequired.Value = vbChecked
Else
    chkFoodRequired.Value = vbUnchecked
End If
If GetSetting(APPNAME, "frmPromptThresh", "Supply", False) Then
    chkSupply.Value = vbChecked
Else
    chkSupply.Value = vbUnchecked
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False

If chkFoodRequired.Value <> vbUnchecked Then
    SaveSetting APPNAME, "frmPromptThresh", "Food Required", True
Else
    SaveSetting APPNAME, "frmPromptThresh", "Food Required", False
End If
If chkSupply.Value <> vbUnchecked Then
    SaveSetting APPNAME, "frmPromptThresh", "Supply", True
Else
    SaveSetting APPNAME, "frmPromptThresh", "Supply", False
End If
End Sub

Private Sub txtNum_KeyPress(KeyAscii As Integer)
cbCom_KeyPress KeyAscii
End Sub

Private Sub txtMultOrigin_Change()
SetThreshLabel
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

Private Sub SetThreshLabel()
Dim n As Integer
Dim sx As Integer
Dim sy As Integer
' Dim owner As Integer    8/2003 efj  removed
On Error GoTo SetThreshLabel_Exit

lblThresh.Caption = vbNullString
If ParseSectors(sx, sy, txtMultOrigin) Then

    'Get Sector Information
    rsSectors.Seek "=", sx, sy
    If Not rsSectors.NoMatch Then
        n = rsSectors.Fields(Left$(cbCom.List(cbCom.ListIndex), 1) _
                    + "_dist")
        lblThresh.Caption = "Current Threshold: " + CStr(n)
        '102303 rjk: Added sector label
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

'112503 rjk: Added the ability to set the food levels based on food required
If cbCom.List(cbCom.ListIndex) = "food" And InStr(txtMultOrigin, "*") = 0 And _
   InStr(txtMultOrigin, "#") = 0 And InStr(txtMultOrigin, ":") = 0 Then
    chkFoodRequired.Visible = True
    If chkFoodRequired = vbChecked Then
        lblFoodPercent.Visible = True
        Label2.Visible = False
    Else
        lblFoodPercent.Visible = False
        Label2.Visible = True
    End If
Else
    chkFoodRequired.Visible = False
    lblFoodPercent.Visible = False
    Label2.Visible = True
End If
SetThreshLabel_Exit:
End Sub

'112503 rjk: Added the ability to set the food levels based on food required
Public Sub AdjustedFoodLevel(strStart As String, lThresh As Long)
Dim sx As Integer
Dim sy As Integer

If ParseSectors(sx, sy, strStart) Then
    rsSectors.Seek "=", sx, sy
    If Not rsSectors.NoMatch Then
        lThresh = CStr(Round(FoodRequired(rsSectors) * (1 + (lThresh / 100)), 0))
    End If
End If
End Sub

