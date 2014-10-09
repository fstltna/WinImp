VERSION 5.00
Begin VB.Form frmPromptSimpleTerritory 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1935
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTool 
      Cancel          =   -1  'True
      Caption         =   "Tool"
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox txtTerrNumber 
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.ComboBox cbField 
      Height          =   315
      ItemData        =   "frmPromptSimpleTerritory.frx":0000
      Left            =   1560
      List            =   "frmPromptSimpleTerritory.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtMultOrigin 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   1335
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
      Left            =   3720
      TabIndex        =   0
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblTerrNumber 
      Alignment       =   1  'Right Justify
      Caption         =   "Territory Number"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblSector 
      Alignment       =   1  'Right Justify
      Caption         =   "Sectors"
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblTerr 
      Alignment       =   1  'Right Justify
      Caption         =   "Territory Field"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Territory"
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
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmPromptSimpleTerritory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log
'250904 rjk: Created to provide simple interface for territory.
'210306 rjk: Switched SendFullDumpCommand to GetSectors

Private Sub cmdTool_Click()
Set frmDrawMap.PromptForm = frmPromptTerritory
'Put form in proper place
frmDrawMap.PromptForm.Left = frmDrawMap.Left + (frmDrawMap.Width - frmDrawMap.PromptForm.Width) \ 2
frmDrawMap.PromptForm.top = frmDrawMap.top + frmDrawMap.Height - frmDrawMap.PromptForm.Height
'Dialog box will take it from here
frmDrawMap.PromptForm.Show vbModeless
Unload Me
End Sub

Private Sub Form_Load()
' Set parent for the toolbar to display on top of:
' Dim success As Long  removed 8/12/03 efj
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
cbField.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

Private Sub cmdOK_Click()
SetMultipleSectorTerritory txtMultOrigin.Text, CStr(cbField.ItemData(cbField.ListIndex)), txtTerrNumber

'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
GetSectors
GetCurrentStrength tsSectors
frmEmpCmd.SubmitEmpireCommand "db2", False

Unload Me
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
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

Private Sub Form_Activate()
txtMultOrigin.SetFocus
End Sub


Private Sub SetMultipleSectorTerritory(strMultOrigin As String, strTerr As String, strTerrNumber As String)
Dim beginPos As Integer
Dim endPos As Integer

If InStr(strMultOrigin, "\") > 0 Then
    beginPos = 1
    endPos = InStr(strMultOrigin, "\")
    While endPos > 0
        SetTerritory Mid$(strMultOrigin, beginPos, endPos - beginPos), strTerr, strTerrNumber
        beginPos = endPos + 1
        endPos = InStr(beginPos, strMultOrigin, "\")
    Wend
    SetTerritory Mid$(strMultOrigin, beginPos), strTerr, strTerrNumber
Else
    SetTerritory strMultOrigin, strTerr, strTerrNumber
End If
End Sub

Private Sub SetTerritory(strStart As String, strTerr As String, strTerrNumber As String)
frmEmpCmd.SubmitEmpireCommand "territory " + strStart + " " + strTerrNumber + " " + strTerr, True
End Sub

