VERSION 5.00
Begin VB.Form frmPromptTerritory 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2790
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cbField 
      Height          =   315
      ItemData        =   "frmPromptTerritory.frx":0000
      Left            =   1560
      List            =   "frmPromptTerritory.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1080
      Width           =   1095
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
      Left            =   5640
      TabIndex        =   9
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdAddNew 
      Cancel          =   -1  'True
      Caption         =   "Add New Territory"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   1560
      Width           =   2175
   End
   Begin VB.ComboBox cbTerr 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.ListBox lstSectors 
      Height          =   2205
      Left            =   3000
      TabIndex        =   4
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblField 
      Caption         =   "Select Territory Field"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Edit Territory"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Sector List"
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label2 
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
      Left            =   720
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmPromptTerritory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'081103 efj: commented out dead Sub ToggleHex
'091103 rjk: added back ToggleHex, used when marking territories
'092703 rjk: General Reformatting
'100103 rjk: Clear list when switching territories so lists do not get combined.
'101703 rjk: Added Strength fields to Sector database
'102303 rjk: Deleted lblDescSect, not used
'            Added Field combobox for selecting between terr, terr1, terr2 and terr3
'            Added batch markers so all the commands are combined on the command display
'112003 rjk: Added option to control strength updates
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
'210306 rjk: Switched SendFullDumpCommand to GetSectors

Dim strTerr As String

Public Sub ToggleHex(sx As Integer, sy As Integer, desc As String)
Dim searchindex As Long
Dim n As Integer

'seargh for item - if found delete it and exit
searchindex = (2048 + CLng(sx)) * &HFFF + (sy + 2048)
For n = 0 To lstSectors.ListCount - 1
    If lstSectors.ItemData(n) = searchindex Then
        lstSectors.RemoveItem n
        Exit Sub
    End If
Next n

'Since entry was not found, add it
lstSectors.AddItem SectorString(sx, sy) + vbTab + desc
lstSectors.ItemData(lstSectors.NewIndex) = searchindex
End Sub

Private Sub cbField_Click()
FillEditTerritoryCombo
End Sub

Private Sub cbTerr_Click()
Dim strTerrField As String '102303 rjk: Added for access to the different territory fields

lstSectors.Clear '100103 rjk: Clear list when switching territories so list do not get combined.

strTerrField = GetTerritoryFieldDatabaseName '102303 rjk: Added for access to the different territory fields
frmDrawMap.StartTerritory
rsSectors.MoveFirst
While Not rsSectors.EOF
    If rsSectors.Fields(strTerrField) = cbTerr.List(cbTerr.ListIndex) Then
        'draw hex in red
        frmDrawMap.ToggleTerritory rsSectors.Fields("x"), rsSectors.Fields("y")
        
        'add to list
        lstSectors.AddItem SectorString(rsSectors.Fields("x").Value, rsSectors.Fields("y").Value) _
                + vbTab + CStr(rsSectors.Fields("eff").Value) + "% " + colDes.Item(rsSectors.Fields("des").Value)
        lstSectors.ItemData(lstSectors.NewIndex) = (2048 + CLng(rsSectors.Fields("x"))) * &HFFF + rsSectors.Fields("y") + 2048
    End If
    rsSectors.MoveNext
Wend
End Sub

Private Sub cmdAddNew_Click()
Dim n As Integer

'find first unused territory

n = 1
While InStr(strTerr, " " + CStr(n) + " ") > 0
    n = n + 1
Wend

'now add that territory to the box
strTerr = strTerr + " " + CStr(n) + " "
cbTerr.AddItem CStr(n)
cbTerr.ListIndex = cbTerr.NewIndex
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Private Sub cmdOK_Click()
Dim strCmd As String
Dim n As Integer

'102303 rjk: Added so all the commands are combined on the command display
frmEmpCmd.SubmitEmpireCommand "bf1", False

'first wipe out the exsisting territory
'102303 rjk: Modified for Multiple Territory field support
'strCmd = "territory * ?terr=" + CStr(cbTerr.List(cbTerr.ListIndex)) + " 0"
strCmd = "territory * ?terr" + CStr(cbField.ItemData(cbField.ListIndex)) + "=" + _
    CStr(cbTerr.List(cbTerr.ListIndex)) + " 0 " + CStr(cbField.ItemData(cbField.ListIndex))

frmEmpCmd.SubmitEmpireCommand strCmd, False

'then add the new one
For n = 0 To lstSectors.ListCount - 1
    strCmd = Left$(lstSectors.List(n), InStr(lstSectors.List(n), vbTab) - 1)
    strCmd = "territory " + strCmd + " " + CStr(cbTerr.List(cbTerr.ListIndex))
    '102303 rjk: Modified for Multiple Territory field support
    strCmd = strCmd + " " + CStr(cbField.ItemData(cbField.ListIndex))
    frmEmpCmd.SubmitEmpireCommand strCmd, False
Next n
'102303 rjk: Added so all the commands are combined on the command display
frmEmpCmd.SubmitEmpireCommand "bf2", False

'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
GetSectors
'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
GetCurrentStrength tsSectors
frmEmpCmd.SubmitEmpireCommand "db2", False

Unload Me
End Sub

Private Sub Form_Load()
'Make Stay always on top
' Dim success As Long       removed 3/2003 efj
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
cbField.ListIndex = 0
FillEditTerritoryCombo '102303 rjk: Code moved to separate function to accomodate multiple territory fields
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmDrawMap.MarkingTerritory = False
frmDrawMap.DrawHexes
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

'102303 rjk: Added for access to the different territory fields
Private Function GetTerritoryFieldDatabaseName() As String
If cbField.ListIndex = 0 Then
    GetTerritoryFieldDatabaseName = "terr"
Else
    GetTerritoryFieldDatabaseName = "terr" + CStr(cbField.ItemData(cbField.ListIndex))
End If
End Function

'102303 rjk: Moved the filling of edit territory combo box to separate so it can also be called
'            cbField when difference territory field is selected.
Private Sub FillEditTerritoryCombo()
Dim strTerrField As String

strTerrField = GetTerritoryFieldDatabaseName
strTerr = vbNullString
rsSectors.MoveFirst
cbTerr.Clear
lstSectors.Clear
While Not rsSectors.EOF
    If rsSectors.Fields(strTerrField) <> 0 Then
        If InStr(strTerr, " " + CStr(rsSectors.Fields(strTerrField)) + " ") = 0 Then
            strTerr = strTerr + " " + CStr(rsSectors.Fields(strTerrField)) + " "
            cbTerr.AddItem CStr(rsSectors.Fields(strTerrField))
        End If
    End If
    rsSectors.MoveNext
Wend
End Sub
