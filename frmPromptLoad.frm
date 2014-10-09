VERSION 5.00
Begin VB.Form frmPromptLoad 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2424
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   8424
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2424
   ScaleWidth      =   8424
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdHelp 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   40
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox txtUnit3 
      Height          =   285
      Left            =   5880
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Land(s)"
      Height          =   255
      Left            =   7200
      TabIndex        =   20
      Top             =   1680
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Plane(s)"
      Height          =   255
      Left            =   5880
      TabIndex        =   19
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtThresh 
      Height          =   285
      Index           =   3
      Left            =   1320
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtThresh 
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txtThresh 
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtThresh 
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtThresh 
      Height          =   285
      Index           =   4
      Left            =   3240
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtThresh 
      Height          =   285
      Index           =   5
      Left            =   3240
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtThresh 
      Height          =   285
      Index           =   6
      Left            =   1320
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtThresh 
      Height          =   285
      Index           =   7
      Left            =   3240
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtThresh 
      Height          =   285
      Index           =   8
      Left            =   3240
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtThresh 
      Height          =   285
      Index           =   9
      Left            =   3240
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txtThresh 
      Height          =   285
      Index           =   10
      Left            =   5160
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtThresh 
      Height          =   285
      Index           =   11
      Left            =   5160
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtThresh 
      Height          =   285
      Index           =   12
      Left            =   5160
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txtThresh 
      Height          =   285
      Index           =   13
      Left            =   5160
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtUnit2 
      Height          =   285
      Left            =   5760
      TabIndex        =   4
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      TabIndex        =   21
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7080
      TabIndex        =   22
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtUnit 
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   5760
      TabIndex        =   41
      Top             =   960
      Width           =   2535
      Begin VB.Label lblMobLeft 
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label lblMcost 
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblPathCost 
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.TextBox txtMultOrigin 
      Height          =   285
      Left            =   7080
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtMultDest 
      Height          =   285
      Left            =   7200
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "to ship"
      Height          =   255
      Left            =   5160
      TabIndex        =   39
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblThresh 
      Caption         =   "Label3"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   38
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblThresh 
      Caption         =   "Label3"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   37
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblThresh 
      Caption         =   "Label3"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   36
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblThresh 
      Caption         =   "Label3"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   35
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblThresh 
      Caption         =   "Label3"
      Height          =   255
      Index           =   4
      Left            =   2040
      TabIndex        =   34
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblThresh 
      Caption         =   "Label3"
      Height          =   255
      Index           =   5
      Left            =   2040
      TabIndex        =   33
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblThresh 
      Caption         =   "Label3"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   32
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblThresh 
      Caption         =   "Label3"
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   31
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblThresh 
      Caption         =   "Label3"
      Height          =   255
      Index           =   8
      Left            =   2040
      TabIndex        =   30
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblThresh 
      Caption         =   "Label3"
      Height          =   255
      Index           =   9
      Left            =   2040
      TabIndex        =   29
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblThresh 
      Caption         =   "Label3"
      Height          =   255
      Index           =   10
      Left            =   3960
      TabIndex        =   28
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblThresh 
      Caption         =   "Label3"
      Height          =   255
      Index           =   11
      Left            =   3960
      TabIndex        =   27
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblThresh 
      Caption         =   "Label3"
      Height          =   255
      Index           =   12
      Left            =   3960
      TabIndex        =   26
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblThresh 
      Caption         =   "Label3"
      Height          =   255
      Index           =   13
      Left            =   3960
      TabIndex        =   25
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "on ship"
      Height          =   255
      Left            =   1440
      TabIndex        =   23
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmPromptLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'081203 efj: removed dead variables
'091803 rjk: Added Unit grid selection when activating this form
'092103 rjk: Updated the information when the origin changes.
'092803 rjk: Made OK button the default.
'            Switched grid selection to use SetUnitType and DisplayFirstSectorPanel.
'            General reformatting.
'            Move the field selection logic from frmDrawMap to the form.
'            Fixed to logic to account for ltend loading lunit from a ship
'            Added logic for multiple units being loaded, ensure there are all at the same location
'100903 rjk: Added bar to the list of database fields that is used to get the current quantities from the database
'            Reduce the civilian value by one so the sector is not abandoned when double clicking.
'101703 rjk: Added Strength fields to Sector database
'111803 rjk: Added Multiple Sector Selection to the MoveAll command. Added -/+ logic to move quantities.
'            Added Commodity Total display using this form.  Rearranged form to make it smaller.
'            Added -/+ quantity logic to load/unload/lload/lunload.
'112003 rjk: Added option to control strength updates
'112103 rjk: Fixed so the quantities are cleared for CommodityTotal or MoveAll when no sector is present.
'121303 rjk: Switch from GetCommodityValueFromDB to Item.DatabaseValue and Item class
'240104 rjk: Fixed the commodities to use a single letter for command.
'280204 rjk: Fixed Mob., Work. and Avail. for setsector, changed to lower case letters.
'090404 rjk: Added all sectors option to Commodity Total
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
'270904 rjk: Enabled land unit loading for 2K4.
'180206 rjk: Replace pdump, ldump, and sdump with GetPlanes, GetLandUnits and GetShips.
'210306 rjk: Switched SendFullDumpCommand to GetSectors
'110307 rjk: Filter the load/unload/lload/lunload commands to relevant units only.
'190508 rjk: Added sector total to CommodityTotal
'310508 rjk: Allow loading of land units on trains for Heavy Metal II
'020608 rjk: Allow unloading of land units on trains for Heavy Metal II

Public strCmd As String
Public strSector As String

Private Sub cmdCancel_Click()
frmDrawMap.DisplayFirstSectorPanel
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Public Sub cmdOK_Click()
Dim n As Integer
Dim X As Integer
Dim strCmd2 As String
Dim strCmd3 As String
Dim strItem As String
Dim beginPos As Integer
Dim endPos As Integer
On Error Resume Next

'If this is a tend command, this holds the ship to be tended to
If txtUnit3.Visible = True Then
    strCmd3 = txtUnit3.Text
Else
    strCmd3 = vbNullString
End If

frmEmpCmd.SubmitEmpireCommand "bf1", False
For n = 0 To 13
    If Len(txtThresh(n)) > 0 Then
        strItem = lblThresh(n).Caption
        
        'Remove the amount portion of the caption
        X = InStr(strItem, " (")
        If X > 0 Then
            strItem = Trim$(Left$(strItem, X))
        End If
                    
        'If Left$(strItem, 1) = "u" Then
        '    strItem = "un"
        'End If
        Dim eiItem As EmpItem
        Set eiItem = Items.FindByFormLabel(strItem)
        If Not (eiItem Is Nothing) Then
            strItem = eiItem.Letter
        End If
        If strCmd = "move" Then
            If InStr(txtMultOrigin.Text, "\") > 0 Then '111803 Multiple Sectors Selected
                beginPos = 1
                endPos = InStr(txtMultOrigin.Text, "\")
                While endPos > 0
                    frmPromptMove.MoveCommodity Mid$(txtMultOrigin.Text, beginPos, endPos - beginPos), txtMultDest.Text, strItem, txtThresh(n).Text, False
                    beginPos = endPos + 1
                    endPos = InStr(beginPos, txtMultOrigin.Text, "\")
                Wend
                frmPromptMove.MoveCommodity Mid$(txtMultOrigin.Text, beginPos), txtMultDest.Text, Items.FindByFormLabel(strItem).Letter, txtThresh(n).Text, False
            Else
                frmPromptMove.MoveCommodity txtMultOrigin.Text, txtMultDest.Text, strItem, txtThresh(n).Text, False
            End If
            'strCmd2 = Trim$(strCmd + " " + strItem + " " + txtMultOrigin.Text + " " + txtThresh(n).Text + " " + txtMultDest.Text)
        ElseIf strCmd = "give" Or strCmd = "setresource" Or strCmd = "setsector" Then
            If InStr(txtMultOrigin.Text, "\") > 0 Then '111803 Multiple Sectors Selected
                beginPos = 1
                endPos = InStr(txtMultOrigin.Text, "\")
                While endPos > 0
                    SetDeityLevel strCmd, Mid$(txtMultOrigin.Text, beginPos, endPos - beginPos), strItem, txtThresh(n).Text
                    beginPos = endPos + 1
                    endPos = InStr(beginPos, txtMultOrigin.Text, "\")
                Wend
                SetDeityLevel strCmd, Mid$(txtMultOrigin.Text, beginPos), strItem, txtThresh(n).Text
            Else
                SetDeityLevel strCmd, txtMultOrigin.Text, strItem, txtThresh(n).Text
            End If
            'strCmd2 = Trim$(strCmd + " " + strItem + " " + txtMultOrigin.Text + " " + txtThresh(n).Text)
        ElseIf strCmd = "load" Or strCmd = "unload" Or strCmd = "lload" Or strCmd = "lunload" Then
            If InStr(txtThresh(n).Text, "-") > 0 Or InStr(txtThresh(n).Text, "+") > 0 Then '111803 -/+ quantity logic
                LoadCommodity strCmd, txtMultOrigin.Text, txtUnit.Text, strItem, txtThresh(n).Text
            Else
                strCmd2 = Trim$(strCmd + " " + strItem + " " + txtUnit.Text + " " + txtThresh(n).Text)
                frmEmpCmd.SubmitEmpireCommand strCmd2, True
            End If
            'strCmd2 = Trim$(strCmd + " " + strItem + " " + txtUnit.Text + " " + txtThresh(n).Text + " " + strCmd3)
        Else
            strCmd2 = Trim$(strCmd + " " + strItem + " " + txtUnit.Text + " " + txtThresh(n).Text + " " + txtUnit3.Text)
            frmEmpCmd.SubmitEmpireCommand strCmd2, True
        End If
    End If
Next n

If strCmd <> "move" Then
    If Option1.Value And Len(txtUnit2) > 0 Then
        strCmd2 = Trim$(strCmd + " plane " + txtUnit.Text + " " + txtUnit2.Text + " " + strCmd3)
        frmEmpCmd.SubmitEmpireCommand strCmd2, True
    End If
    
    If Option2.Value And Len(txtUnit2) > 0 Then
        strCmd2 = Trim$(strCmd + " land " + txtUnit.Text + " " + txtUnit2.Text + " " + strCmd3)
        frmEmpCmd.SubmitEmpireCommand strCmd2, True
    End If
End If
frmEmpCmd.SubmitEmpireCommand "bf2", False

'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
If strCmd <> "tend" Then
    GetSectors
    '101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
    GetCurrentStrength tsSectors
End If
If strCmd = "lload" Or strCmd = "lunload" Or strCmd = "ltend" Then
    GetLandUnits
ElseIf strCmd = "load" Or strCmd = "unload" Or strCmd = "tend" Or strCmd = "ltend" Then
    GetShips
    If (Option2.Value And Len(txtUnit2) > 0) Then
        GetLandUnits
    End If
End If

'dump plane changes if any where loaded/unloaded
If strCmd <> "move" Then
    If Option1.Value And Len(txtUnit2) > 0 Then
        GetPlanes
    End If
End If

frmEmpCmd.SubmitEmpireCommand "db2", False

frmDrawMap.DisplayFirstSectorPanel
Unload Me
End Sub

Private Sub Form_Activate()
If (strCmd = "lload" And Not (frmOptions.b2K4 Or frmOptions.bHeavyMetal)) Or (strCmd = "lunload" And Not frmOptions.bHeavyMetal) Or strCmd = "ltend" Then
    Option2.Visible = False
End If

If strCmd = "move" Then
    Frame1.Visible = True
Else
    Frame1.Visible = False
End If
End Sub

Private Sub Form_Load()
' Set parent for the toolbar to display on top of:
' Dim success As Long  removed 8/12/03 efj
SetLabels

Dim n As Integer
For n = 0 To 13
    txtThresh(n).Text = vbNullString
Next n

Select Case strCmd
Case "move"
    txtMultOrigin.Move lblThresh(5).Left, txtUnit.top
    txtMultDest.Move lblThresh(10).Left, txtUnit3.top
    Label3.Move txtThresh(5).Left, txtUnit3.top
    Label3.Alignment = 2
    txtUnit.Visible = False
    txtUnit3.Visible = False
    txtUnit2.Visible = False
    Option1.Visible = False
    Option2.Visible = False
    txtMultOrigin.Visible = True
    txtMultDest.Visible = True
    Label1.Caption = "from"
    Label3.Caption = "to"
    Label2.Caption = "Move"
    Label3.Visible = True
    'txtMultOrigin.Text = SectorString(MxPos, MyPos) 092103 rjk: Switched to SelX and SelY for top level menu
    txtMultOrigin.Text = strSector
    
    SetSectorLabels
    Left = frmDrawMap.Left + (frmDrawMap.Width - Width) \ 2
    top = frmDrawMap.top + frmDrawMap.Height - Height
Case "unload"
    Label2 = "Unload"
    Label1 = "from ship"
Case "give", "setsector", "setresource"
    Select Case strCmd
    Case "give"
        Label2.Caption = "Give"
    Case "setsector"
        Label2.Caption = "Set Sector"
    Case "setresource"
        Label2.Caption = "Set Resource"
    End Select
    txtMultOrigin.Move lblThresh(5).Left, txtUnit.top
    txtMultDest.Move lblThresh(10).Left, txtUnit3.top
    Label3.Move txtThresh(5).Left, txtUnit3.top
    Label3.Alignment = 2
    txtUnit.Visible = False
    txtUnit3.Visible = False
    txtUnit2.Visible = False
    Option1.Visible = False
    Option2.Visible = False
    txtMultOrigin.Visible = True
    txtMultDest.Visible = False
    Label1.Caption = "to"
    Label3.Caption = vbNullString
    Label3.Visible = True
    'txtMultOrigin.Text = SectorString(MxPos, MyPos) 092103 rjk: Switched to SelX and SelY for top level menu
    txtMultOrigin.Text = strSector
    SetSectorLabels
Case "lload"
    Label1.Caption = "on land"
    SetSectorLabels
Case "ltend"
    Label2.Caption = "Tend"
    Label1.Caption = "from ship"
    Label3.Caption = "to land"
    Label3.Visible = True
    txtUnit3.Visible = True
Case "lunload"
    Label2 = "Unload"
    Label1 = "from land"
Case "load"
    SetSectorLabels
Case "tend"
    Label2.Caption = "Tend"
    Label1.Caption = "from ship"
    Label3.Visible = True
    txtUnit3.Visible = True
Case "commoditytotal"
    Label1.Move lblThresh(5).Left, txtUnit.top
    txtMultOrigin.Move Label1.Left + Label1.Width, txtUnit.top
    txtUnit.Visible = False
    txtUnit3.Visible = False
    txtUnit2.Visible = False
    Option1.Visible = False
    Option2.Visible = False
    txtMultOrigin.Visible = True
    txtMultDest.Visible = False
    Label1.Caption = "for"
    Label2.Caption = "Commodity Total"
    Label3.Visible = True
    Label3.Width = 1000
    Label3.Caption = "0 Sectors"
    txtMultOrigin.Text = strSector
    cmdOK.Visible = False
    cmdCancel.Caption = "Close"
    For n = 0 To 13
        txtThresh(n).Enabled = False
        txtThresh(n).Left = txtThresh(n).Left - txtThresh(n).Width
        txtThresh(n).Width = txtThresh(n).Width + txtThresh(n).Width
    Next n
    
    SetSectorLabels
    Left = frmDrawMap.Left + (frmDrawMap.Width - Width) \ 2
    top = frmDrawMap.top + frmDrawMap.Height - Height
End Select
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
'Option2.Visible = True 092103 rjk: should be deal with above

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

Public Sub SetSectorLabels() '8/2003 efj  removed,  091103 rjk added back - it is used
Dim sx As Integer
Dim sy As Integer
Dim t As Variant, i As Integer
'100903 rjk: Added bar to the list so it is filled as well
t = Array("mil", "gun", "shell", "food", "civ", "uw", "iron", "lcm", "hcm", "oil", "dust", "pet", "rad", "bar")

SetLabels

'Command: setsector
'Give What (iron, gold, oil, uranium, fertility, owner, eff., mob., work, avail., oldown, mines)?

'If the sector info is correct, fill with what available in sector (for Load, lload)
If strCmd = "commoditytotal" Then '111803 Commodity Total
    Dim commodityTotal() As Long
    Dim beginPos As Integer, endPos As Integer
    Dim sectorCount As Integer
    
    t = Array("mil", "gun", "shell", "food", "civ", "uw", "iron", "lcm", "hcm", "oil", "dust", "pet", "rad", "bar")
    ReDim commodityTotal(UBound(t))
    If InStr(txtMultOrigin.Text, "\") > 0 Then
        beginPos = 1
        endPos = InStr(txtMultOrigin.Text, "\")
        While endPos > 0
            If ParseSectors(sx, sy, Mid$(txtMultOrigin.Text, beginPos, endPos - beginPos)) Then
                rsSectors.Seek "=", sx, sy
                If Not rsSectors.NoMatch Then
                    sectorCount = sectorCount + 1
                    For i = LBound(t) To UBound(t)
                        commodityTotal(i) = commodityTotal(i) + rsSectors.Fields(t(i)).Value
                    Next
                End If
            End If
            beginPos = endPos + 1
            endPos = InStr(beginPos, txtMultOrigin.Text, "\")
        Wend
        If ParseSectors(sx, sy, Mid$(txtMultOrigin.Text, beginPos)) Then
            rsSectors.Seek "=", sx, sy
            If Not rsSectors.NoMatch Then
                sectorCount = sectorCount + 1
                For i = LBound(t) To UBound(t)
                    commodityTotal(i) = commodityTotal(i) + rsSectors.Fields(t(i)).Value
                Next
            End If
        End If
    ElseIf ParseSectors(sx, sy, txtMultOrigin.Text) Then
        rsSectors.Seek "=", sx, sy
        If Not rsSectors.NoMatch Then
            sectorCount = sectorCount + 1
            For i = LBound(t) To UBound(t)
                commodityTotal(i) = commodityTotal(i) + rsSectors.Fields(t(i)).Value
            Next
        End If
    ElseIf txtMultOrigin.Text = "*" Then '090404 rjk: Added an option for all sectors
        If Not (rsSectors.EOF And rsSectors.BOF) Then
            rsSectors.MoveFirst
            While Not rsSectors.EOF
                If Not rsSectors.NoMatch Then
                    sectorCount = sectorCount + 1
                    For i = LBound(t) To UBound(t)
                        commodityTotal(i) = commodityTotal(i) + rsSectors.Fields(t(i)).Value
                    Next
                End If
                rsSectors.MoveNext
            Wend
        End If
    End If
    For i = LBound(t) To UBound(t)
        txtThresh(i).Text = CStr(commodityTotal(i))
    Next
    Label3.Caption = CStr(sectorCount) + " Sectors"
ElseIf ParseSectors(sx, sy, strSector) Then
    rsSectors.Seek "=", sx, sy
    If Not rsSectors.NoMatch Then
        If strCmd = "setresource" Or strCmd = "setsector" Then
            lblThresh(0).Caption = lblThresh(0).Caption + " (" + CStr(rsSectors.Fields("min").Value) + ")"
            lblThresh(1).Caption = lblThresh(1).Caption + " (" + CStr(rsSectors.Fields("gold").Value) + ")"
            lblThresh(2).Caption = lblThresh(2).Caption + " (" + CStr(rsSectors.Fields("fert").Value) + ")"
            lblThresh(3).Caption = lblThresh(3).Caption + " (" + CStr(rsSectors.Fields("oil").Value) + ")"
            lblThresh(6).Caption = lblThresh(6).Caption + " (" + CStr(rsSectors.Fields("uran").Value) + ")"

            If strCmd = "setsector" Then
                lblThresh(4).Caption = lblThresh(4).Caption + " (" + CStr(rsSectors.Fields("country").Value) + ")"
                lblThresh(5).Caption = lblThresh(5).Caption + " (" + CStr(rsSectors.Fields("eff").Value) + ")"
                lblThresh(9).Caption = lblThresh(9).Caption + " (" + CStr(rsSectors.Fields("mob").Value) + ")"
                lblThresh(8).Caption = lblThresh(8).Caption + " (" + CStr(rsSectors.Fields("work").Value) + ")"
                lblThresh(7).Caption = lblThresh(7).Caption + " (" + CStr(rsSectors.Fields("avail").Value) + ")"
            End If
        Else
            Dim lblTh As Label
            
            For Each lblTh In lblThresh
                lblTh.Caption = lblTh.Caption + " (" + CStr(Items.FindByFormLabel(lblTh.Caption).DatabaseValue(rsSectors)) + ")"
            Next lblTh
        End If
    End If
End If
End Sub

Public Sub SetShipLabels()
Dim rsTemp As Recordset
Dim hldIndex As String
Dim isLand As Boolean
Dim bLast As Boolean
Dim bSecondFound As Boolean
Dim strFirstSector As String
Dim strUnit As String

SetLabels

'check for ship or land, and choose the correct recordset
If strCmd = "lload" Or strCmd = "lunload" Then '092803 move ltend to ship as the origin is a ship
    Set rsTemp = rsLand
    hldIndex = rsLand.Index
    rsTemp.Index = "PrimaryKey"
    isLand = True
Else
    Set rsTemp = rsShip
    hldIndex = rsShip.Index
    rsTemp.Index = "PrimaryKey"
    isLand = False
End If
    
On Error GoTo Error_Exit:

'If there is a single unit in the box, set the labels based on its current load
If Len(txtUnit.Text) > 0 And InStr(txtUnit.Text, "/") = 0 Then
    rsTemp.Seek "=", txtUnit.Text
    If Not rsTemp.NoMatch Then
        If strCmd = "load" Or strCmd = "lload" Then
            strSector = CStr(rsTemp.Fields("x").Value) + " , " + CStr(rsTemp.Fields("y").Value)
            SetSectorLabels
        Else
            lblThresh(0).Caption = lblThresh(0).Caption + " (" + CStr(rsTemp.Fields("mil").Value) + ")"
            lblThresh(1).Caption = lblThresh(1).Caption + " (" + CStr(rsTemp.Fields("gun").Value) + ")"
            lblThresh(2).Caption = lblThresh(2).Caption + " (" + CStr(rsTemp.Fields("shell").Value) + ")"
            lblThresh(3).Caption = lblThresh(3).Caption + " (" + CStr(rsTemp.Fields("food").Value) + ")"
            
            'Land units don't have civs
            If isLand Then
                lblThresh(4).Caption = lblThresh(4).Caption + " (N/A)"
                lblThresh(5).Caption = lblThresh(5).Caption + " (N/A)"
            Else
                lblThresh(4).Caption = lblThresh(4).Caption + " (" + CStr(rsTemp.Fields("civ").Value) + ")"
                lblThresh(5).Caption = lblThresh(5).Caption + " (" + CStr(rsTemp.Fields("uw").Value) + ")"
            End If
            lblThresh(6).Caption = lblThresh(6).Caption + " (" + CStr(rsTemp.Fields("iron").Value) + ")"
            lblThresh(7).Caption = lblThresh(7).Caption + " (" + CStr(rsTemp.Fields("lcm").Value) + ")"
            lblThresh(8).Caption = lblThresh(8).Caption + " (" + CStr(rsTemp.Fields("hcm").Value) + ")"
            lblThresh(9).Caption = lblThresh(9).Caption + " (" + CStr(rsTemp.Fields("oil").Value) + ")"
            lblThresh(10).Caption = lblThresh(10).Caption + " (" + CStr(rsTemp.Fields("dust").Value) + ")"
            lblThresh(11).Caption = lblThresh(11).Caption + " (" + CStr(rsTemp.Fields("petrol").Value) + ")"
            lblThresh(12).Caption = lblThresh(12).Caption + " (" + CStr(rsTemp.Fields("rad").Value) + ")"
            lblThresh(13).Caption = lblThresh(13).Caption + " (" + CStr(rsTemp.Fields("bar").Value) + ")"
        End If
    End If
'092803 rjk: added logic for multiple units being loaded, ensure there are all at the same location
ElseIf (strCmd = "load" Or strCmd = "lload") And Len(txtUnit.Text) > 0 Then
    strUnit = txtUnit.Text
    bSecondFound = False
    Do
        If InStr(strUnit, "/") <> 0 Then
            rsTemp.Seek "=", Left$(strUnit, InStr(strUnit, "/") - 1)
            strUnit = Mid$(strUnit, InStr(strUnit, "/") + 1)
            bLast = False
        Else
            rsTemp.Seek "=", strUnit
            bLast = True
        End If
        If Not rsTemp.NoMatch Then
            rsSectors.Seek "=", rsTemp.Fields("x").Value, rsTemp.Fields("y").Value
            If Not rsSectors.NoMatch Then
                If (strCmd = "load" And rsSectors.Fields("des") = "h") Or strCmd = "lload" Then
                    If strFirstSector = "" Then
                        strFirstSector = CStr(rsTemp.Fields("x").Value) + " , " + CStr(rsTemp.Fields("y").Value)
                    ElseIf strFirstSector <> (CStr(rsTemp.Fields("x").Value) + " , " + CStr(rsTemp.Fields("y").Value)) Then
                        bSecondFound = True
                        Exit Do 'found a second source location for loading
                    End If
                End If
            End If
        End If
    Loop While Not bLast
    If (Not bSecondFound) And strFirstSector <> "" Then 'only display load quantities if one loading point is being used
        strSector = strFirstSector
        SetSectorLabels
    Else
        SetLabels
    End If
End If

Error_Exit:

rsTemp.Index = hldIndex
Set rsTemp = Nothing
End Sub

Private Sub SetLabels()
'Give What (iron, gold, oil, uranium, fertility, owner, eff., mob., work, avail., oldown, mines)?
If strCmd = "setresource" Or strCmd = "setsector" Then
    lblThresh(0).Caption = "iron"
    lblThresh(1).Caption = "gold"
    lblThresh(2).Caption = "fertility"
    lblThresh(3).Caption = "oil"
    lblThresh(6).Caption = "uranium"
        
    If strCmd = "setresource" Then
        lblThresh(4).Visible = False
        lblThresh(5).Visible = False
        lblThresh(7).Visible = False
        lblThresh(8).Visible = False
        lblThresh(9).Visible = False
        lblThresh(10).Visible = False
        lblThresh(13).Visible = False
        txtThresh(4).Visible = False
        txtThresh(5).Visible = False
        txtThresh(7).Visible = False
        txtThresh(8).Visible = False
        txtThresh(9).Visible = False
        txtThresh(10).Visible = False
        txtThresh(13).Visible = False
    Else
        lblThresh(4).Caption = "owner"
        lblThresh(5).Caption = "eff"
        lblThresh(9).Caption = "mob"
        lblThresh(8).Caption = "work"
        lblThresh(7).Caption = "avail"
        lblThresh(10).Caption = "oldown"
        lblThresh(13).Caption = "mines"
    End If
    lblThresh(11).Visible = False
    lblThresh(12).Visible = False
    txtThresh(11).Visible = False
    txtThresh(12).Visible = False
Else
    lblThresh(0).Caption = Items.FindByLetter("m").FormLabel
    lblThresh(1).Caption = Items.FindByLetter("g").FormLabel
    lblThresh(2).Caption = Items.FindByLetter("s").FormLabel
    lblThresh(3).Caption = Items.FindByLetter("f").FormLabel
    lblThresh(4).Caption = Items.FindByLetter("c").FormLabel
    lblThresh(5).Caption = Items.FindByLetter("u").FormLabel
    lblThresh(6).Caption = Items.FindByLetter("i").FormLabel
    lblThresh(7).Caption = Items.FindByLetter("l").FormLabel
    lblThresh(8).Caption = Items.FindByLetter("h").FormLabel
    lblThresh(9).Caption = Items.FindByLetter("o").FormLabel
    lblThresh(10).Caption = Items.FindByLetter("d").FormLabel
    lblThresh(11).Caption = Items.FindByLetter("p").FormLabel
    lblThresh(12).Caption = Items.FindByLetter("r").FormLabel
    lblThresh(13).Caption = Items.FindByLetter("b").FormLabel
End If
End Sub

Private Sub Option1_Click()
' added by thomas lullier
' to prevent extra click when loading planes

'Call frmDrawMap.ChangeUnitDisplay("Plane", True) 091803 rjk done in GotFocus now
txtUnit2.SetFocus
End Sub

Private Sub Option2_Click()
' added by thomas lullier
' to prevent extra click when loading land units

'Call frmDrawMap.ChangeUnitDisplay("Land", True) '091803 rjk done in GotFocus now
txtUnit2.SetFocus
End Sub

Private Sub txtMultDest_Change()
ComputePathCost
End Sub

Private Sub txtMultDest_DblClick()
If strCmd = "move" Then
    Dim strSector As String
    
    strSector = GetConditionalSectors
    If Len(strSector) > 0 Then
        txtMultDest.Text = strSector
    End If
End If
End Sub

Private Sub txtMultOrigin_Change()
'Only allow OK if a ship is selected
cmdOK.Enabled = (Len(txtMultOrigin) > 0)
'112103 rjk: Removed so the quantity are cleared for CommodityTotal or MoveAll when no sector is present.
'If Len(txtMultOrigin) > 0 Then ' 092103 rjk: Update Labels(quantity) when origin changes
    strSector = txtMultOrigin
    SetSectorLabels
    ComputePathCost
'End If
End Sub

Private Sub txtMultOrigin_DblClick()
If strCmd = "give" Or strCmd = "setresource" Or strCmd = "setsector" Then
    Load frmPromptSectors
    frmPromptSectors.strSectors = txtMultOrigin.Text
    frmPromptSectors.SetControls
    frmPromptSectors.Caption = "Select Sectors"
    frmPromptSectors.Left = frmDrawMap.Left + frmDrawMap.Width - frmPromptSectors.Width
    frmPromptSectors.top = frmDrawMap.top + (frmDrawMap.Height - frmPromptSectors.Height) \ 2
    frmPromptSectors.Show vbModeless
ElseIf strCmd = "move" Or strCmd = "commoditytotal" Then
    Dim strSector As String
    
    strSector = GetConditionalSectors
    If Len(strSector) > 0 Then
        txtMultOrigin.Text = strSector
    End If
End If
End Sub

'added drk 5/31/03.  munged from the PromptMove form in order to display mobility
'cost for the MoveAll command
Private Sub ComputePathCost()
On Error Resume Next
Dim pvar As Variant
Dim pcost As Single
Dim mcost As Single
Dim sx As Integer
Dim sy As Integer
Dim pbdes As String, i As Integer

lblMobLeft.Caption = vbNullString
lblMcost.Caption = vbNullString
lblPathCost.Caption = vbNullString

If LCase$(Trim$(Label2.Caption)) <> "move" Then
    Exit Sub
End If

'111803 rjk: Added Multiple Sector Selection Logic
If InStr(txtMultDest.Text, "\") > 0 Then
    If InStr(txtMultOrigin, "\") > 0 Then
        MsgBox "You can not have multiple sectors in start and end sectors"
        Exit Sub
    End If
End If

'If frmDrawMap.DisplayingPath Then
'    frmDrawMap.DisplayingPath = False
'    frmDrawMap.FinishPath
'End If

'Compute the path cost
If ParseSectors(sx, sy, txtMultDest.Text) Then
    pvar = BestPath(txtMultOrigin.Text, txtMultDest.Text, MT_COMMODITY)
    pcost = pvar(2)
    If Len(CStr(pvar(1))) > 0 Then
        'lblBestPath.Caption = "best path: " + CStr(pvar(1))
        lblPathCost.Caption = "path cost: " + CStr(CSng(CLng(CSng(pcost) * 1000) / 1000))
        'If chkDisplayPath.Value = vbChecked Then
        '    DisplayPath txtMultOrigin.Text, CStr(pvar(1))
        'End If
    Else
        'lblBestPath.Caption = "best path: ??"
        lblPathCost.Caption = "path cost: ??"
    End If
Else
    pcost = PathCost(txtMultOrigin.Text, txtMultDest.Text, MT_COMMODITY)
    If pcost > 0 Then
        'lblBestPath.Caption = "path: " + txtPath.Text
        lblPathCost.Caption = "path cost: " + CStr(CSng(CLng(CSng(pcost) * 1000) / 1000))
        'If chkDisplayPath.Value = vbChecked Then
        '    DisplayPath txtMultOrigin.Text, txtPath.Text
        'End If
    End If
End If

'Compute the mobility cost
If ParseSectors(sx, sy, txtMultOrigin.Text) = False Or pcost <= 0 Then Exit Sub

mcost = 0
'Get Sector Information
rsSectors.Seek "=", sx, sy
If rsSectors.NoMatch Then Exit Sub

If rsSectors.Fields("eff") < 60 Then
    pbdes = vbNullString
Else
    pbdes = rsSectors.Fields("des")
End If

For i = 0 To 13
    If Val(txtThresh(i).Text) > 0 Then
        mcost = mcost + pcost * MobilityWeight(Left$(lblThresh(i).Caption, 1), Val(txtThresh(i).Text), pbdes)
    End If
Next

lblMcost.Caption = "mob. cost: " + CStr(CSng(CLng(CSng(mcost) * 1000) / 1000))
If mcost > rsSectors.Fields("mob") Then
    lblMobLeft.Caption = "Insuff. mobility"
Else
    lblMobLeft.Caption = "mob. left: " + CStr(Int(rsSectors.Fields("mob") - mcost))
End If
End Sub

Private Sub txtThresh_Change(Index As Integer)
ComputePathCost
End Sub

Private Sub txtThresh_DblClick(Index As Integer)
Dim X As Integer
Dim Y As Integer

'put in the max in a sector
X = InStr(lblThresh(Index).Caption, "(")
If X > 0 Then
    Y = InStr(lblThresh(Index).Caption, ")")
    If Y > 0 Then
        '100903 rjk: Reduce the civilian value by one so the sector is not abandoned when double clicking.
        If InStr(lblThresh(Index).Caption, "civilians") > 0 And _
           (strCmd = "move" Or strCmd = "load" Or strCmd = "lload") Then
            If Val(Mid$(lblThresh(Index).Caption, X + 1, Y - X - 1)) > 1 Then
                txtThresh(Index).Text = CStr(Val(Mid$(lblThresh(Index).Caption, X + 1, Y - X - 1)) - 1)
            Else
                txtThresh(Index).Text = "0"
            End If
        Else
            txtThresh(Index).Text = Mid$(lblThresh(Index).Caption, X + 1, Y - X - 1)
        End If
    End If
End If

ComputePathCost
End Sub

Private Sub txtUnit_Change()
'New unit - recalc whats on it for labels
'If Not (strCmd = "load" Or strCmd = "lload") Then 092103 rjk: SetShipLabel now understands load and lload as well
    SetShipLabels
'End If

'Only allow OK if a ship is selected
cmdOK.Enabled = (Len(txtUnit) > 0)

End Sub

Private Sub txtUnit_GotFocus()
Select Case strCmd
Case "load", "unload", "tend", "ltend"
    frmDrawMap.SetUnitDisplay udSHIP, TYPE_ALL, False, False, False
Case "lload", "lunload"
    frmDrawMap.SetUnitDisplay udLAND, TYPE_ALL, False, False, False
Case "give", "setsector", "setresource", "move"
    frmDrawMap.DisplayFirstSectorPanel
End Select
End Sub

Private Sub txtUnit3_GotFocus()
Select Case strCmd
Case "load", "unload", "tend"
    frmDrawMap.SetUnitDisplay udSHIP, TYPE_ALL, False, False, False
Case "lload", "lunload", "ltend"
    frmDrawMap.SetUnitDisplay udLAND, TYPE_ALL, False, False, False
Case "give", "setsector", "setresource", "move"
    frmDrawMap.DisplayFirstSectorPanel
End Select
End Sub

Private Sub txtUnit2_GotFocus()
Dim utType As enumUnitType

If strCmd = "load" Then
    utType = TYPE_SHIP_UNLOADED
ElseIf strCmd = "unload" Then
    utType = TYPE_SHIP_LOADED
ElseIf strCmd = "lload" Then
    utType = TYPE_LAND_UNLOADED
ElseIf strCmd = "lunload" Then
    utType = TYPE_LAND_LOADED
Else
    utType = TYPE_ALL
End If
If Option1 Then
    frmDrawMap.SetUnitDisplay udPLANE, utType, False, False, False
ElseIf Option2 Then
    frmDrawMap.SetUnitDisplay udLAND, utType, False, False, False
End If
End Sub

Private Sub SetDeityLevel(strCmd As String, strOrigin As String, strItem As String, strQuantity As String)
Dim strCmd2 As String

strCmd2 = Trim$(strCmd + " " + strItem + " " + strOrigin + " " + strQuantity)
frmEmpCmd.SubmitEmpireCommand strCmd2, True
End Sub

Private Sub LoadCommodity(strCmd As String, strOrigin As String, strUnits As String, strItem As String, strQuantity As String)
Dim strCmd2 As String
Dim strUnit As String
Dim rsTemp As Recordset
Dim AdjustedQuantity As String
Dim bLast As Boolean
Dim hldIndex As String

'check for ship or land, and choose the correct recordset
If strCmd = "lload" Or strCmd = "lunload" Then
    Set rsTemp = rsLand
    hldIndex = rsLand.Index
    rsTemp.Index = "PrimaryKey"
Else
    Set rsTemp = rsShip
    hldIndex = rsShip.Index
    rsTemp.Index = "PrimaryKey"
End If

Do
    If InStr(strUnits, "/") <> 0 Then
        strUnit = Left$(strUnits, InStr(strUnits, "/") - 1)
        strUnits = Mid$(strUnits, InStr(strUnits, "/") + 1)
        bLast = False
    Else
        strUnit = strUnits
        bLast = True
    End If
    rsTemp.Seek "=", strUnit
    If Not rsTemp.NoMatch Then
        If strCmd = "load" Or strCmd = "lload" Then
            Select Case Left$(strQuantity, 1)
            Case "-"
                rsSectors.Seek "=", rsTemp.Fields("x").Value, rsTemp.Fields("y").Value
                If Not rsSectors.NoMatch Then
                    If (Items.FindByLetter(strItem).DatabaseValue(rsSectors) - Abs(Val(strQuantity))) > 0 Then
                        AdjustedQuantity = CStr(Items.FindByLetter(strItem).DatabaseValue(rsSectors) - Abs(Val(strQuantity)))
                    Else
                        AdjustedQuantity = "0"
                    End If
                Else
                    MsgBox "Ship or Land unit not in a loadable sector"
                    Exit Sub
                End If
            Case "+"
                If (Abs(Val(strQuantity)) - Items.FindByLetter(strItem).DatabaseValue(rsTemp)) > 0 Then
                    AdjustedQuantity = CStr(Abs(Val(strQuantity)) - Items.FindByLetter(strItem).DatabaseValue(rsTemp))
                Else
                    AdjustedQuantity = "0"
                End If
            End Select
        ElseIf strCmd = "unload" Or strCmd = "lunload" Then
            Select Case Left$(strQuantity, 1)
            Case "-"
                If (Items.FindByLetter(Left$(strItem, 1)).DatabaseValue(rsTemp) - Abs(Val(strQuantity))) > 0 Then
                    AdjustedQuantity = CStr(Items.FindByLetter(strItem).DatabaseValue(rsTemp) - Abs(Val(strQuantity)))
                Else
                    AdjustedQuantity = "0"
                End If
            Case "+"
                rsSectors.Seek "=", rsTemp.Fields("x").Value, rsTemp.Fields("y").Value
                If Not rsSectors.NoMatch Then
                    If (Abs(Val(strQuantity)) - Items.FindByLetter(strItem).DatabaseValue(rsSectors)) > 0 Then
                        AdjustedQuantity = CStr(Abs(Val(strQuantity)) - Items.FindByLetter(strItem).DatabaseValue(rsSectors))
                    Else
                        AdjustedQuantity = "0"
                    End If
                Else
                    MsgBox "Ship or Land unit not in a loadable sector"
                    Exit Sub
                End If
            End Select
        End If
        strCmd2 = Trim$(strCmd + " " + strItem + " " + strUnit + " " + AdjustedQuantity)
        frmEmpCmd.SubmitEmpireCommand strCmd2, True
    End If
Loop While Not bLast

End Sub
