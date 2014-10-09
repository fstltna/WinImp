VERSION 5.00
Begin VB.Form frmPromptMove 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2424
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   7944
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2424
   ScaleWidth      =   7944
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdWayPoints 
      Caption         =   "Waypoints"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   20
      Top             =   840
      Width           =   855
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
      Left            =   7560
      TabIndex        =   19
      ToolTipText     =   "Click for Help"
      Top             =   0
      Width           =   375
   End
   Begin VB.CheckBox chkShowAll 
      Caption         =   "Show All"
      Height          =   255
      Left            =   6120
      TabIndex        =   18
      Top             =   2040
      Width           =   975
   End
   Begin VB.CheckBox chkDisplayPath 
      Caption         =   "Display Path"
      Height          =   255
      Left            =   4680
      TabIndex        =   17
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ListBox cbCom 
      Columns         =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      IntegralHeight  =   0   'False
      ItemData        =   "frmPromptMove.frx":0000
      Left            =   4680
      List            =   "frmPromptMove.frx":0002
      TabIndex        =   11
      Top             =   120
      Width           =   2880
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtMultOrigin 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtMultPath 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Move"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblPathCost 
      Caption         =   "path cost:"
      Height          =   252
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   1332
   End
   Begin VB.Label lblBestPath 
      Caption         =   "best path:"
      Height          =   252
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   4452
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "to"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblLeft 
      Caption         =   "mob left:"
      Height          =   252
      Left            =   3240
      TabIndex        =   16
      Top             =   1200
      Width           =   1212
   End
   Begin VB.Label lblCost 
      Caption         =   "mob cost:"
      Height          =   252
      Left            =   1680
      TabIndex        =   15
      Top             =   1200
      Width           =   1452
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblDest 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label lblOrigin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Move"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "from"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   600
      TabIndex        =   6
      Top             =   480
      Width           =   492
   End
End
Attribute VB_Name = "frmPromptMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'081203 efj: added Option Explicit and removed dead variables
'081203 efj: added missing variable definitions
'090703 rjk: ch needs to be a string not a Integer for cbCom_KeyPress
'092803 rjk: General reformatting. Added a check to make sure a commodity was select
'100203 rjk: Erase the sector description when not avaliable or valid
'            Moved the default commodity for deliver to the form.
'100903 rjk: Reduce the civilian value by one so the sector is not abandoned when double clicking.
'            Added bulk move capability in both directions, one sector to many and many to one sector.
'            The many is selected by the condition set in the display mode.
'            You can not test in the cond move mode and it will not compute costs either.
'            The commands are batched up.
'101703 rjk: Added Strength fields to Sector database
'101903 rjk: Removed Conditional Move code, replaced with Multiple Sector Selection code
'102303 rjk: Modified Move to compute the desired quantity based on the following logic
'            ### moves that many
'            +### moves as many as needed to result in ### total in the To sector
'            -### leaves ### behind in the From sector and moves the excess
'            Added sector count to Sector Label when in Multiple Sector Selection Mode
'103103 rjk: Changed the initial field to txtNum field.
'111203 rjk: Fixed the moving of uw with -/+ logic.
'111403 rjk: Set the Move key to be the default key.
'111803 rjk: Moved DisplayPath to frmDrawMap
'            Moved GetCommodityValueFromDB to EmpDataBase.bas
'112003 rjk: Added option to control strength updates
'121303 rjk: Switch from GetCommodityValueFromDB to Item.DatabaseValue and Item class
'240104 rjk: Switched to use item object and single letter for commodities, fixed the selecting of the default commodity.
'270104 rjk: Change the selection of the default commodity for deliver to new sector designation
'180304 rjk: Fixed selecting the default to deal with sdes being a space
'010704 rjk: Switched dump * to SendFullDumpCommand, supports suppressing sea sectors for Deity dumps.
'230704 rjk: Added Radius to Title for delivery for StarWars.
'210306 rjk: Switched SendFullDumpCommand to GetSectors
'301007 rjk: Rearrange to make more room for all long paths, fixed "from" label

'This static var contains the last commodity moved so the box
'can be reset to that commodity when it pops up again
Private LastCommodityMoved As String
Private bFirstLoad As Boolean

Dim strAmount As String

Private Sub cbCom_Click()

SetCurrentQuantityLabel
If txtNum.Visible Then
    txtNum.SetFocus
End If
ComputePathCost
End Sub

Private Sub SetCurrentQuantityLabel()
Label4.Caption = cbCom.List(cbCom.ListIndex)

Dim str As String
Dim sx As Integer
Dim sy As Integer
On Error Resume Next

If ParseSectors(sx, sy, txtMultOrigin.Text) Then
    'Get Sector Information
    If LCase$(Trim$(Label2.Caption)) = "move" Then
        rsSectors.Seek "=", sx, sy
        If Not rsSectors.NoMatch Then
            '102303 rjk: Moved common code a function
            strAmount = CStr(Items.FindByFormName(Label4.Caption).DatabaseValue(rsSectors))
            '100903 rjk: Suppress the avail when for cond mode: many to one
            '            Changed to total avail for cond mode: one to many
            '101903 rjk: Changed to work for Multiple Sector Selection Instead
            'If chkCondMove = vbUnchecked Then
            If InStr(txtMultPath.Text, "\") = 0 Then
                Label4.Caption = Label4.Caption + " [" + strAmount + " avail]"
            'ElseIf optCondMoveTo.Value = True Then
            Else
                Label4.Caption = Label4.Caption + " [" + strAmount + " tot. avail]"
            End If
        End If
    Else
        'must be a delivery - put in cutoff
        Label4.Caption = "cutoff for " + Label4.Caption
    End If
End If
End Sub
Private Sub cbCom_KeyPress(KeyAscii As Integer)
Dim ch As String
Dim i As Integer

On Error Resume Next
ch = LCase$(Chr(KeyAscii))
If ch >= "a" And ch < "z" Then
    i = FindByLetterCommodityBox(ch)
    If i <> -1 Then
        cbCom.ListIndex = i
    End If
    KeyAscii = 0
End If
End Sub

Private Function FindByLetterCommodityBox(ch As String) As Integer
Dim i As Integer

For i = 0 To cbCom.ListCount - 1
    If Left$(cbCom.List(i), 1) = Left$(ch, 1) Then
        FindByLetterCommodityBox = i
        Exit Function
    End If
Next i
FindByLetterCommodityBox = -1
End Function

Private Sub chkDisplayPath_Click()
ComputePathCost
End Sub

Private Sub chkShowAll_Click()
txtMultOrigin_Change
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
frmDrawMap.DisplayPromptHelp Label2.Caption
End Sub

Public Sub cmdOK_Click()
Dim strItem As String
Dim beginPos As Integer
Dim endPos As Integer

'092103 rjk: Added check to ensure commodity has been selected
If Len(cbCom.Text) = 0 Then
    MsgBox "No Commodity selected"
    Exit Sub
End If

strItem = Items.FindByFormName(cbCom.Text).Letter

'101903 rjk: Rework logic for Multiple Sector Selection, also supports function SetDeliveryRoute and MoveCommodity
Select Case Label2.Caption
Case "Deliver"
    If InStr(txtMultOrigin.Text, "\") > 0 Then '101803 Multiple Sectors Selected
        beginPos = 1
        endPos = InStr(txtMultOrigin.Text, "\")
        While endPos > 0
            If Not SetDeliveryRoute(Mid$(txtMultOrigin.Text, beginPos, endPos - beginPos), txtMultPath.Text, strItem, txtNum.Text) Then
                Exit Sub
            End If
            beginPos = endPos + 1
            endPos = InStr(beginPos, txtMultOrigin.Text, "\")
        Wend
        If Not SetDeliveryRoute(Mid$(txtMultOrigin.Text, beginPos), txtMultPath.Text, strItem, txtNum.Text) Then
            Exit Sub
        End If
    Else
        If Not SetDeliveryRoute(txtMultOrigin.Text, txtMultPath.Text, strItem, txtNum.Text) Then
            Exit Sub
        End If
    End If
Case "Move"
    'Save the last commodity moved for next time
    LastCommodityMoved = cbCom.List(cbCom.ListIndex)
    If InStr(txtMultOrigin.Text, "\") > 0 Then '101803 Multiple Sectors Selected
        beginPos = 1
        endPos = InStr(txtMultOrigin.Text, "\")
        While endPos > 0
            MoveCommodity Mid$(txtMultOrigin.Text, beginPos, endPos - beginPos), txtMultPath.Text, strItem, txtNum.Text, False
            beginPos = endPos + 1
            endPos = InStr(beginPos, txtMultOrigin.Text, "\")
        Wend
        MoveCommodity Mid$(txtMultOrigin.Text, beginPos), txtMultPath.Text, strItem, txtNum.Text, False
    Else
        MoveCommodity txtMultOrigin.Text, txtMultPath.Text, strItem, txtNum.Text, False
    End If
End Select
    

'database update
frmEmpCmd.SubmitEmpireCommand "db1", False
GetSectors
'101703 rjk: Added Strength fields to Sector database 112003 rjk: Added option to control
GetCurrentStrength tsSectors
frmEmpCmd.SubmitEmpireCommand "db2", False

Unload Me
End Sub

'101903 rjk: Added for Multiple Sector Selection
Private Function SetDeliveryRoute(strStart As String, strPath As String, strItem As String, strThreshold As String)
Dim sx As Integer
Dim sy As Integer
Dim sx2 As Integer
Dim sy2 As Integer

If ParseSectors(sx, sy, strPath) Then
    If ParseSectors(sx2, sy2, strStart) Then
        If sx = sx2 And sy = sy2 Then
            strPath = "h"  'Same hex means cancel
        Else
            strPath = EmpirePathDirection(sx - sx2, sy - sy2)
            If Len(strPath) > 0 Then
                If Right$(strPath, 1) = "h" Then
                    strPath = Left$(strPath, Len(strPath) - 1)
                End If
                If Len(strPath) <> 1 Then
                    MsgBox "Must choose an adjacent sector for deliver"
                    SetDeliveryRoute = False
                    Exit Function
                End If
            End If
        End If
    End If
End If
frmEmpCmd.SubmitEmpireCommand "deliver " + strItem + " " + strStart + " " + strThreshold + " " + strPath, True
SetDeliveryRoute = True
End Function

'101903 rjk: Added for Multiple Sector Selection
Public Sub MoveCommodity(strStart As String, strPath As String, strItem As String, strQuantity As String, bTest As Boolean)
Dim wp As Variant
Dim X As Integer
Dim beginPos As Integer
Dim endPos As Integer
Dim strAdjQuantity As String
    
wp = ParseWaypoints(strPath)
If IsArray(wp) Then
    Dim strDest As String
    strDest = wp(UBound(wp))
    strAdjQuantity = AdjustedQuantity(strStart, strDest, strItem, strQuantity)
    If strAdjQuantity <= 0 Then
        Exit Sub
    End If
    frmEmpCmd.SubmitEmpireCommand "bf1", True
    For X = 1 To UBound(wp)
        If bTest Then
            frmEmpCmd.SubmitEmpireCommand "test " + strItem + " " + strStart + " " + strAdjQuantity + " " + wp(X), True
        Else
            frmEmpCmd.SubmitEmpireCommand "move " + strItem + " " + strStart + " " + strAdjQuantity + " " + wp(X), True
        End If
        strStart = wp(X)
    Next X
    frmEmpCmd.SubmitEmpireCommand "bf2", True
ElseIf InStr(strPath, "\") > 0 Then
    beginPos = 1
    endPos = InStr(strPath, "\")
    While endPos > 0
        MoveCommodity strStart, Mid$(strPath, beginPos, endPos - beginPos), strItem, strQuantity, bTest
        beginPos = endPos + 1
        endPos = InStr(beginPos, strPath, "\")
    Wend
    MoveCommodity strStart, Mid$(strPath, beginPos), strItem, strQuantity, bTest
Else
    strAdjQuantity = AdjustedQuantity(strStart, strPath, strItem, strQuantity)
    If Val(strAdjQuantity) <= 0 Then
        Exit Sub
    End If
    If bTest Then
        frmEmpCmd.SubmitEmpireCommand "test " + strItem + " " + strStart + " " + strAdjQuantity + " " + strPath, True
    Else
        frmEmpCmd.SubmitEmpireCommand "move " + strItem + " " + strStart + " " + strAdjQuantity + " " + strPath, True
    End If
End If
End Sub

Private Sub cmdTest_Click()
Dim strItem As String
Dim beginPos As Integer
Dim endPos As Integer

strItem = cbCom.Text
If Left$(strItem, 1) = "u" Then
    strItem = "un"
End If

'101903 rjk: Switched to Multiple Sector Selection
If InStr(txtMultOrigin.Text, "\") > 0 Then
    beginPos = 1
    endPos = InStr(txtMultOrigin.Text, "\")
    While endPos > 0
        MoveCommodity Mid$(txtMultOrigin.Text, beginPos, endPos - beginPos), txtMultPath.Text, strItem, txtNum.Text, True
        beginPos = endPos + 1
        endPos = InStr(beginPos, txtMultOrigin.Text, "\")
    Wend
    MoveCommodity Mid$(txtMultOrigin.Text, beginPos), txtMultPath.Text, strItem, txtNum.Text, True
Else
    MoveCommodity txtMultOrigin.Text, txtMultPath.Text, strItem, txtNum.Text, True
End If
End Sub

Private Sub cmdWayPoints_Click()
txtMultPath.SetFocus
txtMultPath_KeyPress (32) 'simulate a space
End Sub

Private Sub Form_Activate()
If bFirstLoad Then
    bFirstLoad = False
    Select Case LCase$(Trim$(Label2.Caption))
    Case "deliver"
        cmdCancel.Move cmdTest.Left, cmdTest.top
        cmdTest.Visible = False
        chkDisplayPath.Visible = False
        chkShowAll.Visible = False '100203 rjk: Deliver already shows all.
        cmdOK.Caption = "OK"
        txtNum.Text = "0"
        cmdWayPoints.Visible = False
        'txtMultPath.SetFocus 103103 rjk: Change to txtUnit, as txtMultPath is filled in from txtNum
        txtNum.SetFocus
        If frmOptions.bStarWars Then
            Label1.Caption = "Direction/ Radius"
            lblPathCost.Visible = False
        End If
    '101903 rjk: Removed the chkCond code, Added txtMultOrigin_Change to get initial quantity label correct
    Case "move"
        txtMultOrigin_Change
    End Select
End If
End Sub

Private Sub Form_Load()
bFirstLoad = True
' Set parent for the toolbar to display on top of:
' Dim sucess As Long  removed 8/12/03 efj
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)

'Put form in proper place
Left = frmDrawMap.Left + (frmDrawMap.Width - Width) \ 2
top = frmDrawMap.top + frmDrawMap.Height - Height

If LastCommodityMoved = "" Then 'check for first time
    LastCommodityMoved = "civilians"
End If

LoadCommodityBox cbCom
ComputePathCost
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
If frmDrawMap.DisplayingPath Then
    frmDrawMap.DisplayingPath = False
    frmDrawMap.FinishPath
End If
End Sub

Private Sub txtMultOrigin_DblClick()
Dim strSector As String

strSector = GetConditionalSectors
If Len(strSector) > 0 Then
    txtMultOrigin.Text = strSector
End If
End Sub

Private Sub txtNum_Change()
ComputePathCost
End Sub

Private Sub txtNum_DblClick()
'Put the full amount in the text box
If Len(strAmount) > 0 Then
    If cbCom.Text = "civilians" Then '100903 rjk: Reduce the civilian value by one so the sector is not abandoned.
        If Val(strAmount) = 1 Then
            txtNum = "0"
        Else
            txtNum = CStr(Val(strAmount) - 1)
        End If
    Else
        txtNum = strAmount
    End If
End If
End Sub

Private Sub txtNum_KeyPress(KeyAscii As Integer)
cbCom_KeyPress KeyAscii
End Sub

Private Sub txtMultPath_Change()
' Dim p1 As Integer    8/2003 efj  removed
Dim sx As Integer
Dim sy As Integer
On Error Resume Next

If Len(txtMultPath.Text) = 0 Then
    lblDest.Caption = vbNullString
    Exit Sub
End If

If ParseSectors(sx, sy, txtMultPath.Text) Then
    'Get Sector Information
    rsSectors.Seek "=", sx, sy
    If Not rsSectors.NoMatch Then
        lblDest = CStr(rsSectors.Fields("eff").Value) + "% "
        lblDest = lblDest + colDes.Item(rsSectors.Fields("des").Value)
    Else '100203 rjk: Erase the sector description when not avaliable or valid
        lblDest = ""
    End If
Else '100203 rjk: Erase the sector description when not avaliable or valid
    lblDest = ""
End If

'101903 rjk: Added Multiple Sector Selection Logic
If InStr(txtMultPath.Text, "\") > 0 Then
    If InStr(txtMultOrigin, "\") > 0 Then
        MsgBox "You can not have multiple sectors in start and end sectors"
        cmdOK.Enabled = False
        cmdTest.Enabled = False
    Else
        cmdOK.Enabled = True
        cmdTest.Enabled = True
    End If
    chkDisplayPath.Enabled = False
    cmdWayPoints.Enabled = False
    lblDest = CStr(CountMultipleSectors(txtMultPath.Text)) + " Sectors"
Else
    cmdWayPoints.Enabled = True
    If InStr(txtMultOrigin.Text, "\") = 0 Then
        chkDisplayPath.Enabled = True
    Else
        chkDisplayPath.Enabled = False
    End If
    cmdOK.Enabled = True
    cmdTest.Enabled = True
End If
SetCurrentQuantityLabel
ComputePathCost
End Sub

Private Sub txtMultPath_DblClick()
'Make sure starting sector is valid before prompting
Dim sx As Integer
Dim sy As Integer
If Not ParseSectors(sx, sy, txtMultOrigin.Text) Then
    Exit Sub
End If
Load frmPromptPath
frmPromptPath.lblSector.Caption = txtMultOrigin.Text
frmPromptPath.lblEndSector.Caption = txtMultOrigin.Text
frmPromptPath.Caption = "Select Movement Path"
frmPromptPath.Left = frmDrawMap.Left + frmDrawMap.Width - frmPromptPath.Width
frmPromptPath.top = frmDrawMap.top + (frmDrawMap.Height - frmPromptPath.Height) \ 2
frmPromptPath.Show vbModeless
End Sub

Private Sub txtMultOrigin_Change()
' Dim p1 As Integer    8/2003 efj  removed
Dim sx As Integer
Dim sy As Integer
Dim n As Integer
Dim str As String
On Error Resume Next

If ParseSectors(sx, sy, txtMultOrigin.Text) Then
    'Get Sector Information
    rsSectors.Seek "=", sx, sy
    If Not rsSectors.NoMatch Then
        lblOrigin = CStr(rsSectors.Fields("eff").Value) + "% "
        lblOrigin = lblOrigin + colDes.Item(rsSectors.Fields("des").Value)
        
        'remove all items that are not in sector
        LoadCommodityBox cbCom
        If LCase$(Trim$(Label2.Caption)) = "move" Then
            If Not chkShowAll.Value = vbChecked Then
                For n = cbCom.ListCount - 1 To 0 Step -1
                    '102303 rjk: Move common code to a function
                    If Items.FindByFormName(cbCom.List(n)).DatabaseValue(rsSectors) = 0 Then
                        cbCom.RemoveItem n
                    End If
                Next n
            End If
            For n = 0 To cbCom.ListCount - 1
                If cbCom.List(n) = LastCommodityMoved Then
                    If cbCom.Visible Then
                        cbCom.ListIndex = n 'set to last commodity
                    End If
                    Exit For
                End If
            Next n
            If n = cbCom.ListCount Then
                If cbCom.Visible Then
                    cbCom.ListIndex = 0
                End If
            End If
        ElseIf LCase$(Trim$(Label2.Caption)) = "deliver" Then
            '270104 rjk: Change the selection of the default commodity for deliver to new sector designation
            '180304 rjk: Fixed to deal with sdes being a space
            If Trim$(rsSectors.Fields("sdes").Value) <> "" And dspDesignation <> DD_CURRENT Then
                SetDefaultCommodity rsSectors.Fields("sdes").Value
            Else
                SetDefaultCommodity rsSectors.Fields("des").Value
            End If
        End If
        chkShowAll.Enabled = True
        If InStr(txtMultPath.Text, "\") = 0 Then
            chkDisplayPath.Enabled = True
        Else
            chkDisplayPath.Enabled = False
        End If
    Else '100203 rjk: Erase the sector description when not avaliable or valid
        lblOrigin = ""
        cbCom.Clear
        lblBestPath.Caption = ""
        lblPathCost.Caption = ""
        lblCost.Caption = ""
        lblLeft.Caption = ""
        chkShowAll.Enabled = False
        chkDisplayPath.Enabled = False
    End If
    cmdOK.Enabled = True
    cmdTest.Enabled = True
Else '100203 rjk: Erase the sector description when not avaliable or valid
    '101903 rjk: Added Multiple Sector Selection Logic, removed condition logic
    lblOrigin = ""
    If InStr(txtMultOrigin.Text, "\") > 0 Then
        Dim strCurrentComm As String
        
        strCurrentComm = cbCom.Text
        LoadCommodityBox cbCom
        For n = 0 To cbCom.ListCount - 1
            If cbCom.List(n) = strCurrentComm Then
                cbCom.ListIndex = n 'set to last commodity
                Exit For
            End If
        Next n
        If n = cbCom.ListCount Then
            cbCom.ListIndex = 0
        End If
        txtMultOrigin.SetFocus
        If InStr(txtMultPath.Text, "\") > 0 Then
            MsgBox "You can not have multiple sectors in start and end sectors"
            cmdOK.Enabled = False
            cmdTest.Enabled = False
        Else
            cmdOK.Enabled = True
            cmdTest.Enabled = True
        End If
        lblOrigin = CStr(CountMultipleSectors(txtMultOrigin.Text)) + " Sectors"
    Else
        cbCom.Clear
        cmdOK.Enabled = True
        cmdTest.Enabled = True
    End If
    lblBestPath.Caption = ""
    lblPathCost.Caption = ""
    lblCost.Caption = ""
    lblLeft.Caption = ""
    chkShowAll.Enabled = False
    chkDisplayPath.Enabled = False
End If
SetCurrentQuantityLabel
ComputePathCost
End Sub

Private Sub txtMultPath_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 And (LCase$(Trim$(Label2.Caption)) = "move") Then
    KeyAscii = 0
    Load frmPromptWaypoints
    frmPromptWaypoints.Left = frmDrawMap.Left + frmDrawMap.Width - frmPromptWaypoints.Width
    frmPromptWaypoints.top = frmDrawMap.top + (frmDrawMap.Height - frmPromptWaypoints.Height) \ 2
    frmPromptWaypoints.Show vbModeless
    frmPromptWaypoints.txtOrigin.SetFocus
End If
End Sub

Private Sub ComputePathCost()
On Error Resume Next
Dim pvar As Variant
Dim pcost As Single
Dim sx As Integer
Dim sy As Integer
Dim pbdes As String

lblLeft.Caption = vbNullString
lblCost.Caption = vbNullString
lblBestPath.Caption = vbNullString
lblPathCost.Caption = vbNullString

If LCase$(Trim$(Label2.Caption)) <> "move" Then
    Exit Sub
End If

If frmDrawMap.DisplayingPath Then
    frmDrawMap.DisplayingPath = False
    frmDrawMap.FinishPath
End If

'Compute the path cost
If ParseSectors(sx, sy, txtMultPath.Text) Then
    pvar = BestPath(txtMultOrigin.Text, txtMultPath.Text, MT_COMMODITY)
    pcost = pvar(2)
    If Len(CStr(pvar(1))) > 0 Then
        lblBestPath.Caption = "best path: " + CStr(pvar(1))
        lblPathCost.Caption = "path cost: " + CStr(CSng(CLng(CSng(pcost) * 1000) / 1000))
        If chkDisplayPath.Value = vbChecked Then
            frmDrawMap.DisplayPath txtMultOrigin.Text, CStr(pvar(1))
        End If
    Else
        lblBestPath.Caption = "best path: ??"
        lblPathCost.Caption = "path cost: ??"
    End If
Else
    pcost = PathCost(txtMultOrigin.Text, txtMultPath.Text, MT_COMMODITY)
    If pcost > 0 Then
        lblBestPath.Caption = "path: " + txtMultPath.Text
        lblPathCost.Caption = "path cost: " + CStr(CSng(CLng(CSng(pcost) * 1000) / 1000))
        If chkDisplayPath.Value = vbChecked Then
            frmDrawMap.DisplayPath txtMultOrigin.Text, txtMultPath.Text
        End If
    End If
End If

'Compute the mobility cost
'102303 rjk: Changed to Mobility cost to use AdjustedQuantity
Dim nAdjQuantity As Integer
nAdjQuantity = Val(AdjustedQuantity(txtMultOrigin.Text, txtMultPath.Text, Items.FindByFormName(cbCom.List(cbCom.ListIndex)).Letter, txtNum.Text))
'If pcost > 0 And Val(txtNum.Text) > 0 Then
If pcost > 0 And nAdjQuantity > 0 Then
    If ParseSectors(sx, sy, txtMultOrigin.Text) Then
        'Get Sector Information
        rsSectors.Seek "=", sx, sy
        If Not rsSectors.NoMatch Then
            If rsSectors.Fields("eff") < 60 Then
                pbdes = vbNullString
            Else
                pbdes = rsSectors.Fields("des")
            End If
            'pcost = pcost * MobilityWeight(Left$(cbCom.Text, 1), Val(txtNum.Text), pbdes)
            pcost = pcost * MobilityWeight(Left$(cbCom.Text, 1), nAdjQuantity, pbdes)
            lblCost.Caption = "mob. cost: " + CStr(CSng(CLng(CSng(pcost) * 1000) / 1000))
            If pcost > rsSectors.Fields("mob") Then
                lblLeft.Caption = "Insuff. mobility"
            Else
                lblLeft.Caption = "mob. left: " + CStr(Int(rsSectors.Fields("mob") - pcost))
            End If
        End If
    End If
End If
End Sub

Private Sub SetDefaultCommodity(strDes As String)
Dim i As Integer

i = -1
'determine com box type based on sector
'092103 rjk: rsSectors.Seek "=", MxPos, MyPos switched to SelX and SelY to work from Top Level Menu
'100203 rjk: Moved com box type based on sector logic to the form.
Select Case strDes
    Case "j" 'set to lcm
        i = FindByLetterCommodityBox("l")
    Case "k" 'set to hcm
        i = FindByLetterCommodityBox("h")
    Case "m" 'set to iron
        i = FindByLetterCommodityBox("i")
    Case "g" 'set to dust
        i = FindByLetterCommodityBox("d")
    Case "o" 'set to oil
        i = FindByLetterCommodityBox("o")
    Case "u" 'set to rads
        i = FindByLetterCommodityBox("r")
    Case "%" 'set to petrol
        i = FindByLetterCommodityBox("p")
    Case "a" 'set to food
        i = FindByLetterCommodityBox("f")
    Case "b" 'set to bars
        i = FindByLetterCommodityBox("b")
    Case "d" 'set to guns
        i = FindByLetterCommodityBox("g")
    Case "i" 'set to shells
        i = FindByLetterCommodityBox("s")
    Case "y" '240104 rjk: Added for St@r W@rs (uw85 robotics)
        i = FindByLetterCommodityBox("u")
End Select
If i <> -1 Then
    cbCom.ListIndex = i
End If
End Sub

'102303 rjk: Function to compute the desired quantity based on the following logic
'            ### moves that many
'            +### moves as many as needed to result in ### total in the To sector
'            -### leaves ### behind in the From sector and moves the excess
Private Function AdjustedQuantity(strStart As String, strEnd As String, strItem As String, strQuantity As String) As String
Dim sx As Integer
Dim sy As Integer

If Len(strQuantity) = 0 Then
    AdjustedQuantity = strQuantity
    Exit Function
End If
Select Case Left$(strQuantity, 1)
Case "-"
    If Not ParseSectors(sx, sy, strStart) Then
        AdjustedQuantity = strQuantity
        Exit Function
    End If

    rsSectors.Seek "=", sx, sy
    If rsSectors.NoMatch Then
        AdjustedQuantity = strQuantity
        Exit Function
    End If
    
    If (Items.FindByLetter(strItem).DatabaseValue(rsSectors) - Abs(Val(strQuantity))) > 0 Then
        AdjustedQuantity = CStr(Items.FindByLetter(strItem).DatabaseValue(rsSectors) - Abs(Val(strQuantity)))
    Else
        AdjustedQuantity = "0"
    End If
Case "+"
    If Not ParseSectors(sx, sy, strEnd) Then
        AdjustedQuantity = strQuantity
        Exit Function
    End If

    rsSectors.Seek "=", sx, sy
    If rsSectors.NoMatch Then
        AdjustedQuantity = strQuantity
        Exit Function
    End If
    
    If (Val(strQuantity) - Items.FindByLetter(strItem).DatabaseValue(rsSectors)) > 0 Then
        AdjustedQuantity = CStr(Val(strQuantity) - Items.FindByLetter(strItem).DatabaseValue(rsSectors))
    Else
        AdjustedQuantity = "0"
    End If
Case Else
    AdjustedQuantity = strQuantity
End Select
End Function
