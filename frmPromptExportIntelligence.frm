VERSION 5.00
Begin VB.Form frmPromptExportIntelligence 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   10185
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkHeader 
      Caption         =   "Include Headers"
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Frame frameData 
      Caption         =   "Data Fields"
      Height          =   2775
      Left            =   8280
      TabIndex        =   31
      Top             =   120
      Width           =   1815
      Begin VB.CommandButton cmdClearAllData 
         Caption         =   "Clear All"
         Height          =   375
         Left            =   960
         TabIndex        =   49
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton cmdSetAllData 
         Caption         =   "Set All"
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   2280
         Width           =   735
      End
      Begin VB.CheckBox chkData 
         Caption         =   "Bar"
         Height          =   255
         Index           =   16
         Left            =   960
         TabIndex        =   43
         Top             =   960
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkData 
         Caption         =   "Food"
         Height          =   255
         Index           =   14
         Left            =   960
         TabIndex        =   32
         Top             =   720
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkData 
         Caption         =   "Petrol"
         Height          =   255
         Index           =   13
         Left            =   960
         TabIndex        =   33
         Top             =   480
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkData 
         Caption         =   "Iron"
         Height          =   255
         Index           =   15
         Left            =   960
         TabIndex        =   42
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkData 
         Caption         =   "Wake"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   46
         Top             =   1920
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkData 
         Caption         =   "Tech"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   44
         Top             =   1920
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkData 
         Caption         =   "Def."
         Height          =   255
         Index           =   8
         Left            =   960
         TabIndex        =   41
         ToolTipText     =   "Defense"
         Top             =   1440
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkData 
         Caption         =   "Rail"
         Height          =   255
         Index           =   7
         Left            =   960
         TabIndex        =   40
         Top             =   1680
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkData 
         Caption         =   "Road"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   34
         Top             =   1680
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkData 
         Caption         =   "Eff."
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   35
         ToolTipText     =   "Efficiency"
         Top             =   1440
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkData 
         Caption         =   "Old Owner"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   45
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkData 
         Caption         =   "Civ."
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   36
         Top             =   960
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkData 
         Caption         =   "Guns"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   37
         Top             =   720
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkData 
         Caption         =   "Shells"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   38
         Top             =   480
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkData 
         Caption         =   "Mil."
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
   End
   Begin VB.Frame frameOffset 
      Caption         =   "Offset"
      Height          =   615
      Left            =   120
      TabIndex        =   25
      Top             =   1920
      Width           =   2775
      Begin VB.TextBox txtOffsetY 
         Height          =   285
         Left            =   2160
         TabIndex        =   28
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtOffsetX 
         Height          =   285
         Left            =   840
         TabIndex        =   26
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblOffsetY 
         Alignment       =   1  'Right Justify
         Caption         =   "Y Offset"
         Height          =   255
         Left            =   1440
         TabIndex        =   29
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblOffsetX 
         Alignment       =   1  'Right Justify
         Caption         =   "X Offset"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame frameType 
      Caption         =   "Items"
      Height          =   2775
      Left            =   6120
      TabIndex        =   9
      Top             =   120
      Width           =   1935
      Begin VB.CheckBox chkItems 
         Caption         =   "Our Nukes"
         Enabled         =   0   'False
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   20
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CheckBox chkItems 
         Caption         =   "Our Land Units"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   21
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CheckBox chkItems 
         Caption         =   "Our Planes"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   22
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CheckBox chkItems 
         Caption         =   "Our Ships"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   23
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CheckBox chkItems 
         Caption         =   "Our Sectors"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   24
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox chkItems 
         Caption         =   "Enemy Nukes"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox chkItems 
         Caption         =   "Enemy Land Units"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkItems 
         Caption         =   "Enemy Planes"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkItems 
         Caption         =   "Enemy Ships"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkItems 
         Caption         =   "Enemy Sectors"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Value           =   1  'Checked
         Width           =   1695
      End
   End
   Begin VB.Frame frameFilter 
      Caption         =   "Filters"
      Height          =   1695
      Left            =   3120
      TabIndex        =   5
      Top             =   480
      Width           =   2775
      Begin VB.TextBox txtMultOrigin 
         Height          =   285
         Left            =   1200
         TabIndex        =   17
         Text            =   "*"
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtDesignations 
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ComboBox cbNations 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblOrigin 
         Alignment       =   1  'Right Justify
         Caption         =   "Sectors (sect,range,*)"
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblCountry 
         Caption         =   "Enemy Nation to include in export"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblDesignations 
         Alignment       =   1  'Right Justify
         Caption         =   "Designations"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Frame frameOutput 
      Caption         =   "Output"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.ComboBox cbTelegram 
         Height          =   315
         Left            =   360
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1080
         Width           =   2295
      End
      Begin VB.OptionButton optDestination 
         Caption         =   "File"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optDestination 
         Caption         =   "Telegram"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Label lblTelegram 
         Caption         =   "Nation to telegram"
         Height          =   255
         Left            =   1200
         TabIndex        =   19
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Export Intelligence"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   30
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmPromptExportIntelligence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'103003 rjk: Created.  Added parameters to ExportIntelligence.
'            Moved ExportIntelligence to the form.
'110403 rjk: Reorganized SendTelegram to send headers for each telegram and to deal with long lines better.
'111203 rjk: Added some safety checks before sending the telegram.
'111403 rjk: Do not send enemy intelligence about a sector we now own.
'111803 rjk: Added Header and control for it.
'            Added Data field selection and expand the sectors to dump everything possible.
'111903 rjk: Added SetAll and ClearAll for data fields.
'070304 rjk: Switch Enemy timestamp to UTC
'110304 rjk: Replace null Enemy Sector designation with BMap if known otherwise "?"
'090607 rjk: Remove function call chkComm_Click.
'270711 rjk: Switched the relationships to use the xdump nationrelationships instead.

Dim filenumber As Integer
Dim FileName As String
Dim strTelegram As String
    
Enum enumSectorSearch
    ssFound
    ssNotFound
    ssError
End Enum

Private Sub cbTelegram_Click()
If cbTelegram.Text <> "" Then
    cmdOK.Enabled = True
Else
    cmdOK.Enabled = False
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdClearAllData_Click()
chkData(0).Value = vbUnchecked
chkData(1).Value = vbUnchecked
chkData(4).Value = vbUnchecked
chkData(5).Value = vbUnchecked
chkData(6).Value = vbUnchecked
chkData(7).Value = vbUnchecked
chkData(8).Value = vbUnchecked
chkData(9).Value = vbUnchecked
chkData(10).Value = vbUnchecked
chkData(11).Value = vbUnchecked
chkData(12).Value = vbUnchecked
chkData(13).Value = vbUnchecked
chkData(14).Value = vbUnchecked
chkData(15).Value = vbUnchecked
chkData(16).Value = vbUnchecked
End Sub

Private Sub cmdOK_Click()
ExportIntelligence
Unload Me
End Sub

Private Sub cmdSetAllData_Click()
chkData(0).Value = vbChecked
chkData(1).Value = vbChecked
chkData(4).Value = vbChecked
chkData(5).Value = vbChecked
chkData(6).Value = vbChecked
chkData(7).Value = vbChecked
chkData(8).Value = vbChecked
chkData(9).Value = vbChecked
chkData(10).Value = vbChecked
chkData(11).Value = vbChecked
chkData(12).Value = vbChecked
chkData(13).Value = vbChecked
chkData(14).Value = vbChecked
chkData(15).Value = vbChecked
chkData(16).Value = vbChecked
End Sub

Private Sub Form_Load()
Dim Index As Integer

Nations.FillListBox cbNations
Nations.FillListBox cbTelegram

For Index = 0 To cbNations.ListCount - 1
    If cbNations.ItemData(Index) = CountryNumber Then
        cbNations.RemoveItem (Index)
        Exit For
    End If
Next Index

cbNations.AddItem "All", 0
cbNations.ItemData(cbNations.NewIndex) = -1
cbNations.ListIndex = 0

cmdOK.Enabled = False 'Enabled either file is selected or destination for telegram is selected.


'Put form in proper place
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
Left = frmDrawMap.Left + (frmDrawMap.Width - Width) \ 2
top = frmDrawMap.top + frmDrawMap.Height - Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDrawMap.PromptForm = Nothing
frmDrawMap.PromptUp = False
End Sub

Private Sub optDestination_Click(Index As Integer)
If optDestination(0) Then
    cmdOK.Enabled = True
    cbTelegram.Enabled = False
ElseIf optDestination(1) Then
    cbTelegram.Enabled = True
    If cbTelegram.Text <> "" Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If
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
Public Sub ExportIntelligence()
On Error GoTo Empire_Error


Dim n As Integer
Dim nSect As Integer
Dim nUnit As Integer
Dim secx As Integer
Dim secy As Integer
Dim searchResult As enumSectorSearch

filenumber = -1

'111203 rjk: Do some safety checks before sending the telegram.
If optDestination(1) Then 'telegram
    If Nations.YourRelation(CStr(cbTelegram.ItemData(cbTelegram.ListIndex))) <> iREL_ALLIED And _
       Nations.YourRelation(CStr(cbTelegram.ItemData(cbTelegram.ListIndex))) <> iREL_FRIENDLY Then
        If MsgBox("This export is not being sent to an ally or an friend, do you what to continue", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    If cbNations.ItemData(cbNations.ListIndex) = -1 And _
       txtMultOrigin.Text = "*" Then
        If MsgBox("You have selected to export all your enemy intelligence about the items selected, is not correct?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    If Nations.YourRelation(CStr(cbTelegram.ItemData(cbTelegram.ListIndex))) <> iREL_ALLIED And _
        (chkItems(5) Or chkItems(6) Or chkItems(7) Or chkItems(8) Or chkItems(9)) Then
        If MsgBox("You are sending your own sector information to non-ally, do you what to continue", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
End If

If Not OpenDestination Then
    Exit Sub
End If

If chkHeader.Value = vbChecked Then '111803 rjk: Added header option
    OutputIntelligence "+++ Sector Header x,y,des,owner,oldowner,eff,road,rail,defense,civ,mil,shell,gun,pet,food,iron,bar,timestamp", True
End If
    
'now write out the sector file
If chkItems(0) Then
    If Not (rsEnemySect.BOF And rsEnemySect.EOF) Then
        rsEnemySect.MoveFirst
        While Not rsEnemySect.EOF
            If (cbNations.ItemData(cbNations.ListIndex) = -1) Or _
               (cbNations.ItemData(cbNations.ListIndex) = rsEnemySect.Fields("owner")) _
               Then
                If (InStr(txtDesignations.Text, rsEnemySect.Fields("des")) > 0) Or _
                   (Len(txtDesignations.Text) = 0) _
                   Then
                    secx = rsEnemySect.Fields("x").Value
                    secy = rsEnemySect.Fields("y").Value
                    searchResult = SectorWithInSectorList(secx, secy)
                    Select Case searchResult
                    Case ssFound
                        rsSectors.Seek "=", secx, secy
                        If rsSectors.NoMatch Then '111403 rjk: Do not send enemy intelligence about a sector we now own.
                            OffsetSectorLocation secx, secy, Val(txtOffsetX.Text), Val(txtOffsetY.Text)
                            'Write #filenumber, secX;
                            OutputIntelligence CStr(secx), False
                            'Write #filenumber, secY;
                            OutputIntelligence CStr(secy), False
                            nSect = nSect + 1
                            For n = 2 To rsEnemySect.Fields.Count - 2
                                'Write #filenumber, rsEnemySect.Fields(n);
                                If n >= 5 And n <= 16 Then
                                    If chkData(n).Value = vbChecked Then
                                        OutputIntelligence rsEnemySect.Fields(n), False
                                    Else
                                        OutputIntelligence "-1", False
                                    End If
                                ElseIf n = 2 Then
                                    If IsNull(rsEnemySect.Fields(2)) Then
                                        rsBmap.Seek "=", secx, secy
                                        If rsBmap.NoMatch Then
                                            OutputIntelligence "?", False
                                        Else
                                            OutputIntelligence rsBmap.Fields("des"), False
                                        End If
                                    Else
                                        OutputIntelligence rsEnemySect.Fields(n), False
                                    End If
                                Else
                                    OutputIntelligence rsEnemySect.Fields(n), False
                                End If
                            Next n
                            'Write #filenumber, rsEnemySect.Fields(rsEnemySect.Fields.Count - 1)
                            OutputIntelligence rsEnemySect.Fields(rsEnemySect.Fields.Count - 1), True
                        End If
                    Case ssError
                        CloseDestination False
                        Exit Sub
                    End Select
                End If
            End If
            rsEnemySect.MoveNext
        Wend
    End If
End If
If chkItems(5) Then
    If Not (rsSectors.BOF And rsSectors.EOF) Then
        rsSectors.MoveFirst
        While Not rsSectors.EOF
            If (InStr(txtDesignations.Text, rsSectors.Fields("des")) > 0) Or _
               (Len(txtDesignations.Text) = 0) _
               Then
                secx = rsSectors.Fields("x").Value
                secy = rsSectors.Fields("y").Value
                searchResult = SectorWithInSectorList(secx, secy)
                Select Case searchResult
                Case ssFound
                    OffsetSectorLocation secx, secy, Val(txtOffsetX.Text), Val(txtOffsetY.Text)
                    nSect = nSect + 1
                    OutputIntelligence CStr(secx), False
                    OutputIntelligence CStr(secy), False
                    OutputIntelligence rsSectors.Fields("des"), False
                    OutputIntelligence rsSectors.Fields("country"), False
                    OutputIntelligence "-1", False
                    If chkData(4).Value = vbChecked Then
                        OutputIntelligence rsSectors.Fields("eff"), False
                    Else
                        OutputIntelligence "-1", False
                    End If
                    If chkData(6).Value = vbChecked Then
                        OutputIntelligence rsSectors.Fields("road"), False
                    Else
                        OutputIntelligence "-1", False
                    End If
                    If chkData(7).Value = vbChecked Then
                        OutputIntelligence rsSectors.Fields("rail"), False
                    Else
                        OutputIntelligence "-1", False
                    End If
                    If chkData(8).Value = vbChecked Then
                        OutputIntelligence rsSectors.Fields("defense"), False
                    Else
                        OutputIntelligence "-1", False
                    End If
                    If chkData(9).Value = vbChecked Then
                        OutputIntelligence rsSectors.Fields("civ"), False
                    Else
                        OutputIntelligence "-1", False
                    End If
                    If chkData(10).Value = vbChecked Then
                        OutputIntelligence rsSectors.Fields("mil"), False
                    Else
                        OutputIntelligence "-1", False
                    End If
                    If chkData(11).Value = vbChecked Then
                        OutputIntelligence rsSectors.Fields("shell"), False
                    Else
                        OutputIntelligence "-1", False
                    End If
                    If chkData(12).Value = vbChecked Then
                        OutputIntelligence rsSectors.Fields("gun"), False
                    Else
                        OutputIntelligence "-1", False
                    End If
                    If chkData(13).Value = vbChecked Then
                        OutputIntelligence rsSectors.Fields("pet"), False
                    Else
                        OutputIntelligence "-1", False
                    End If
                    If chkData(14).Value = vbChecked Then
                        OutputIntelligence rsSectors.Fields("food"), False
                    Else
                        OutputIntelligence "-1", False
                    End If
                    If chkData(15).Value = vbChecked Then
                        OutputIntelligence rsSectors.Fields("iron"), False
                    Else
                        OutputIntelligence "-1", False
                    End If
                    If chkData(16).Value = vbChecked Then
                        OutputIntelligence rsSectors.Fields("bar"), False
                    Else
                        OutputIntelligence "-1", False
                    End If
                    OutputIntelligence GetWinACEUTC, True
                Case ssError
                    CloseDestination False
                    Exit Sub
                End Select
            End If
            rsSectors.MoveNext
        Wend
    End If
End If
OutputIntelligence "+++ " + CStr(nSect) + " Sectors exported", True

If chkHeader.Value = vbChecked Then '111803 rjk: Added header option
    OutputIntelligence "+++ Unit Header id,type,x,y,owner,eff,mil,tech,timestamp,wake", True
End If

'now write out the unit file
If Not (rsEnemyUnit.BOF And rsEnemyUnit.EOF) Then
    rsEnemyUnit.MoveFirst
    While Not rsEnemyUnit.EOF
        If (rsEnemyUnit.Fields("class") = "s" And chkItems(1)) Or _
           (rsEnemyUnit.Fields("class") = "p" And chkItems(2)) Or _
           (rsEnemyUnit.Fields("class") = "l" And chkItems(3)) _
           Then
            If (cbNations.ItemData(cbNations.ListIndex) = -1) Or _
               (cbNations.ItemData(cbNations.ListIndex) = rsEnemyUnit.Fields("owner")) _
               Then
                secx = rsEnemyUnit.Fields("x").Value
                secy = rsEnemyUnit.Fields("y").Value
                searchResult = SectorWithInSectorList(secx, secy)
                Select Case searchResult
                Case ssFound
                    OffsetSectorLocation secx, secy, Val(txtOffsetX.Text), Val(txtOffsetY.Text)
                    nUnit = nUnit + 1
                    OutputIntelligence rsEnemyUnit.Fields("id"), False
                    OutputIntelligence rsEnemyUnit.Fields("type"), False
                    OutputIntelligence CStr(secx), False
                    OutputIntelligence CStr(secy), False
                    OutputIntelligence rsEnemyUnit.Fields("owner"), False
                    If chkData(5).Value = vbChecked Then
                        OutputIntelligence rsEnemyUnit.Fields("eff"), False
                    Else
                        OutputIntelligence "-1", False
                    End If
                    If chkData(10).Value = vbChecked Then
                        OutputIntelligence rsEnemyUnit.Fields("mil"), False
                    Else
                        OutputIntelligence "-1", False
                    End If
                    If chkData(1).Value = vbChecked Then
                        OutputIntelligence rsEnemyUnit.Fields("tech"), False
                    Else
                        OutputIntelligence "-1", False
                    End If
                    OutputIntelligence rsEnemyUnit.Fields("timestamp"), False
                    OutputIntelligence rsEnemyUnit.Fields("class"), False
                    If chkData(1).Value = vbChecked Then
                        OutputIntelligence rsEnemyUnit.Fields("wake"), True
                    Else
                        OutputIntelligence "", True
                    End If
                Case ssError
                    CloseDestination False
                    Exit Sub
                End Select
            End If
        End If
        rsEnemyUnit.MoveNext
    Wend
End If

'now write out the our ship file
If chkItems(6) Then
    If Not (rsShip.BOF And rsShip.EOF) Then
        rsShip.MoveFirst
        While Not rsShip.EOF
            secx = rsShip.Fields("x").Value
            secy = rsShip.Fields("y").Value
            searchResult = SectorWithInSectorList(secx, secy)
            Select Case searchResult
            Case ssFound
                OffsetSectorLocation secx, secy, Val(txtOffsetX.Text), Val(txtOffsetY.Text)
                nUnit = nUnit + 1
                OutputIntelligence rsShip.Fields("id"), False
                OutputIntelligence rsShip.Fields("type"), False
                OutputIntelligence CStr(secx), False
                OutputIntelligence CStr(secy), False
                OutputIntelligence rsShip.Fields("country"), False
                If chkData(5).Value = vbChecked Then
                    OutputIntelligence rsShip.Fields("eff"), False
                Else
                    OutputIntelligence "-1", False
                End If
                If chkData(10).Value = vbChecked Then
                    OutputIntelligence rsShip.Fields("mil"), False
                Else
                    OutputIntelligence "-1", False
                End If
                If chkData(1).Value = vbChecked Then
                    OutputIntelligence rsShip.Fields("tech"), False
                Else
                    OutputIntelligence "-1", False
                End If
                OutputIntelligence GetWinACEUTC, False
                OutputIntelligence "s", False 'Class
                OutputIntelligence "", True 'Wake
            Case ssError
                CloseDestination False
                Exit Sub
            End Select
            rsShip.MoveNext
        Wend
    End If
End If

'now write out the our plane file
If chkItems(7) Then
    If Not (rsPlane.BOF And rsPlane.EOF) Then
        rsPlane.MoveFirst
        While Not rsPlane.EOF
            secx = rsPlane.Fields("x").Value
            secy = rsPlane.Fields("y").Value
            searchResult = SectorWithInSectorList(secx, secy)
            Select Case searchResult
            Case ssFound
                OffsetSectorLocation secx, secy, Val(txtOffsetX.Text), Val(txtOffsetY.Text)
                nUnit = nUnit + 1
                OutputIntelligence rsPlane.Fields("id"), False
                OutputIntelligence rsPlane.Fields("type"), False
                OutputIntelligence CStr(secx), False
                OutputIntelligence CStr(secy), False
                OutputIntelligence rsPlane.Fields("country"), False
                If chkData(5).Value = vbChecked Then
                    OutputIntelligence rsPlane.Fields("eff"), False
                Else
                    OutputIntelligence "-1", False
                End If
                OutputIntelligence "-1", False
                If chkData(1).Value = vbChecked Then
                    OutputIntelligence rsPlane.Fields("tech"), False
                Else
                    OutputIntelligence "-1", False
                End If
                OutputIntelligence GetWinACEUTC, False
                OutputIntelligence "p", False 'Class
                OutputIntelligence "", True 'Wake
            Case ssError
                CloseDestination False
                Exit Sub
            End Select
            rsPlane.MoveNext
        Wend
    End If
End If

'now write out the our land unit file
If chkItems(8) Then
    If Not (rsLand.BOF And rsLand.EOF) Then
        rsLand.MoveFirst
        While Not rsLand.EOF
            secx = rsLand.Fields("x").Value
            secy = rsLand.Fields("y").Value
            searchResult = SectorWithInSectorList(secx, secy)
            Select Case searchResult
            Case ssFound
                OffsetSectorLocation secx, secy, Val(txtOffsetX.Text), Val(txtOffsetY.Text)
                nUnit = nUnit + 1
                OutputIntelligence rsLand.Fields("id"), False
                OutputIntelligence rsLand.Fields("type"), False
                OutputIntelligence CStr(secx), False
                OutputIntelligence CStr(secy), False
                OutputIntelligence rsLand.Fields("country"), False
                If chkData(5).Value = vbChecked Then
                    OutputIntelligence rsLand.Fields("eff"), False
                Else
                    OutputIntelligence "-1", False
                End If
                If chkData(10).Value = vbChecked Then
                    OutputIntelligence rsLand.Fields("mil"), False
                Else
                    OutputIntelligence "-1", False
                End If
                If chkData(1).Value = vbChecked Then
                    OutputIntelligence rsLand.Fields("tech"), False
                Else
                    OutputIntelligence "-1", False
                End If
                OutputIntelligence GetWinACEUTC, False
                OutputIntelligence "l", False 'Class
                OutputIntelligence "", True 'Wake
            Case ssError
                CloseDestination False
                Exit Sub
            End Select
            rsLand.MoveNext
        Wend
    End If
End If

OutputIntelligence "+++ " + CStr(nUnit) + " Units exported", True
CloseDestination True

If optDestination(0) Then 'file
    MsgBox CStr(nSect) + " Sectors, " + CStr(nUnit) + " Units exported to file " + FileName
ElseIf optDestination(1) Then 'telegram
    MsgBox CStr(nSect) + " Sectors, " + CStr(nUnit) + " Units exported to " + cbTelegram.Text
End If

Exit Sub
Empire_Error:
CloseDestination False
EmpireError "ExportIntelligence", vbNullString, vbNullString
End Sub


Private Function OpenDestination() As Boolean
If optDestination(0) Then 'file
    If VerifySubDirectory("Exports", True) Then
        FileName = App.path + "\Exports\" + GameName + " Intelligence.txt"
    Else
        FileName = App.path + "\" + GameName + " Intelligence.txt"
    End If
    filenumber = FreeFile   ' Get unused file number.
    Open FileName For Output As #filenumber
    OpenDestination = True
ElseIf optDestination(1) Then 'telegram
    strTelegram = vbNullString
    OpenDestination = True
End If
End Function

Private Sub OutputIntelligence(vField As Variant, bEOL As Boolean)
If optDestination(0) Then 'file
    If filenumber <> -1 Then
    End If
        If bEOL Then
            Write #filenumber, vField
        Else
            Write #filenumber, vField;
        End If
ElseIf optDestination(1) Then 'telegram
    strTelegram = strTelegram + Chr$(34) + CStr(vField) + Chr$(34)
    If bEOL Then
        strTelegram = strTelegram + vbLf
    Else
    strTelegram = strTelegram + ","
    End If
End If
End Sub

Private Function CloseDestination(bSend) As Boolean
If optDestination(0) Then 'file
    Close #filenumber
ElseIf optDestination(1) Then 'telegram
    If bSend Then
        SendTelegram (strTelegram)
    End If
    strTelegram = vbNullString
End If
End Function

Private Function SectorWithInSectorList(secx As Integer, secy As Integer) As enumSectorSearch
Dim sx As Integer
Dim sy As Integer
Dim beginPos As Integer
Dim endPos As Integer

SectorWithInSectorList = ssError
If Len(txtMultOrigin.Text) = 0 Then
    Exit Function
ElseIf Len(txtMultOrigin.Text) = 1 And txtMultOrigin.Text = "*" Then
    SectorWithInSectorList = ssFound
ElseIf InStr(txtMultOrigin.Text, ":") Then
    SectorWithInSectorList = SectorWithInRange(txtMultOrigin.Text, secx, secy)
ElseIf InStr(txtMultOrigin.Text, "#") > 0 Then
    MsgBox ("Realms not support in the sector box")
ElseIf InStr(txtMultOrigin.Text, "?") > 0 Then
    MsgBox ("Conditionals not support in the sector box")
ElseIf InStr(txtMultOrigin.Text, "\") Then
    beginPos = 1
    endPos = InStr(txtMultOrigin.Text, "\")
    While endPos > 0
        If ParseSectors(sx, sy, Mid$(txtMultOrigin.Text, beginPos, endPos - beginPos)) Then
            If sx = secx And sy = secy Then
                SectorWithInSectorList = ssFound
                Exit Function
            End If
        Else
            MsgBox ("Invalid sector format")
            Exit Function
        End If
        beginPos = endPos + 1
        endPos = InStr(beginPos, txtMultOrigin.Text, "\")
    Wend
    If ParseSectors(sx, sy, Mid$(txtMultOrigin.Text, beginPos)) Then
        If sx = secx And sy = secy Then
            SectorWithInSectorList = ssFound
            Exit Function
        End If
    Else
        MsgBox ("Invalid sector format")
        Exit Function
    End If
    SectorWithInSectorList = ssNotFound
ElseIf InStr(txtMultOrigin.Text, ",") Then
    If ParseSectors(sx, sy, txtMultOrigin.Text) Then
        If sx = secx And sy = secy Then
            SectorWithInSectorList = ssFound
        Else
            SectorWithInSectorList = ssNotFound
        End If
    Else
        MsgBox ("Invalid sector format")
    End If
Else
    MsgBox ("Invalid sector format")
End If
End Function

Private Function SectorWithInRange(strSector As String, secx As Integer, secy As Integer) As enumSectorSearch
Dim sx1 As Integer
Dim sx2 As Integer
Dim sy1 As Integer
Dim sy2 As Integer

SectorWithInRange = ssNotFound

If ParseSectorRange(sx1, sx2, sy1, sy2, strSector) Then
    If sx1 <= sx2 Then
        If sy1 <= sy2 Then
            If (secx >= sx1 And secx <= sx2) And _
               (secy >= sy1 And secy <= sy2) _
               Then
                SectorWithInRange = ssFound
            End If
        Else
            If (secx >= sx1 And secx <= sx2) And _
               (secy >= sy1 Or secy <= sy2) _
               Then
                SectorWithInRange = ssFound
            End If
        End If
    Else
        If sy1 <= sy2 Then
            If (secx >= sx1 Or secx <= sx2) And _
               (secy >= sy1 And secy <= sy2) _
               Then
                SectorWithInRange = ssFound
            End If
        Else
            If (secx >= sx1 Or secx <= sx2) And _
               (secy >= sy1 Or secy <= sy2) _
               Then
                SectorWithInRange = ssFound
            End If
        End If
    End If
Else
    MsgBox ("Invalid sector range format")
    SectorWithInRange = ssError
End If
End Function

Private Sub SendTelegram(strTelegram As String)
Dim numTelegrams As Integer
Dim strTarget() As String
Dim n As Integer
Dim strTemp As String
Dim strSubject As String

strSubject = "+++ Intelligence Report " + Format$(Now, "ddd mmm dd hh:mm:ss yyyy") + vbLf
ReDim strTarget(1 To 1)
numTelegrams = 1
strTarget(numTelegrams) = strSubject
Do
    n = InStr(strTelegram, vbLf)
    If n = 0 Then
        strTemp = strTelegram
        strTelegram = ""
    Else
        strTemp = Left$(strTelegram, n - 1)
        strTelegram = Mid$(strTelegram, n + 1)
    End If
    
    'just in case we get a real long line, split it up

    If Len(strTemp) > (1021 - Len(strSubject)) Then
        MsgBox "Error in generating telegram - line too long - telegram aborted"
        Exit Sub
    End If
    
    If (Len(strTarget(numTelegrams)) + Len(strTemp)) >= 1022 Then
        numTelegrams = numTelegrams + 1
        ReDim Preserve strTarget(1 To numTelegrams)
        strTarget(numTelegrams) = strSubject
    End If
    
    strTarget(numTelegrams) = strTarget(numTelegrams) + strTemp + vbLf
Loop While Len(strTelegram) > 0

'run thru multiple telegrams if necessay
For n = 1 To numTelegrams
    frmEmpCmd.SubmitEmpireCommand "te1 " + strTarget(n), False
    frmEmpCmd.SubmitEmpireCommand "telegram " + CStr(cbTelegram.ItemData(cbTelegram.ListIndex)), False
Next n
End Sub
