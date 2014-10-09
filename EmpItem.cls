VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmpItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"empItems"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Change Log:
'270904 rjk: Added inefficient packing, xdump chr index.

'local variable(s) to hold property value(s)
Private mvarFormName As String 'local copy
Private mvarLetter As String 'local copy
Private mvarSectorPanelLabel As String 'local copy
Private mvarConditionalName As String 'local copy
Private mvarValue As Integer 'local copy
Private mvarWeight As Integer 'local copy
Public Enum enumItemPacking
    PACKING_INEFFICIENT = 0
    PACKING_REGULAR = 1
    PACKING_WAREHOUSE = 2
    PACKING_URBAN = 3
    PACKING_BANK = 4 'Must be the last one as it is used an array size
End Enum
Private mvarPacking(PACKING_BANK) As Integer 'local copy
Private mvarIntelligenceNames As String 'local copy
Private mvarDatabaseName As String 'local copy
Private mvarSellable As Boolean 'local copy
Private mvarItemName As String 'local copy
Private mvarProductName As String 'local copy
Private mvarFormLabel As String 'local copy
Private mvarDistributionPanelLabel As String 'local copy
Private mvarChrIndex As Integer 'local copy

Public Property Let DistributionPanelLabel(ByVal vData As String)
mvarDistributionPanelLabel = vData
End Property

Public Property Get DistributionPanelLabel() As String
DistributionPanelLabel = mvarDistributionPanelLabel
End Property

Public Property Let FormLabel(ByVal vData As String)
mvarFormLabel = vData
End Property

Public Property Get FormLabel() As String
FormLabel = mvarFormLabel
End Property

Public Property Let ProductName(ByVal vData As String)
mvarProductName = vData
End Property

Public Property Get ProductName() As String
ProductName = mvarProductName
End Property

Public Property Let ChrIndex(ByVal eiChrIndex As Integer)
mvarChrIndex = eiChrIndex
End Property

Public Property Get ChrIndex() As Integer
ChrIndex = mvarChrIndex
End Property

Public Sub SaveItem()
rsItems.Seek "=", mvarLetter
If Not rsItems.NoMatch Then
    With rsItems
        .Edit
        .Fields("letter") = mvarLetter
        .Fields("value") = mvarValue
        .Fields("sell") = mvarSellable
        .Fields("lbs") = mvarWeight
        .Fields("pack_ie") = mvarPacking(PACKING_INEFFICIENT)
        .Fields("pack_rg") = mvarPacking(PACKING_REGULAR)
        .Fields("pack_wh") = mvarPacking(PACKING_WAREHOUSE)
        .Fields("pack_ur") = mvarPacking(PACKING_URBAN)
        .Fields("pack_bk") = mvarPacking(PACKING_BANK)
        .Fields("name") = mvarItemName
        .Fields("p_sname") = mvarProductName
        .Fields("cond_name") = mvarConditionalName
        .Fields("db_name") = mvarDatabaseName
        .Fields("sector_panel_label") = mvarSectorPanelLabel
        .Fields("distribution_panel_label") = mvarDistributionPanelLabel
        .Fields("form_name") = mvarFormName
        .Fields("intelligence_names") = mvarIntelligenceNames
        .Fields("form_label") = mvarFormLabel
        .Fields("chr_i") = mvarChrIndex
        .Update
    End With
End If
End Sub

Public Property Let ItemName(ByVal vData As String)
mvarItemName = vData
End Property

Public Property Get ItemName() As String
ItemName = mvarItemName
End Property

Public Property Let Sellable(bSell As Boolean)
mvarSellable = bSell
End Property

Public Property Get Sellable() As Boolean
Sellable = mvarSellable
End Property

Public Property Let DatabaseValue(rsTable As Recordset, ByVal vData As Long)
rsTable.Fields(mvarDatabaseName) = vData
End Property

Public Property Get DatabaseValue(rsTable As Recordset) As Long
DatabaseValue = rsTable.Fields(mvarDatabaseName).Value
End Property

Public Property Let DatabaseName(ByVal vData As String)
mvarDatabaseName = vData
End Property

Public Property Get DatabaseName() As String
DatabaseName = mvarDatabaseName
End Property

Public Property Let IntelligenceNames(ByVal vData As String)
mvarIntelligenceNames = vData
End Property

Public Property Get IntelligenceNames() As String
IntelligenceNames = mvarIntelligenceNames
End Property

Public Property Let Packing(packingType As enumItemPacking, ByVal vData As Integer)
mvarPacking(packingType) = vData
End Property

Public Property Get Packing(packingType As enumItemPacking) As Integer
Packing = mvarPacking(packingType)
End Property

Public Property Let Weight(ByVal vData As Integer)
mvarWeight = vData
End Property

Public Property Get Weight() As Integer
Weight = mvarWeight
End Property

Public Property Let Value(ByVal vData As Integer)
mvarValue = vData
End Property

Public Property Get Value() As Integer
Value = mvarValue
End Property

Public Property Let ConditionalName(ByVal vData As String)
mvarConditionalName = vData
End Property


Public Property Get ConditionalName() As String
ConditionalName = mvarConditionalName
End Property

Public Property Let SectorPanelLabel(ByVal vData As String)
mvarSectorPanelLabel = vData
End Property

Public Property Get SectorPanelLabel() As String
SectorPanelLabel = mvarSectorPanelLabel
End Property

Public Property Let Letter(ByVal vData As String)
mvarLetter = vData
End Property

Public Property Get Letter() As String
Letter = mvarLetter
End Property

Public Property Let FormName(ByVal vData As String)
mvarFormName = vData
End Property

Public Property Get FormName() As String
FormName = mvarFormName
End Property

