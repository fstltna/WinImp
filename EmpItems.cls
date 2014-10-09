VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmpItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"EmpItem"
Attribute VB_Ext_KEY = "Member0" ,"EmpItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Change Log:
'270904 rjk: Added inefficient packing, xdump chr index.
'011204 rjk: The xdump chr index now starts at zero.
'180206 rjk: Switch from item to eiItem.
'270608 rjk: Stored the update Item table when done loading the xdump item char.

'local variable to hold collection
Private mCol As Collection

Public Function FindByConditionalName(strConditionalName As String) As EmpItem
Dim eiItem As EmpItem

For Each eiItem In mCol
    If eiItem.ConditionalName = strConditionalName Then
        Set FindByConditionalName = eiItem
        Exit Function
    End If
Next eiItem
Set FindByConditionalName = Nothing
End Function

Public Function FindByDatabaseName(strDatabaseName As String) As EmpItem
Dim eiItem As EmpItem

For Each eiItem In mCol
    If eiItem.DatabaseName = strDatabaseName Then
        Set FindByDatabaseName = eiItem
        Exit Function
    End If
Next eiItem
Set FindByDatabaseName = Nothing
End Function

Public Function FindByFormLabel(strFormLabel As String) As EmpItem
Dim eiItem As EmpItem

For Each eiItem In mCol
    If eiItem.FormLabel = strFormLabel Then
        Set FindByFormLabel = eiItem
        Exit Function
    End If
Next eiItem
Set FindByFormLabel = Nothing
End Function

Public Function FindByFormName(strFormName As String) As EmpItem
Dim eiItem As EmpItem

For Each eiItem In mCol
    If eiItem.FormName = strFormName Then
        Set FindByFormName = eiItem
        Exit Function
    End If
Next eiItem
Set FindByFormName = Nothing
End Function

Public Function FindByLetter(Letter As String) As EmpItem
Dim eiItem As EmpItem

For Each eiItem In mCol
    If eiItem.Letter = Left$(Letter, 1) Then
        Set FindByLetter = eiItem
        Exit Function
    End If
Next eiItem
Set FindByLetter = Nothing
End Function

Public Function FindByChrIndex(iIndex As Integer) As EmpItem
Dim eiItem As EmpItem

For Each eiItem In mCol
    If eiItem.ChrIndex = iIndex Then
        Set FindByChrIndex = eiItem
        Exit Function
    End If
Next eiItem
Set FindByChrIndex = Nothing
End Function

Public Function Add(FormName As String, Letter As String, SectorPanelLabel As String, ConditionalName As String, Value As Integer, Weight As Integer, Packing() As Integer, IntelligenceNames As String, DatabaseName As String, DatabaseValue As Integer, Sellable As Boolean, ChrIndex As Integer, Optional sKey As String) As EmpItem
'create a new object
Dim objNewMember As EmpItem
Set objNewMember = New EmpItem

'set the properties passed into the method
objNewMember.FormName = FormName
objNewMember.Letter = Letter
objNewMember.SectorPanelLabel = SectorPanelLabel
objNewMember.ConditionalName = ConditionalName
objNewMember.Value = Value
objNewMember.Weight = Weight
objNewMember.Packing(PACKING_INEFFICIENT) = Packing(PACKING_INEFFICIENT)
objNewMember.Packing(PACKING_REGULAR) = Packing(PACKING_REGULAR)
objNewMember.Packing(PACKING_WAREHOUSE) = Packing(PACKING_WAREHOUSE)
objNewMember.Packing(PACKING_URBAN) = Packing(PACKING_URBAN)
objNewMember.Packing(PACKING_BANK) = Packing(PACKING_BANK)
objNewMember.IntelligenceNames = IntelligenceNames
objNewMember.DatabaseName = DatabaseName
'objNewMember.DatabaseValue = DatabaseValue
objNewMember.Sellable = Sellable
objNewMember.ChrIndex = ChrIndex

If Len(sKey) = 0 Then
    mCol.Add objNewMember
Else
    mCol.Add objNewMember, sKey
End If

'return the object created
Set Add = objNewMember
Set objNewMember = Nothing
End Function

Public Sub AddItem(eiItem As EmpItem, Optional sKey As String)
If Len(sKey) = 0 Then
    mCol.Add eiItem
Else
    mCol.Add eiItem, sKey
End If
End Sub

Public Property Get Item(vntIndexKey As Variant) As EmpItem
Attribute Item.VB_UserMemId = 0
  'used when referencing an element in the collection
  'vntIndexKey contains either the Index or Key to the collection,
  'this is why it is declared as a Variant
  'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
'used when retrieving the number of elements in the
'collection. Syntax: Debug.Print x.Count
Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
'used when removing an element from the collection
'vntIndexKey contains either the Index or Key, which is why
'it is declared as a Variant
'Syntax: x.Remove(xyz)

mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'this property allows you to enumerate
'this collection with the For...Each syntax
Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
Dim eiItem As EmpItem
'creates the collection when this class is created
Set mCol = New Collection
With rsItems
    If .RecordCount = 0 Then
        FillItemsTablewithDefaults
    End If
    .MoveFirst
    While Not .EOF
        Set eiItem = New EmpItem
        eiItem.Letter = .Fields("letter")
        eiItem.Value = .Fields("value")
        eiItem.Sellable = .Fields("sell")
        eiItem.Weight = .Fields("lbs")
        eiItem.Packing(PACKING_INEFFICIENT) = .Fields("pack_ie")
        eiItem.Packing(PACKING_REGULAR) = .Fields("pack_rg")
        eiItem.Packing(PACKING_WAREHOUSE) = .Fields("pack_wh")
        eiItem.Packing(PACKING_URBAN) = .Fields("pack_ur")
        eiItem.Packing(PACKING_BANK) = .Fields("pack_bk")
        eiItem.FormName = .Fields("form_name")
        eiItem.ItemName = .Fields("name")
        eiItem.IntelligenceNames = .Fields("intelligence_names")
        eiItem.ConditionalName = .Fields("cond_name")
        eiItem.DatabaseName = .Fields("db_name")
        eiItem.SectorPanelLabel = .Fields("sector_panel_label")
        eiItem.DistributionPanelLabel = .Fields("distribution_panel_label")
        eiItem.ProductName = .Fields("p_sname")
        eiItem.FormLabel = .Fields("form_label")
        eiItem.ChrIndex = .Fields("chr_i")
        mCol.Add eiItem
        .MoveNext
    Wend
End With
End Sub

Private Sub FillItemsTablewithDefaults()
With rsItems
'     c     1   no   1 10 10 10 10 civilians
    .AddNew
    .Fields("letter") = "c"
    .Fields("value") = 1
    .Fields("sell") = False
    .Fields("lbs") = 1
    .Fields("pack_ie") = 1
    .Fields("pack_rg") = 10
    .Fields("pack_wh") = 10
    .Fields("pack_ur") = 10
    .Fields("pack_bk") = 10
    .Fields("name") = "civilians"
    .Fields("intelligence_names") = "civilians"
    .Fields("p_sname") = "civilians"
    .Fields("cond_name") = "civ"
    .Fields("db_name") = "civ"
    .Fields("sector_panel_label") = "Civ:"
    .Fields("distribution_panel_label") = "Civ"
    .Fields("form_name") = "civilians"
    .Fields("form_label") = "Civilians"
    .Fields("chr_i") = 0
    .Update
'     m     0   no   1  1  1  1  1 military
    .AddNew
    .Fields("letter") = "m"
    .Fields("value") = 0
    .Fields("sell") = False
    .Fields("lbs") = 1
    .Fields("pack_ie") = 1
    .Fields("pack_rg") = 1
    .Fields("pack_wh") = 1
    .Fields("pack_ur") = 1
    .Fields("pack_bk") = 1
    .Fields("name") = "military"
    .Fields("intelligence_names") = "military"
    .Fields("p_sname") = "military"
    .Fields("cond_name") = "mil"
    .Fields("db_name") = "mil"
    .Fields("sector_panel_label") = "Mil:"
    .Fields("distribution_panel_label") = "Mil"
    .Fields("form_name") = "military"
    .Fields("form_label") = "Military"
    .Fields("chr_i") = 1
    .Update
'     s     5  yes   1  1 10  1  1 shells
    .AddNew
    .Fields("letter") = "s"
    .Fields("value") = 5
    .Fields("sell") = True
    .Fields("lbs") = 1
    .Fields("pack_ie") = 1
    .Fields("pack_rg") = 1
    .Fields("pack_wh") = 10
    .Fields("pack_ur") = 1
    .Fields("pack_bk") = 1
    .Fields("name") = "shells"
    .Fields("intelligence_names") = "shells"
    .Fields("p_sname") = "shells"
    .Fields("cond_name") = "shell"
    .Fields("db_name") = "shell"
    .Fields("sector_panel_label") = "Shell"
    .Fields("distribution_panel_label") = "Shell"
    .Fields("form_name") = "shells"
    .Fields("form_label") = "Shells"
    .Fields("chr_i") = 2
    .Update
'     g    60  yes  10  1 10  1  1 guns
    .AddNew
    .Fields("letter") = "g"
    .Fields("value") = 60
    .Fields("sell") = True
    .Fields("lbs") = 10
    .Fields("pack_ie") = 1
    .Fields("pack_rg") = 1
    .Fields("pack_wh") = 10
    .Fields("pack_ur") = 1
    .Fields("pack_bk") = 1
    .Fields("name") = "guns"
    .Fields("intelligence_names") = "guns"
    .Fields("p_sname") = "guns"
    .Fields("cond_name") = "gun"
    .Fields("db_name") = "gun"
    .Fields("sector_panel_label") = "Guns"
    .Fields("distribution_panel_label") = "Guns"
    .Fields("form_name") = "guns"
    .Fields("form_label") = "Guns"
    .Fields("chr_i") = 3
    .Update
'     p     4  yes   1  1 10  1  1 petrol
    .AddNew
    .Fields("letter") = "p"
    .Fields("value") = 4
    .Fields("sell") = True
    .Fields("lbs") = 1
    .Fields("pack_ie") = 1
    .Fields("pack_rg") = 1
    .Fields("pack_wh") = 10
    .Fields("pack_ur") = 1
    .Fields("pack_bk") = 1
    .Fields("name") = "petrol"
    .Fields("intelligence_names") = "petrol"
    .Fields("p_sname") = "petrol"
    .Fields("cond_name") = "pet"
    .Fields("db_name") = "pet"
    .Fields("sector_panel_label") = "Pet"
    .Fields("distribution_panel_label") = "Pet"
    .Fields("form_name") = "petrol"
    .Fields("form_label") = "Pet"
    .Fields("chr_i") = 4
    .Update
'     i     2  yes   1  1 10  1  1 iron ore
    .AddNew
    .Fields("letter") = "i"
    .Fields("value") = 2
    .Fields("sell") = True
    .Fields("lbs") = 1
    .Fields("pack_ie") = 1
    .Fields("pack_rg") = 1
    .Fields("pack_wh") = 10
    .Fields("pack_ur") = 1
    .Fields("pack_bk") = 1
    .Fields("name") = "iron ore"
    .Fields("intelligence_names") = "iron ore"
    .Fields("p_sname") = "iron"
    .Fields("cond_name") = "iron"
    .Fields("db_name") = "iron"
    .Fields("sector_panel_label") = "Iron"
    .Fields("distribution_panel_label") = "Iron"
    .Fields("form_name") = "iron"
    .Fields("form_label") = "Iron"
    .Fields("chr_i") = 5
    .Update
'     d    20  yes   5  1 10  1  1 dust (gold)
    .AddNew
    .Fields("letter") = "d"
    .Fields("value") = 20
    .Fields("sell") = True
    .Fields("lbs") = 5
    .Fields("pack_ie") = 1
    .Fields("pack_rg") = 1
    .Fields("pack_wh") = 10
    .Fields("pack_ur") = 1
    .Fields("pack_bk") = 1
    .Fields("name") = "dust (gold)"
    .Fields("intelligence_names") = "dust"
    .Fields("p_sname") = "dust"
    .Fields("cond_name") = "dust"
    .Fields("db_name") = "dust"
    .Fields("sector_panel_label") = "Dust"
    .Fields("distribution_panel_label") = "Dust"
    .Fields("form_name") = "dust"
    .Fields("form_label") = "Dust"
    .Fields("chr_i") = 6
    .Update
'     b   280  yes  50  1  5  1  4 bars of gold
    .AddNew
    .Fields("letter") = "b"
    .Fields("value") = 280
    .Fields("sell") = True
    .Fields("lbs") = 50
    .Fields("pack_ie") = 1
    .Fields("pack_rg") = 1
    .Fields("pack_wh") = 5
    .Fields("pack_ur") = 1
    .Fields("pack_bk") = 4
    .Fields("name") = "bars of gold"
    .Fields("intelligence_names") = "bars"
    .Fields("p_sname") = "bars"
    .Fields("cond_name") = "bar"
    .Fields("db_name") = "bar"
    .Fields("sector_panel_label") = "Bars"
    .Fields("distribution_panel_label") = "Bars"
    .Fields("form_name") = "bars"
    .Fields("form_label") = "Bars"
    .Fields("chr_i") = 7
    .Update
'     f     0  yes   1  1 10  1  1 food
    .AddNew
    .Fields("letter") = "f"
    .Fields("value") = 0
    .Fields("sell") = True
    .Fields("lbs") = 1
    .Fields("pack_ie") = 1
    .Fields("pack_rg") = 1
    .Fields("pack_wh") = 10
    .Fields("pack_ur") = 1
    .Fields("pack_bk") = 1
    .Fields("name") = "food"
    .Fields("intelligence_names") = "food"
    .Fields("p_sname") = "food"
    .Fields("cond_name") = "food"
    .Fields("db_name") = "food"
    .Fields("sector_panel_label") = "Food"
    .Fields("distribution_panel_label") = "Food"
    .Fields("form_name") = "food"
    .Fields("form_label") = "Food"
    .Fields("chr_i") = 8
    .Update
'     o     8  yes   1  1 10  1  1 oil
    .AddNew
    .Fields("letter") = "o"
    .Fields("value") = 8
    .Fields("sell") = True
    .Fields("lbs") = 1
    .Fields("pack_ie") = 1
    .Fields("pack_rg") = 1
    .Fields("pack_wh") = 10
    .Fields("pack_ur") = 1
    .Fields("pack_bk") = 1
    .Fields("name") = "oil"
    .Fields("intelligence_names") = "oil"
    .Fields("p_sname") = "oil"
    .Fields("cond_name") = "oil"
    .Fields("db_name") = "oil"
    .Fields("sector_panel_label") = "Oil"
    .Fields("distribution_panel_label") = "Oil"
    .Fields("form_name") = "oil"
    .Fields("form_label") = "Oil"
    .Fields("chr_i") = 9
    .Update
'     l     2  yes   1  1 10  1  1 light products
    .AddNew
    .Fields("letter") = "l"
    .Fields("value") = 2
    .Fields("sell") = True
    .Fields("lbs") = 1
    .Fields("pack_ie") = 1
    .Fields("pack_rg") = 1
    .Fields("pack_wh") = 10
    .Fields("pack_ur") = 1
    .Fields("pack_bk") = 1
    .Fields("name") = "light products"
    .Fields("intelligence_names") = "light products"
    .Fields("p_sname") = "lcm"
    .Fields("cond_name") = "lcm"
    .Fields("db_name") = "lcm"
    .Fields("sector_panel_label") = "Lcms"
    .Fields("distribution_panel_label") = "Lcms"
    .Fields("form_name") = "lcm"
    .Fields("form_label") = "Lcms"
    .Fields("chr_i") = 10
    .Update
'     h     4  yes   1  1 10  1  1 heavy products
    .AddNew
    .Fields("letter") = "h"
    .Fields("value") = 4
    .Fields("sell") = True
    .Fields("lbs") = 1
    .Fields("pack_ie") = 1
    .Fields("pack_rg") = 1
    .Fields("pack_wh") = 10
    .Fields("pack_ur") = 1
    .Fields("pack_bk") = 1
    .Fields("name") = "heavy products"
    .Fields("intelligence_names") = "heavy products"
    .Fields("p_sname") = "hcm"
    .Fields("cond_name") = "hcm"
    .Fields("db_name") = "hcm"
    .Fields("sector_panel_label") = "Hcms"
    .Fields("distribution_panel_label") = "Hcms"
    .Fields("form_name") = "hcm"
    .Fields("form_label") = "Hcms"
    .Fields("chr_i") = 11
    .Update
'     r   150  yes   8  1 10  1  1 radioactive materials
    .AddNew
    .Fields("letter") = "r"
    .Fields("value") = 150
    .Fields("sell") = True
    .Fields("lbs") = 8
    .Fields("pack_ie") = 1
    .Fields("pack_rg") = 1
    .Fields("pack_wh") = 10
    .Fields("pack_ur") = 1
    .Fields("pack_bk") = 1
    .Fields("name") = "radioactive materials"
    .Fields("intelligence_names") = "radioactive materials"
    .Fields("p_sname") = "rads"
    .Fields("cond_name") = "rad"
    .Fields("db_name") = "rad"
    .Fields("sector_panel_label") = "Rads"
    .Fields("distribution_panel_label") = "Rads"
    .Fields("form_name") = "rads"
    .Fields("form_label") = "Rads"
    .Fields("chr_i") = 13
    .Update
'     u     1  yes   2  1  2  1  1 uncompensated workers
    .AddNew
    .Fields("letter") = "u"
    .Fields("value") = 1
    .Fields("sell") = True
    .Fields("lbs") = 2
    .Fields("pack_ie") = 1
    .Fields("pack_rg") = 1
    .Fields("pack_wh") = 2
    .Fields("pack_ur") = 1
    .Fields("pack_bk") = 1
    .Fields("name") = "uncompensated workers"
    .Fields("intelligence_names") = "uncompensated workers"
    .Fields("p_sname") = "uw"
    .Fields("cond_name") = "uw"
    .Fields("db_name") = "uw"
    .Fields("sector_panel_label") = "Uw:"
    .Fields("distribution_panel_label") = "Uws"
    .Fields("form_name") = "uw"
    .Fields("form_label") = "unc."
    .Fields("chr_i") = 12
    .Update
End With
End Sub

Public Sub UpdateDatabase()
Dim eiItem As EmpItem
DeleteAllRecords rsItems
For Each eiItem In mCol
    With rsItems
    .AddNew
    .Fields("letter") = eiItem.Letter
    .Fields("value") = eiItem.Value
    .Fields("sell") = eiItem.Sellable
    .Fields("lbs") = eiItem.Weight
    .Fields("pack_ie") = eiItem.Packing(PACKING_INEFFICIENT)
    .Fields("pack_rg") = eiItem.Packing(PACKING_REGULAR)
    .Fields("pack_wh") = eiItem.Packing(PACKING_WAREHOUSE)
    .Fields("pack_ur") = eiItem.Packing(PACKING_URBAN)
    .Fields("pack_bk") = eiItem.Packing(PACKING_BANK)
    .Fields("name") = eiItem.ItemName
    .Fields("intelligence_names") = eiItem.IntelligenceNames
    .Fields("p_sname") = "eiItem.ProductName"
    .Fields("cond_name") = eiItem.ConditionalName
    .Fields("db_name") = eiItem.DatabaseName
    .Fields("sector_panel_label") = eiItem.SectorPanelLabel
    .Fields("distribution_panel_label") = eiItem.DistributionPanelLabel
    .Fields("form_name") = eiItem.FormName
    .Fields("form_label") = eiItem.FormLabel
    .Fields("chr_i") = eiItem.ChrIndex
    .Update
    End With
Next eiItem
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub
