Option Strict Off
Option Explicit On
Friend Class EmpItems
	Implements System.Collections.IEnumerable
	
	'Change Log:
	'270904 rjk: Added inefficient packing, xdump chr index.
	'011204 rjk: The xdump chr index now starts at zero.
	'180206 rjk: Switch from item to eiItem.
	'270608 rjk: Stored the update Item table when done loading the xdump item char.
	
	'local variable to hold collection
	Private mCol As Collection
	
	Public Function FindByConditionalName(ByRef strConditionalName As String) As EmpItem
		Dim eiItem As EmpItem
		
		For	Each eiItem In mCol
			If eiItem.ConditionalName = strConditionalName Then
				FindByConditionalName = eiItem
				Exit Function
			End If
		Next eiItem
		'UPGRADE_NOTE: Object FindByConditionalName may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		FindByConditionalName = Nothing
	End Function
	
	Public Function FindByDatabaseName(ByRef strDatabaseName As String) As EmpItem
		Dim eiItem As EmpItem
		
		For	Each eiItem In mCol
			If eiItem.DatabaseName = strDatabaseName Then
				FindByDatabaseName = eiItem
				Exit Function
			End If
		Next eiItem
		'UPGRADE_NOTE: Object FindByDatabaseName may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		FindByDatabaseName = Nothing
	End Function
	
	Public Function FindByFormLabel(ByRef strFormLabel As String) As EmpItem
		Dim eiItem As EmpItem
		
		For	Each eiItem In mCol
			If eiItem.FormLabel = strFormLabel Then
				FindByFormLabel = eiItem
				Exit Function
			End If
		Next eiItem
		'UPGRADE_NOTE: Object FindByFormLabel may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		FindByFormLabel = Nothing
	End Function
	
	Public Function FindByFormName(ByRef strFormName As String) As EmpItem
		Dim eiItem As EmpItem
		
		For	Each eiItem In mCol
			If eiItem.FormName = strFormName Then
				FindByFormName = eiItem
				Exit Function
			End If
		Next eiItem
		'UPGRADE_NOTE: Object FindByFormName may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		FindByFormName = Nothing
	End Function
	
	Public Function FindByLetter(ByRef Letter As String) As EmpItem
		Dim eiItem As EmpItem
		
		For	Each eiItem In mCol
			If eiItem.Letter = Left(Letter, 1) Then
				FindByLetter = eiItem
				Exit Function
			End If
		Next eiItem
		'UPGRADE_NOTE: Object FindByLetter may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		FindByLetter = Nothing
	End Function
	
	Public Function FindByChrIndex(ByRef iIndex As Short) As EmpItem
		Dim eiItem As EmpItem
		
		For	Each eiItem In mCol
			If eiItem.ChrIndex = iIndex Then
				FindByChrIndex = eiItem
				Exit Function
			End If
		Next eiItem
		'UPGRADE_NOTE: Object FindByChrIndex may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		FindByChrIndex = Nothing
	End Function
	
	Public Function Add(ByRef FormName As String, ByRef Letter As String, ByRef SectorPanelLabel As String, ByRef ConditionalName As String, ByRef Value As Short, ByRef Weight As Short, ByRef Packing() As Short, ByRef IntelligenceNames As String, ByRef DatabaseName As String, ByRef DatabaseValue As Short, ByRef Sellable As Boolean, ByRef ChrIndex As Short, Optional ByRef sKey As String = "") As EmpItem
		'create a new object
		Dim objNewMember As EmpItem
		objNewMember = New EmpItem
		
		'set the properties passed into the method
		objNewMember.FormName = FormName
		objNewMember.Letter = Letter
		objNewMember.SectorPanelLabel = SectorPanelLabel
		objNewMember.ConditionalName = ConditionalName
		objNewMember.Value = Value
		objNewMember.Weight = Weight
		objNewMember.Packing(EmpItem.enumItemPacking.PACKING_INEFFICIENT) = Packing(EmpItem.enumItemPacking.PACKING_INEFFICIENT)
		objNewMember.Packing(EmpItem.enumItemPacking.PACKING_REGULAR) = Packing(EmpItem.enumItemPacking.PACKING_REGULAR)
		objNewMember.Packing(EmpItem.enumItemPacking.PACKING_WAREHOUSE) = Packing(EmpItem.enumItemPacking.PACKING_WAREHOUSE)
		objNewMember.Packing(EmpItem.enumItemPacking.PACKING_URBAN) = Packing(EmpItem.enumItemPacking.PACKING_URBAN)
		objNewMember.Packing(EmpItem.enumItemPacking.PACKING_BANK) = Packing(EmpItem.enumItemPacking.PACKING_BANK)
		objNewMember.IntelligenceNames = IntelligenceNames
		objNewMember.DatabaseName = DatabaseName
		'objNewMember.DatabaseValue = DatabaseValue
		objNewMember.Sellable = Sellable
		objNewMember.ChrIndex = ChrIndex
		
		If Len(sKey) = 0 Then
			mCol.Add(objNewMember)
		Else
			mCol.Add(objNewMember, sKey)
		End If
		
		'return the object created
		Add = objNewMember
		'UPGRADE_NOTE: Object objNewMember may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objNewMember = Nothing
	End Function
	
	Public Sub AddItem(ByRef eiItem As EmpItem, Optional ByRef sKey As String = "")
		If Len(sKey) = 0 Then
			mCol.Add(eiItem)
		Else
			mCol.Add(eiItem, sKey)
		End If
	End Sub
	
	Default Public ReadOnly Property Item(ByVal vntIndexKey As Object) As EmpItem
		Get
			'used when referencing an element in the collection
			'vntIndexKey contains either the Index or Key to the collection,
			'this is why it is declared as a Variant
			'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
			Item = mCol.Item(vntIndexKey)
		End Get
	End Property
	
	Public ReadOnly Property Count() As Integer
		Get
			'used when retrieving the number of elements in the
			'collection. Syntax: Debug.Print x.Count
			Count = mCol.Count()
		End Get
	End Property
	
	'UPGRADE_NOTE: NewEnum property was commented out. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B3FC1610-34F3-43F5-86B7-16C984F0E88E"'
	'Public ReadOnly Property NewEnum() As stdole.IUnknown
		'Get
			'this property allows you to enumerate
			'this collection with the For...Each syntax
			'NewEnum = mCol._NewEnum
		'End Get
	'End Property
	
	Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
		'UPGRADE_TODO: Uncomment and change the following line to return the collection enumerator. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="95F9AAD0-1319-4921-95F0-B9D3C4FF7F1C"'
		'GetEnumerator = mCol.GetEnumerator
	End Function
	
	Public Sub Remove(ByRef vntIndexKey As Object)
		'used when removing an element from the collection
		'vntIndexKey contains either the Index or Key, which is why
		'it is declared as a Variant
		'Syntax: x.Remove(xyz)
		
		mCol.Remove(vntIndexKey)
	End Sub
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		Dim eiItem As EmpItem
		'creates the collection when this class is created
		mCol = New Collection
		With rsItems
			If .RecordCount = 0 Then
				FillItemsTablewithDefaults()
			End If
			.MoveFirst()
			While Not .EOF
				eiItem = New EmpItem
				eiItem.Letter = .Fields("letter").Value
				eiItem.Value = .Fields("value").Value
				eiItem.Sellable = .Fields("sell").Value
				eiItem.Weight = .Fields("lbs").Value
				eiItem.Packing(EmpItem.enumItemPacking.PACKING_INEFFICIENT) = .Fields("pack_ie").Value
				eiItem.Packing(EmpItem.enumItemPacking.PACKING_REGULAR) = .Fields("pack_rg").Value
				eiItem.Packing(EmpItem.enumItemPacking.PACKING_WAREHOUSE) = .Fields("pack_wh").Value
				eiItem.Packing(EmpItem.enumItemPacking.PACKING_URBAN) = .Fields("pack_ur").Value
				eiItem.Packing(EmpItem.enumItemPacking.PACKING_BANK) = .Fields("pack_bk").Value
				eiItem.FormName = .Fields("form_name").Value
				eiItem.ItemName = .Fields("name").Value
				eiItem.IntelligenceNames = .Fields("intelligence_names").Value
				eiItem.ConditionalName = .Fields("cond_name").Value
				eiItem.DatabaseName = .Fields("db_name").Value
				eiItem.SectorPanelLabel = .Fields("sector_panel_label").Value
				eiItem.DistributionPanelLabel = .Fields("distribution_panel_label").Value
				eiItem.ProductName = .Fields("p_sname").Value
				eiItem.FormLabel = .Fields("form_label").Value
				eiItem.ChrIndex = .Fields("chr_i").Value
				mCol.Add(eiItem)
				.MoveNext()
			End While
		End With
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	Private Sub FillItemsTablewithDefaults()
		With rsItems
			'     c     1   no   1 10 10 10 10 civilians
			.AddNew()
			.Fields("letter").Value = "c"
			.Fields("value").Value = 1
			.Fields("sell").Value = False
			.Fields("lbs").Value = 1
			.Fields("pack_ie").Value = 1
			.Fields("pack_rg").Value = 10
			.Fields("pack_wh").Value = 10
			.Fields("pack_ur").Value = 10
			.Fields("pack_bk").Value = 10
			.Fields("name").Value = "civilians"
			.Fields("intelligence_names").Value = "civilians"
			.Fields("p_sname").Value = "civilians"
			.Fields("cond_name").Value = "civ"
			.Fields("db_name").Value = "civ"
			.Fields("sector_panel_label").Value = "Civ:"
			.Fields("distribution_panel_label").Value = "Civ"
			.Fields("form_name").Value = "civilians"
			.Fields("form_label").Value = "Civilians"
			.Fields("chr_i").Value = 0
			.Update()
			'     m     0   no   1  1  1  1  1 military
			.AddNew()
			.Fields("letter").Value = "m"
			.Fields("value").Value = 0
			.Fields("sell").Value = False
			.Fields("lbs").Value = 1
			.Fields("pack_ie").Value = 1
			.Fields("pack_rg").Value = 1
			.Fields("pack_wh").Value = 1
			.Fields("pack_ur").Value = 1
			.Fields("pack_bk").Value = 1
			.Fields("name").Value = "military"
			.Fields("intelligence_names").Value = "military"
			.Fields("p_sname").Value = "military"
			.Fields("cond_name").Value = "mil"
			.Fields("db_name").Value = "mil"
			.Fields("sector_panel_label").Value = "Mil:"
			.Fields("distribution_panel_label").Value = "Mil"
			.Fields("form_name").Value = "military"
			.Fields("form_label").Value = "Military"
			.Fields("chr_i").Value = 1
			.Update()
			'     s     5  yes   1  1 10  1  1 shells
			.AddNew()
			.Fields("letter").Value = "s"
			.Fields("value").Value = 5
			.Fields("sell").Value = True
			.Fields("lbs").Value = 1
			.Fields("pack_ie").Value = 1
			.Fields("pack_rg").Value = 1
			.Fields("pack_wh").Value = 10
			.Fields("pack_ur").Value = 1
			.Fields("pack_bk").Value = 1
			.Fields("name").Value = "shells"
			.Fields("intelligence_names").Value = "shells"
			.Fields("p_sname").Value = "shells"
			.Fields("cond_name").Value = "shell"
			.Fields("db_name").Value = "shell"
			.Fields("sector_panel_label").Value = "Shell"
			.Fields("distribution_panel_label").Value = "Shell"
			.Fields("form_name").Value = "shells"
			.Fields("form_label").Value = "Shells"
			.Fields("chr_i").Value = 2
			.Update()
			'     g    60  yes  10  1 10  1  1 guns
			.AddNew()
			.Fields("letter").Value = "g"
			.Fields("value").Value = 60
			.Fields("sell").Value = True
			.Fields("lbs").Value = 10
			.Fields("pack_ie").Value = 1
			.Fields("pack_rg").Value = 1
			.Fields("pack_wh").Value = 10
			.Fields("pack_ur").Value = 1
			.Fields("pack_bk").Value = 1
			.Fields("name").Value = "guns"
			.Fields("intelligence_names").Value = "guns"
			.Fields("p_sname").Value = "guns"
			.Fields("cond_name").Value = "gun"
			.Fields("db_name").Value = "gun"
			.Fields("sector_panel_label").Value = "Guns"
			.Fields("distribution_panel_label").Value = "Guns"
			.Fields("form_name").Value = "guns"
			.Fields("form_label").Value = "Guns"
			.Fields("chr_i").Value = 3
			.Update()
			'     p     4  yes   1  1 10  1  1 petrol
			.AddNew()
			.Fields("letter").Value = "p"
			.Fields("value").Value = 4
			.Fields("sell").Value = True
			.Fields("lbs").Value = 1
			.Fields("pack_ie").Value = 1
			.Fields("pack_rg").Value = 1
			.Fields("pack_wh").Value = 10
			.Fields("pack_ur").Value = 1
			.Fields("pack_bk").Value = 1
			.Fields("name").Value = "petrol"
			.Fields("intelligence_names").Value = "petrol"
			.Fields("p_sname").Value = "petrol"
			.Fields("cond_name").Value = "pet"
			.Fields("db_name").Value = "pet"
			.Fields("sector_panel_label").Value = "Pet"
			.Fields("distribution_panel_label").Value = "Pet"
			.Fields("form_name").Value = "petrol"
			.Fields("form_label").Value = "Pet"
			.Fields("chr_i").Value = 4
			.Update()
			'     i     2  yes   1  1 10  1  1 iron ore
			.AddNew()
			.Fields("letter").Value = "i"
			.Fields("value").Value = 2
			.Fields("sell").Value = True
			.Fields("lbs").Value = 1
			.Fields("pack_ie").Value = 1
			.Fields("pack_rg").Value = 1
			.Fields("pack_wh").Value = 10
			.Fields("pack_ur").Value = 1
			.Fields("pack_bk").Value = 1
			.Fields("name").Value = "iron ore"
			.Fields("intelligence_names").Value = "iron ore"
			.Fields("p_sname").Value = "iron"
			.Fields("cond_name").Value = "iron"
			.Fields("db_name").Value = "iron"
			.Fields("sector_panel_label").Value = "Iron"
			.Fields("distribution_panel_label").Value = "Iron"
			.Fields("form_name").Value = "iron"
			.Fields("form_label").Value = "Iron"
			.Fields("chr_i").Value = 5
			.Update()
			'     d    20  yes   5  1 10  1  1 dust (gold)
			.AddNew()
			.Fields("letter").Value = "d"
			.Fields("value").Value = 20
			.Fields("sell").Value = True
			.Fields("lbs").Value = 5
			.Fields("pack_ie").Value = 1
			.Fields("pack_rg").Value = 1
			.Fields("pack_wh").Value = 10
			.Fields("pack_ur").Value = 1
			.Fields("pack_bk").Value = 1
			.Fields("name").Value = "dust (gold)"
			.Fields("intelligence_names").Value = "dust"
			.Fields("p_sname").Value = "dust"
			.Fields("cond_name").Value = "dust"
			.Fields("db_name").Value = "dust"
			.Fields("sector_panel_label").Value = "Dust"
			.Fields("distribution_panel_label").Value = "Dust"
			.Fields("form_name").Value = "dust"
			.Fields("form_label").Value = "Dust"
			.Fields("chr_i").Value = 6
			.Update()
			'     b   280  yes  50  1  5  1  4 bars of gold
			.AddNew()
			.Fields("letter").Value = "b"
			.Fields("value").Value = 280
			.Fields("sell").Value = True
			.Fields("lbs").Value = 50
			.Fields("pack_ie").Value = 1
			.Fields("pack_rg").Value = 1
			.Fields("pack_wh").Value = 5
			.Fields("pack_ur").Value = 1
			.Fields("pack_bk").Value = 4
			.Fields("name").Value = "bars of gold"
			.Fields("intelligence_names").Value = "bars"
			.Fields("p_sname").Value = "bars"
			.Fields("cond_name").Value = "bar"
			.Fields("db_name").Value = "bar"
			.Fields("sector_panel_label").Value = "Bars"
			.Fields("distribution_panel_label").Value = "Bars"
			.Fields("form_name").Value = "bars"
			.Fields("form_label").Value = "Bars"
			.Fields("chr_i").Value = 7
			.Update()
			'     f     0  yes   1  1 10  1  1 food
			.AddNew()
			.Fields("letter").Value = "f"
			.Fields("value").Value = 0
			.Fields("sell").Value = True
			.Fields("lbs").Value = 1
			.Fields("pack_ie").Value = 1
			.Fields("pack_rg").Value = 1
			.Fields("pack_wh").Value = 10
			.Fields("pack_ur").Value = 1
			.Fields("pack_bk").Value = 1
			.Fields("name").Value = "food"
			.Fields("intelligence_names").Value = "food"
			.Fields("p_sname").Value = "food"
			.Fields("cond_name").Value = "food"
			.Fields("db_name").Value = "food"
			.Fields("sector_panel_label").Value = "Food"
			.Fields("distribution_panel_label").Value = "Food"
			.Fields("form_name").Value = "food"
			.Fields("form_label").Value = "Food"
			.Fields("chr_i").Value = 8
			.Update()
			'     o     8  yes   1  1 10  1  1 oil
			.AddNew()
			.Fields("letter").Value = "o"
			.Fields("value").Value = 8
			.Fields("sell").Value = True
			.Fields("lbs").Value = 1
			.Fields("pack_ie").Value = 1
			.Fields("pack_rg").Value = 1
			.Fields("pack_wh").Value = 10
			.Fields("pack_ur").Value = 1
			.Fields("pack_bk").Value = 1
			.Fields("name").Value = "oil"
			.Fields("intelligence_names").Value = "oil"
			.Fields("p_sname").Value = "oil"
			.Fields("cond_name").Value = "oil"
			.Fields("db_name").Value = "oil"
			.Fields("sector_panel_label").Value = "Oil"
			.Fields("distribution_panel_label").Value = "Oil"
			.Fields("form_name").Value = "oil"
			.Fields("form_label").Value = "Oil"
			.Fields("chr_i").Value = 9
			.Update()
			'     l     2  yes   1  1 10  1  1 light products
			.AddNew()
			.Fields("letter").Value = "l"
			.Fields("value").Value = 2
			.Fields("sell").Value = True
			.Fields("lbs").Value = 1
			.Fields("pack_ie").Value = 1
			.Fields("pack_rg").Value = 1
			.Fields("pack_wh").Value = 10
			.Fields("pack_ur").Value = 1
			.Fields("pack_bk").Value = 1
			.Fields("name").Value = "light products"
			.Fields("intelligence_names").Value = "light products"
			.Fields("p_sname").Value = "lcm"
			.Fields("cond_name").Value = "lcm"
			.Fields("db_name").Value = "lcm"
			.Fields("sector_panel_label").Value = "Lcms"
			.Fields("distribution_panel_label").Value = "Lcms"
			.Fields("form_name").Value = "lcm"
			.Fields("form_label").Value = "Lcms"
			.Fields("chr_i").Value = 10
			.Update()
			'     h     4  yes   1  1 10  1  1 heavy products
			.AddNew()
			.Fields("letter").Value = "h"
			.Fields("value").Value = 4
			.Fields("sell").Value = True
			.Fields("lbs").Value = 1
			.Fields("pack_ie").Value = 1
			.Fields("pack_rg").Value = 1
			.Fields("pack_wh").Value = 10
			.Fields("pack_ur").Value = 1
			.Fields("pack_bk").Value = 1
			.Fields("name").Value = "heavy products"
			.Fields("intelligence_names").Value = "heavy products"
			.Fields("p_sname").Value = "hcm"
			.Fields("cond_name").Value = "hcm"
			.Fields("db_name").Value = "hcm"
			.Fields("sector_panel_label").Value = "Hcms"
			.Fields("distribution_panel_label").Value = "Hcms"
			.Fields("form_name").Value = "hcm"
			.Fields("form_label").Value = "Hcms"
			.Fields("chr_i").Value = 11
			.Update()
			'     r   150  yes   8  1 10  1  1 radioactive materials
			.AddNew()
			.Fields("letter").Value = "r"
			.Fields("value").Value = 150
			.Fields("sell").Value = True
			.Fields("lbs").Value = 8
			.Fields("pack_ie").Value = 1
			.Fields("pack_rg").Value = 1
			.Fields("pack_wh").Value = 10
			.Fields("pack_ur").Value = 1
			.Fields("pack_bk").Value = 1
			.Fields("name").Value = "radioactive materials"
			.Fields("intelligence_names").Value = "radioactive materials"
			.Fields("p_sname").Value = "rads"
			.Fields("cond_name").Value = "rad"
			.Fields("db_name").Value = "rad"
			.Fields("sector_panel_label").Value = "Rads"
			.Fields("distribution_panel_label").Value = "Rads"
			.Fields("form_name").Value = "rads"
			.Fields("form_label").Value = "Rads"
			.Fields("chr_i").Value = 13
			.Update()
			'     u     1  yes   2  1  2  1  1 uncompensated workers
			.AddNew()
			.Fields("letter").Value = "u"
			.Fields("value").Value = 1
			.Fields("sell").Value = True
			.Fields("lbs").Value = 2
			.Fields("pack_ie").Value = 1
			.Fields("pack_rg").Value = 1
			.Fields("pack_wh").Value = 2
			.Fields("pack_ur").Value = 1
			.Fields("pack_bk").Value = 1
			.Fields("name").Value = "uncompensated workers"
			.Fields("intelligence_names").Value = "uncompensated workers"
			.Fields("p_sname").Value = "uw"
			.Fields("cond_name").Value = "uw"
			.Fields("db_name").Value = "uw"
			.Fields("sector_panel_label").Value = "Uw:"
			.Fields("distribution_panel_label").Value = "Uws"
			.Fields("form_name").Value = "uw"
			.Fields("form_label").Value = "unc."
			.Fields("chr_i").Value = 12
			.Update()
		End With
	End Sub
	
	Public Sub UpdateDatabase()
		Dim eiItem As EmpItem
		DeleteAllRecords(rsItems)
		For	Each eiItem In mCol
			With rsItems
				.AddNew()
				.Fields("letter").Value = eiItem.Letter
				.Fields("value").Value = eiItem.Value
				.Fields("sell").Value = eiItem.Sellable
				.Fields("lbs").Value = eiItem.Weight
				.Fields("pack_ie").Value = eiItem.Packing(EmpItem.enumItemPacking.PACKING_INEFFICIENT)
				.Fields("pack_rg").Value = eiItem.Packing(EmpItem.enumItemPacking.PACKING_REGULAR)
				.Fields("pack_wh").Value = eiItem.Packing(EmpItem.enumItemPacking.PACKING_WAREHOUSE)
				.Fields("pack_ur").Value = eiItem.Packing(EmpItem.enumItemPacking.PACKING_URBAN)
				.Fields("pack_bk").Value = eiItem.Packing(EmpItem.enumItemPacking.PACKING_BANK)
				.Fields("name").Value = eiItem.ItemName
				.Fields("intelligence_names").Value = eiItem.IntelligenceNames
				.Fields("p_sname").Value = "eiItem.ProductName"
				.Fields("cond_name").Value = eiItem.ConditionalName
				.Fields("db_name").Value = eiItem.DatabaseName
				.Fields("sector_panel_label").Value = eiItem.SectorPanelLabel
				.Fields("distribution_panel_label").Value = eiItem.DistributionPanelLabel
				.Fields("form_name").Value = eiItem.FormName
				.Fields("form_label").Value = eiItem.FormLabel
				.Fields("chr_i").Value = eiItem.ChrIndex
				.Update()
			End With
		Next eiItem
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'destroys collection when this class is terminated
		'UPGRADE_NOTE: Object mCol may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mCol = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class