Option Strict Off
Option Explicit On
Friend Class EmpItem
	
	'Change Log:
	'270904 rjk: Added inefficient packing, xdump chr index.
	
	'local variable(s) to hold property value(s)
	Private mvarFormName As String 'local copy
	Private mvarLetter As String 'local copy
	Private mvarSectorPanelLabel As String 'local copy
	Private mvarConditionalName As String 'local copy
	Private mvarValue As Short 'local copy
	Private mvarWeight As Short 'local copy
	Public Enum enumItemPacking
		PACKING_INEFFICIENT = 0
		PACKING_REGULAR = 1
		PACKING_WAREHOUSE = 2
		PACKING_URBAN = 3
		PACKING_BANK = 4 'Must be the last one as it is used an array size
	End Enum
	Private mvarPacking(enumItemPacking.PACKING_BANK) As Short 'local copy
	Private mvarIntelligenceNames As String 'local copy
	Private mvarDatabaseName As String 'local copy
	Private mvarSellable As Boolean 'local copy
	Private mvarItemName As String 'local copy
	Private mvarProductName As String 'local copy
	Private mvarFormLabel As String 'local copy
	Private mvarDistributionPanelLabel As String 'local copy
	Private mvarChrIndex As Short 'local copy
	
	
	Public Property DistributionPanelLabel() As String
		Get
			DistributionPanelLabel = mvarDistributionPanelLabel
		End Get
		Set(ByVal Value As String)
			mvarDistributionPanelLabel = Value
		End Set
	End Property
	
	
	Public Property FormLabel() As String
		Get
			FormLabel = mvarFormLabel
		End Get
		Set(ByVal Value As String)
			mvarFormLabel = Value
		End Set
	End Property
	
	
	Public Property ProductName() As String
		Get
			ProductName = mvarProductName
		End Get
		Set(ByVal Value As String)
			mvarProductName = Value
		End Set
	End Property
	
	
	Public Property ChrIndex() As Short
		Get
			ChrIndex = mvarChrIndex
		End Get
		Set(ByVal Value As Short)
			mvarChrIndex = Value
		End Set
	End Property
	
	
	Public Property ItemName() As String
		Get
			ItemName = mvarItemName
		End Get
		Set(ByVal Value As String)
			mvarItemName = Value
		End Set
	End Property
	
	
	Public Property Sellable() As Boolean
		Get
			Sellable = mvarSellable
		End Get
		Set(ByVal Value As Boolean)
			mvarSellable = Value
		End Set
	End Property
	
	
	Public Property DatabaseValue(ByVal rsTable As DAO.Recordset) As Integer
		Get
			DatabaseValue = rsTable.Fields(mvarDatabaseName).Value
		End Get
		Set(ByVal Value As Integer)
			rsTable.Fields(mvarDatabaseName).Value = Value
		End Set
	End Property
	
	
	Public Property DatabaseName() As String
		Get
			DatabaseName = mvarDatabaseName
		End Get
		Set(ByVal Value As String)
			mvarDatabaseName = Value
		End Set
	End Property
	
	
	Public Property IntelligenceNames() As String
		Get
			IntelligenceNames = mvarIntelligenceNames
		End Get
		Set(ByVal Value As String)
			mvarIntelligenceNames = Value
		End Set
	End Property
	
	
	Public Property Packing(ByVal packingType As enumItemPacking) As Short
		Get
			Packing = mvarPacking(packingType)
		End Get
		Set(ByVal Value As Short)
			mvarPacking(packingType) = Value
		End Set
	End Property
	
	
	Public Property Weight() As Short
		Get
			Weight = mvarWeight
		End Get
		Set(ByVal Value As Short)
			mvarWeight = Value
		End Set
	End Property
	
	
	Public Property Value() As Short
		Get
			Value = mvarValue
		End Get
		Set(ByVal Value As Short)
			mvarValue = Value
		End Set
	End Property
	
	
	
	Public Property ConditionalName() As String
		Get
			ConditionalName = mvarConditionalName
		End Get
		Set(ByVal Value As String)
			mvarConditionalName = Value
		End Set
	End Property
	
	
	Public Property SectorPanelLabel() As String
		Get
			SectorPanelLabel = mvarSectorPanelLabel
		End Get
		Set(ByVal Value As String)
			mvarSectorPanelLabel = Value
		End Set
	End Property
	
	
	Public Property Letter() As String
		Get
			Letter = mvarLetter
		End Get
		Set(ByVal Value As String)
			mvarLetter = Value
		End Set
	End Property
	
	
	Public Property FormName() As String
		Get
			FormName = mvarFormName
		End Get
		Set(ByVal Value As String)
			mvarFormName = Value
		End Set
	End Property
	
	Public Sub SaveItem()
		rsItems.Seek("=", mvarLetter)
		If Not rsItems.NoMatch Then
			With rsItems
				.Edit()
				.Fields("letter").Value = mvarLetter
				.Fields("value").Value = mvarValue
				.Fields("sell").Value = mvarSellable
				.Fields("lbs").Value = mvarWeight
				.Fields("pack_ie").Value = mvarPacking(enumItemPacking.PACKING_INEFFICIENT)
				.Fields("pack_rg").Value = mvarPacking(enumItemPacking.PACKING_REGULAR)
				.Fields("pack_wh").Value = mvarPacking(enumItemPacking.PACKING_WAREHOUSE)
				.Fields("pack_ur").Value = mvarPacking(enumItemPacking.PACKING_URBAN)
				.Fields("pack_bk").Value = mvarPacking(enumItemPacking.PACKING_BANK)
				.Fields("name").Value = mvarItemName
				.Fields("p_sname").Value = mvarProductName
				.Fields("cond_name").Value = mvarConditionalName
				.Fields("db_name").Value = mvarDatabaseName
				.Fields("sector_panel_label").Value = mvarSectorPanelLabel
				.Fields("distribution_panel_label").Value = mvarDistributionPanelLabel
				.Fields("form_name").Value = mvarFormName
				.Fields("intelligence_names").Value = mvarIntelligenceNames
				.Fields("form_label").Value = mvarFormLabel
				.Fields("chr_i").Value = mvarChrIndex
				.Update()
			End With
		End If
	End Sub
End Class