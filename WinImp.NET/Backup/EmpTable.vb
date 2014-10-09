Option Strict Off
Option Explicit On
Friend Class EmpTable
	Implements System.Collections.IEnumerable
	
	'Change Log:
	'050206 rjk: Created
	'220506 rjk: Changed NSC_STRINGY constant to variable extracted from meta-type table.
	'150606 Convert iNscSTRINGY to Symbol table.
	
	Private mCol As Collection
	Private mstrName As String
	Private mxdType As XdumpParse.enumXdumpType
	
	Public ReadOnly Property Column(ByVal vntIndexKey As Object) As EmpTableColumn
		Get
			'used when referencing an element in the collection
			'vntIndexKey contains either the Index or Key to the collection,
			'this is why it is declared as a Variant
			'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
			Column = mCol.Item(vntIndexKey)
		End Get
	End Property
	
	Public ReadOnly Property Count() As Integer
		Get
			'used when retrieving the number of elements in the
			'collection. Syntax: Debug.Print x.Count
			Count = mCol.Count()
		End Get
	End Property
	
	Public ReadOnly Property ParameterCount() As Integer
		Get
			Dim Item As EmpTableColumn
			
			ParameterCount = 0
			For	Each Item In mCol
				If estsTable(XdumpParse.enumXdumpType.XD_META_TYPE).Count >= Item.ColType Then
					If estsTable(XdumpParse.enumXdumpType.XD_META_TYPE)((Item.ColType)).Name <> "c" And Item.Length > 1 Then
						ParameterCount = ParameterCount + Item.Length
					Else
						ParameterCount = ParameterCount + 1
					End If
				Else
					ParameterCount = ParameterCount + 1
				End If
			Next Item
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
	
	
	Public Property Name() As String
		Get
			Name = mstrName
		End Get
		Set(ByVal Value As String)
			mstrName = Value
		End Set
	End Property
	
	
	Public Property TblType() As XdumpParse.enumXdumpType
		Get
			TblType = mxdType
		End Get
		Set(ByVal Value As XdumpParse.enumXdumpType)
			mxdType = Value
		End Set
	End Property
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		'creates the collection when this class is created
		mCol = New Collection
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	Public Sub AddUpdateColumn(ByRef strLine As String)
		Dim etTableColumn As EmpTableColumn
		Dim strParams() As String
		Dim strName As String
		
		strParams = Split(strLine, " ")
		strName = RemoveEscapeSequencesAndQuotes(strParams(0))
		On Error GoTo Add_New
		etTableColumn = mCol.Item(strName)
		etTableColumn.ColType = CShort(strParams(1))
		etTableColumn.Flags = CShort(strParams(2))
		etTableColumn.Length = CShort(strParams(3))
		etTableColumn.Table = CShort(strParams(4))
		Exit Sub
Add_New: 
		etTableColumn = New EmpTableColumn
		etTableColumn.Name = strName
		etTableColumn.ColType = CShort(strParams(1))
		etTableColumn.Flags = CShort(strParams(2))
		etTableColumn.Length = CShort(strParams(3))
		etTableColumn.Table = CShort(strParams(4))
		mCol.Add(etTableColumn, strName)
	End Sub
	
	Public Function GetZeroIndex(ByRef strName As String) As Short
		Dim Item As EmpTableColumn
		Dim i As Short
		
		i = 0
		For	Each Item In mCol
			If Item.Name = strName Then
				GetZeroIndex = i
				Exit Function
			End If
			i = i + 1
		Next Item
		GetZeroIndex = -1
	End Function
	
	Public Function GetOneIndex(ByRef strName As String) As Short
		GetOneIndex = GetZeroIndex(strName)
		If GetOneIndex <> -1 Then
			GetOneIndex = GetOneIndex + 1
		End If
	End Function
End Class