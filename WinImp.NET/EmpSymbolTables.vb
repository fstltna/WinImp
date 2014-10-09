Option Strict Off
Option Explicit On
Friend Class EmpSymbolTables
	Implements System.Collections.IEnumerable
	
	'Change Log:
	'110606 rjk: Created
	
	Private mCol As Collection
	
	Default Public ReadOnly Property Table(ByVal xdType As XdumpParse.enumXdumpType) As EmpSymbolTable
		Get
			Table = mCol.Item(CStr(xdType))
		End Get
	End Property
	
	Public ReadOnly Property Count() As Integer
		Get
			Count = mCol.Count()
		End Get
	End Property
	
	'UPGRADE_NOTE: NewEnum property was commented out. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B3FC1610-34F3-43F5-86B7-16C984F0E88E"'
	'Public ReadOnly Property NewEnum() As stdole.IUnknown
		'Get
			'NewEnum = mCol._NewEnum
		'End Get
	'End Property
	
	Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
		'UPGRADE_TODO: Uncomment and change the following line to return the collection enumerator. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="95F9AAD0-1319-4921-95F0-B9D3C4FF7F1C"'
		'GetEnumerator = mCol.GetEnumerator
	End Function
	
	Public Sub AddUpdate(ByRef xdType As XdumpParse.enumXdumpType, ByRef strName As String)
		Dim etTable As EmpSymbolTable
		
		On Error GoTo Add_New
		etTable = mCol.Item(CStr(xdType))
		etTable.Name = strName
		Exit Sub
Add_New: 
		etTable = New EmpSymbolTable
		etTable.TblType = xdType
		etTable.Name = strName
		mCol.Add(etTable, CStr(xdType))
	End Sub
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		'creates the collection when this class is created
		mCol = New Collection
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
End Class