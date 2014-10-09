Option Strict Off
Option Explicit On
Friend Class EmpTableColumn
	
	'Change Log:
	'050206 rjk: Created
	
	Private mstrName As String
	Private miType As Short
	Private miFlags As Short
	Private miLength As Short
	Private miTable As Short
	
	
	Public Property Name() As String
		Get
			Name = mstrName
		End Get
		Set(ByVal Value As String)
			mstrName = Value
		End Set
	End Property
	
	
	Public Property ColType() As Short
		Get
			ColType = miType
		End Get
		Set(ByVal Value As Short)
			miType = Value
		End Set
	End Property
	
	
	Public Property Flags() As Short
		Get
			Flags = miType
		End Get
		Set(ByVal Value As Short)
			miFlags = Value
		End Set
	End Property
	
	
	Public Property Length() As Short
		Get
			Length = miLength
		End Get
		Set(ByVal Value As Short)
			miLength = Value
		End Set
	End Property
	
	
	Public Property Table() As Short
		Get
			Table = miType
		End Get
		Set(ByVal Value As Short)
			miTable = Value
		End Set
	End Property
End Class