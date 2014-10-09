Option Strict Off
Option Explicit On
Friend Class EmpSymbolEntry
	
	'Change Log:
	'100606 rjk: Created
	'190606 rjk: Convert value to long to support ship-chr-flags.
	
	Private mstrName As String
	Private mlValue As Integer
	
	
	Public Property Name() As String
		Get
			Name = mstrName
		End Get
		Set(ByVal Value As String)
			mstrName = Value
		End Set
	End Property
	
	
	Public Property Value() As Integer
		Get
			Value = mlValue
		End Get
		Set(ByVal Value As Integer)
			mlValue = Value
		End Set
	End Property
End Class