VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmpSymbolEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'100606 rjk: Created
'190606 rjk: Convert value to long to support ship-chr-flags.

Private mstrName As String
Private mlValue As Long

Public Property Let Name(ByVal vData As String)
mstrName = vData
End Property

Public Property Get Name() As String
Name = mstrName
End Property

Public Property Let Value(ByVal vData As Long)
mlValue = vData
End Property

Public Property Get Value() As Long
Value = mlValue
End Property

