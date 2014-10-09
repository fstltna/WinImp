VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmpTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'050206 rjk: Created
'120606 rjk: Switched table key to enumXdumpType from a string

Private mCol As Collection

Public Property Get Table(xdType As enumXdumpType) As EmpTable
Attribute Table.VB_UserMemId = 0
  'used when referencing an element in the collection
  'vntIndexKey contains either the Index or Key to the collection,
  'this is why it is declared as a Variant
  'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
Set Table = mCol(CStr(xdType))
End Property

Public Property Get Count() As Long
'used when retrieving the number of elements in the
'collection. Syntax: Debug.Print x.Count
Count = mCol.Count
End Property

Public Property Get NewEnum() As IUnknown
'this property allows you to enumerate
'this collection with the For...Each syntax
Set NewEnum = mCol.[_NewEnum]
End Property

Public Sub AddUpdate(xdType As enumXdumpType, strName As String)
Dim etTable As EmpTable

On Error GoTo Add_New
Set etTable = mCol(CStr(xdType))
etTable.Name = strName
Exit Sub
Add_New:
Set etTable = New EmpTable
etTable.TblType = xdType
etTable.Name = strName
mCol.Add etTable, CStr(xdType)
End Sub

Private Sub Class_Initialize()
'creates the collection when this class is created
Set mCol = New Collection

End Sub

