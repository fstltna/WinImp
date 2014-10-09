VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmpSymbolTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'110606 rjk: Created

Private mCol As Collection

Public Property Get Table(xdType As enumXdumpType) As EmpSymbolTable
Attribute Table.VB_UserMemId = 0
Set Table = mCol(CStr(xdType))
End Property

Public Property Get Count() As Long
Count = mCol.Count
End Property

Public Property Get NewEnum() As IUnknown
Set NewEnum = mCol.[_NewEnum]
End Property

Public Sub AddUpdate(xdType As enumXdumpType, strName As String)
Dim etTable As EmpSymbolTable

On Error GoTo Add_New
Set etTable = mCol.Item(CStr(xdType))
etTable.Name = strName
Exit Sub
Add_New:
Set etTable = New EmpSymbolTable
etTable.TblType = xdType
etTable.Name = strName
mCol.Add etTable, CStr(xdType)
End Sub

Private Sub Class_Initialize()
'creates the collection when this class is created
Set mCol = New Collection

End Sub

