VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmpUnitView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Change Log:
'112803 rjk: Created
'120103 rjk: Added a check for unknown relations (offline) to the enemyinfo generation
'120303 rjk: Moved friendlyUnit to frmOptions
'010404 rjk: Switched to screeninfo to a type
'170404 rjk: Fixed the displaying of the Units when some units are disabled.
'140704 rjk: Added Expired Units.
'290606 rjk: Added Our Nuke Units.

Enum enumViewingUnits
    VU_OUR_SHIPS = 0
    VU_OUR_LAND_UNITS = 1
    VU_OUR_PLANES = 2
    VU_OUR_NUKES = 3
    VU_ENEMY_SHIPS = 4
    VU_ENEMY_LAND_UNITS = 5
    VU_ENEMY_PLANES = 6
    VU_ALLIED_SHIPS = 7
    VU_ALLIED_LAND_UNITS = 8
    VU_ALLIED_PLANES = 9
    VU_NEUTRAL_SHIPS = 10
    VU_NEUTRAL_LAND_UNITS = 11
    VU_NEUTRAL_PLANES = 12
    VU_EXPIRED_SHIPS = 13
    VU_EXPIRED_LAND_UNITS = 14
    VU_EXPIRED_PLANES = 15
    'if more units are added, need to change lastViewingUnit and array dimensions
    '0100404 rjk: need to set the size in the iUnitCount in tScreenInfo to match the num enumViewingUnits
End Enum

Private Const VU_NUM_UNITS = VU_EXPIRED_PLANES + 1
Private bViewUnit(VU_NUM_UNITS - 1) As Boolean
Private vuPriorityList(VU_NUM_UNITS - 1) As enumViewingUnits
Private vuPriorityListCount As Integer
Private picArray(VU_NUM_UNITS - 1) As Picture
Private m_iExpiry As Integer

Public Property Get UnitPicture(iUnitCounts() As Integer, bFirstUnit As Boolean) As Picture
Static passCount As Integer
Static priorityListIndex As Integer
Dim Index As Integer
Dim modValue As Integer
Dim Value As Integer

If bFirstUnit Then
    passCount = 1
    priorityListIndex = -1
End If

For Index = 0 To vuPriorityListCount - 1
    priorityListIndex = priorityListIndex + 1
    If priorityListIndex >= priorityListCount Then
        priorityListIndex = 0
        passCount = passCount + 1
    End If
    If iUnitCounts(vuPriorityList(priorityListIndex)) >= passCount Then
        Set UnitPicture = picArray(vuPriorityList(priorityListIndex))
        Exit Property
    End If
Next Index
Set UnitPicture = Nothing
End Property

Public Sub Save()
Dim Index As enumViewingUnits
Dim strTemp As String

For Index = 0 To vuPriorityListCount - 1
    strTemp = strTemp + CStr(vuPriorityList(Index)) + ","
Next Index
If Len(strTemp) > 0 Then
    strTemp = Left$(strTemp, Len(strTemp) - 1)
End If
SaveSetting APPNAME, "UnitView", "PriorityList", strTemp
SaveSetting APPNAME, "UnitView", "ExpiryDate", m_iExpiry
End Sub

Public Sub Load()
Dim Index As enumViewingUnits
Dim strPriorityItems() As String
Dim strTemp As String

LoadUnitPictures

vuPriorityListCount = 0

strTemp = GetSetting(APPNAME, "UnitView", "PriorityList", vbNullString)
If Len(strTemp) = 0 Then
    For Index = 0 To lastViewingUnit
        bViewUnit(Index) = True
        vuPriorityList(Index) = Index
        vuPriorityListCount = vuPriorityListCount + 1
    Next Index
Else
    For Index = 0 To lastViewingUnit
        bViewUnit(Index) = False
    Next Index
    strPriorityItems = Split(strTemp, ",")
    For Index = 0 To UBound(strPriorityItems)
        bViewUnit(CInt(strPriorityItems(Index))) = True
        vuPriorityList(Index) = CInt(strPriorityItems(Index))
        vuPriorityListCount = vuPriorityListCount + 1
    Next Index
End If
m_iExpiry = GetSetting(APPNAME, "UnitView", "ExpiryDate", 0)
End Sub

Public Property Get bViewUnits(Index As enumViewingUnits) As Boolean
    bViewUnits = bViewUnit(Index)
End Property

Public Property Let bViewUnits(Index As enumViewingUnits, Value As Boolean)
    bViewUnit(Index) = Value
End Property

Public Property Get lastViewingUnit() As enumViewingUnits
    lastViewingUnit = VU_NUM_UNITS - 1
End Property

Public Property Get iExpiry() As Integer
    iExpiry = m_iExpiry
End Property

Public Property Let iExpiry(Value As Integer)
    m_iExpiry = Value
End Property

Private Sub LoadUnitPictures()
On Error GoTo Empire_Error
Set picArray(VU_OUR_LAND_UNITS) = LoadPicture(App.path + "\land.gif")
Set picArray(VU_OUR_SHIPS) = LoadPicture(App.path + "\ship.gif")
Set picArray(VU_OUR_PLANES) = LoadPicture(App.path + "\plane.gif")
Set picArray(VU_OUR_NUKES) = LoadPicture(App.path + "\nuke.gif")
Set picArray(VU_ENEMY_LAND_UNITS) = LoadPicture(App.path + "\eland.gif")
Set picArray(VU_NEUTRAL_LAND_UNITS) = LoadPicture(App.path + "\nland.gif")
Set picArray(VU_ALLIED_LAND_UNITS) = LoadPicture(App.path + "\fland.gif")
Set picArray(VU_EXPIRED_LAND_UNITS) = LoadPicture(App.path + "\xland.gif")
Set picArray(VU_ENEMY_SHIPS) = LoadPicture(App.path + "\eship.gif")
Set picArray(VU_NEUTRAL_SHIPS) = LoadPicture(App.path + "\nship.gif")
Set picArray(VU_ALLIED_SHIPS) = LoadPicture(App.path + "\fship.gif")
Set picArray(VU_EXPIRED_SHIPS) = LoadPicture(App.path + "\xship.gif")
Set picArray(VU_ENEMY_PLANES) = LoadPicture(App.path + "\eplane.gif")
Set picArray(VU_NEUTRAL_PLANES) = LoadPicture(App.path + "\nplane.gif")
Set picArray(VU_ALLIED_PLANES) = LoadPicture(App.path + "\fplane.gif")
Set picArray(VU_EXPIRED_PLANES) = LoadPicture(App.path + "\xplane.gif")
Exit Sub

Empire_Error:
EmpireError "LoadUnitPictures", vbNullString, vbNullString
End Sub

Public Property Let priorityList(Index As enumViewingUnits, Value As enumViewingUnits)
vuPriorityList(Index) = Value
End Property

Public Property Get priorityList(Index As enumViewingUnits) As enumViewingUnits
priorityList = vuPriorityList(Index)
End Property

Public Property Let priorityListCount(Value As Integer)
vuPriorityListCount = Value
End Property

Public Property Get priorityListCount() As Integer
priorityListCount = vuPriorityListCount
End Property

Public Sub ClearAllUnits()
Dim Index As enumViewingUnits

For Index = 0 To lastViewingUnit
    bViewUnit(Index) = False
Next Index
vuPriorityListCount = 0
End Sub

Public Function VisibleUnits(iUnitCounts() As Integer) As Boolean
Dim Index As Integer
Dim modValue As Integer
Dim Value As Integer

For Index = 0 To vuPriorityListCount - 1
    If iUnitCounts(vuPriorityList(Index)) > 0 Then
        VisibleUnits = True
        Exit Function
    End If
Next Index
VisibleUnits = False
End Function

