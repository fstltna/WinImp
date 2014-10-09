Attribute VB_Name = "NukeDamage"
Option Explicit

'Change Log:
'200604 rjk: Created

Public NukeDamageType As String
Public NukeTargetSector As String
Public bNukeAirBurst As Boolean

Public Sub ShowNukeDamage(picMap As PictureBox)
Dim MaxRadius As Integer
Dim StartX As Integer
Dim StartY As Integer
Dim iDistance As Integer
Dim iDirection As Integer
Dim iSpoke As Integer
Dim sx As Integer
Dim sy As Integer
Dim X As Integer
Dim Y As Integer
Dim PosX As Single
Dim PosY As Single
Dim xa, ya, xas, yas As Variant
Dim strID As String
Dim bFission As Boolean
Dim iPos As Integer
Dim hldIndex As String
Dim iBlast As Integer
Dim iDamage As Integer
Dim HldColor As Long
Dim HldFontSize As Variant
Dim strDamage As String

iPos = InStr(NukeDamageType, " - ")

If iPos = 0 Then
    MsgBox "Invalid Nuke Type"
    Exit Sub
End If

strID = Left$(NukeDamageType, iPos - 1)
If InStr(NukeDamageType, "fission") > 0 Then
    bFission = True
Else
    bFission = False
End If

hldIndex = rsBuild.Index
rsBuild.Index = "PrimaryKey"
rsBuild.Seek "=", "n", strID

If rsBuild.NoMatch Then
    MsgBox "Nuke Type information not found"
    Exit Sub
End If
       
iBlast = rsBuild.Fields("stat1")
iDamage = rsBuild.Fields("stat2")
rsBuild.Index = hldIndex

If bNukeAirBurst Then
    MaxRadius = iBlast
Else
    MaxRadius = Fix(iBlast * 2 / 3)
End If

xa = Array(-2, -1, 1, 2, 1, -1)
ya = Array(0, -1, -1, 0, 1, 1)
xas = Array(1, 2, 1, -1, -2, -1)
yas = Array(-1, 0, 1, 1, 0, -1)

If Not ParseSectors(StartX, StartY, NukeTargetSector) Then
    Exit Sub
End If

HldColor = picMap.ForeColor
picMap.ForeColor = vbWhite
HldFontSize = picMap.Font.Size
picMap.Font.Size = picMap.Font.Size * 0.7

Coord StartX, StartY, PosX, PosY
strDamage = CStr(NukeSectorDamage(iBlast, iDamage, 0, bNukeAirBurst))
sx = PosX - (picMap.TextHeight(strDamage) / 2)
sy = PosY - (picMap.TextWidth(strDamage) / 2)
If sx >= 0 And sy >= 0 Then
    picMap.CurrentX = sx
    picMap.CurrentY = sy
    picMap.Print strDamage
End If

'Now put out the nav markers
For iDistance = 1 To MaxRadius
    For iDirection = 0 To 5
        For iSpoke = 0 To MaxRadius - 1
            X = (iDistance * xa(iDirection)) + (iSpoke * xas(iDirection))
            Y = (iDistance * ya(iDirection)) + (iSpoke * yas(iDirection))
            Coord StartX + X, StartY + Y, PosX, PosY
            strDamage = CStr(NukeSectorDamage(iBlast, iDamage, SectorDistance(0, 0, X, Y), bNukeAirBurst))
            sx = PosX - (picMap.TextHeight(strDamage) / 2)
            sy = PosY - (picMap.TextWidth(strDamage) / 2)
            If sx >= 0 And sy >= 0 Then
                picMap.CurrentX = sx
                picMap.CurrentY = sy
                picMap.Print strDamage
            End If
        Next iSpoke
    Next iDirection
Next iDistance

picMap.ForeColor = HldColor
picMap.Font.Size = HldFontSize
End Sub

'int
'nukedamage(struct nchrstr *ncp, int range, int airburst)
'{
'    int dam;
'    int rad;
'
'    rad = ncp->n_blast;
'    if (airburst)
'    rad = (int)(rad * 1.5);
'    if (rad < range)
'    return 0;
'    if (airburst) {
'    /* larger area, less center damage */
'    dam = (int)((ncp->n_dam * 0.75) - (range * 20));
'    } else {
'    /* smaller area, more center damage */
'    dam = (int)(ncp->n_dam / (range + 1.0));
'    }
'    if (dam < 5)
'    dam = 0;
'    return dam;
'}

Private Function NukeSectorDamage(iBlast As Integer, iDamage As Integer, iDistance As Integer, bAirBurst As Boolean) As Integer
Dim iRad As Integer

iRad = iBlast
If bAirBurst Then
    iRad = Fix(iRad * 1.5)
End If
If iRad < iDistance Then
    NukeSectorDamage = 0
    Exit Function
End If
If bAirBurst Then
    NukeSectorDamage = Fix((iDamage * 0.75) - (iDistance * 20))
Else
    NukeSectorDamage = Fix(iDamage / (iDistance + 1#))
End If
If NukeSectorDamage < 5 Then
    NukeSectorDamage = 0
End If
End Function

