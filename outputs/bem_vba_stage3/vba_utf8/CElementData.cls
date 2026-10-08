VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CElementData"
Option Explicit
Public D As Long
Public ElementOrder As Long
Public K As Long
Public NG As Long
Public Coords As Variant
Public Length As Double
Public ModelScale As Double
Public Planar As Boolean
Public Regular As CElementRule
Public Sub Initialize(ByVal r As CBemRegion, ByVal e As Long)
    Dim ids As Variant, p As Variant, i As Long, j As Long, edge As Long, length2 As Double, difference As Double
    ReDim Coords(1 To 6, 1 To 3) As Double
    D = r.Dimension: ElementOrder = r.ElementOrder: K = r.FieldCount: ModelScale = r.ModelScale
    ids = r.Elements(e + 1): NG = UBound(ids) + 1
    For i = 1 To NG
        p = r.Vertices(ids(i - 1) + 1)
        For j = 1 To D: Coords(i, j) = p(j - 1): Next j
    Next i
    Length = MaxEdge(r, e): Planar = True
    If ElementOrder = 2 Then
        If D = 2 Then
            For j = 1 To D: difference = MaxDouble(difference, Abs(Coords(2, j) - (Coords(1, j) + Coords(3, j)) / 2)): Next j
        Else
            For i = 1 To 3
                edge = i Mod 3 + 1
                For j = 1 To D: difference = MaxDouble(difference, Abs(Coords(i + 3, j) - (Coords(i, j) + Coords(edge, j)) / 2)): Next j
            Next i
        End If
        Planar = difference < ModelScale * 1E-12
    End If
End Sub
