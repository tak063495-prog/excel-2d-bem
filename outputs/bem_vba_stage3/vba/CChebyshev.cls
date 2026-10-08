VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CChebyshev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public P As Long, D As Long, Q As Long
Private Z() As Double, BW() As Double, GridData() As Double
Public Sub Initialize(ByVal order As Long, ByVal dimension As Long)
    Dim i As Long, j As Long, index As Long, digit As Long, angle As Double
    P = order: D = dimension: Q = P ^ D
    ReDim Z(1 To P): ReDim BW(1 To P): ReDim GridData(1 To Q, 1 To D)
    For i = 1 To P
        angle = PI * (i - 0.5) / P: Z(i) = Cos(angle): BW(i) = (-1#) ^ (i - 1) * Sin(angle)
    Next i
    For i = 1 To Q
        index = i - 1
        For j = D To 1 Step -1: digit = index Mod P: GridData(i, j) = Z(digit + 1): index = index \ P: Next j
    Next i
End Sub
Public Function Grid() As Variant
    Grid = GridData
End Function
Public Sub Basis(ByRef x() As Double, ByRef center() As Double, ByVal halfWidth As Double, ByRef result() As Double)
    Dim weights(1 To 3, 1 To 8) As Double, i As Long, axis As Long, matchNode As Long, index As Long, digit As Long
    Dim delta As Double, total As Double, value As Double
    ReDim result(1 To Q)
    For axis = 1 To D
        matchNode = 0: total = 0
        For i = 1 To P
            delta = (x(axis) - center(axis)) / halfWidth - Z(i)
            If Abs(delta) < 0.0000000000001 Then matchNode = i: Exit For
            weights(axis, i) = BW(i) / delta: total = total + weights(axis, i)
        Next i
        If matchNode > 0 Then
            For i = 1 To P: weights(axis, i) = 0: Next i
            weights(axis, matchNode) = 1
        Else
            For i = 1 To P: weights(axis, i) = weights(axis, i) / total: Next i
        End If
    Next axis
    For i = 1 To Q
        index = i - 1: value = 1
        For axis = D To 1 Step -1: digit = index Mod P: value = value * weights(axis, digit + 1): index = index \ P: Next axis
        result(i) = value
    Next i
End Sub

