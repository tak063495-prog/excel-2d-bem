VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CIntegrationRule"
Option Explicit
Public Count As Long
Public Capacity As Long
Public A As Variant
Public B As Variant
Public W As Variant
Public Sub Allocate(ByVal size As Long)
    Capacity = size: Count = 0
    ReDim A(1 To size) As Double: ReDim B(1 To size) As Double: ReDim W(1 To size) As Double
End Sub
Public Sub Add(ByVal s As Double, ByVal t As Double, ByVal weight As Double)
    If weight <= 0 Then Exit Sub
    Count = Count + 1
    If Count > Capacity Then Fail "積分点配列の容量が不足しています。"
    A(Count) = s: B(Count) = t: W(Count) = weight
End Sub
Public Sub StorePoint(ByVal index As Long, ByVal s As Double, ByVal t As Double, ByVal weight As Double)
    A(index) = s: B(index) = t: W(index) = weight
End Sub
