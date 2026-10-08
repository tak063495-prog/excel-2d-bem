VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CSolveRegion"
Option Explicit
Public Ref As CBemRegion
Public Geometries As Collection
Public Offset As Long
Public NF As Long
Public ND As Long
Public Mu As Double
Public NuEff As Double
Public Lambda As Double
Public MeshScale As Double
Public IsPlaneStrain As Boolean
Public X As Variant
Public Normal As Variant
Public G As Variant
Public H As Variant
Public U As Variant
Public T As Variant
Public U0 As Variant
Public T0 As Variant
Public UMap As Variant
Public TMap As Variant
Public UFactor As Variant
Public TFactor As Variant
Public Paired As Variant
Public PrescribedU As Variant
Public Backend As String
Public Fmm As CFmmOperator
Public DiagG As Variant, DiagH As Variant
Public SelectionJSON As String
Public EstimatedMB As Double

Public Sub AllocateFields()
    Dim d As Long: d = Ref.Dimension
    ReDim X(1 To NF, 1 To d) As Double: ReDim Normal(1 To NF, 1 To d) As Double
    ReDim U(1 To ND) As Double: ReDim T(1 To ND) As Double
    ReDim U0(1 To ND) As Double: ReDim T0(1 To ND) As Double
    ReDim UMap(1 To ND) As Long: ReDim TMap(1 To ND) As Long
    ReDim UFactor(1 To ND) As Double: ReDim TFactor(1 To ND) As Double
    ReDim Paired(1 To NF) As Boolean: ReDim PrescribedU(1 To ND) As Boolean
End Sub
Public Sub AllocateMatrices()
    ReDim G(1 To ND, 1 To ND) As Double: ReDim H(1 To ND, 1 To ND) As Double
End Sub
Public Sub SetField(ByVal field As Long, ByVal axis As Long, ByVal position As Double, ByVal direction As Double)
    X(field, axis) = position: Normal(field, axis) = direction
End Sub
Public Sub SetInterface(ByVal field As Long, ByVal firstU As Long, ByVal firstT As Long, ByVal factor As Double)
    Dim axis As Long, index As Long
    Paired(field) = True
    For axis = 1 To Ref.Dimension
        index = (field - 1) * Ref.Dimension + axis
        UMap(index) = firstU + axis: UFactor(index) = 1
        TMap(index) = firstT + axis: TFactor(index) = factor
    Next axis
End Sub
Public Sub SetExternal(ByVal index As Long, ByVal fixedU As Boolean, ByVal value As Double, ByVal column As Long, ByVal factor As Double)
    PrescribedU(index) = fixedU
    If fixedU Then
        U0(index) = value: TMap(index) = column: TFactor(index) = factor
    Else
        T0(index) = value: UMap(index) = column: UFactor(index) = 1
    End If
End Sub
