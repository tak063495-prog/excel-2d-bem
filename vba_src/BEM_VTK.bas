Attribute VB_Name = "BEM_VTK"
Option Explicit
Public Function ResultantJSON(ByVal sr As CSolveRegion) As String
    Dim rule As CIntegrationRule, geom As CElementData, prepared As CElementRule, weights() As Double, shape() As Double, t() As Double
    Dim force(1 To 3) As Double, e As Long, i As Long, j As Long, l As Long, d As Long, value As Double, parts As Collection
    d = sr.Ref.Dimension: t = sr.T: Set rule = RegularRule(d, IIf(d = 2, 32, 16))
    For e = 0 To sr.Ref.Elements.Count - 1
        Set geom = sr.Geometries(e + 1): Set prepared = New CElementRule: prepared.Initialize geom, rule: weights = prepared.Weight: shape = prepared.N
        For i = 1 To prepared.Count: For j = 1 To d
            value = 0
            For l = 1 To sr.Ref.FieldCount: value = value + shape(i, l) * t((e * sr.Ref.FieldCount + l - 1) * d + j): Next l
            force(j) = force(j) + weights(i) * value
        Next j: Next i
    Next e
    Set parts = New Collection: For j = 1 To d: parts.Add JNum(force(j)): Next j
    ResultantJSON = "[" & JoinCollection(parts, ",") & "]"
End Function
Public Sub SaveBoundaryVTK(ByVal sr As CSolveRegion, ByVal folder As String)
    Dim r As CBemRegion, lines As Collection, cols As Collection, ids As Variant, point As Variant, geomParams As Variant
    Dim shape(1 To 6) As Double, u() As Double, t() As Double, values() As Double, data(1 To 3) As Double
    Dim e As Long, i As Long, j As Long, l As Long, ng As Long, d As Long, cellType As Long, position As Long, axis As Long, kind As Long, permutation As Variant
    Dim fso As Object, a As Double, b As Double, vectorTitle As String
    Set r = sr.Ref: d = r.Dimension: ids = r.Elements(1): ng = UBound(ids) + 1: u = sr.U: t = sr.T: Set lines = New Collection
    lines.Add "# vtk DataFile Version 3.0": lines.Add "Unified elasticity boundary, discontinuous fields": lines.Add "ASCII": lines.Add "DATASET UNSTRUCTURED_GRID"
    lines.Add "POINTS " & r.Elements.Count * ng & " double"
    For Each ids In r.Elements
        For i = 0 To UBound(ids)
            point = r.Vertices(ids(i) + 1): lines.Add JNum(point(0)) & " " & JNum(point(1)) & " " & JNum(point(2))
        Next i
    Next ids
    permutation = Array(0, 1, 2, 3, 4, 5)
    If d = 2 Then
        cellType = IIf(r.ElementOrder = 2, 21, 3)
        If r.ElementOrder = 2 Then permutation = Array(0, 2, 1)
    Else
        cellType = IIf(r.ElementOrder = 2, 22, 5)
    End If
    lines.Add "CELLS " & r.Elements.Count & " " & (ng + 1) * r.Elements.Count
    For e = 0 To r.Elements.Count - 1
        Set cols = New Collection: cols.Add CStr(ng)
        For i = 0 To ng - 1: cols.Add CStr(e * ng + permutation(i)): Next i
        lines.Add JoinCollection(cols, " ")
    Next e
    lines.Add "CELL_TYPES " & r.Elements.Count
    For e = 0 To r.Elements.Count - 1: lines.Add CStr(cellType): Next e
    lines.Add "POINT_DATA " & ng * r.Elements.Count
    geomParams = Array(Array(0#, 0#), Array(1#, 0#), Array(0#, 1#), Array(0.5, 0#), Array(0.5, 0.5), Array(0#, 0.5))
    For kind = 0 To 1
        If kind = 0 Then values = u: vectorTitle = "displacement_m" Else values = t: vectorTitle = "traction_kPa"
        lines.Add "VECTORS " & vectorTitle & " double"
        For e = 0 To r.Elements.Count - 1
            For i = 0 To ng - 1
                If d = 2 Then a = -1 + 2 * i / (ng - 1): b = 0 Else a = geomParams(i)(0): b = geomParams(i)(1)
                FieldShape d, r.ElementOrder, a, b, shape
                For j = 1 To 3
                    data(j) = 0
                    If j <= d Then For l = 1 To r.FieldCount: data(j) = data(j) + shape(l) * values((e * r.FieldCount + l - 1) * d + j): Next l
                Next j
                lines.Add JNum(data(1)) & " " & JNum(data(2)) & " " & JNum(data(3))
            Next i
        Next e
    Next kind
    Set fso = CreateObject("Scripting.FileSystemObject"): WriteUTF8 fso.BuildPath(folder, r.RegionName & "_boundary.vtk"), JoinCollection(lines, vbLf) & vbLf
End Sub
