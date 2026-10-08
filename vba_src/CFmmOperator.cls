VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CFmmOperator"
Option Explicit
Private Nodes As Collection, Leaves As Collection, FarPairs As Collection, NearPairs As Collection, Groups As Collection
Private NearGroups As Object, GroupLookup As Object
Private Cheb As CChebyshev, Mat As CSolveRegion
Private D As Long, NF As Long, K As Long, Q As Long, NC As Long, Root As Long
Private XYZ() As Double, Pad As Double, RootHalf As Double
Private RowPtr() As Long, SourceIds() As Long, NearG() As Double, NearH() As Double
Private Correction() As Double, DG() As Double, DH() As Double
Private CacheLimit As Double, CacheBytes As Double, ClockValue As Long
Private Hits As Long, Misses As Long, NearEntries As Double, Prepared As Boolean
Public Property Get NodeCount() As Long
    NodeCount = Nodes.Count
End Property
Public Property Get LeafCount() As Long
    LeafCount = Leaves.Count
End Property
Public Property Get FarCount() As Long
    FarCount = FarPairs.Count
End Property
Public Property Get NearCount() As Long
    NearCount = NearPairs.Count
End Property
Public Property Get TranslationCount() As Long
    TranslationCount = Groups.Count
End Property
Public Property Get EntryCount() As Double
    EntryCount = NearEntries
End Property
Public Function PlanNodes() As Variant
    Dim out() As Variant, node As CFmmNode, ids() As Long, c() As Double, i As Long, j As Long
    ReDim out(1 To Nodes.Count, 1 To 8)
    For Each node In Nodes
        c = node.Center: ids = node.Ids: i = node.Number
        out(i, 1) = i - 1: out(i, 2) = node.Level: out(i, 3) = UBound(ids): out(i, 4) = node.Half0: out(i, 5) = node.HalfWidth
        For j = 1 To D: out(i, 5 + j) = c(j): Next j
    Next node
    PlanNodes = out
End Function
Public Function PlanPairs(ByVal far As Boolean) As Variant
    Dim pairs As Collection, out() As Variant, pair As Variant, i As Long
    If far Then Set pairs = FarPairs Else Set pairs = NearPairs
    If pairs.Count = 0 Then PlanPairs = Empty: Exit Function
    ReDim out(1 To pairs.Count, 1 To 2)
    For Each pair In pairs: i = i + 1: out(i, 1) = pair(0) - 1: out(i, 2) = pair(1) - 1: Next pair
    PlanPairs = out
End Function
Private Function BuildNode(ByRef ids() As Long, ByRef center() As Double, ByVal half0 As Double, ByVal level As Long, ByVal parent As Long) As Long
    Dim node As CFmmNode, count(0 To 7) As Long, codes() As Long, selected() As Long, childCenter(1 To 3) As Double
    Dim i As Long, axis As Long, octant As Long, number As Long, cursor As Long, child As Long
    Set node = New CFmmNode: node.Number = Nodes.Count + 1: node.Level = level: node.Parent = parent
    node.Half0 = half0: node.HalfWidth = half0 + Pad: node.Center = center: node.Ids = ids: Nodes.Add node
    If UBound(ids) > FmmLeaf And level < 20 Then
        ReDim codes(1 To UBound(ids))
        For i = 1 To UBound(ids)
            For axis = 1 To D
                If XYZ(ids(i), axis) > center(axis) Then codes(i) = codes(i) + 2 ^ (axis - 1)
            Next axis
            count(codes(i)) = count(codes(i)) + 1
        Next i
        For octant = 0 To 2 ^ D - 1
            If count(octant) > 0 Then
                ReDim selected(1 To count(octant)): cursor = 0
                For i = 1 To UBound(ids): If codes(i) = octant Then cursor = cursor + 1: selected(cursor) = ids(i)
                Next i
                For axis = 1 To D: childCenter(axis) = center(axis) + IIf((octant And 2 ^ (axis - 1)) <> 0, 1#, -1#) * half0 / 2: Next axis
                child = BuildNode(selected, childCenter, half0 / 2, level + 1, node.Number): node.Children.Add child
            End If
        Next octant
    Else
        Leaves.Add node.Number
    End If
    BuildNode = node.Number
End Function
Private Function TranslationKey(ByVal a As CFmmNode, ByVal b As CFmmNode) As String
    Dim ac() As Double, bc() As Double, axis As Long, key As String
    ac = a.Center: bc = b.Center: key = a.Level & ":" & b.Level
    For axis = 1 To D: key = key & ":" & Format$((bc(axis) - ac(axis)) / RootHalf, "0.0000000000000"): Next axis
    TranslationKey = key
End Function
Private Sub Walk(ByVal target As Long, ByVal source As Long)
    Dim a As CFmmNode, b As CFmmNode, ac() As Double, bc() As Double, distance As Double, axis As Long, child As Variant
    Dim key As String, group As CFmmTranslation, list As Collection, ia() As Long, ib() As Long
    Set a = Nodes(target): Set b = Nodes(source): ac = a.Center: bc = b.Center
    For axis = 1 To D: distance = distance + (ac(axis) - bc(axis)) ^ 2: Next axis
    distance = Sqr(distance)
    If distance > 0 And Sqr(D) * (a.HalfWidth + b.HalfWidth) < FmmTheta * distance Then
        FarPairs.Add Array(target, source): key = TranslationKey(a, b)
        If GroupLookup.Exists(key) Then
            Set group = GroupLookup(key)
        Else
            Set group = New CFmmTranslation: group.Key = key: group.TargetNode = target: group.SourceNode = source
            GroupLookup.Add key, group: Groups.Add group
        End If
        group.Targets.Add target: group.Sources.Add source
    ElseIf a.Children.Count = 0 And b.Children.Count = 0 Then
        NearPairs.Add Array(target, source): ia = a.Ids: ib = b.Ids: NearEntries = NearEntries + CDbl(UBound(ia)) * UBound(ib)
        key = CStr(target)
        If NearGroups.Exists(key) Then Set list = NearGroups(key) Else Set list = New Collection: NearGroups.Add key, list
        list.Add source
    ElseIf a.Children.Count > 0 And (b.Children.Count = 0 Or a.Half0 >= b.Half0) Then
        For Each child In a.Children: Walk CLng(child), source: Next child
    Else
        For Each child In b.Children: Walk target, CLng(child): Next child
    End If
End Sub
Public Sub InitializePlan(ByVal sr As CSolveRegion)
    Dim geom As CElementData, coords() As Double, control() As Double, ids() As Long, lower(1 To 3) As Double, upper(1 To 3) As Double, center(1 To 3) As Double
    Dim e As Long, i As Long, j As Long, l As Long, axis As Long, nextVertex As Long, ng As Long, localLower As Double, localUpper As Double, half0 As Double
    D = sr.Ref.Dimension: NF = sr.NF: K = sr.Ref.FieldCount: NC = D + D * D: XYZ = sr.X
    Set Mat = New CSolveRegion: Set Mat.Ref = sr.Ref: Mat.Mu = sr.Mu: Mat.NuEff = sr.NuEff: Mat.MeshScale = sr.MeshScale
    Set Cheb = New CChebyshev: Cheb.Initialize FmmOrder, D: Q = Cheb.Q
    Set Nodes = New Collection: Set Leaves = New Collection: Set FarPairs = New Collection: Set NearPairs = New Collection: Set Groups = New Collection
    Set NearGroups = CreateObject("Scripting.Dictionary"): Set GroupLookup = CreateObject("Scripting.Dictionary")
    For axis = 1 To D: lower(axis) = 1E+250: upper(axis) = -1E+250: Next axis
    For e = 1 To sr.Geometries.Count
        Set geom = sr.Geometries(e): coords = geom.Coords: ng = geom.NG: ReDim control(1 To ng, 1 To D)
        For i = 1 To ng: For axis = 1 To D: control(i, axis) = coords(i, axis): Next axis: Next i
        If sr.Ref.ElementOrder = 2 Then
            If D = 2 Then
                For axis = 1 To D: control(2, axis) = 2 * coords(2, axis) - (coords(1, axis) + coords(3, axis)) / 2: Next axis
            Else
                For i = 1 To 3: nextVertex = i Mod 3 + 1
                    For axis = 1 To D: control(i + 3, axis) = 2 * coords(i + 3, axis) - (coords(i, axis) + coords(nextVertex, axis)) / 2: Next axis
                Next i
            End If
        End If
        For axis = 1 To D
            localLower = control(1, axis): localUpper = localLower
            For i = 2 To ng: localLower = MinDouble(localLower, control(i, axis)): localUpper = MaxDouble(localUpper, control(i, axis)): Next i
            lower(axis) = MinDouble(lower(axis), localLower): upper(axis) = MaxDouble(upper(axis), localUpper)
            For l = 1 To K
                i = (e - 1) * K + l
                Pad = MaxDouble(Pad, MaxDouble(Abs(localLower - XYZ(i, axis)), Abs(localUpper - XYZ(i, axis))))
            Next l
        Next axis
    Next e
    Pad = Pad * (1 + 1E-12)
    For axis = 1 To D: center(axis) = (lower(axis) + upper(axis)) / 2: half0 = MaxDouble(half0, (upper(axis) - lower(axis)) / 2): Next axis
    RootHalf = half0 * (1 + 1E-12): ReDim ids(1 To NF)
    For i = 1 To NF: ids(i) = i: Next i
    Root = BuildNode(ids, center, RootHalf, 0, 0)
    Walk Root, Root
End Sub
Public Function DenseEstimateMB() As Double
    DenseEstimateMB = (16# * D * D * NF * NF + 40000# * NF) / 1048576
End Function
Public Function FmmEstimateMB(ByVal cacheMB As Double) As Double
    Dim total As Double
    total = (16# * D * D + 24) * NearEntries + 16# * Nodes.Count * Q * Q + 8# * NF * Q * (D + 1) + 8# * Nodes.Count * Q * (NC + D) + 40000# * NF
    total = total + MinDouble(cacheMB * 1048576, 8# * D * NC * Q * Q * Groups.Count)
    If FarPairs.Count > 0 Then total = total + 32# * D * NC * Q * Q
    FmmEstimateMB = total / 1048576
End Function
Private Sub PrepareNodes()
    Dim node As CFmmNode, parent As CFmmNode, grid() As Double, center() As Double, knots() As Double, transfer() As Double
    Dim pc() As Double, x(1 To 3) As Double, basis() As Double, i As Long, j As Long, axis As Long
    grid = Cheb.Grid
    For Each node In Nodes
        center = node.Center: ReDim knots(1 To Q, 1 To D)
        For i = 1 To Q: For axis = 1 To D: knots(i, axis) = center(axis) + node.HalfWidth * grid(i, axis): Next axis: Next i
        node.Knots = knots
    Next node
    For Each node In Nodes
        If node.Parent > 0 Then
            Set parent = Nodes(node.Parent): pc = parent.Center: knots = node.Knots: ReDim transfer(1 To Q, 1 To Q)
            For i = 1 To Q
                For axis = 1 To D: x(axis) = knots(i, axis): Next axis
                Cheb.Basis x, pc, parent.HalfWidth, basis
                For j = 1 To Q: transfer(i, j) = basis(j): Next j
            Next i
            node.Transfer = transfer
        End If
    Next node
End Sub
Private Sub PrepareLeaves(ByVal sr As CSolveRegion)
    Dim number As Variant, node As CFmmNode, ids() As Long, center() As Double, elements As Object, element As Variant
    Dim geom As CElementData, prepared As CElementRule, rule As CIntegrationRule
    Dim st() As Double, sun() As Double, target() As Double, positions() As Double, directions() As Double, weights() As Double, shapes() As Double
    Dim x(1 To 3) As Double, basis() As Double, i As Long, a As Long, axis As Long, localField As Long, qpoint As Long, e As Long, weight As Double, value As Double
    Set rule = RegularRule(D, IIf(D = 2, MaxLong(24, 2 * FmmOrder + 4), MaxLong(12, 2 * FmmOrder + 4)))
    For Each number In Leaves
        Set node = Nodes(CLng(number)): ids = node.Ids: center = node.Center
        ReDim st(1 To Q, 1 To UBound(ids)): ReDim sun(1 To Q, 1 To UBound(ids), 1 To D): ReDim target(1 To UBound(ids), 1 To Q)
        Set elements = CreateObject("Scripting.Dictionary")
        For i = 1 To UBound(ids)
            e = (ids(i) - 1) \ K
            If Not elements.Exists(CStr(e)) Then elements.Add CStr(e), e
            For axis = 1 To D: x(axis) = XYZ(ids(i), axis): Next axis
            Cheb.Basis x, center, node.HalfWidth, basis
            For a = 1 To Q: target(i, a) = basis(a): Next a
        Next i
        For Each element In elements.Items
            e = CLng(element): Set geom = sr.Geometries(e + 1): Set prepared = New CElementRule: prepared.Initialize geom, rule
            positions = prepared.XYZ: directions = prepared.Normals: weights = prepared.Weight: shapes = prepared.N
            For qpoint = 1 To prepared.Count
                For axis = 1 To D: x(axis) = positions(qpoint, axis): Next axis
                Cheb.Basis x, center, node.HalfWidth, basis
                For i = 1 To UBound(ids)
                    If (ids(i) - 1) \ K = e Then
                        localField = (ids(i) - 1) Mod K + 1: weight = weights(qpoint) * shapes(qpoint, localField)
                        For a = 1 To Q
                            value = basis(a) * weight: st(a, i) = st(a, i) + value
                            For axis = 1 To D: sun(a, i, axis) = sun(a, i, axis) + value * directions(qpoint, axis): Next axis
                        Next a
                    End If
                Next i
            Next qpoint
        Next element
        node.SourceT = st: node.SourceUN = sun: node.Target = target
        If CLng(number) Mod 8 = 0 Then Application.StatusBar = "FMM葉の積分 " & sr.Ref.RegionName & "/" & number & "/" & Nodes.Count: DoEvents
    Next number
End Sub
Private Sub PrepareNear(ByVal sr As CSolveRegion)
    Dim targetKey As Variant, sourceNode As Variant, a As CFmmNode, b As CFmmNode, sourceLeaves As Collection
    Dim selected() As Boolean, sources() As Long, ids() As Long, bid() As Long, counts() As Long, position() As Long, elements As Object, element As Variant
    Dim i As Long, j As Long, l As Long, row As Long, field As Long, pos As Long, n As Long, e As Long, axis As Long, component As Long
    Dim gb() As Double, hb() As Double, completed As Long
    ReDim counts(1 To NF)
    For Each targetKey In NearGroups.Keys
        Set a = Nodes(CLng(targetKey)): ids = a.Ids: Set sourceLeaves = NearGroups(targetKey): ReDim selected(1 To NF): n = 0
        For Each sourceNode In sourceLeaves
            Set b = Nodes(CLng(sourceNode)): bid = b.Ids
            For i = 1 To UBound(bid): If Not selected(bid(i)) Then selected(bid(i)) = True: n = n + 1
            Next i
        Next sourceNode
        ReDim sources(1 To n): j = 0
        For i = 1 To NF: If selected(i) Then j = j + 1: sources(j) = i
        Next i
        a.NearSources = sources
        For i = 1 To UBound(ids): counts(ids(i)) = n: Next i
    Next targetKey
    ReDim RowPtr(1 To NF + 1): RowPtr(1) = 1
    For i = 1 To NF: RowPtr(i + 1) = RowPtr(i) + counts(i): Next i
    If NearEntries > 2147483646# Then Fail "FMM近傍行列の添字上限を超えます。メッシュまたはleafを変更してください。"
    n = CLng(NearEntries): ReDim SourceIds(1 To n): ReDim NearG(1 To n, 1 To D, 1 To D): ReDim NearH(1 To n, 1 To D, 1 To D)
    ReDim DG(1 To NF, 1 To D, 1 To D): ReDim DH(1 To NF, 1 To D, 1 To D)
    For Each targetKey In NearGroups.Keys
        Set a = Nodes(CLng(targetKey)): ids = a.Ids: sources = a.NearSources: ReDim position(1 To NF): Set elements = CreateObject("Scripting.Dictionary")
        For j = 1 To UBound(sources)
            position(sources(j)) = j: e = (sources(j) - 1) \ K
            If Not elements.Exists(CStr(e)) Then elements.Add CStr(e), e
        Next j
        For i = 1 To UBound(ids)
            row = ids(i)
            For j = 1 To UBound(sources): SourceIds(RowPtr(row) + j - 1) = sources(j): Next j
            For Each element In elements.Items
                e = CLng(element): ElementInfluence sr, row, e, gb, hb
                For l = 1 To K
                    field = e * K + l
                    If position(field) > 0 Then
                        pos = RowPtr(row) + position(field) - 1
                        For axis = 1 To D: For component = 1 To D
                            NearG(pos, axis, component) = gb(l, axis, component): NearH(pos, axis, component) = hb(l, axis, component)
                            If field = row Then DG(row, axis, component) = gb(l, axis, component): DH(row, axis, component) = hb(l, axis, component)
                        Next component: Next axis
                    End If
                Next l
            Next element
            completed = completed + 1
            If completed Mod 4 = 0 Then Application.StatusBar = "FMM近傍積分 " & sr.Ref.RegionName & " " & completed & "/" & NF: DoEvents
        Next i
    Next targetKey
End Sub
Private Sub MakeTranslation(ByVal group As CFmmTranslation, ByRef matrix() As Double)
    Dim a As CFmmNode, b As CFmmNode, ak() As Double, bk() As Double, x(1 To 3) As Double, y(1 To 3) As Double
    Dim uk(1 To 3, 1 To 3) As Double, tensor(1 To 3, 1 To 3, 1 To 3) As Double
    Dim i As Long, j As Long, axis As Long, component As Long, direction As Long, row As Long, col As Long
    Set a = Nodes(group.TargetNode): Set b = Nodes(group.SourceNode): ak = a.Knots: bk = b.Knots
    ReDim matrix(1 To D * Q, 1 To NC * Q)
    For i = 1 To Q
        For axis = 1 To D: x(axis) = ak(i, axis): Next axis
        For j = 1 To Q
            For axis = 1 To D: y(axis) = bk(j, axis): Next axis
            KelvinTensor Mat, x, y, uk, tensor
            For axis = 1 To D
                row = (i - 1) * D + axis
                For component = 1 To D
                    matrix(row, (j - 1) * NC + component) = uk(axis, component)
                    For direction = 1 To D: matrix(row, (j - 1) * NC + D + (component - 1) * D + direction) = tensor(axis, component, direction): Next direction
                Next component
            Next axis
        Next j
    Next i
End Sub
Private Sub GetTranslation(ByVal group As CFmmTranslation, ByRef matrix() As Double)
    Dim size As Double, other As CFmmTranslation, oldest As CFmmTranslation
    ClockValue = ClockValue + 1: group.LastUse = ClockValue
    If Not IsEmpty(group.Coefficients) Then Hits = Hits + 1: matrix = group.Coefficients: Exit Sub
    Misses = Misses + 1: MakeTranslation group, matrix: size = 8# * D * NC * Q * Q
    If size <= CacheLimit Then
        Do While CacheBytes + size > CacheLimit
            Set oldest = Nothing
            For Each other In Groups
                If Not IsEmpty(other.Coefficients) Then
                    If oldest Is Nothing Then Set oldest = other Else If other.LastUse < oldest.LastUse Then Set oldest = other
                End If
            Next other
            If oldest Is Nothing Then Exit Do
            oldest.Coefficients = Empty: CacheBytes = CacheBytes - size
        Loop
        group.Coefficients = matrix: CacheBytes = CacheBytes + size
    End If
End Sub
Public Sub Prepare(ByVal sr As CSolveRegion, ByVal cacheMB As Double)
    Dim traction() As Double, displacement() As Double, out() As Double, axis As Long, i As Long, j As Long
    CacheLimit = cacheMB * 1048576
    Call PrepareNodes
    PrepareLeaves sr
    PrepareNear sr
    ReDim Correction(1 To NF, 1 To D, 1 To D): ReDim traction(1 To NF * D): ReDim displacement(1 To NF * D)
    For axis = 1 To D
        For i = 1 To NF: displacement((i - 1) * D + axis) = 1: Next i
        Raw traction, displacement, out
        For i = 1 To NF: For j = 1 To D
            Correction(i, j, axis) = -out((i - 1) * D + j): DH(i, j, axis) = DH(i, j, axis) + Correction(i, j, axis)
        Next j: Next i
        For i = 1 To NF: displacement((i - 1) * D + axis) = 0: Next i
    Next axis
    sr.DiagG = DG: sr.DiagH = DH: Prepared = True
End Sub
Public Sub Raw(ByRef traction() As Double, ByRef displacement() As Double, ByRef out() As Double)
    Dim moments() As Double, localValues() As Double, ids() As Long, st() As Double, sun() As Double, transfer() As Double, target() As Double, matrix() As Double
    Dim node As CFmmNode, group As CFmmTranslation, number As Variant, a As Long, b As Long, i As Long, j As Long, axis As Long, component As Long, direction As Long, index As Long, parent As Long
    Dim pair As Long, targetNode As Long, sourceNode As Long, value As Double, row As Long, col As Long, pos As Long
    ReDim moments(1 To Nodes.Count, 1 To Q, 1 To NC): ReDim localValues(1 To Nodes.Count, 1 To Q, 1 To D): ReDim out(1 To NF * D)
    For Each number In Leaves
        Set node = Nodes(CLng(number)): ids = node.Ids: st = node.SourceT: sun = node.SourceUN
        For i = 1 To UBound(ids)
            For component = 1 To D
                index = (ids(i) - 1) * D + component
                For a = 1 To Q
                    moments(node.Number, a, component) = moments(node.Number, a, component) + st(a, i) * traction(index)
                    For direction = 1 To D
                        col = D + (component - 1) * D + direction
                        moments(node.Number, a, col) = moments(node.Number, a, col) + sun(a, i, direction) * displacement(index)
                    Next direction
                Next a
            Next component
        Next i
    Next number
    For i = Nodes.Count To 2 Step -1
        Set node = Nodes(i): parent = node.Parent: transfer = node.Transfer
        For a = 1 To Q: For b = 1 To Q
            value = transfer(a, b)
            For component = 1 To NC: moments(parent, b, component) = moments(parent, b, component) + value * moments(i, a, component): Next component
        Next b: Next a
    Next i
    For Each group In Groups
        GetTranslation group, matrix
        For pair = 1 To group.Targets.Count
            targetNode = group.Targets(pair): sourceNode = group.Sources(pair)
            For a = 1 To Q: For axis = 1 To D
                row = (a - 1) * D + axis: value = 0
                For b = 1 To Q: For component = 1 To NC: value = value + matrix(row, (b - 1) * NC + component) * moments(sourceNode, b, component): Next component: Next b
                localValues(targetNode, a, axis) = localValues(targetNode, a, axis) + value
            Next axis: Next a
        Next pair
    Next group
    For i = 2 To Nodes.Count
        Set node = Nodes(i): parent = node.Parent: transfer = node.Transfer
        For a = 1 To Q: For b = 1 To Q
            value = transfer(a, b)
            For axis = 1 To D: localValues(i, a, axis) = localValues(i, a, axis) + value * localValues(parent, b, axis): Next axis
        Next b: Next a
    Next i
    For Each number In Leaves
        Set node = Nodes(CLng(number)): ids = node.Ids: target = node.Target
        For i = 1 To UBound(ids): For axis = 1 To D
            value = 0
            For a = 1 To Q: value = value + target(i, a) * localValues(node.Number, a, axis): Next a
            out((ids(i) - 1) * D + axis) = value
        Next axis: Next i
    Next number
    For i = 1 To NF
        For pos = RowPtr(i) To RowPtr(i + 1) - 1
            j = SourceIds(pos)
            For axis = 1 To D
                row = (i - 1) * D + axis
                For component = 1 To D
                    col = (j - 1) * D + component
                    out(row) = out(row) + NearG(pos, axis, component) * traction(col) + NearH(pos, axis, component) * displacement(col)
                Next component
            Next axis
        Next pos
    Next i
End Sub
Public Sub Apply(ByRef traction() As Double, ByRef displacement() As Double, ByRef out() As Double)
    Dim i As Long, axis As Long, component As Long, row As Long
    Raw traction, displacement, out
    For i = 1 To NF: For axis = 1 To D
        row = (i - 1) * D + axis
        For component = 1 To D: out(row) = out(row) + Correction(i, axis, component) * displacement((i - 1) * D + component): Next component
    Next axis: Next i
End Sub
Public Function StatsJSON() As String
    StatsJSON = "{""backend"":""fmm"",""elements"":" & NF \ K & ",""field_nodes"":" & NF & ",""p"":" & FmmOrder & ",""nodes"":" & Nodes.Count & ",""leaves"":" & Leaves.Count & ",""far_pairs"":" & FarPairs.Count & ",""near_pairs"":" & NearPairs.Count
    StatsJSON = StatsJSON & ",""near_entries"":" & JNum(NearEntries) & ",""translation_groups"":" & Groups.Count & ",""near_matrix_MB"":" & JNum(((16# * D * D + 4) * NearEntries + 4# * (NF + 1)) / 1048576)
    StatsJSON = StatsJSON & ",""m2l_cache_MB"":" & JNum(CacheBytes / 1048576) & ",""cache_limit_MB"":" & JNum(CacheLimit / 1048576) & ",""cache_hits"":" & Hits & ",""cache_misses"":" & Misses & "}"
End Function
Public Function CalibrationSources(ByVal index As Long) As Variant
    Dim keys As Variant, node As CFmmNode, sources() As Long, selected() As Boolean, sourceNode As Variant, b As CFmmNode, bid() As Long, i As Long, count As Long
    keys = NearGroups.Keys: Set node = Nodes(CLng(keys(index)))
    ReDim selected(1 To NF)
    For Each sourceNode In NearGroups(keys(index))
        Set b = Nodes(CLng(sourceNode)): bid = b.Ids
        For i = 1 To UBound(bid): selected(bid(i)) = True: Next i
    Next sourceNode
    For i = 1 To NF: If selected(i) Then count = count + 1
    Next i
    ReDim sources(1 To count): count = 0
    For i = 1 To NF: If selected(i) Then count = count + 1: sources(count) = i
    Next i
    CalibrationSources = sources
End Function
Public Function CalibrationTargets(ByVal index As Long) As Variant
    Dim keys As Variant, node As CFmmNode
    keys = NearGroups.Keys: Set node = Nodes(CLng(keys(index))): CalibrationTargets = node.Ids
End Function
Public Function NearGroupCount() As Long
    NearGroupCount = NearGroups.Count
End Function
