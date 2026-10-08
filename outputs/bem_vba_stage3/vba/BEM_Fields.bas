Attribute VB_Name = "BEM_Fields"
Option Explicit
Public FieldReports As Object
Private Function InteriorOrder(ByVal d As Long, ByVal level As Long) As Long
    Dim values As Variant
    If d = 2 Then values = Array(8, 12, 20, 32, 48, 72, 112) Else values = Array(4, 6, 8, 12, 16, 24, 32)
    InteriorOrder = values(MinLong(level, 6)) * 2 ^ MaxLong(level - 6, 0)
End Function
Private Function EvaluateElement(ByVal sr As CSolveRegion, ByVal e As Long, ByRef x() As Double, ByRef uref() As Double, ByVal s As Double, ByVal t As Double, ByVal distance As Double, ByVal family As String, ByVal level As Long, ByRef count As Long, ByRef values() As Double, ByRef floor() As Double) As Boolean
    Dim geom As CElementData, rule As CIntegrationRule, prepared As CElementRule, upper As Double, order As Long
    Dim q As Long, d As Long, k As Long, ns As Long, nvalues As Long, i As Long, j As Long, l As Long, index As Long, a As Long, b As Long
    Dim y(1 To 3) As Double, normal(1 To 3) As Double, uy(1 To 3) As Double, ty(1 To 3) As Double
    Dim uk(1 To 3, 1 To 3) As Double, tk(1 To 3, 1 To 3) As Double, du(1 To 3, 1 To 3, 1 To 3) As Double, dt(1 To 3, 1 To 3, 1 To 3) As Double
    Dim grad(1 To 3, 1 To 3) As Double, f(1 To 12) As Double, compensation(1 To 12) As Double
    Dim trace As Double, value As Double, weighted As Double, adjusted As Double, total As Double
    Dim positions() As Double, directions() As Double, weights() As Double, shapes() As Double, boundaryU() As Double, boundaryT() As Double
    d = sr.Ref.Dimension: k = sr.Ref.FieldCount: ns = d * d: If d = 2 Then ns = ns + 1
    nvalues = d + ns: Set geom = sr.Geometries(e + 1): order = InteriorOrder(d, level)
    If family = "gauss" Then upper = CDbl(order) ^ (d - 1) Else upper = IIf(d = 2, 2# * (16 * 2 ^ level + 1), 3# * (16 * 2 ^ level + 1) * 12 * 2 ^ MinLong(level, 4))
    If count + upper > MaxPointBudget Then Exit Function
    If family = "gauss" Then Set rule = RegularRule(d, order) Else Set rule = NearRule(geom, s, t, distance, level)
    If count + rule.count > MaxPointBudget Then Exit Function
    Set prepared = New CElementRule: prepared.Initialize geom, rule
    positions = prepared.xyz: directions = prepared.Normals: weights = prepared.weight: shapes = prepared.n
    boundaryU = sr.u: boundaryT = sr.t
    ReDim values(1 To nvalues): ReDim floor(1 To nvalues)
    For q = 1 To prepared.count
        For i = 1 To d
            y(i) = positions(q, i): normal(i) = directions(q, i): uy(i) = -uref(i): ty(i) = 0
            For l = 1 To k
                index = (e * k + l - 1) * d + i
                uy(i) = uy(i) + shapes(q, l) * boundaryU(index): ty(i) = ty(i) + shapes(q, l) * boundaryT(index)
            Next l
        Next i
        Kelvin sr, x, y, normal, uk, tk, du, dt, True
        For i = 1 To d
            f(i) = 0
            For j = 1 To d: f(i) = f(i) + uk(i, j) * ty(j) - tk(i, j) * uy(j): Next j
            For l = 1 To d
                grad(i, l) = 0
                For j = 1 To d: grad(i, l) = grad(i, l) + du(i, j, l) * ty(j) - dt(i, j, l) * uy(j): Next j
            Next l
        Next i
        trace = 0: For i = 1 To d: trace = trace + grad(i, i): Next i
        index = d
        For i = 1 To d
            For j = 1 To d
                index = index + 1: f(index) = sr.Mu * (grad(i, j) + grad(j, i))
                If i = j Then f(index) = f(index) + sr.Lambda * trace
            Next j
        Next i
        If d = 2 Then
            f(nvalues) = 0
            If sr.IsPlaneStrain Then f(nvalues) = sr.Lambda * trace
        End If
        For i = 1 To nvalues
            weighted = weights(q) * f(i): adjusted = weighted - compensation(i): total = values(i) + adjusted
            compensation(i) = (total - values(i)) - adjusted: values(i) = total
            floor(i) = floor(i) + Abs(weighted)
        Next i
    Next q
    For i = 1 To nvalues: floor(i) = floor(i) * 64 * 2.22044604925031E-16: Next i
    count = count + prepared.count: EvaluateElement = True
End Function

Private Function EvaluatePoint(ByVal sr As CSolveRegion, ByVal pointIndex As Long, ByRef row As Variant) As Boolean
    Dim d As Long, k As Long, e As Long, i As Long, j As Long, l As Long, nearest As Long, levels As Long, level As Long, familyIndex As Long, count As Long
    Dim x(1 To 3) As Double, uref(1 To 3) As Double, shape(1 To 6) As Double, point As Variant, geom As CElementData
    Dim s() As Double, t() As Double, distances() As Double, nearestDistance As Double, target As Double, family As String, firstFamily As String
    Dim values() As Double, floor() As Double, previous() As Double, lastValue() As Double, estimates() As Double, accumulated() As Double, errors() As Double
    Dim good As Boolean, hasValue As Boolean, hasPrevious As Boolean, hasEstimate As Boolean, passed As Long, available As Boolean, allGood As Boolean, nvalues As Long, stress(1 To 3, 1 To 3) As Double
    Dim uError As Double, stressError As Double, elementError As Double, errorCode As String, reason As String, details As Collection, estimated As Collection, elementsDE As Long, switched As Boolean
    d = sr.Ref.Dimension: k = sr.Ref.FieldCount: point = sr.Ref.points(pointIndex): nvalues = d + d * d + IIf(d = 2, 1, 0)
    For i = 1 To d: x(i) = point(i - 1): Next i
    ReDim s(0 To sr.Ref.elements.count - 1): ReDim t(0 To sr.Ref.elements.count - 1): ReDim distances(0 To sr.Ref.elements.count - 1)
    ReDim accumulated(1 To nvalues): ReDim errors(1 To nvalues): nearestDistance = 1E+250
    row = Array(sr.Ref.RegionName, sr.Ref.PointIds(pointIndex), x(1), x(2), Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, False, "INVALID_INTERIOR_POINT", False)
    If d = 3 Then row(4) = x(3)
    For e = 0 To sr.Ref.elements.count - 1
        Set geom = sr.Geometries(e + 1): ClosestPoint geom, x, s(e), t(e), distances(e)
        If distances(e) / geom.length <= InteriorMinDistanceRatio Then
            For i = 5 To 13: row(i) = "éÊìæïsâ¬": Next i
            FieldReports(sr.Ref.RegionName & "/" & sr.Ref.PointIds(pointIndex)) = "{""converged"":false,""error_code"":""INVALID_INTERIOR_POINT"",""value_available"":false}"
            Exit Function
        End If
        If distances(e) < nearestDistance Then nearestDistance = distances(e): nearest = e
    Next e
    FieldShape d, sr.Ref.ElementOrder, s(nearest), t(nearest), shape
    For i = 1 To d
        For l = 1 To k: uref(i) = uref(i) + shape(l) * sr.u((nearest * k + l - 1) * d + i): Next l
    Next i
    available = True: allGood = True
    Set details = New Collection
    For e = 0 To sr.Ref.elements.count - 1
        Set geom = sr.Geometries(e + 1): firstFamily = "gauss"
        If QuadratureMode = "de" Or (QuadratureMode = "auto" And distances(e) / geom.length < InteriorNearRatio) Then firstFamily = "de"
        count = 0: hasValue = False: hasEstimate = False: good = False
        reason = "refinement_limit_or_roundoff": switched = False
        ReDim lastValue(1 To nvalues): ReDim estimates(1 To nvalues)
        For familyIndex = 0 To IIf(firstFamily = "gauss" And QuadratureMode = "auto", 1, 0)
            family = firstFamily: If familyIndex = 1 Then family = "de"
            If familyIndex = 1 Then switched = True
            levels = IIf(family = "gauss", MaxGaussLevel, MaxDELevel): hasPrevious = False: passed = 0
            ReDim previous(1 To nvalues)
            For level = 0 To levels
                If Not EvaluateElement(sr, e, x, uref, s(e), t(e), distances(e), family, level, count, values, floor) Then reason = "point_budget": Exit For
                hasValue = True
                For i = 1 To nvalues: lastValue(i) = values(i): Next i
                If hasPrevious Then
                    hasEstimate = True: passed = passed + 1
                    For i = 1 To nvalues
                        estimates(i) = MaxDouble(Abs(values(i) - previous(i)), floor(i))
                        target = IIf(i <= d, DisplacementTolerance, StressTolerance) / sr.Ref.elements.count
                        If estimates(i) > target Then passed = 0
                    Next i
                    If passed >= 2 Then good = True: reason = "two_refinements": Exit For
                End If
                For i = 1 To nvalues: previous(i) = values(i): Next i
                hasPrevious = True
            Next level
            If good Then Exit For
        Next familyIndex
        If Not hasValue Then available = False
        If Not good Then allGood = False
        For i = 1 To nvalues
            accumulated(i) = accumulated(i) + lastValue(i)
            If hasEstimate Then errors(i) = errors(i) + estimates(i) Else errors(i) = 1E+200
        Next i
        Set estimated = New Collection
        For i = 1 To nvalues: If hasEstimate Then estimated.Add JNum(estimates(i)) Else estimated.Add "null"
        Next i
        If family = "de" Then elementsDE = elementsDE + 1
        details.Add "{""element"":" & e & ",""method"":" & JQuote(family) & ",""converged"":" & LCase$(CStr(good)) & ",""reason"":" & JQuote(reason) & ",""evaluations"":" & count & ",""gauss_to_de"":" & LCase$(CStr(switched)) & ",""distance_ratio"":" & JNum(distances(e) / geom.length) & ",""error_estimate"": [" & JoinCollection(estimated, ",") & "]}"
        If e Mod 4 = 0 Then Application.StatusBar = "ì‡ïîèÍêœï™ " & sr.Ref.RegionName & "/" & sr.Ref.PointIds(pointIndex) & " " & e + 1 & "/" & sr.Ref.elements.count: DoEvents
    Next e
    For i = 1 To nvalues
        target = IIf(i <= d, DisplacementTolerance, StressTolerance)
        If errors(i) > target Then allGood = False
        If i <= d Then uError = MaxDouble(uError, errors(i)) Else stressError = MaxDouble(stressError, errors(i))
    Next i
    allGood = allGood And available
    If available Then
        For i = 1 To d: row(4 + i) = accumulated(i) + uref(i): Next i
        l = d
        For i = 1 To d: For j = 1 To d: l = l + 1: stress(i, j) = accumulated(l): Next j: Next i
        If d = 2 Then stress(3, 3) = accumulated(nvalues)
        row(8) = stress(1, 1): row(9) = stress(2, 2): row(10) = stress(3, 3): row(11) = stress(1, 2): row(12) = stress(2, 3): row(13) = stress(3, 1)
    Else
        For i = 5 To 13: row(i) = "éÊìæïsâ¬": Next i
    End If
    errorCode = IIf(allGood, "OK", "QUADRATURE_TOLERANCE_UNMET")
    row(14) = allGood: row(15) = errorCode: row(16) = available
    Set estimated = New Collection
    For i = 1 To nvalues: If errors(i) < 1E+190 Then estimated.Add JNum(errors(i)) Else estimated.Add "null"
    Next i
    FieldReports(sr.Ref.RegionName & "/" & sr.Ref.PointIds(pointIndex)) = "{""point_id"":" & JQuote(sr.Ref.PointIds(pointIndex)) & ",""converged"":" & LCase$(CStr(allGood)) & ",""error_code"":" & JQuote(errorCode) & ",""value_available"":" & LCase$(CStr(available)) & ",""de_elements"":" & elementsDE & ",""error_estimate"": [" & JoinCollection(estimated, ",") & "],""element_reports"": [" & JoinCollection(details, ",") & "],""error_scope"":""quadrature estimate only; boundary discretization/FMM/algebraic errors excluded""}"
    LogAction "ì‡ì_êœï™", errorCode, sr.Ref.RegionName & "/" & sr.Ref.PointIds(pointIndex) & " / êÑíËïœà åÎç∑[m] " & JNum(uError) & " / êÑíËâûóÕåÎç∑[kPa] " & JNum(stressError)
    EvaluatePoint = allGood
End Function
Public Sub EvaluateAllFields(ByVal rows As Collection, ByRef allGood As Boolean)
    Dim name As Variant, sr As CSolveRegion, i As Long, row As Variant
    Set FieldReports = CreateObject("Scripting.Dictionary")
    For Each name In SolveData.keys
        Set sr = SolveData(name)
        For i = 1 To sr.Ref.points.count
            If Not EvaluatePoint(sr, i, row) Then allGood = False
            rows.Add row
        Next i
    Next name
End Sub

