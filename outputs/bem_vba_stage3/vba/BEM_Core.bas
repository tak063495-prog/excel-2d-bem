Attribute VB_Name = "BEM_Core"
Option Explicit

Public Models As Object
Public InterfacePairs As Collection
Public MeshFingerprint As String
Public LastImportFingerprint As String
Public Busy As Boolean
Public Const FIRST_ROW As Long = 6

Public Sub Fail(ByVal message As String)
    Err.Raise vbObjectError + 2048, "ElasticBEM", message
End Sub
Public Function WS(ByVal name As String) As Worksheet
    Set WS = ThisWorkbook.Worksheets(name)
End Function
Public Function TextAt(ByVal s As Worksheet, ByVal row As Long, ByVal col As Long) As String
    Dim value As Variant: value = s.Cells(row, col).Value2
    If IsError(value) Then Fail s.name & "!" & s.Cells(row, col).address(False, False) & ": セルがエラーです。"
    TextAt = Trim$(CStr(value))
End Function
Public Function NumberAt(ByVal s As Worksheet, ByVal row As Long, ByVal col As Long) As Double
    Dim value As Variant, address As String
    value = s.Cells(row, col).Value2: address = s.name & "!" & s.Cells(row, col).address(False, False)
    If IsError(value) Then Fail address & ": セルがエラーです。"
    If IsEmpty(value) Or VarType(value) = vbBoolean Then Fail address & ": 数値が必要です。"
    If Len(Trim$(CStr(value))) = 0 Or Not IsNumeric(value) Then Fail address & ": 数値が必要です。"
    NumberAt = CDbl(value)
    If Abs(NumberAt) > 1E+100 Then Fail address & ": 数値の絶対値が大きすぎます。"
End Function
Public Function ZeroIfBlank(ByVal s As Worksheet, ByVal row As Long, ByVal col As Long) As Double
    If TextAt(s, row, col) <> "" Then ZeroIfBlank = NumberAt(s, row, col)
End Function
Public Function LastRow(ByVal s As Worksheet, ByVal ncols As Long) As Long
    Dim col As Long, row As Long: LastRow = FIRST_ROW - 1
    For col = 1 To ncols
        row = s.Cells(s.rows.count, col).End(xlUp).row
        If row > LastRow Then LastRow = row
    Next col
    If LastRow > 10005 Then Fail s.name & ": 入力行は10000行までです。"
End Function
Public Function ActiveRow(ByVal s As Worksheet, ByVal row As Long, ByVal ncols As Long) As Boolean
    Dim col As Long, used As Boolean
    For col = 1 To ncols
        If TextAt(s, row, col) <> "" Then used = True
    Next col
    If used And TextAt(s, row, 1) = "" Then Fail s.name & " 行" & row & ": 領域IDまたはIDが空欄です。"
    ActiveRow = used
End Function
Public Function ValidId(ByVal value As String) As Boolean
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[A-Za-z0-9_-]+$": ValidId = re.Test(value)
End Function
Public Function CurrentDimension() As Long
    Select Case TextAt(WS("操作"), 5, 2)
    Case "2D": CurrentDimension = 2
    Case "3D": CurrentDimension = 3
    Case Else: Fail "操作!B5: 2Dまたは3Dを選択してください。"
    End Select
End Function
Public Function CurrentOrder() As Long
    Dim v As Double: v = NumberAt(WS("操作"), 7, 2)
    If v <> 0 And v <> 2 Then Fail "操作!B7: 要素次数は0または2です。"
    CurrentOrder = CLng(v)
End Function
Public Function MeshSize() As Double
    MeshSize = NumberAt(WS("操作"), 8, 2)
    If MeshSize <= 0 Then Fail "操作!B8: hは正の数です。"
End Function
Public Function Ceiling(ByVal value As Double) As Long
    If value > 100000 Or value < 0 Then Fail "分割数が上限を超えています。hを大きくしてください。"
    Ceiling = CLng(-Int(-value))
    If Ceiling < 1 Then Ceiling = 1
End Function
Public Function DivisionAt(ByVal row As Long, ByVal col As Long, ByVal fallback As Long, Optional ByVal minimum As Long = 1) As Long
    Dim v As Double: v = ZeroIfBlank(WS("領域"), row, col)
    If v < 0 Or v <> Fix(v) Or v > 100000 Then Fail "領域 行" & row & ": 分割数は0または正の整数です。"
    If v = 0 Then v = MaxLong(fallback, minimum)
    If v < minimum Then Fail "領域 行" & row & ": 分割数は" & minimum & "以上です。"
    DivisionAt = CLng(v)
End Function
Public Function DivisionFromSize(ByVal row As Long, ByVal col As Long, ByVal ratio As Double, Optional ByVal minimum As Long = 1) As Long
    If ZeroIfBlank(WS("領域"), row, col) = 0 Then
        DivisionFromSize = DivisionAt(row, col, MaxLong(minimum, Ceiling(ratio)), minimum)
    Else
        DivisionFromSize = DivisionAt(row, col, minimum, minimum)
    End If
End Function

Public Function InputFingerprint() As String
    Dim name As Variant, s As Worksheet, data As Variant, nr As Long, nc As Long, i As Long, j As Long
    Dim parts As Collection: Set parts = New Collection
    data = WS("操作").Range("B5:B13").Value2
    For i = 1 To UBound(data, 1)
        If IsError(data(i, 1)) Then Fail "操作シートにセルエラーがあります。"
        parts.Add JQuote(CStr(data(i, 1)))
    Next i
    data = WS("操作").Range("B34:B61").Value2
    For i = 1 To UBound(data, 1)
        If IsError(data(i, 1)) Then Fail "解析設定にセルエラーがあります。"
        parts.Add JQuote(CStr(data(i, 1)))
    Next i
    For Each name In Array("材料", "領域", "輪郭2D", "境界条件", "接合", "評価点")
        Set s = WS(CStr(name))
        Select Case CStr(name)
        Case "領域": nc = 13
        Case "境界条件": nc = 8
        Case "評価点": nc = 5
        Case Else: nc = 4
        End Select
        nr = LastRow(s, nc): parts.Add CStr(name)
        If nr >= FIRST_ROW Then
            data = s.Range(s.Cells(FIRST_ROW, 1), s.Cells(nr, nc)).Value2
            For i = 1 To UBound(data, 1)
                For j = 1 To nc
                    If IsError(data(i, j)) Then Fail CStr(name) & ": セルエラーがあります。"
                    parts.Add JQuote(CStr(data(i, j)))
                Next j
                parts.Add ";"
            Next i
        End If
    Next name
    InputFingerprint = JoinCollection(parts, ",")
End Function

Public Sub BuildModels()
    Dim oldEvents As Boolean, mats As Object, newModels As Object, s As Worksheet, row As Long, id As String
    Dim r As CBemRegion, v As Variant, h As Double, d As Long, order As Long, state As String
    Dim estimate As Double, count As Double, cap As Double, nx As Long, ny As Long, nz As Long
    Dim corners As Variant, i As Long, span As Variant, lower As Variant, upper As Variant
    Dim fingerprint As String, description As String, errnum As Long
    On Error GoTo Failed
    d = CurrentDimension(): order = CurrentOrder(): h = MeshSize()
    state = TextAt(WS("操作"), 6, 2)
    If d = 2 And state <> "plane_strain" And state <> "plane_stress" Then Fail "2D解析状態を選択してください。"
    cap = NumberAt(WS("操作"), 9, 2)
    If cap <> Fix(cap) Or cap < 4 Or cap > 100000 Then Fail "最大要素数は4～100000の整数です。"
    fingerprint = InputFingerprint()
    Set mats = CreateObject("Scripting.Dictionary"): Set newModels = CreateObject("Scripting.Dictionary")
    Set s = WS("材料")
    For row = FIRST_ROW To LastRow(s, 4)
        If ActiveRow(s, row, 4) Then
            id = TextAt(s, row, 1)
            If Not ValidId(id) Or mats.Exists(id) Then Fail "材料 行" & row & ": IDが不正または重複しています。"
            v = Array(NumberAt(s, row, 2), NumberAt(s, row, 3))
            If v(0) <= 0 Or v(1) <= -1 Or v(1) >= 0.5 Then Fail "材料 行" & row & ": E>0、-1<ν<0.5が必要です。"
            mats.Add id, v
        End If
    Next row
    Set s = WS("領域")
    For row = FIRST_ROW To LastRow(s, 13)
        If ActiveRow(s, row, 13) Then
            Set r = New CBemRegion: r.RegionName = TextAt(s, row, 1)
            If Not ValidId(r.RegionName) Or newModels.Exists(r.RegionName) Then Fail "領域 行" & row & ": IDが不正または重複しています。"
            r.Dimension = d: r.ElementOrder = order: r.ShapeName = UCase$(TextAt(s, row, 2)): r.MaterialId = TextAt(s, row, 3)
            If Not mats.Exists(r.MaterialId) Then Fail "領域 行" & row & ": 材料IDが見つかりません。"
            v = mats(r.MaterialId): r.Young = v(0): r.Poisson = v(1)
            If r.ShapeName = "POLYGON" Then r.Origin = Array(0#, 0#, 0#) Else r.Origin = Array(NumberAt(s, row, 4), NumberAt(s, row, 5), 0#)
            If d = 3 Then r.Origin = Array(r.Origin(0), r.Origin(1), NumberAt(s, row, 6))
            r.Sizes = Array(0#, 0#, 0#): nx = 0: ny = 0: nz = 0
            Select Case r.ShapeName
            Case "RECT", "BOX"
                If (d = 2 And r.ShapeName <> "RECT") Or (d = 3 And r.ShapeName <> "BOX") Then Fail "領域 行" & row & ": 解析次元に対応する形状を選択してください。"
                span = Array(NumberAt(s, row, 7), NumberAt(s, row, 8), 0#)
                If d = 3 Then span(2) = NumberAt(s, row, 9)
                For i = 0 To d - 1
                    If span(i) <= 0 Then Fail "領域 行" & row & ": 寸法は正の数です。"
                Next i
                r.Sizes = span: r.ModelScale = MaxDouble(span(0), MaxDouble(span(1), span(2)))
                nx = DivisionFromSize(row, 11, span(0) / h): ny = DivisionFromSize(row, 12, span(1) / h)
                If d = 3 Then nz = DivisionFromSize(row, 13, span(2) / h)
                If d = 2 Then count = 2# * (nx + ny) Else count = 4# * (CDbl(nx) * ny + CDbl(ny) * nz + CDbl(nz) * nx)
            Case "CIRCLE", "SPHERE"
                If (d = 2 And r.ShapeName <> "CIRCLE") Or (d = 3 And r.ShapeName <> "SPHERE") Then Fail "領域 行" & row & ": 解析次元に対応する形状を選択してください。"
                r.Radius = NumberAt(s, row, 10)
                If r.Radius <= 0 Then Fail "領域 行" & row & ": 半径は正の数です。"
                r.ModelScale = 2 * r.Radius: nx = DivisionFromSize(row, 11, 2 * PI * r.Radius / h, 8)
                If d = 2 Then
                    count = nx
                Else
                    ny = DivisionFromSize(row, 12, PI * r.Radius / h, 4): count = 2# * nx * (ny - 1)
                End If
            Case "POLYGON"
                If d <> 2 Then Fail "POLYGONは2D専用です。"
                corners = ReadPolygon(r): lower = corners(0): upper = corners(0)
                For i = 0 To UBound(corners)
                    lower = Array(MinDouble(lower(0), corners(i)(0)), MinDouble(lower(1), corners(i)(1)), 0#)
                    upper = Array(MaxDouble(upper(0), corners(i)(0)), MaxDouble(upper(1), corners(i)(1)), 0#)
                Next i
                r.Origin = lower: span = Sub3(upper, lower): r.ModelScale = MaxDouble(span(0), span(1))
                nx = DivisionAt(row, 11, 0): If ZeroIfBlank(s, row, 11) = 0 Then nx = 0
                count = 0
                For i = 0 To UBound(corners)
                    If nx > 0 Then count = count + nx Else count = count + Ceiling(Norm3(Sub3(corners(i), corners((i + 1) Mod (UBound(corners) + 1)))) / h)
                Next i
            Case Else: Fail "領域 行" & row & ": 形状を選択してください。"
            End Select
            If r.ModelScale < 0.0000000001 Or r.ModelScale > 1000000000000# Then Fail r.RegionName & ": モデル寸法は1E-10～1E12 mです。"
            For i = 0 To d - 1
                If Abs(r.Origin(i)) > r.ModelScale * 100000000 Then Fail r.RegionName & ": 原点をモデルの近くに移してください（丸め誤差対策）。"
            Next i
            estimate = estimate + count
            If estimate > cap Then Fail "要素数の見積もり " & Format$(estimate, "0") & " が上限 " & cap & " を超えます。分割数を減らしてください。"
            r.Divisions = Array(nx, ny, nz): newModels.Add r.RegionName, r
        End If
    Next row
    If newModels.count = 0 Then Fail "少なくとも1領域を入力してください。"
    For Each v In newModels.keys
        Set r = newModels(v): GenerateGeometry r
    Next v
    'Only commit models after geometry generation; failed later checks invalidate the in-memory model.
    Set Models = newModels
    ReadBoundaryRules
    MatchInterfaces
    ReadEvaluationPoints
    CheckRigidModes
    MeshFingerprint = fingerprint
    Exit Sub
Failed:
    errnum = Err.number: description = Err.description
    Set Models = Nothing: MeshFingerprint = ""
    Err.Raise errnum, "ElasticBEM", description
End Sub

Public Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function
Public Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function

Public Function PolygonArea(ByVal points As Variant) As Double
    Dim i As Long, a As Variant, b As Variant, base As Variant
    base = points(0)
    For i = 0 To UBound(points)
        a = Sub3(points(i), base): b = Sub3(points((i + 1) Mod (UBound(points) + 1)), base)
        PolygonArea = PolygonArea + a(0) * b(1) - a(1) * b(0)
    Next i
    PolygonArea = PolygonArea / 2
End Function
Public Function ReadPolygon(ByVal r As CBemRegion) As Variant
    Dim s As Worksheet, items As Object, row As Long, id As String, seq As Double, i As Long, j As Long
    Dim p() As Variant, a As Variant, b As Variant, n As Long, ModelScale As Double, tol As Double
    Set s = WS("輪郭2D"): Set items = CreateObject("Scripting.Dictionary")
    For row = FIRST_ROW To LastRow(s, 4)
        If ActiveRow(s, row, 4) Then
            id = TextAt(s, row, 1)
            If id = r.RegionName Then
                seq = NumberAt(s, row, 2)
                If seq < 1 Or seq <> Fix(seq) Or seq > 10000 Then Fail "輪郭2D: 順番は1～10000の整数です。"
                If items.Exists(CStr(seq)) Then Fail "輪郭2D: 順番が重複しています。"
                items.Add CStr(seq), Array(NumberAt(s, row, 3), NumberAt(s, row, 4), 0#)
            End If
        End If
    Next row
    n = items.count
    If n < 3 Then Fail r.RegionName & ": 多角形は3頂点以上必要です。"
    ReDim p(0 To n - 1)
    For i = 1 To n
        If Not items.Exists(CStr(i)) Then Fail r.RegionName & ": 頂点の順番を1から連番で入力してください。"
        p(i - 1) = items(CStr(i)): ModelScale = MaxDouble(ModelScale, Norm3(Sub3(p(i - 1), items("1"))))
    Next i
    tol = MaxDouble(ModelScale * 0.00000000001, 1E-20)
    If Abs(PolygonArea(p)) <= tol * ModelScale Then Fail r.RegionName & ": 多角形の面積が0です。"
    For i = 0 To n - 1
        a = p(i): b = p((i + 1) Mod n)
        If Norm3(Sub3(b, a)) <= tol Then Fail r.RegionName & ": 重複頂点があります。先頭を末尾に再入力しないでください。"
        If Abs(Orient2(a, b, p((i + 2) Mod n))) <= tol * Norm3(Sub3(b, a)) Then
            If Dot3(Sub3(b, a), Sub3(p((i + 2) Mod n), b)) < 0 Then Fail r.RegionName & ": 隣接する辺が折り返して重なっています。"
        End If
        For j = i + 1 To n - 1
            If j <> i + 1 And Not (i = 0 And j = n - 1) Then
                If SegmentsIntersect(a, b, p(j), p((j + 1) Mod n), tol) Then Fail r.RegionName & ": 多角形が自己交差または接触しています。"
            End If
        Next j
    Next i
    ReadPolygon = p
End Function
Private Function Orient2(ByVal a As Variant, ByVal b As Variant, ByVal c As Variant) As Double
    Orient2 = (b(0) - a(0)) * (c(1) - a(1)) - (b(1) - a(1)) * (c(0) - a(0))
End Function
Private Function OnSegment(ByVal a As Variant, ByVal b As Variant, ByVal p As Variant, ByVal tol As Double) As Boolean
    OnSegment = Abs(Orient2(a, b, p)) <= tol * Norm3(Sub3(b, a)) And p(0) >= MinDouble(a(0), b(0)) - tol And p(0) <= MaxDouble(a(0), b(0)) + tol And p(1) >= MinDouble(a(1), b(1)) - tol And p(1) <= MaxDouble(a(1), b(1)) + tol
End Function
Private Function SegmentsIntersect(ByVal a As Variant, ByVal b As Variant, ByVal c As Variant, ByVal d As Variant, ByVal tol As Double) As Boolean
    Dim o1 As Double, o2 As Double, o3 As Double, o4 As Double
    o1 = Orient2(a, b, c): o2 = Orient2(a, b, d): o3 = Orient2(c, d, a): o4 = Orient2(c, d, b)
    SegmentsIntersect = ((o1 > 0 And o2 < 0) Or (o1 < 0 And o2 > 0)) And ((o3 > 0 And o4 < 0) Or (o3 < 0 And o4 > 0))
    If OnSegment(a, b, c, tol) Or OnSegment(a, b, d, tol) Or OnSegment(c, d, a, tol) Or OnSegment(c, d, b, tol) Then SegmentsIntersect = True
End Function

Private Sub ReadBoundaryRules()
    Dim s As Worksheet, row As Long, r As CBemRegion, id As String, boundary As String, kinds As Variant, values As Variant, j As Long
    Dim key As Variant, col As Long, kind As String
    Set s = WS("境界条件")
    For row = FIRST_ROW To LastRow(s, 8)
        If ActiveRow(s, row, 8) Then
            id = TextAt(s, row, 1): If Not Models.Exists(id) Then Fail "境界条件 行" & row & ": 領域IDが見つかりません。"
            Set r = Models(id): boundary = UCase$(TextAt(s, row, 2))
            If boundary <> "ALL" And Not r.groups.Exists(boundary) Then Fail id & ": 境界名 " & boundary & " がありません。"
            If r.Rules.Exists(boundary) Then Fail id & "/" & boundary & ": 境界条件が重複しています。"
            kinds = Array("t", "t", "t"): values = Array(0#, 0#, 0#)
            For j = 0 To r.Dimension - 1
                col = 3 + 2 * j: kind = LCase$(TextAt(s, row, col))
                If kind <> "u" And kind <> "t" Then Fail "境界条件 行" & row & ": 種別はuまたはtです。"
                kinds(j) = kind: values(j) = NumberAt(s, row, col + 1)
            Next j
            r.Rules.Add boundary, Array(kinds, values)
        End If
    Next row
End Sub

Private Function ElementKey(ByVal r As CBemRegion, ByVal e As Long, ByVal tol As Double) As String
    Dim ids As Variant, p As Variant, keys() As String, i As Long, j As Long, temp As String
    ids = r.elements(e + 1): ReDim keys(0 To UBound(ids))
    For i = 0 To UBound(ids)
        p = r.Vertices(ids(i) + 1)
        keys(i) = Format$(p(0) / tol, "0") & "," & Format$(p(1) / tol, "0") & "," & Format$(p(2) / tol, "0")
    Next i
    For i = 1 To UBound(keys)
        temp = keys(i): j = i - 1
        Do While j >= 0
            If StrComp(keys(j), temp, vbBinaryCompare) <= 0 Then Exit Do
            keys(j + 1) = keys(j): j = j - 1
        Loop
        keys(j + 1) = temp
    Next i
    ElementKey = Join(keys, ";")
End Function
Private Sub MatchInterfaces()
    Dim s As Worksheet, row As Long, a As CBemRegion, b As CBemRegion, ba As String, bb As String
    Dim key As String, lookup As Object, ea As Collection, eb As Collection, id As Variant, e As Variant, f As Long
    Dim na As Variant, nb As Variant, jac As Double, tol As Double, idsA As Variant, idsB As Variant, pairs As Collection
    Dim i As Long, j As Long, pa As Variant, pb As Variant, good As Boolean
    Set InterfacePairs = New Collection: Set s = WS("接合")
    For row = FIRST_ROW To LastRow(s, 4)
        If ActiveRow(s, row, 4) Then
            If Not Models.Exists(TextAt(s, row, 1)) Or Not Models.Exists(TextAt(s, row, 3)) Then Fail "接合 行" & row & ": 領域IDが見つかりません。"
            Set a = Models(TextAt(s, row, 1)): Set b = Models(TextAt(s, row, 3))
            If a.RegionName = b.RegionName Then Fail "接合は異なる2領域を指定してください。"
            ba = UCase$(TextAt(s, row, 2)): bb = UCase$(TextAt(s, row, 4))
            If Not a.groups.Exists(ba) Or Not b.groups.Exists(bb) Then Fail "接合 行" & row & ": 境界名が見つかりません。"
            If a.InterfaceGroups.Exists(ba) Or b.InterfaceGroups.Exists(bb) Then Fail "接合境界が重複しています。"
            If a.Rules.Exists(ba) Or b.Rules.Exists(bb) Then Fail "接合面の個別境界条件を削除してください。ALLは外部境界にだけ適用されます。"
            Set ea = a.groups(ba): Set eb = b.groups(bb)
            If ea.count <> eb.count Then Fail "接合面の要素数が異なります。分割数を揃えてください。"
            tol = MaxDouble(a.ModelScale, b.ModelScale) * 0.00000000001
            Set lookup = CreateObject("Scripting.Dictionary"): Set pairs = New Collection
            For Each e In eb
                key = ElementKey(b, CLng(e), tol)
                If lookup.Exists(key) Then Fail "接合面に重複要素があります。"
                lookup.Add key, CLng(e)
            Next e
            For Each e In ea
                key = ElementKey(a, CLng(e), tol)
                If Not lookup.Exists(key) Then Fail "接合面の節点・分割が一致しません。"
                f = lookup(key): lookup.Remove key
                'Compare actual coordinates after hashing, including quadratic midside nodes.
                idsA = a.elements(CLng(e) + 1): idsB = b.elements(f + 1)
                For i = 0 To UBound(idsA)
                    pa = a.Vertices(idsA(i) + 1): good = False
                    For j = 0 To UBound(idsB)
                        pb = b.Vertices(idsB(j) + 1)
                        If Norm3(Sub3(pa, pb)) <= tol * 2 Then good = True
                    Next j
                    If Not good Then Fail "接合面の幾何節点が一致しません。"
                Next i
                na = a.Differential(CLng(e), IIf(a.Dimension = 2, 0, 1 / 3), 1 / 3, jac)
                nb = b.Differential(f, IIf(b.Dimension = 2, 0, 1 / 3), 1 / 3, jac)
                If Norm3(Add3(na, nb)) > 0.00000001 Then Fail "接合面の法線が反対向きではありません。"
                pairs.Add f
            Next e
            a.InterfaceGroups.Add ba, b.RegionName: b.InterfaceGroups.Add bb, a.RegionName
            InterfacePairs.Add Array(a.RegionName, ba, b.RegionName, bb, pairs)
        End If
    Next row
    For Each id In Models.keys
        Set a = Models(id)
        For Each e In a.groups.keys
            If Not a.InterfaceGroups.Exists(e) Then pa = a.RuleFor(CStr(e))
        Next e
    Next id
End Sub

Private Sub ReadEvaluationPoints()
    Dim s As Worksheet, row As Long, r As CBemRegion, id As String, pointId As String, p As Variant, allIds As Object
    Set s = WS("評価点"): Set allIds = CreateObject("Scripting.Dictionary")
    For row = FIRST_ROW To LastRow(s, 5)
        If ActiveRow(s, row, 5) Then
            id = TextAt(s, row, 1): If Not Models.Exists(id) Then Fail "評価点 行" & row & ": 領域IDが見つかりません。"
            Set r = Models(id): pointId = TextAt(s, row, 2)
            If Not ValidId(pointId) Or allIds.Exists(id & "/" & pointId) Then Fail "評価点: 点IDが不正または重複しています。"
            allIds.Add id & "/" & pointId, True
            p = Array(NumberAt(s, row, 3), NumberAt(s, row, 4), 0#)
            If r.Dimension = 3 Then p(2) = NumberAt(s, row, 5)
            If r.Dimension = 2 And ZeroIfBlank(s, row, 5) <> 0 Then Fail "2Dの評価点のZは空欄または0です。"
            If Not StrictlyInside(r, p) Then Fail id & "/" & pointId & ": 点が領域内部にありません（境界上も不可）。"
            r.points.Add p: r.PointIds.Add pointId
        End If
    Next row
End Sub
Public Function StrictlyInside(ByVal r As CBemRegion, ByVal p As Variant) As Boolean
    Dim v As Variant, i As Long, j As Long, q As Variant, a As Variant, b As Variant, tol As Double, inside As Boolean
    Dim proj As Double, dist As Double, denom As Double
    tol = r.ModelScale * 0.000000001: q = Sub3(p, r.Origin)
    Select Case r.ShapeName
    Case "RECT", "BOX"
        StrictlyInside = True
        For i = 0 To r.Dimension - 1
            If q(i) <= tol Or q(i) >= r.Sizes(i) - tol Then StrictlyInside = False
        Next i
    Case "CIRCLE", "SPHERE"
        StrictlyInside = InsideRoundMesh(r, p)
    Case "POLYGON"
        v = ReadPolygon(r): j = UBound(v)
        For i = 0 To UBound(v)
            a = v(j): b = v(i)
            If OnSegment(a, b, p, tol) Then Exit Function
            If (a(1) > p(1)) <> (b(1) > p(1)) Then
                If p(0) < (b(0) - a(0)) * (p(1) - a(1)) / (b(1) - a(1)) + a(0) Then inside = Not inside
            End If
            j = i
        Next i
        StrictlyInside = inside
    End Select
End Function

Private Sub CheckRigidModes()
    Dim done As Object, members As Object, queue As Collection, name As Variant, item As Variant, boundary As Variant
    Dim r As CBemRegion, p As Variant, lower As Variant, upper As Variant, center As Variant, ModelScale As Double
    Dim vectors As Collection, e As Variant, j As Long, k As Long, rule As Variant, row As Variant, rank As Long, modes As Long
    Set done = CreateObject("Scripting.Dictionary")
    For Each name In Models.keys
        If Not done.Exists(name) Then
            Set members = CreateObject("Scripting.Dictionary"): Set queue = New Collection
            queue.Add CStr(name): members.Add CStr(name), True
            Do While queue.count > 0
                item = queue(1): queue.Remove 1: Set r = Models(item)
                For Each boundary In r.InterfaceGroups.keys
                    If Not members.Exists(r.InterfaceGroups(boundary)) Then members.Add r.InterfaceGroups(boundary), True: queue.Add r.InterfaceGroups(boundary)
                Next boundary
            Loop
            Set r = Models(name): lower = r.Vertices(1): upper = lower
            For Each item In members.keys
                done.Add item, True: Set r = Models(item)
                For Each p In r.Vertices
                    For j = 0 To r.Dimension - 1
                        lower(j) = MinDouble(lower(j), p(j)): upper(j) = MaxDouble(upper(j), p(j))
                    Next j
                Next p
            Next item
            center = Mul3(Add3(lower, upper), 0.5): ModelScale = MaxDouble(upper(0) - lower(0), MaxDouble(upper(1) - lower(1), upper(2) - lower(2)))
            Set vectors = New Collection: modes = IIf(r.Dimension = 2, 3, 6)
            For Each item In members.keys
                Set r = Models(item)
                For Each boundary In r.groups.keys
                    If Not r.InterfaceGroups.Exists(boundary) Then
                        rule = r.RuleFor(CStr(boundary))
                        For Each e In r.groups(boundary)
                            For k = 0 To r.FieldCount - 1
                                p = Mul3(Sub3(r.CollocationPoint(CLng(e), k), center), 1 / ModelScale)
                                For j = 0 To r.Dimension - 1
                                    If rule(0)(j) = "u" Then
                                        If r.Dimension = 2 Then
                                            If j = 0 Then row = Array(1#, 0#, -p(1)) Else row = Array(0#, 1#, p(0))
                                        Else
                                            Select Case j
                                            Case 0: row = Array(1#, 0#, 0#, 0#, p(2), -p(1))
                                            Case 1: row = Array(0#, 1#, 0#, -p(2), 0#, p(0))
                                            Case 2: row = Array(0#, 0#, 1#, p(1), -p(0), 0#)
                                            End Select
                                        End If
                                        AddIndependent vectors, row
                                    End If
                                Next j
                            Next k
                        Next e
                    End If
                Next boundary
            Next item
            If vectors.count < modes Then Fail CStr(name) & "を含む連結領域: 剛体運動の拘束が不足しています（独立拘束 " & vectors.count & "/" & modes & "）。"
        End If
    Next name
End Sub
Private Sub AddIndependent(ByVal basis As Collection, ByVal row As Variant)
    Dim v As Variant, i As Long, pass As Long, dot As Double, length As Double
    For pass = 1 To 2
        For Each v In basis
            dot = 0
            For i = 0 To UBound(row): dot = dot + row(i) * v(i): Next i
            For i = 0 To UBound(row): row(i) = row(i) - dot * v(i): Next i
        Next v
    Next pass
    For i = 0 To UBound(row): length = length + row(i) * row(i): Next i
    length = Sqr(length)
    If length > 0.00000001 Then
        For i = 0 To UBound(row): row(i) = row(i) / length: Next i
        basis.Add row
    End If
End Sub

