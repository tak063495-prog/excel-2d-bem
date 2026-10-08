Attribute VB_Name = "BEM_IO"
Option Explicit

Public Function JQuote(ByVal value As String) As String
    Dim i As Long, ch As String, code As Long, result As String
    For i = 1 To Len(value)
        ch = mid$(value, i, 1): code = AscW(ch)
        Select Case ch
        Case """": result = result & "\"""
        Case "\": result = result & "\\"
        Case vbCr: result = result & "\r"
        Case vbLf: result = result & "\n"
        Case vbTab: result = result & "\t"
        Case Else
            If code >= 0 And code < 32 Then result = result & "\u" & right$("0000" & Hex$(code), 4) Else result = result & ch
        End Select
    Next i
    JQuote = """" & result & """"
End Function
Public Function JNum(ByVal value As Double) As String
    'Str uses a period independently of Excel's decimal separator.
    JNum = Trim$(Str$(value))
    If left$(JNum, 1) = "." Then JNum = "0" & JNum
    If left$(JNum, 2) = "-." Then JNum = "-0" & mid$(JNum, 2)
End Function
Public Function JoinCollection(ByVal parts As Collection, ByVal separator As String) As String
    Dim values() As String, i As Long
    If parts.count = 0 Then Exit Function
    ReDim values(0 To parts.count - 1)
    For i = 1 To parts.count: values(i - 1) = CStr(parts(i)): Next i
    JoinCollection = Join(values, separator)
End Function
Private Function VectorJSON(ByVal p As Variant, ByVal d As Long, Optional ByVal strings As Boolean = False) As String
    Dim i As Long, parts As Collection: Set parts = New Collection
    For i = 0 To d - 1
        If strings Then parts.Add JQuote(CStr(p(i))) Else parts.Add JNum(CDbl(p(i)))
    Next i
    VectorJSON = "[" & JoinCollection(parts, ",") & "]"
End Function
Private Function IndexJSON(ByVal indices As Collection) As String
    Dim item As Variant, parts As Collection: Set parts = New Collection
    For Each item In indices: parts.Add CStr(CLng(item)): Next item
    IndexJSON = "[" & JoinCollection(parts, ",") & "]"
End Function

Public Function ModelJSON() As String
    Dim allRegions As Collection, list As Collection, coords As Collection, elems As Collection, rules As Collection
    Dim r As CBemRegion, name As Variant, p As Variant, ids As Variant, boundary As Variant, rule As Variant
    Dim content As String, backend As String, i As Long, item As Variant, iface As Collection
    If Models Is Nothing Then Fail "先にメッシュを生成してください。"
    backend = TextAt(WS("操作"), 13, 2)
    If backend <> "auto" And backend <> "dense" And backend <> "fmm" Then Fail "演算方式はauto、dense、fmmです。"
    Set allRegions = New Collection
    For Each name In Models.keys
        Set r = Models(name): Set coords = New Collection: Set elems = New Collection: Set rules = New Collection
        For Each p In r.Vertices: coords.Add VectorJSON(p, r.Dimension): Next p
        For Each ids In r.elements
            Set list = New Collection
            For i = 0 To UBound(ids): list.Add CStr(CLng(ids(i))): Next i
            elems.Add "[" & JoinCollection(list, ",") & "]"
        Next ids
        For Each boundary In r.groups.keys
            If Not r.InterfaceGroups.Exists(boundary) Then
                rule = r.RuleFor(CStr(boundary))
                rules.Add "{""elements"":" & IndexJSON(r.groups(boundary)) & ",""type"":" & VectorJSON(rule(0), r.Dimension, True) & ",""values"":" & VectorJSON(rule(1), r.Dimension) & "}"
            End If
        Next boundary
        Set list = New Collection
        For Each p In r.points: list.Add VectorJSON(p, r.Dimension): Next p
        content = "{""name"":" & JQuote(r.RegionName) & ",""vertices"": [" & JoinCollection(coords, ",") & "],""elements"": [" & JoinCollection(elems, ",") & "]"
        content = content & ",""material"":{""E"":" & JNum(r.Young) & ",""nu"":" & JNum(r.Poisson) & "},""element_order"":" & r.ElementOrder
        content = content & ",""boundary_conditions"": [" & JoinCollection(rules, ",") & "],""interior_points"": [" & JoinCollection(list, ",") & "]}"
        allRegions.Add content
    Next name
    Set iface = New Collection
    For Each item In InterfacePairs
        Set r = Models(item(0))
        iface.Add "{""region_a"":" & JQuote(CStr(item(0))) & ",""elements_a"":" & IndexJSON(r.groups(item(1))) & ",""region_b"":" & JQuote(CStr(item(2))) & ",""elements_b"":" & IndexJSON(item(4)) & "}"
    Next item
    content = "three_dimensional": If CurrentDimension() = 2 Then content = TextAt(WS("操作"), 6, 2)
    ModelJSON = "{""dimension"":" & CurrentDimension() & ",""state"":" & JQuote(content) & ",""element_order"":" & CurrentOrder()
    Call ReadSolveOptions
    Call ReadFmmOptions
    ModelJSON = ModelJSON & ",""regions"": [" & JoinCollection(allRegions, ",") & "],""interfaces"": [" & JoinCollection(iface, ",") & "],""solver"":{""backend"":" & JQuote(backend) & ",""p"":" & FmmOrder & ",""leaf"":" & FmmLeaf & ",""theta"":" & JNum(FmmTheta) & ",""cache_mb"":" & JNum(FmmCacheMB) & ",""expected_iterations"":" & ExpectedIterations & ",""memory_mb"":" & JNum(SolveSetting(53, 512)) & "}"
    ModelJSON = ModelJSON & ",""integration"":{""boundary_order"":" & BoundaryOrder & ",""boundary_de_level"":" & BoundaryDE & ",""self_order"":" & SelfOrder & "}"
    ModelJSON = ModelJSON & ",""quadrature"":{""mode"":" & JQuote(QuadratureMode) & ",""near_ratio"":" & JNum(InteriorNearRatio) & ",""min_distance_ratio"":" & JNum(InteriorMinDistanceRatio) & ",""displacement_tol"":" & JNum(DisplacementTolerance) & ",""stress_tol"":" & JNum(StressTolerance) & ",""max_de_level"":" & MaxDELevel & ",""max_gauss_level"":" & MaxGaussLevel & ",""max_points"":" & MaxPointBudget & "}"
    ModelJSON = ModelJSON & ",""linear_solver"":{""rtol"":" & JNum(LinearTolerance) & ",""restart"":" & GMRESRestart & ",""maxiter"":" & GMRESMaxCycles & ",""vba_method"":" & JQuote(LinearMethod) & "}}"
End Function
Public Sub WriteUTF8(ByVal path As String, ByVal value As String)
    Dim stream As Object, binary As Object: Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.WriteText value: stream.position = 0: stream.Type = 1: stream.position = 3
    Set binary = CreateObject("ADODB.Stream"): binary.Type = 1: binary.Open
    binary.Write stream.Read: binary.SaveToFile path, 2: binary.Close: stream.Close
End Sub
Public Function ReadUTF8(ByVal path As String) As String
    Dim stream As Object, fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Fail "ファイルがありません: " & path
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open: stream.LoadFromFile path
    ReadUTF8 = stream.ReadText: stream.Close
End Function
Private Function CleanSnapshot(ByVal value As String) As String
    value = Replace(Replace(value, vbCr, ""), vbLf, "")
    If Len(value) > 0 Then
        If AscW(left$(value, 1)) = -257 Then value = mid$(value, 2)
    End If
    CleanSnapshot = Trim$(value)
End Function
Public Function OutputDirectory() As String
    Dim subfolder As String, stem As String, fso As Object
    If ThisWorkbook.path = "" Then Fail "先にこのブックを保存してください。"
    subfolder = TextAt(WS("操作"), 12, 2): stem = TextAt(WS("操作"), 10, 2)
    If Not ValidId(subfolder) Or Not ValidId(stem) Then Fail "出力サブフォルダとモデル名はASCII英数字・_・-で指定してください。"
    Set fso = CreateObject("Scripting.FileSystemObject")
    OutputDirectory = fso.BuildPath(fso.BuildPath(ThisWorkbook.path, subfolder), stem)
    If Not fso.FolderExists(fso.BuildPath(ThisWorkbook.path, subfolder)) Then fso.CreateFolder fso.BuildPath(ThisWorkbook.path, subfolder)
    If Not fso.FolderExists(OutputDirectory) Then fso.CreateFolder OutputDirectory
End Function

Public Function ExportCore(Optional ByVal folder As String = "") As String
    Dim fso As Object, s As Worksheet, nr As Long, nc As Long, lines As Collection, cols As Collection
    Dim data As Variant, i As Long, j As Long, name As Variant, value As Variant, snapshot As String
    If Busy Then Fail "ほかの処理を実行中です。完了を待ってください。"
    EnsureMesh
    If folder = "" Then folder = OutputDirectory()
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folder) Then Fail "出力フォルダがありません。"
    snapshot = ModelJSON()
    If fso.FileExists(fso.BuildPath(folder, "model.json")) Then
        If CleanSnapshot(ReadUTF8(fso.BuildPath(folder, "model.json"))) <> snapshot Then ArchivePreviousResults folder
    End If
    For Each name In Array("節点", "要素")
        Set s = WS(CStr(name)): nc = IIf(CStr(name) = "節点", 5, 18)
        nr = s.Cells(s.rows.count, 1).End(xlUp).row
        data = s.Range(s.Cells(5, 1), s.Cells(nr, nc)).Value2: Set lines = New Collection
        For i = 1 To UBound(data, 1)
            Set cols = New Collection
            For j = 1 To nc
                value = data(i, j)
                If IsEmpty(value) Then
                    cols.Add ""
                ElseIf VarType(value) <> vbString And IsNumeric(value) Then
                    cols.Add JNum(CDbl(value))
                Else
                    cols.Add """" & Replace(CStr(value), """", """""") & """"
                End If
            Next j
            lines.Add JoinCollection(cols, ",")
        Next i
        WriteUTF8 fso.BuildPath(folder, IIf(CStr(name) = "節点", "mesh_nodes.csv", "mesh_elements.csv")), JoinCollection(lines, vbCrLf) & vbCrLf
    Next name
    'model.json is written last so incomplete exports cannot be mistaken for a new model.
    If fso.FileExists(fso.BuildPath(folder, "model.json")) Then
        If CleanSnapshot(ReadUTF8(fso.BuildPath(folder, "model.json"))) <> snapshot Then WriteUTF8 fso.BuildPath(folder, "model.json"), snapshot & vbCrLf
    Else
        WriteUTF8 fso.BuildPath(folder, "model.json"), snapshot & vbCrLf
    End If
    WriteUTF8 fso.BuildPath(folder, "解析方法.txt"), "ExcelのVBA解析実行で解析できます。解析後のJSON・CSV出力は境界・内部場・report.jsonも保存します。" & vbCrLf & "元Pythonで計算する場合: python run_bem.py """ & fso.BuildPath(folder, "model.json") & """ --out """ & folder & """" & vbCrLf
    If LastSolvedFingerprint = InputFingerprint() And Not SolveData Is Nothing Then
        SaveSolvedResults folder
        WS("操作").Range("B23").Value2 = "モデル・VBA解析結果CSV出力済み"
    Else
        WS("操作").Range("B23").Value2 = "JSON・メッシュCSV出力済み"
    End If
    LogAction "モデル出力", "OK", folder
    ExportCore = folder
End Function

Private Function CSVLines(ByVal path As String) As Variant
    Dim text As String
    text = Replace(ReadUTF8(path), vbCr, "")
    Do While right$(text, 1) = vbLf: text = left$(text, Len(text) - 1): Loop
    If Len(text) = 0 Then Fail "CSVが空です: " & path
    CSVLines = Split(text, vbLf)
End Function

Public Sub ArchivePreviousResults(ByVal folder As String)
    Dim fso As Object, file As Object, paths As Collection, path As Variant, archive As String, suffix As Long, stem As String
    Set fso = CreateObject("Scripting.FileSystemObject"): Set paths = New Collection
    folder = fso.GetAbsolutePathName(folder)
    For Each file In fso.GetFolder(folder).Files
        If file.name = "report.json" Or file.name = "solution.npz" Or file.name Like "*_boundary.csv" Or file.name Like "*_interior.csv" Or file.name Like "*_boundary.vtk" Then paths.Add file.path
    Next file
    If paths.count = 0 Then Exit Sub
    stem = "previous_results_" & Format$(Now, "yyyymmdd_hhnnss"): archive = fso.BuildPath(folder, stem)
    Do While fso.FolderExists(archive)
        suffix = suffix + 1: archive = fso.BuildPath(folder, stem & "_" & suffix)
    Loop
    fso.CreateFolder archive
    If fso.FileExists(fso.BuildPath(folder, "model.json")) Then fso.CopyFile fso.BuildPath(folder, "model.json"), fso.BuildPath(archive, "model.json"), True
    For Each path In paths
        fso.MoveFile CStr(path), fso.BuildPath(archive, fso.GetFileName(CStr(path)))
    Next path
    LogAction "旧結果の保存", "OK", archive
End Sub
Private Sub CheckResultDate(ByVal path As String, ByVal modelDate As Date)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Fail "結果ファイルがありません: " & path
    If fso.GetFile(path).DateLastModified < modelDate Then Fail "モデル更新前の結果です。現在のmodel.jsonを再解析してください: " & path
End Sub
Private Function CSVNumber(ByVal text As String, ByVal context As String) As Double
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[+-]?([0-9]+(\.[0-9]*)?|\.[0-9]+)([eE][+-]?[0-9]+)?$"
    If Not re.Test(text) Then Fail context & ": 有限数ではありません。"
    CSVNumber = Val(text)
End Function
Private Sub CoordinatesMatch(ByVal values As Variant, ByVal expected As Variant, ByVal r As CBemRegion, ByVal context As String)
    Dim j As Long
    For j = 0 To r.Dimension - 1
        If Abs(CSVNumber(values(j), context) - expected(j)) > r.ModelScale * 0.000000001 Then Fail context & ": 座標が現在のモデルと一致しません。"
    Next j
End Sub
Private Function StatusBool(ByVal text As String, ByVal context As String) As String
    Select Case LCase$(text)
    Case "true": StatusBool = "True"
    Case "false": StatusBool = "False"
    Case Else: Fail context & ": TrueまたはFalseが必要です。"
    End Select
End Function

Public Function ImportCore(ByVal folder As String) As String
    Dim fso As Object, r As CBemRegion, name As Variant, lines As Variant, values As Variant, p As Variant
    Dim bound As Collection, interior As Collection, row As Variant, e As Long, fieldIndex As Long, i As Long, j As Long, d As Long
    Dim expectedHeader As String, headers As Collection, axes As Variant, ctx As String, stats As String, passed As Boolean
    Dim report As String, re As Object, matches As Object, boundaryData As Variant, interiorData As Variant, modelDate As Date
    If Busy Then Fail "ほかの処理を実行中です。完了を待ってください。"
    Call EnsureMesh
    Set fso = CreateObject("Scripting.FileSystemObject")
    If CleanSnapshot(ReadUTF8(fso.BuildPath(folder, "model.json"))) <> ModelJSON() Then Fail "model.jsonが現在の入力と一致しません。同じモデルの解析結果を選んでください。"
    modelDate = fso.GetFile(fso.BuildPath(folder, "model.json")).DateLastModified
    CheckResultDate fso.BuildPath(folder, "report.json"), modelDate
    Set bound = New Collection: Set interior = New Collection: d = CurrentDimension(): axes = Array("x", "y", "z")
    For Each name In Models.keys
        Set r = Models(name): Set headers = New Collection
        For i = 0 To d - 1: headers.Add axes(i) & "_m": Next i
        For i = 0 To d - 1: headers.Add "u" & axes(i) & "_m": Next i
        For i = 0 To d - 1: headers.Add "t" & axes(i) & "_kPa": Next i
        CheckResultDate fso.BuildPath(folder, r.RegionName & "_boundary.csv"), modelDate
        lines = CSVLines(fso.BuildPath(folder, r.RegionName & "_boundary.csv"))
        If lines(0) <> JoinCollection(headers, ",") Then Fail r.RegionName & ": 境界CSVの列名が一致しません。"
        If UBound(lines) <> r.elements.count * r.FieldCount Then Fail r.RegionName & ": 境界CSVの行数が一致しません。"
        For i = 1 To UBound(lines)
            ctx = r.RegionName & " boundary 行" & i: values = Split(lines(i), ",")
            If UBound(values) + 1 <> 3 * d Then Fail ctx & ": 列数が不正です。"
            e = (i - 1) \ r.FieldCount: fieldIndex = (i - 1) Mod r.FieldCount
            p = r.CollocationPoint(e, fieldIndex): CoordinatesMatch values, p, r, ctx
            row = Array(r.RegionName, i, e + 1, fieldIndex + 1, p(0), p(1), Empty, Empty, Empty, Empty, Empty, Empty, Empty)
            If d = 3 Then row(6) = p(2)
            For j = 0 To d - 1
                row(7 + j) = CSVNumber(values(d + j), ctx): row(10 + j) = CSVNumber(values(2 * d + j), ctx)
            Next j
            bound.Add row
        Next i
        If r.points.count > 0 Then
            Set headers = New Collection
            For i = 0 To d - 1: headers.Add axes(i) & "_m": Next i
            For i = 0 To d - 1: headers.Add "u" & axes(i) & "_m": Next i
            For Each p In Array("sxx_kPa", "syy_kPa", "szz_kPa", "sxy_kPa", "syz_kPa", "szx_kPa", "quadrature_converged", "error_code", "value_available"): headers.Add CStr(p): Next p
            CheckResultDate fso.BuildPath(folder, r.RegionName & "_interior.csv"), modelDate
            lines = CSVLines(fso.BuildPath(folder, r.RegionName & "_interior.csv"))
            If lines(0) <> JoinCollection(headers, ",") Then Fail r.RegionName & ": 内点CSVの列名が一致しません。"
            If UBound(lines) <> r.points.count Then Fail r.RegionName & ": 内点CSVの行数が一致しません。"
            For i = 1 To UBound(lines)
                ctx = r.RegionName & " interior 行" & i: values = Split(lines(i), ",")
                If UBound(values) + 1 <> 2 * d + 9 Then Fail ctx & ": 列数が不正です。"
                p = r.points(i): CoordinatesMatch values, p, r, ctx
                row = Array(r.RegionName, r.PointIds(i), p(0), p(1), Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty)
                If d = 3 Then row(4) = p(2)
                row(14) = StatusBool(values(2 * d + 6), ctx)
                row(16) = StatusBool(values(2 * d + 8), ctx)
                Set re = CreateObject("VBScript.RegExp"): re.Pattern = "^[A-Za-z0-9_-]+$"
                If Not re.Test(values(2 * d + 7)) Then Fail ctx & ": エラーコードが不正です。"
                row(15) = values(2 * d + 7)
                For j = 0 To d - 1
                    If row(16) = "True" Then row(5 + j) = CSVNumber(values(d + j), ctx) Else row(5 + j) = "取得不可"
                Next j
                For j = 0 To 5
                    If row(16) = "True" Then row(8 + j) = CSVNumber(values(2 * d + j), ctx) Else row(8 + j) = "取得不可"
                Next j
                interior.Add row
            Next i
        End If
    Next name
    'Check report before committing any worksheet changes.
    report = ReadUTF8(fso.BuildPath(folder, "report.json"))
    Set re = CreateObject("VBScript.RegExp"): re.Pattern = """all_requested_checks_passed""\s*:\s*(true|false)"
    If Not re.Test(report) Then Fail "report.jsonに収束判定がありません。"
    Set matches = re.Execute(report): passed = (matches(0).SubMatches(0) = "true")
    boundaryData = RowsMatrix(bound, 13): interiorData = RowsMatrix(interior, 17)
    ClearOutput "境界結果", 13: ClearOutput "内点結果", 17
    If bound.count > 0 Then WS("境界結果").Range("A6").Resize(bound.count, 13).Value2 = boundaryData
    If interior.count > 0 Then WS("内点結果").Range("A6").Resize(interior.count, 17).Value2 = interiorData
    WS("境界結果").Range("E6:M" & bound.count + 5).NumberFormat = "0.000000E+00"
    If interior.count > 0 Then WS("内点結果").Range("C6:N" & interior.count + 5).NumberFormat = "0.000000E+00"
    WS("操作").Range("B22").Value2 = IIf(passed, "取込済み（収束判定OK）", "取込済み（収束未達あり）")
    WS("操作").Range("B23").Value2 = "境界・内点CSV取込済み"
    LastImportFingerprint = MeshFingerprint
    LastSolvedFingerprint = "": Set SolveData = Nothing
    WS("境界結果").Range("A3").Value2 = "現在のモデルの解析結果。二次要素の場の節点と幾何節点は異なります。"
    WS("内点結果").Range("A3").Value2 = "現在のモデルの解析結果。積分収束・エラーコード・値取得可否を保持。"
    LogAction "結果取込", IIf(passed, "OK", "未収束"), folder
    ImportCore = "境界 " & bound.count & " 行、内点 " & interior.count & " 行"
End Function

Public Function RowsMatrix(ByVal rows As Collection, ByVal ncols As Long) As Variant
    Dim data() As Variant, i As Long, j As Long, row As Variant
    If rows.count = 0 Then RowsMatrix = Empty: Exit Function
    ReDim data(1 To rows.count, 1 To ncols)
    For i = 1 To rows.count
        row = rows(i)
        For j = 1 To ncols: data(i, j) = row(j - 1): Next j
    Next i
    RowsMatrix = data
End Function
Public Sub ClearOutput(ByVal name As String, ByVal ncols As Long)
    Dim last As Long, col As Long, s As Worksheet: Set s = WS(name)
    For col = 1 To ncols: last = MaxLong(last, s.Cells(s.rows.count, col).End(xlUp).row): Next col
    If last >= FIRST_ROW Then s.Range(s.Cells(FIRST_ROW, 1), s.Cells(last, ncols)).ClearContents
End Sub

