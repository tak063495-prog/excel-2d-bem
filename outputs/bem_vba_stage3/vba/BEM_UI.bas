Attribute VB_Name = "BEM_UI"
Option Explicit

Public Sub LogAction(ByVal action As String, ByVal status As String, ByVal detail As String)
    Dim s As Worksheet, row As Long: Set s = WS("ログ")
    row = MaxLong(FIRST_ROW, s.Cells(s.rows.count, 1).End(xlUp).row + 1)
    s.Cells(row, 1).value = Now: s.Cells(row, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    s.Cells(row, 2).Value2 = action: s.Cells(row, 3).Value2 = status: s.Cells(row, 4).Value2 = detail
End Sub
Public Sub MarkChanged()
    Dim oldEvents As Boolean: oldEvents = Application.EnableEvents
    On Error GoTo Finished
    Application.EnableEvents = False
    MeshFingerprint = "": LastImportFingerprint = "": Set Models = Nothing
    LastSolvedFingerprint = "": Set SolveData = Nothing
    Set FieldReports = Nothing
    WS("操作").Range("D35").Value2 = "真の相対残差: 再解析が必要"
    WS("操作").Range("D36").Value2 = "計算時間: 未解析"
    WS("操作").Range("D49").Value2 = "GMRES反復数: 再解析が必要"
    WS("操作").Range("D50").Value2 = "線形判定: 未解析"
    WS("操作").Range("B17").Value2 = "入力変更・再生成が必要"
    If WS("境界結果").Range("A6").Value2 <> "" Then WS("操作").Range("B22").Value2 = "旧入力の結果・再解析が必要"
    WS("メッシュ表示").Range("A3").Value2 = "入力が変更されました。メッシュ生成で表示を更新してください。"
    WS("境界結果").Range("A3").Value2 = "旧入力の結果です。現在のモデルをVBAで再解析するか、対応する結果を取り込んでください。"
    WS("内点結果").Range("A3").Value2 = "旧入力の結果です。現在のモデルをVBAで再解析するか、対応する結果を取り込んでください。"
Finished:
    Application.EnableEvents = oldEvents
End Sub

Public Sub ApplyDimension()
    Dim d As Long, oldEvents As Boolean, color As Long, s As Worksheet, options As String, i As Long
    oldEvents = Application.EnableEvents: Application.EnableEvents = False
    On Error GoTo Finished
    d = CurrentDimension()
    If d = 2 Then
        WS("領域").Range("B6:B10005").Validation.Delete
        WS("領域").Range("B6:B10005").Validation.Add xlValidateList, xlValidAlertStop, xlBetween, "RECT,CIRCLE,POLYGON"
        WS("輪郭2D").Visible = xlSheetVisible
        options = "BOTTOM,RIGHT,TOP,LEFT,OUTER"
        For i = 1 To 20: options = options & ",EDGE" & i: Next i
    Else
        WS("領域").Range("B6:B10005").Validation.Delete
        WS("領域").Range("B6:B10005").Validation.Add xlValidateList, xlValidAlertStop, xlBetween, "BOX,SPHERE"
        If ActiveSheet.name = "輪郭2D" And ActiveWorkbook Is ThisWorkbook Then WS("操作").Activate
        WS("輪郭2D").Visible = xlSheetHidden
        options = "XMIN,XMAX,YMIN,YMAX,ZMIN,ZMAX,OUTER"
    End If
    With WS("境界条件").Range("B6:B10005").Validation
        .Delete: .Add xlValidateList, xlValidAlertStop, xlBetween, "ALL," & options
        .ShowError = False
    End With
    With WS("接合").Range("B6:B10005,D6:D10005").Validation
        .Delete: .Add xlValidateList, xlValidAlertStop, xlBetween, options
        .ShowError = False
    End With
    color = IIf(d = 2, RGB(234, 237, 242), RGB(255, 246, 216))
    WS("領域").Range("F6:F25,I6:I25,M6:M25").interior.color = color
    WS("境界条件").Range("G6:H25").interior.color = color
    WS("評価点").Range("E6:E25").interior.color = color
    WS("操作").Range("B6").interior.color = IIf(d = 3, RGB(234, 237, 242), RGB(255, 240, 188))
    WS("操作").Calculate
Finished:
    Application.EnableEvents = oldEvents
End Sub

Public Sub EnsureMesh()
    Dim signature As String: signature = InputFingerprint()
    If Models Is Nothing Then
        GenerateCore False
    ElseIf MeshFingerprint <> signature Then
        GenerateCore False
    End If
End Sub
Public Function GenerateCore(Optional ByVal preview As Boolean = True) As String
    Dim oldEvents As Boolean, oldScreen As Boolean, description As String, errnum As Long
    If Busy Then Fail "ほかの処理を実行中です。完了を待ってください。"
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    Busy = True: Application.EnableEvents = False: Application.ScreenUpdating = False
    Call ApplyDimension
    Call BuildModels
    Call WriteMeshTables
    If preview Then DrawPreview
    WS("操作").Range("B17").Value2 = "生成済み・入力確認OK"
    WS("操作").Range("B23").Value2 = "メッシュ生成済み"
    If LastImportFingerprint <> MeshFingerprint Then
        If WS("境界結果").Range("A6").Value2 <> "" Then
            WS("操作").Range("B22").Value2 = "旧入力の結果・再解析が必要"
        Else
            WS("操作").Range("B22").Value2 = "解析未実施"
        End If
    End If
    LogAction "メッシュ生成", "OK", CStr(WS("操作").Range("B19").Value2) & " 要素"
    GenerateCore = "OK: " & WS("操作").Range("B18").Value2 & " 節点 / " & WS("操作").Range("B19").Value2 & " 要素"
    GoTo Finished
Failed:
    errnum = Err.number: description = Err.description
    WS("操作").Range("B17").Value2 = "生成失敗・入力を修正"
    WS("操作").Range("B23").Value2 = "詳細はログを確認"
    LogAction "メッシュ生成", "ERROR", description
Finished:
    Busy = False: Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    If errnum <> 0 Then Err.Raise errnum, "ElasticBEM", description
End Function
Public Sub BEM_Generate()
    On Error GoTo Failed
    GenerateCore True: WS("メッシュ表示").Activate: Exit Sub
Failed:
    MsgBox Err.description, vbExclamation, "BEM入力チェック"
End Sub
Public Sub BEM_Check()
    On Error GoTo Failed
    GenerateCore False
    WS("操作").Range("B23").Value2 = "入力チェックOK"
    LogAction "入力チェック", "OK", "閉境界・接合・拘束・評価点を確認": Exit Sub
Failed:
    WS("操作").Range("B23").Value2 = "入力エラー・ログを確認"
    LogAction "入力チェック", "ERROR", Err.description
    MsgBox Err.description, vbExclamation, "BEM入力チェック"
End Sub
Public Sub BEM_Export()
    Dim folder As String
    On Error GoTo Failed
    folder = ExportCore(): MsgBox "モデルを出力しました。" & vbCrLf & folder, vbInformation, "BEMモデル出力": Exit Sub
Failed:
    LogAction "モデル出力", "ERROR", Err.description
    MsgBox Err.description, vbExclamation, "BEMモデル出力"
End Sub
Public Sub BEM_Import()
    Dim folder As String, dialog As FileDialog, result As String
    On Error GoTo Failed
    Set dialog = Application.FileDialog(msoFileDialogFolderPicker)
    dialog.Title = "model.jsonと解析CSVが入ったフォルダを選択"
    If dialog.Show <> -1 Then Exit Sub
    folder = dialog.SelectedItems(1): result = ImportCore(folder)
    WS("境界結果").Activate: MsgBox result, vbInformation, "結果取込": Exit Sub
Failed:
    LogAction "結果取込", "ERROR", Err.description
    MsgBox Err.description, vbExclamation, "結果取込"
End Sub
Public Sub BEM_Sample2D()
    LoadSample "2D_RECT": GenerateCore True: WS("操作").Activate
End Sub
Public Sub BEM_Sample3D()
    LoadSample "3D_BOX": GenerateCore True: WS("操作").Activate
End Sub
Public Sub BEM_Preview()
    On Error GoTo Failed
    Call EnsureMesh
    Call DrawPreview
    WS("メッシュ表示").Activate: Exit Sub
Failed:
    MsgBox Err.description, vbExclamation, "メッシュ表示"
End Sub

Private Sub WriteMeshTables()
    Dim nodes As Collection, elements As Collection, r As CBemRegion, name As Variant, boundary As Variant
    Dim p As Variant, n As Variant, ids As Variant, row As Variant, e As Variant, i As Long, j As Long, jac As Double, fields As Long
    Set nodes = New Collection: Set elements = New Collection
    For Each name In Models.keys
        Set r = Models(name)
        For i = 1 To r.vertices.count
            p = r.vertices(i): row = Array(r.RegionName, i, p(0), p(1), Empty)
            If r.dimension = 3 Then row(4) = p(2)
            nodes.Add row
        Next i
        For Each boundary In r.groups.keys
            For Each e In r.groups(boundary)
                ids = r.elements(CLng(e) + 1)
                p = r.position(CLng(e), IIf(r.dimension = 2, 0, 1 / 3), 1 / 3)
                n = r.Differential(CLng(e), IIf(r.dimension = 2, 0, 1 / 3), 1 / 3, jac)
                row = Array(r.RegionName, CLng(e) + 1, boundary, r.ElementOrder, Empty, Empty, Empty, Empty, Empty, Empty, p(0), p(1), Empty, n(0), n(1), Empty, ElementMeasure(r, CLng(e)), MaxEdge(r, CLng(e)))
                For j = 0 To UBound(ids): row(4 + j) = ids(j) + 1: Next j
                If r.dimension = 3 Then row(12) = p(2): row(15) = n(2)
                elements.Add row
            Next e
        Next boundary
        fields = fields + r.elements.count * r.FieldCount
    Next name
    ClearOutput "節点", 5: ClearOutput "要素", 18
    WS("節点").Range("A6").Resize(nodes.count, 5).Value2 = RowsMatrix(nodes, 5)
    WS("要素").Range("A6").Resize(elements.count, 18).Value2 = RowsMatrix(elements, 18)
    WS("節点").Range("C6:E" & nodes.count + 5).NumberFormat = "0.000000E+00"
    WS("要素").Range("K6:R" & elements.count + 5).NumberFormat = "0.000000E+00"
    WS("操作").Range("B18:B21").Value2 = Application.Transpose(Array(nodes.count, elements.count, fields, CurrentDimension() * fields))
End Sub

Public Sub DrawPreview()
    Dim s As Worksheet, chartObj As ChartObject, chart As chart, series As series, colors As Variant
    Dim lines As Collection, seen As Object, r As CBemRegion, name As Variant, boundary As Variant, e As Variant
    Dim ids As Variant, a As Long, b As Long, i As Long, j As Long, start As Long, count As Long, nrows As Long
    Dim p As Variant, q As Variant, u As Variant, v As Variant, key As String, row As Variant, data As Variant
    Dim minx As Double, maxx As Double, miny As Double, maxy As Double, span As Double, cx As Double, cy As Double, aspect As Double
    Dim total As Long, stride As Long, stepNo As Long, curveSteps As Long, segments As Long, hasBounds As Boolean, stage As String
    Dim errnum As Long, description As String
    On Error GoTo Failed
    stage = "表示の初期化"
    Set s = WS("メッシュ表示")
    If s.ChartObjects.count > 0 Then s.ChartObjects.Delete
    s.Range("T1:V" & MaxLong(1, s.Cells(s.rows.count, 20).End(xlUp).row)).ClearContents
    colors = Array(RGB(37, 82, 132), RGB(206, 108, 36), RGB(58, 136, 110), RGB(139, 81, 164), RGB(167, 79, 92), RGB(90, 123, 151))
    stage = "グラフ作成"
    Set chartObj = s.ChartObjects.Add(left:=s.Range("A6").left, Top:=s.Range("A6").Top, Width:=740, height:=600)
    chartObj.name = "BEMMesh": Set chart = chartObj.chart
    chart.ChartType = xlXYScatterLinesNoMarkers: chart.HasTitle = True
    chart.ChartTitle.text = IIf(CurrentDimension() = 2, "2D境界メッシュ（XY）", "3D境界メッシュ（等角投影）")
    chart.HasLegend = True: chart.Legend.position = xlLegendPositionBottom
    chart.ChartArea.Font.name = "Yu Gothic": chart.ChartArea.Font.Size = 10
    chart.ChartArea.Format.Line.Visible = msoFalse: chart.DisplayBlanksAs = xlNotPlotted
    total = CLng(WS("操作").Range("B19").Value2)
    stride = MaxLong(1, Ceiling(total / 3000#)): start = 1: i = 0
    For Each name In Models.keys
        stage = "線分データ作成"
        Set r = Models(name): Set seen = CreateObject("Scripting.Dictionary")
        For Each boundary In r.groups.keys
            Set lines = New Collection
            For Each e In r.groups(boundary)
                count = count + 1
                If (count - 1) Mod stride = 0 Then
                    ids = r.elements(CLng(e) + 1)
                    If r.dimension = 2 Then
                        curveSteps = IIf(r.ElementOrder = 2, 6, 1)
                        For stepNo = 0 To curveSteps - 1
                            p = r.position(CLng(e), -1 + 2 * stepNo / curveSteps)
                            q = r.position(CLng(e), -1 + 2 * (stepNo + 1) / curveSteps)
                            AppendPlotLine lines, ProjectPoint(p, r.dimension), ProjectPoint(q, r.dimension), minx, maxx, miny, maxy, hasBounds
                        Next stepNo
                    Else
                        For j = 0 To 2
                            a = ids(j): b = ids((j + 1) Mod 3)
                            key = CStr(MinLong(a, b)) & ":" & CStr(MaxLong(a, b))
                            If Not seen.Exists(key) Then
                                seen.Add key, True: p = r.vertices(a + 1): q = r.vertices(b + 1)
                                If r.ElementOrder = 2 Then
                                    v = r.vertices(ids(3 + j) + 1)
                                    For stepNo = 0 To 3
                                        u = QuadEdge(p, v, q, -1 + stepNo / 2)
                                        row = QuadEdge(p, v, q, -0.5 + stepNo / 2)
                                        AppendPlotLine lines, ProjectPoint(u, 3), ProjectPoint(row, 3), minx, maxx, miny, maxy, hasBounds
                                    Next stepNo
                                Else
                                    AppendPlotLine lines, ProjectPoint(p, 3), ProjectPoint(q, 3), minx, maxx, miny, maxy, hasBounds
                                End If
                            End If
                        Next j
                    End If
                End If
            Next e
            If lines.count > 0 Then
                stage = "系列作成 " & r.RegionName & "/" & CStr(boundary)
                data = RowsMatrix(lines, 2): nrows = lines.count
                s.Cells(start, 20).Resize(nrows, 2).Value2 = data
                Set series = chart.SeriesCollection.NewSeries
                series.name = r.RegionName & "/" & CStr(boundary)
                series.XValues = s.Cells(start, 20).Resize(nrows, 1)
                series.values = s.Cells(start, 21).Resize(nrows, 1)
                series.Format.Line.ForeColor.RGB = colors(i Mod 6): series.Format.Line.weight = 1
                start = start + nrows: i = i + 1
            End If
        Next boundary
        If r.dimension = 2 Then
            Set lines = New Collection
            For j = 1 To r.vertices.count
                p = r.vertices(j): lines.Add Array(p(0), p(1))
            Next j
            s.Cells(start, 20).Resize(lines.count, 2).Value2 = RowsMatrix(lines, 2)
            Set series = chart.SeriesCollection.NewSeries
            series.name = r.RegionName & "/幾何節点"
            series.XValues = s.Cells(start, 20).Resize(lines.count, 1)
            series.values = s.Cells(start, 21).Resize(lines.count, 1)
            series.ChartType = xlXYScatter
            series.MarkerStyle = xlMarkerStyleCircle: series.MarkerSize = 4
            series.MarkerForegroundColor = RGB(65, 80, 96): series.MarkerBackgroundColor = RGB(65, 80, 96)
            series.Format.Line.Visible = msoFalse
            start = start + lines.count
        End If
    Next name
    s.columns("T:V").Hidden = True: chart.PlotVisibleOnly = False
    stage = "座標軸の設定"
    chart.axes(xlCategory).HasTitle = True: chart.axes(xlValue).HasTitle = True
    chart.axes(xlCategory).AxisTitle.text = IIf(CurrentDimension() = 2, "X [m]", "投影X [m]")
    chart.axes(xlValue).AxisTitle.text = IIf(CurrentDimension() = 2, "Y [m]", "投影Y [m]")
    chart.axes(xlCategory).HasMajorGridlines = False: chart.axes(xlValue).HasMajorGridlines = False
    chart.axes(xlCategory).TickLabels.NumberFormat = "0.###"
    chart.axes(xlValue).TickLabels.NumberFormat = "0.###"
    chart.Refresh
    aspect = chart.PlotArea.InsideWidth / chart.PlotArea.InsideHeight
    span = MaxDouble(maxx - minx, (maxy - miny) * aspect) * 1.1
    cx = (minx + maxx) / 2: cy = (miny + maxy) / 2
    chart.axes(xlCategory).MinimumScale = cx - span / 2: chart.axes(xlCategory).MaximumScale = cx + span / 2
    chart.axes(xlValue).MinimumScale = cy - span / (2 * aspect): chart.axes(xlValue).MaximumScale = cy + span / (2 * aspect)
    chart.Refresh
    'A second size read accounts for tick-label layout after the explicit axis limits.
    aspect = chart.PlotArea.InsideWidth / chart.PlotArea.InsideHeight
    chart.axes(xlValue).MinimumScale = cy - span / (2 * aspect): chart.axes(xlValue).MaximumScale = cy + span / (2 * aspect)
    s.Range("A3").Value2 = IIf(stride = 1, "生成済み。全境界要素を表示。", "生成済み。表示は" & stride & "要素ごとに間引き。節点・要素表は全件を保持。")
    s.Range("A5").Value2 = "節点 " & WS("操作").Range("B18").Value2 & " / 要素 " & WS("操作").Range("B19").Value2 & " / 未知数 " & WS("操作").Range("B21").Value2
    s.Range("A6").ClearContents
    Exit Sub
Failed:
    errnum = Err.number: description = Err.description
    Err.Raise errnum, "ElasticBEM", stage & ": " & description
End Sub
Private Function QuadEdge(ByVal a As Variant, ByVal m As Variant, ByVal b As Variant, ByVal t As Double) As Variant
    QuadEdge = Add3(Add3(Mul3(a, t * (t - 1) / 2), Mul3(m, 1 - t * t)), Mul3(b, t * (t + 1) / 2))
End Function
Private Function ProjectPoint(ByVal p As Variant, ByVal d As Long) As Variant
    If d = 2 Then ProjectPoint = Array(p(0), p(1)) Else ProjectPoint = Array((p(0) - p(1)) / Sqr(2), (2 * p(2) - p(0) - p(1)) / Sqr(6))
End Function
Private Sub AppendPlotLine(ByVal lines As Collection, ByVal a As Variant, ByVal b As Variant, ByRef minx As Double, ByRef maxx As Double, ByRef miny As Double, ByRef maxy As Double, ByRef hasBounds As Boolean)
    Dim p As Variant
    lines.Add a: lines.Add b: lines.Add Array(CVErr(xlErrNA), CVErr(xlErrNA))
    For Each p In Array(a, b)
        If Not hasBounds Then minx = p(0): maxx = p(0): miny = p(1): maxy = p(1): hasBounds = True
        minx = MinDouble(minx, p(0)): maxx = MaxDouble(maxx, p(0)): miny = MinDouble(miny, p(1)): maxy = MaxDouble(maxy, p(1))
    Next p
End Sub

Public Sub LoadSample(ByVal kind As String)
    Dim oldEvents As Boolean, d As Long, shape As String, name As Variant, bcs As Collection, labels As Variant, i As Long
    Dim r As Variant, regRows As Collection, count As Long, poly As Variant, j As Long
    Dim description As String, errnum As Long
    If Busy Then Fail "ほかの処理を実行中です。完了を待ってください。"
    oldEvents = Application.EnableEvents: Application.EnableEvents = False
    On Error GoTo Failed
    Select Case kind
    Case "2D_RECT": d = 2: shape = "RECT"
    Case "2D_CIRCLE": d = 2: shape = "CIRCLE"
    Case "2D_POLYGON": d = 2: shape = "POLYGON"
    Case "3D_BOX": d = 3: shape = "BOX"
    Case "3D_SPHERE": d = 3: shape = "SPHERE"
    Case "2D_JOIN": d = 2: shape = "RECT"
    Case "3D_JOIN": d = 3: shape = "BOX"
    Case Else: Fail "サンプル名が不正です。"
    End Select
    For Each name In Array("材料", "領域", "輪郭2D", "境界条件", "接合", "評価点")
        Select Case CStr(name)
        Case "領域": count = 13
        Case "境界条件": count = 8
        Case "評価点": count = 5
        Case Else: count = 4
        End Select
        ClearOutput CStr(name), count
    Next name
    WS("操作").Range("B5").Value2 = CStr(d) & "D"
    WS("操作").Range("B6").Value2 = "plane_strain": WS("操作").Range("B7").Value2 = 2
    WS("操作").Range("B8").Value2 = 0.25: WS("操作").Range("B9").Value2 = 20000
    WS("材料").Range("A6:D6").Value2 = Array("MAT1", 30000, 0.3, "元Pythonのblock例")
    Set regRows = New Collection: r = Array("block", shape, "MAT1", 0#, 0#, 0#, 1#, 1#, 1#, 0.5, 4, 4, 2)
    If d = 3 Then r(10) = 2: r(11) = 2
    If shape = "CIRCLE" Or shape = "SPHERE" Then r(10) = 16: r(11) = 8
    If shape = "POLYGON" Then r(10) = Empty
    regRows.Add r
    If InStr(kind, "JOIN") > 0 Then
        r(0) = "left": regRows.Remove 1: regRows.Add r
        r(0) = "right": r(2) = "MAT2": r(3) = 1: regRows.Add r
        WS("材料").Range("A7:D7").Value2 = Array("MAT2", 60000, 0.3, "接合サンプル")
        WS("接合").Range("A6:D6").Value2 = Array("left", IIf(d = 2, "RIGHT", "XMAX"), "right", IIf(d = 2, "LEFT", "XMIN"))
    End If
    WS("領域").Range("A6").Resize(regRows.count, 13).Value2 = RowsMatrix(regRows, 13)
    Set bcs = New Collection
    For i = 1 To regRows.count
        r = regRows(i): name = r(0)
        bcs.Add Array(name, "ALL", "t", 0#, "t", 0#, "t", 0#)
        If shape = "CIRCLE" Or shape = "SPHERE" Then
            bcs.Remove bcs.count: bcs.Add Array(name, "OUTER", "u", 0#, "u", 0#, "u", 0#)
        Else
            If shape = "POLYGON" Then
                bcs.Remove bcs.count: bcs.Add Array(name, "ALL", "u", 0#, "u", 0#, "u", 0#)
            Else
                If CStr(name) <> "right" Then bcs.Add Array(name, IIf(d = 2, "LEFT", "XMIN"), "u", 0#, "t", 0#, "t", 0#)
                bcs.Add Array(name, IIf(d = 2, "BOTTOM", "YMIN"), "t", 0#, "u", 0#, "t", 0#)
                If d = 3 Then bcs.Add Array(name, "ZMIN", "t", 0#, "t", 0#, "u", 0#)
                If CStr(name) <> "left" Then bcs.Add Array(name, IIf(d = 2, "RIGHT", "XMAX"), "t", 100#, "t", 0#, "t", 0#)
            End If
        End If
        WS("評価点").Cells(5 + i, 1).Resize(1, 5).Value2 = Array(name, "P1", IIf(shape = "CIRCLE" Or shape = "SPHERE", 0#, CDbl(r(3)) + 0.5), IIf(shape = "CIRCLE" Or shape = "SPHERE", 0#, 0.5), IIf(d = 3 And shape <> "SPHERE", 0.5, 0#))
    Next i
    WS("境界条件").Range("A6").Resize(bcs.count, 8).Value2 = RowsMatrix(bcs, 8)
    If shape = "POLYGON" Then
        poly = Array(Array(0#, 0#), Array(1#, 0#), Array(1#, 0.6), Array(0.6, 1#), Array(0#, 1#))
        For i = 0 To UBound(poly)
            WS("輪郭2D").Cells(FIRST_ROW + i, 1).Resize(1, 4).Value2 = Array("block", i + 1, poly(i)(0), poly(i)(1))
        Next i
    End If
    Call MarkChanged
    Call ApplyDimension
    LogAction "サンプル読込", "OK", kind
    GoTo Finished
Failed:
    errnum = Err.number: description = Err.description
Finished:
    Application.EnableEvents = oldEvents
    If errnum <> 0 Then Err.Raise errnum, "ElasticBEM", description
End Sub

