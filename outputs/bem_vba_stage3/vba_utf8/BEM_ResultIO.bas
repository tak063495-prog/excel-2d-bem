Attribute VB_Name = "BEM_ResultIO"
Option Explicit
Private Function JSONBool(ByVal value As Boolean) As String
    JSONBool = IIf(value, "true", "false")
End Function
Private Function CSVBool(ByVal value As Variant) As String
    CSVBool = IIf(LCase$(CStr(value)) = "true", "True", "False")
End Function
Public Sub SaveSolvedResults(ByVal folder As String)
    Dim fso As Object, name As Variant, sr As CSolveRegion, lines As Collection, cols As Collection, header As Collection, reports As Collection, checks As Collection
    Dim axes As Variant, stresses As Variant, i As Long, j As Long, d As Long, index As Long, interior As Variant, row As Long, last As Long
    Dim report As String, summary As String, uv() As Double, tv() As Double, xyz() As Double, available As Boolean, opStats As String
    If LastSolvedFingerprint <> InputFingerprint() Or SolveData Is Nothing Then Fail "現在の入力をVBAで解析してから結果を出力してください。"
    ArchivePreviousResults folder
    Set fso = CreateObject("Scripting.FileSystemObject"): axes = Array("x", "y", "z")
    stresses = Array("sxx", "syy", "szz", "sxy", "syz", "szx")
    last = WS("内点結果").Cells(WS("内点結果").Rows.Count, 1).End(xlUp).Row
    If last >= FIRST_ROW Then interior = WS("内点結果").Range("A6:Q" & last).Value2
    Set reports = New Collection
    For Each name In SolveData.Keys
        Set sr = SolveData(name): d = sr.Ref.Dimension: uv = sr.U: tv = sr.T: xyz = sr.X
        Set lines = New Collection: Set header = New Collection
        For j = 0 To d - 1: header.Add axes(j) & "_m": Next j
        For j = 0 To d - 1: header.Add "u" & axes(j) & "_m": Next j
        For j = 0 To d - 1: header.Add "t" & axes(j) & "_kPa": Next j
        lines.Add JoinCollection(header, ",")
        For i = 1 To sr.NF
            Set cols = New Collection
            For j = 1 To d: cols.Add JNum(sr.WorldCoordinate(i, j)): Next j
            For j = 1 To d: cols.Add JNum(uv((i - 1) * d + j)): Next j
            For j = 1 To d: cols.Add JNum(tv((i - 1) * d + j)): Next j
            lines.Add JoinCollection(cols, ",")
        Next i
        WriteUTF8 fso.BuildPath(folder, CStr(name) & "_boundary.csv"), JoinCollection(lines, vbCrLf) & vbCrLf
        SaveBoundaryVTK sr, folder
        Set checks = New Collection
        If sr.Ref.Points.Count > 0 Then
            Set lines = New Collection: Set header = New Collection
            For j = 0 To d - 1: header.Add axes(j) & "_m": Next j
            For j = 0 To d - 1: header.Add "u" & axes(j) & "_m": Next j
            For j = 0 To 5: header.Add stresses(j) & "_kPa": Next j
            header.Add "quadrature_converged": header.Add "error_code": header.Add "value_available"
            lines.Add JoinCollection(header, ",")
            For row = 1 To UBound(interior, 1)
                If CStr(interior(row, 1)) = CStr(name) Then
                    Set cols = New Collection: available = CSVBool(interior(row, 17)) = "True"
                    For j = 1 To d: cols.Add JNum(CDbl(interior(row, 2 + j))): Next j
                    For j = 1 To d
                        If available Then cols.Add JNum(CDbl(interior(row, 5 + j))) Else cols.Add "nan"
                    Next j
                    For j = 1 To 6
                        If available Then cols.Add JNum(CDbl(interior(row, 8 + j))) Else cols.Add "nan"
                    Next j
                    cols.Add CSVBool(interior(row, 15)): cols.Add CStr(interior(row, 16)): cols.Add CSVBool(interior(row, 17))
                    lines.Add JoinCollection(cols, ",")
                    checks.Add FieldReports(CStr(name) & "/" & CStr(interior(row, 2)))
                End If
            Next row
            If lines.Count <> sr.Ref.Points.Count + 1 Then Fail "内部場の結果行数が不正です。再解析してください。"
            WriteUTF8 fso.BuildPath(folder, CStr(name) & "_interior.csv"), JoinCollection(lines, vbCrLf) & vbCrLf
        End If
        If sr.Backend = "fmm" Then opStats = sr.Fmm.StatsJSON Else opStats = "{""backend"":""dense"",""elements"":" & sr.Ref.Elements.Count & ",""field_nodes"":" & sr.NF & ",""matrix_MB"":" & JNum(16# * sr.ND * sr.ND / 1048576) & "}"
        reports.Add "{""name"":" & JQuote(CStr(name)) & ",""material"":{""E"":" & JNum(sr.Ref.Young) & ",""nu"":" & JNum(sr.Ref.Poisson) & "},""element_order"":" & sr.Ref.ElementOrder & ",""elements"":" & sr.Ref.Elements.Count & ",""field_nodes"":" & sr.NF & ",""selection"":" & sr.SelectionJSON & ",""operator"":" & opStats & ",""resultant_traction_kN"":" & ResultantJSON(sr) & ",""resultant_units"":" & JQuote(IIf(sr.Ref.Dimension = 2, "kN/m of out-of-plane length", "kN")) & ",""interior_reports"": [" & JoinCollection(checks, ",") & "]}"
    Next name
    summary = "{""backend"":" & JQuote(SolveBackend) & ",""method"":" & JQuote(LinearMethod) & ",""gmres_info"":" & SolveLinearInfo & ",""iterations"":" & SolveIterations & ",""restart_cycles"":" & SolveCycles & ",""error_code"":" & JQuote(IIf(SolveConverged, "OK", "LINEAR_SOLVE_TOLERANCE_UNMET")) & ",""residual_normalization"":""RHS; absolute-term scale when RHS is below arithmetic roundoff"",""converged"":" & JSONBool(SolveConverged) & ",""true_relative_residual"":" & JNum(SolveResidual) & ",""DOF"":" & SolveDOF & ",""seconds"":" & JNum(SolveSeconds)
    summary = summary & ",""interface_displacement_jump_max"":" & JNum(InterfaceJump) & ",""interface_traction_imbalance_max"":" & JNum(InterfaceImbalance) & "}"
    report = "{""dimension"":" & CurrentDimension() & ",""linear_solver"":" & summary & ",""regions"": [" & JoinCollection(reports, ",") & "],""all_requested_checks_passed"":" & JSONBool(SolveAllChecks) & ",""accuracy_scope"":""quadrature and algebraic residual only; discretization/FMM errors require convergence studies""}"
    'Commit report last; the import checks its timestamp after model.json.
    SaveSolutionNPZ folder
    WriteUTF8 fso.BuildPath(folder, "report.json"), report & vbCrLf
    LogAction "VBA結果出力", IIf(SolveAllChecks, "OK", "精度未達"), folder
End Sub
